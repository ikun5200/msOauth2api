const Imap = require('node-imap');
const simpleParser = require('mailparser').simpleParser;
const {
  generateAuthString,
  getAccessToken,
  graphApi,
  normalizeGraphMailbox,
  validatePassword
} = require('./utils');

const ALLOWED_MAILBOXES = ['INBOX', 'Junk'];

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function buildResponseData(mail, seqno) {
  const generatedId = `imap_${seqno}_${Date.now()}`;
  const headerMessageId = mail?.headers?.get?.('message-id');

  return {
    id: generatedId,
    messageId: mail?.messageId || headerMessageId || generatedId,
    send: mail?.from?.text || '',
    subject: mail?.subject || '',
    text: mail?.text || '',
    html: mail?.html || '',
    date: mail?.date || null,
    mode: 'imap',
    _imapSeqno: seqno
  };
}

async function getEmailsGraph(accessToken, mailbox) {
  const response = await fetch(`https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=1&$orderby=receivedDateTime desc&$select=id,from,subject,bodyPreview,body,createdDateTime,internetMessageId`, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to fetch emails: ${response.status}, ${errorText}`);
  }

  const responseData = await response.json();
  return (responseData.value || []).map(item => ({
    id: item.id,
    messageId: item.internetMessageId || item.id,
    send: item.from?.emailAddress?.address || '',
    subject: item.subject || '',
    text: item.bodyPreview || '',
    html: item.body?.content || '',
    date: item.createdDateTime,
    mode: 'graph'
  }));
}

module.exports = async (req, res) => {
  if (!validatePassword(req, res)) return;

  const params = req.method === 'GET' ? req.query : req.body;
  let { refresh_token, client_id, email, mailbox, response_type = 'json' } = params;

  if (!refresh_token || !client_id || !email || !mailbox) {
    return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or mailbox' });
  }

  if (!ALLOWED_MAILBOXES.includes(mailbox)) {
    return res.status(400).json({ error: 'Invalid mailbox. Allowed: INBOX, Junk' });
  }

  try {
    const graphApiResult = await graphApi(refresh_token, client_id);

    if (graphApiResult.status) {
      const result = await getEmailsGraph(graphApiResult.access_token, normalizeGraphMailbox(mailbox));
      return res.status(200).json(result);
    }

    const accessToken = await getAccessToken(refresh_token, client_id);
    const authString = generateAuthString(email, accessToken);

    const imap = new Imap({
      user: email,
      xoauth2: authString,
      host: 'outlook.office365.com',
      port: 993,
      tls: true,
      tlsOptions: {
        rejectUnauthorized: false
      }
    });

    let responseHandled = false;
    const sendJsonResponse = (statusCode, data) => {
      if (responseHandled) return;
      responseHandled = true;
      res.status(statusCode).json(data);
    };
    const sendHtmlResponse = (html) => {
      if (responseHandled) return;
      responseHandled = true;
      res.status(200).send(html);
    };

    imap.once('ready', async () => {
      try {
        await new Promise((resolve, reject) => {
          imap.openBox(mailbox, true, (err, box) => {
            if (err) return reject(err);
            resolve(box);
          });
        });

        const results = await new Promise((resolve, reject) => {
          imap.search(['ALL'], (err, searchResults) => {
            if (err) return reject(err);
            resolve((searchResults || []).slice(-1));
          });
        });

        if (results.length === 0) {
          imap.end();
          return sendJsonResponse(200, null);
        }

        const responseData = await new Promise((resolve, reject) => {
          const parseTasks = [];
          const f = imap.fetch(results, { bodies: '' });

          f.on('message', (msg, seqno) => {
            msg.on('body', stream => {
              parseTasks.push(new Promise((resolveParse, rejectParse) => {
                simpleParser(stream, (err, mail) => {
                  if (err) return rejectParse(new Error(`Failed to parse message ${seqno}: ${err.message}`));

                  try {
                    resolveParse(buildResponseData(mail, seqno));
                  } catch (parseError) {
                    rejectParse(new Error(`Failed to normalize message ${seqno}: ${parseError.message}`));
                  }
                });
              }));
            });
          });

          f.once('error', reject);
          f.once('end', async () => {
            try {
              const parsedMessages = await Promise.all(parseTasks);
              if (parsedMessages.length === 0) {
                reject(new Error('No email content was parsed'));
                return;
              }
              resolve(parsedMessages[0]);
            } catch (err) {
              reject(err);
            }
          });
        });

        imap.end();

        if (response_type === 'json') {
          return sendJsonResponse(200, responseData);
        }

        if (response_type === 'html') {
          const safeSend = escapeHtml(responseData.send);
          const safeSubject = escapeHtml(responseData.subject);
          const safeDate = escapeHtml(responseData.date);
          const safeText = escapeHtml(responseData.text || '').replace(/\n/g, '<br>');
          const htmlResponse = `
            <html>
              <body style="font-family: Arial, sans-serif; line-height: 1.6; margin: 0; padding: 20px; background-color: #f9f9f9;">
                <div style="margin: 0 auto; background: #fff; padding: 20px; border: 1px solid #ddd; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
                  <h1 style="color: #333;">邮件信息</h1>
                  <p><strong>邮件ID:</strong> ${escapeHtml(responseData.id)}</p>
                  <p><strong>Message-ID:</strong> ${escapeHtml(responseData.messageId)}</p>
                  <p><strong>模式:</strong> ${escapeHtml(responseData.mode)}</p>
                  <p><strong>发件人:</strong> ${safeSend}</p>
                  <p><strong>主题:</strong> ${safeSubject}</p>
                  <p><strong>日期:</strong> ${safeDate}</p>
                  <div style="background: #f4f4f4; padding: 10px; border: 1px solid #ddd;">
                    <p><strong>内容:</strong></p>
                    <p>${safeText}</p>
                  </div>
                </div>
              </body>
            </html>
          `;
          return sendHtmlResponse(htmlResponse);
        }

        return sendJsonResponse(400, { error: 'Invalid response_type. Use "json" or "html".' });
      } catch (err) {
        imap.end();
        sendJsonResponse(500, { error: err.message });
      }
    });

    imap.once('error', err => {
      sendJsonResponse(500, { error: err.message });
    });

    imap.connect();
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
};
