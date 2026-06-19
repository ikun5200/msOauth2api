const Imap = require('node-imap');
const simpleParser = require('mailparser').simpleParser;
const {
  generateAuthString,
  getAccessToken,
  graphApi,
  normalizeGraphMailbox,
  validatePassword
} = require('../lib/oauth-utils');

const ALLOWED_MAILBOXES = ['INBOX', 'Junk'];

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
  let nextLink = `https://graph.microsoft.com/v1.0/me/mailFolders/${mailbox}/messages?$top=1000&$select=id,from,subject,bodyPreview,body,createdDateTime,internetMessageId`;
  let emails = [];

  while (nextLink) {
    const response = await fetch(nextLink, {
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
    emails = emails.concat(responseData.value || []);
    nextLink = responseData['@odata.nextLink'] || null;
  }

  return emails.map(item => ({
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
  let { refresh_token, client_id, email, mailbox } = params;

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

    const emailList = [];
    let responseHandled = false;
    const sendResponse = (statusCode, data) => {
      if (responseHandled) return;
      responseHandled = true;
      res.status(statusCode).json(data);
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
            resolve(searchResults || []);
          });
        });

        if (results.length === 0) {
          imap.end();
          return sendResponse(200, []);
        }

        const parseTasks = [];
        const f = imap.fetch(results, { bodies: '' });

        f.on('message', (msg, seqno) => {
          msg.on('body', stream => {
            parseTasks.push(new Promise((resolve, reject) => {
              simpleParser(stream, (err, mail) => {
                if (err) return reject(new Error(`Failed to parse message ${seqno}: ${err.message}`));

                try {
                  emailList.push(buildResponseData(mail, seqno));
                  resolve();
                } catch (parseError) {
                  reject(new Error(`Failed to normalize message ${seqno}: ${parseError.message}`));
                }
              });
            }));
          });
        });

        await new Promise((resolve, reject) => {
          f.once('error', reject);
          f.once('end', resolve);
        });

        await Promise.all(parseTasks);
        sendResponse(200, emailList);
        imap.end();
      } catch (err) {
        imap.end();
        sendResponse(500, { error: err.message });
      }
    });

    imap.once('error', err => {
      sendResponse(500, { error: err.message });
    });

    imap.connect();
  } catch (error) {
    return res.status(500).json({ error: error.message });
  }
};
