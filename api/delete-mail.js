const Imap = require('node-imap');
const {
  generateAuthString,
  getAccessToken,
  graphApi,
  validatePassword
} = require('../lib/oauth-utils');

module.exports = async (req, res) => {
  if (!validatePassword(req, res)) return;

  const params = req.method === 'GET' ? req.query : req.body;
  const { refresh_token, client_id, email, message_id } = params;
  const folderName = params.mailbox || 'INBOX';

  if (!refresh_token || !client_id || !email || !message_id) {
    return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, email, or message_id' });
  }

  try {
    const graphApiResult = await graphApi(refresh_token, client_id);

    if (graphApiResult.status && graphApiResult.scope.includes('https://graph.microsoft.com/Mail.ReadWrite')) {
      return await deleteSingleEmailGraphApi(graphApiResult.access_token, message_id, res);
    }

    return await deleteSingleEmailImap(refresh_token, client_id, email, message_id, folderName, res);
  } catch (error) {
    return res.status(500).json({ error: 'Error', details: error.message });
  }
};

async function deleteSingleEmailGraphApi(accessToken, messageId, res) {
  try {
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}`, {
      method: 'DELETE',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to delete message ${messageId}: ${response.status}, ${errorText}`);
    }

    return res.json({
      success: true,
      message: 'Email deleted successfully via Graph API.',
      mode: 'graph',
      messageId,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    return res.status(500).json({
      success: false,
      error: 'Graph API Error',
      details: error.message,
      mode: 'graph',
      messageId
    });
  }
}

async function deleteSingleEmailImap(refreshToken, clientId, email, messageId, folderName, res) {
  try {
    const accessToken = await getAccessToken(refreshToken, clientId);
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
    const sendResponse = (statusCode, data) => {
      if (responseHandled) return;
      responseHandled = true;
      res.status(statusCode).json(data);
    };

    imap.once('ready', async () => {
      try {
        await new Promise((resolve, reject) => {
          imap.openBox(folderName, false, (err, box) => {
            if (err) return reject(err);
            resolve(box);
          });
        });

        const searchResults = await new Promise((resolve, reject) => {
          imap.search([['HEADER', 'MESSAGE-ID', messageId]], (err, results) => {
            if (err) return reject(err);
            resolve(results || []);
          });
        });

        if (searchResults.length === 0) {
          sendResponse(404, {
            success: false,
            error: 'Email not found',
            mode: 'imap',
            messageId
          });
          imap.end();
          return;
        }

        await new Promise((resolve, reject) => {
          imap.setFlags(searchResults, ['\\Deleted'], err => {
            if (err) return reject(err);
            imap.expunge(expungeErr => {
              if (expungeErr) return reject(expungeErr);
              resolve();
            });
          });
        });

        sendResponse(200, {
          success: true,
          message: 'Email deleted successfully via IMAP.',
          mode: 'imap',
          messageId,
          timestamp: new Date().toISOString()
        });
        imap.end();
      } catch (error) {
        sendResponse(500, {
          success: false,
          error: 'IMAP processing error',
          details: error.message,
          mode: 'imap',
          messageId
        });
        imap.end();
      }
    });

    imap.once('error', err => {
      sendResponse(500, {
        success: false,
        error: 'IMAP connection error',
        details: err.message,
        mode: 'imap',
        messageId
      });
    });

    setTimeout(() => {
      if (!responseHandled) {
        sendResponse(500, {
          success: false,
          error: 'IMAP operation timeout',
          mode: 'imap',
          messageId
        });
        imap.end();
      }
    }, 30000);

    imap.connect();
  } catch (error) {
    return res.status(500).json({
      success: false,
      error: 'Failed to delete email',
      details: error.message,
      mode: 'imap',
      messageId
    });
  }
}
