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
  const { refresh_token, client_id, email } = params;

  if (!refresh_token || !client_id || !email) {
    return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id, or email' });
  }

  try {
    const graphApiResult = await graphApi(refresh_token, client_id);

    if (graphApiResult.status && graphApiResult.scope.includes('https://graph.microsoft.com/Mail.ReadWrite')) {
      return await processFolderGraphApi(graphApiResult.access_token, 'junkemail', res);
    }

    return await processJunkImap(refresh_token, client_id, email, res);
  } catch (error) {
    return res.status(500).json({ error: 'Error', details: error.message });
  }
};

async function processFolderGraphApi(accessToken, folder, res) {
  try {
    const deleteBatchSize = 20;

    async function getAllMessages() {
      let allMessages = [];
      let nextLink = `https://graph.microsoft.com/v1.0/me/mailFolders/${folder}/messages?$select=id&$top=1000`;

      while (nextLink) {
        const response = await fetch(nextLink, {
          method: 'GET',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        });

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to get messages: ${response.status}, ${errorText}`);
        }

        const data = await response.json();
        allMessages = allMessages.concat(data.value || []);
        nextLink = data['@odata.nextLink'] || null;
      }

      return allMessages;
    }

    async function deleteMessage(messageId) {
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
    }

    const messages = await getAllMessages();

    if (messages.length === 0) {
      return res.json({ message: 'No junk emails found.', mode: 'graph', stats: { total: 0, deleted: 0, failed: 0 } });
    }

    let deletedCount = 0;
    let failedCount = 0;

    for (let index = 0; index < messages.length; index += deleteBatchSize) {
      const messageBatch = messages.slice(index, index + deleteBatchSize);
      const deleteResults = await Promise.allSettled(messageBatch.map(message => deleteMessage(message.id)));

      deleteResults.forEach(result => {
        if (result.status === 'fulfilled') deletedCount++;
        else failedCount++;
      });
    }

    return res.json({
      message: 'Junk emails processed successfully via Graph API.',
      mode: 'graph',
      stats: {
        total: messages.length,
        deleted: deletedCount,
        failed: failedCount
      }
    });
  } catch (error) {
    return res.status(500).json({ error: 'Graph API Error', details: error.message });
  }
}

async function processJunkImap(refreshToken, clientId, email, res) {
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
        const junkFolders = ['Junk Email', 'Junk', 'Spam'];
        let selectedFolder = null;
        let totalMessages = 0;

        for (const folder of junkFolders) {
          try {
            const box = await new Promise((resolve, reject) => {
              imap.openBox(folder, false, (err, openBox) => {
                if (err) return reject(err);
                resolve(openBox);
              });
            });
            selectedFolder = folder;
            totalMessages = box.messages.total;
            break;
          } catch (err) {
            selectedFolder = null;
          }
        }

        if (!selectedFolder) {
          sendResponse(404, { error: 'No junk email folder found', mode: 'imap' });
          imap.end();
          return;
        }

        if (totalMessages === 0) {
          sendResponse(200, {
            message: 'No junk emails found.',
            mode: 'imap',
            stats: { total: 0, deleted: 0, failed: 0 }
          });
          imap.end();
          return;
        }

        const results = await new Promise((resolve, reject) => {
          imap.search(['ALL'], (err, searchResults) => {
            if (err) return reject(err);
            resolve(searchResults || []);
          });
        });

        await new Promise((resolve, reject) => {
          imap.setFlags(results, ['\\Deleted'], err => {
            if (err) return reject(err);
            imap.expunge(expungeErr => {
              if (expungeErr) return reject(expungeErr);
              resolve();
            });
          });
        });

        sendResponse(200, {
          message: 'Junk emails processed successfully via IMAP.',
          mode: 'imap',
          stats: {
            total: totalMessages,
            deleted: results.length,
            failed: 0
          }
        });
        imap.end();
      } catch (err) {
        imap.end();
        sendResponse(500, { error: 'IMAP processing error', details: err.message, mode: 'imap' });
      }
    });

    imap.once('error', err => {
      sendResponse(500, { error: 'IMAP connection error', details: err.message, mode: 'imap' });
    });

    setTimeout(() => {
      if (!responseHandled) {
        sendResponse(500, { error: 'IMAP operation timeout', mode: 'imap' });
        imap.end();
      }
    }, 60000);

    imap.connect();
  } catch (error) {
    return res.status(500).json({ error: 'IMAP Error', details: error.message });
  }
}
