const { getAccessToken } = require('./utils');

module.exports = async (req, res) => {
  const params = req.method === 'GET' ? req.query : req.body;
  const expectedSendPassword = process.env.SEND_PASSWORD;
  const expectedPassword = process.env.PASSWORD;
  const providedPassword = params.send_password || params.password;

  if (expectedSendPassword && providedPassword !== expectedSendPassword) {
    return res.status(401).json({ error: '密码验证失败' });
  }

  if (!expectedSendPassword && expectedPassword && providedPassword !== expectedPassword) {
    return res.status(401).json({ error: '密码验证失败' });
  }

  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const {
      refresh_token,
      client_id,
      email,
      to,
      subject,
      text,
      html
    } = params;

    if (!refresh_token || !client_id || !email || !to || !subject || (!text && !html)) {
      return res.status(400).json({ error: 'Missing required parameters' });
    }

    const accessToken = await getAccessToken(refresh_token, client_id);
    const toRecipients = to.split(',').map(recipient => ({
      emailAddress: {
        address: recipient.trim()
      }
    }));

    const emailMessage = {
      message: {
        subject,
        body: {
          contentType: html ? 'HTML' : 'Text',
          content: html || text
        },
        toRecipients
      },
      saveToSentItems: true
    };

    const response = await fetch('https://graph.microsoft.com/v1.0/me/sendMail', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(emailMessage)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to send email: ${response.status}, ${errorText}`);
    }

    const messageId = `${Date.now()}-${Math.random().toString(36).slice(2, 15)}`;
    return res.status(200).json({ message: 'Email sent successfully', messageId });
  } catch (error) {
    return res.status(500).json({ error: 'Failed to send email', details: error.message });
  }
};
