const { refreshTokens } = require('../utils');

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { refresh_token } = req.body;

  if (!refresh_token) {
    return res.status(400).json({ error: 'Missing refresh_token' });
  }

  const clientId = process.env.CLIENT_ID;

  if (!clientId) {
    return res.status(500).json({
      error: 'Server configuration error',
      details: 'Missing CLIENT_ID'
    });
  }

  try {
    const tokens = await refreshTokens(refresh_token, clientId, {
      scope: 'openid profile email Mail.Read Mail.ReadWrite Mail.Send offline_access'
    });

    return res.status(200).json({
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token || refresh_token,
      expires_in: tokens.expires_in
    });
  } catch (err) {
    return res.status(500).json({ error: 'Refresh failed', details: err.message });
  }
};
