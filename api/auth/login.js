const { generateCodeVerifier, generateCodeChallenge } = require('../utils');

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const clientId = process.env.CLIENT_ID;
  const tenantId = process.env.TENANT_ID || 'common';
  const redirectUri = process.env.REDIRECT_URI;

  if (!clientId || !redirectUri) {
    return res.status(500).json({
      error: 'Server configuration error',
      details: 'Missing CLIENT_ID or REDIRECT_URI'
    });
  }

  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  res.setHeader('Set-Cookie', [
    `code_verifier=${codeVerifier}; Path=/; HttpOnly; Secure; Max-Age=300; SameSite=Lax`
  ]);

  const scope = encodeURIComponent('openid profile email Mail.Read Mail.ReadWrite Mail.Send offline_access');
  const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?`
    + `client_id=${encodeURIComponent(clientId)}&`
    + 'response_type=code&'
    + `redirect_uri=${encodeURIComponent(redirectUri)}&`
    + `scope=${scope}&`
    + `code_challenge=${encodeURIComponent(codeChallenge)}&`
    + 'code_challenge_method=S256&'
    + 'response_mode=query';

  return res.redirect(authUrl);
};
