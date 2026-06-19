function parseCookies(cookieHeader = '') {
  return cookieHeader
    .split(';')
    .map(cookie => cookie.trim())
    .filter(Boolean)
    .reduce((cookies, cookie) => {
      const separatorIndex = cookie.indexOf('=');
      if (separatorIndex === -1) return cookies;
      const key = decodeURIComponent(cookie.slice(0, separatorIndex));
      const value = decodeURIComponent(cookie.slice(separatorIndex + 1));
      cookies[key] = value;
      return cookies;
    }, {});
}

module.exports = async function handler(req, res) {
  if (req.method !== 'GET') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  const { code, error, error_description } = req.query;

  if (error) {
    return res.status(400).json({ error, error_description });
  }

  if (!code) {
    return res.status(400).json({ error: 'Missing authorization code' });
  }

  const cookies = parseCookies(req.headers.cookie || '');
  const codeVerifier = cookies.code_verifier;

  if (!codeVerifier) {
    return res.status(400).json({ error: 'Session expired or invalid. Please login again.' });
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

  const tokenEndpoint = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const params = new URLSearchParams({
    client_id: clientId,
    grant_type: 'authorization_code',
    code,
    redirect_uri: redirectUri,
    code_verifier: codeVerifier,
    scope: 'openid profile email Mail.Read Mail.ReadWrite Mail.Send offline_access'
  });

  try {
    const response = await fetch(tokenEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: params.toString()
    });

    const responseText = await response.text();

    if (!response.ok) {
      return res.status(500).json({
        error: 'Failed to exchange token',
        details: responseText
      });
    }

    const tokens = JSON.parse(responseText);

    res.setHeader('Set-Cookie', [
      'code_verifier=; Path=/; HttpOnly; Secure; Max-Age=0; SameSite=Lax'
    ]);

    return res.status(200).json({
      access_token: tokens.access_token,
      refresh_token: tokens.refresh_token,
      expires_in: tokens.expires_in,
      token_type: tokens.token_type
    });
  } catch (err) {
    return res.status(500).json({ error: 'Internal Server Error', details: err.message });
  }
};
