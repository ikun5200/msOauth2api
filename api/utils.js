const crypto = require('crypto');

const TOKEN_ENDPOINT = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token';

function generateCodeVerifier() {
  return crypto.randomBytes(32).toString('base64url');
}

function bufferToBase64Url(buffer) {
  return buffer.toString('base64')
    .replace(/\+/g, '-')
    .replace(/\//g, '_')
    .replace(/=/g, '');
}

async function generateCodeChallenge(verifier) {
  const hash = crypto.createHash('sha256').update(verifier).digest();
  return bufferToBase64Url(hash);
}

function generateAuthString(user, accessToken) {
  const authString = `user=${user}\x01auth=Bearer ${accessToken}\x01\x01`;
  return Buffer.from(authString).toString('base64');
}

function estimateRtExpiresAt() {
  const rtExpiresAt = new Date();
  rtExpiresAt.setDate(rtExpiresAt.getDate() + 90);
  return rtExpiresAt.toISOString();
}

function isRefreshTokenInvalid(error) {
  const message = String(error?.message || error || '');
  return message.includes('invalid_grant')
    || message.includes('AADSTS70008')
    || message.includes('AADSTS700082');
}

function validatePassword(req, res, envName = 'PASSWORD') {
  const params = req.method === 'GET' ? req.query : req.body;
  const expectedPassword = process.env[envName];

  if (expectedPassword && params?.password !== expectedPassword) {
    res.status(401).json({ error: '密码验证失败' });
    return false;
  }

  return true;
}

async function refreshTokens(refresh_token, client_id, extraParams = {}) {
  const response = await fetch(TOKEN_ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body: new URLSearchParams({
      client_id,
      grant_type: 'refresh_token',
      refresh_token,
      ...extraParams
    }).toString()
  });

  const responseText = await response.text();

  if (!response.ok) {
    throw new Error(`HTTP error! status: ${response.status}, response: ${responseText}`);
  }

  try {
    const data = JSON.parse(responseText);
    return {
      access_token: data.access_token,
      refresh_token: data.refresh_token || null,
      expires_in: data.expires_in || null,
      scope: data.scope || '',
      rt_expires_at: estimateRtExpiresAt()
    };
  } catch (parseError) {
    throw new Error(`Failed to parse JSON: ${parseError.message}, response: ${responseText}`);
  }
}

async function getAccessToken(refresh_token, client_id) {
  const tokenResult = await refreshTokens(refresh_token, client_id);
  return tokenResult.access_token;
}

async function graphApi(refresh_token, client_id) {
  const tokenResult = await refreshTokens(refresh_token, client_id, {
    scope: 'https://graph.microsoft.com/.default'
  });
  const scope = tokenResult.scope || '';
  const hasMailReadWrite = scope.includes('https://graph.microsoft.com/Mail.ReadWrite');
  const hasMailRead = scope.includes('https://graph.microsoft.com/Mail.Read');

  return {
    access_token: tokenResult.access_token,
    refresh_token: tokenResult.refresh_token,
    expires_in: tokenResult.expires_in,
    rt_expires_at: tokenResult.rt_expires_at,
    scope,
    status: hasMailReadWrite || hasMailRead
  };
}

function normalizeGraphMailbox(mailbox) {
  if (mailbox === 'INBOX') return 'inbox';
  if (mailbox === 'Junk') return 'junkemail';
  return 'inbox';
}

module.exports = {
  generateAuthString,
  generateCodeVerifier,
  generateCodeChallenge,
  getAccessToken,
  get_access_token: getAccessToken,
  graphApi,
  graph_api: graphApi,
  refreshTokens,
  refresh_tokens: refreshTokens,
  isRefreshTokenInvalid,
  is_refresh_token_invalid: isRefreshTokenInvalid,
  normalizeGraphMailbox,
  validatePassword
};
