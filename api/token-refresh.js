const { refreshTokens, isRefreshTokenInvalid, validatePassword } = require('./utils');

module.exports = async (req, res) => {
  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  if (!validatePassword(req, res)) return;

  const params = req.method === 'GET' ? req.query : req.body;
  const { refresh_token, client_id, email } = params;

  if (!refresh_token || !client_id) {
    return res.status(400).json({
      error: 'Missing required parameters: refresh_token or client_id'
    });
  }

  try {
    const tokenResult = await refreshTokens(refresh_token, client_id);
    const newRefreshToken = tokenResult.refresh_token || refresh_token;

    return res.status(200).json({
      success: true,
      email: email || 'not_provided',
      refresh_token: newRefreshToken,
      token_info: {
        new_refresh_token: newRefreshToken,
        expires_in: tokenResult.expires_in,
        rt_expires_at: tokenResult.rt_expires_at,
        rt_was_refreshed: Boolean(tokenResult.refresh_token),
        rt_reauth_required: false
      }
    });
  } catch (error) {
    const requiresReauth = isRefreshTokenInvalid(error);
    return res.status(requiresReauth ? 401 : 500).json({
      success: false,
      error: requiresReauth
        ? 'Refresh token has expired, re-authorization required'
        : error.message,
      rt_reauth_required: requiresReauth
    });
  }
};
