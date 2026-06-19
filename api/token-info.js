const { refreshTokens, validatePassword } = require('../lib/oauth-utils');

module.exports = async (req, res) => {
  if (req.method !== 'GET' && req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  if (!validatePassword(req, res)) return;

  const params = req.method === 'GET' ? req.query : req.body;
  const { refresh_token, client_id, email } = params;

  if (!refresh_token || !client_id) {
    return res.status(400).json({ error: 'Missing required parameters: refresh_token, client_id' });
  }

  async function checkGraphApiSupport() {
    try {
      const tokenResult = await refreshTokens(refresh_token, client_id, {
        scope: 'https://graph.microsoft.com/.default'
      });
      const scope = tokenResult.scope || '';
      const hasMailReadWrite = scope.includes('https://graph.microsoft.com/Mail.ReadWrite');
      const hasMailRead = scope.includes('https://graph.microsoft.com/Mail.Read');

      return {
        supported: hasMailReadWrite || hasMailRead,
        access_token: tokenResult.access_token,
        scope,
        permissions: {
          mailReadWrite: hasMailReadWrite,
          mailRead: hasMailRead,
          mailSend: scope.includes('https://graph.microsoft.com/Mail.Send'),
          userRead: scope.includes('https://graph.microsoft.com/User.Read')
        }
      };
    } catch (error) {
      return {
        supported: false,
        error: error.message,
        access_token: null,
        scope: null
      };
    }
  }

  async function checkImapSupport() {
    try {
      const tokenResult = await refreshTokens(refresh_token, client_id);
      const scope = tokenResult.scope || '';
      const hasImapAccess = scope.includes('https://outlook.office.com/IMAP.AccessAsUser.All')
        || scope.includes('https://outlook.office.com/POP.AccessAsUser.All');

      return {
        supported: hasImapAccess,
        access_token: tokenResult.access_token,
        scope,
        permissions: {
          imapAccess: scope.includes('https://outlook.office.com/IMAP.AccessAsUser.All'),
          popAccess: scope.includes('https://outlook.office.com/POP.AccessAsUser.All')
        }
      };
    } catch (error) {
      return {
        supported: false,
        error: error.message,
        access_token: null,
        scope: null
      };
    }
  }

  try {
    const [graphResult, imapResult] = await Promise.all([
      checkGraphApiSupport(),
      checkImapSupport()
    ]);

    let primaryMode = 'unknown';
    const supportedModes = [];

    if (graphResult.supported) {
      primaryMode = 'graph';
      supportedModes.push('graph');
    }

    if (imapResult.supported) {
      if (primaryMode === 'unknown') primaryMode = 'imap';
      supportedModes.push('imap');
    }

    const graphCanWriteMail = Boolean(graphResult.permissions?.mailReadWrite);
    const tokenInfo = {
      primaryMode,
      supportedModes,
      email: email || 'not_provided',
      capabilities: {
        graphApi: {
          supported: graphResult.supported,
          permissions: graphResult.permissions || {},
          scope: graphResult.scope,
          error: graphResult.error || null
        },
        imap: {
          supported: imapResult.supported,
          permissions: imapResult.permissions || {},
          scope: imapResult.scope,
          error: imapResult.error || null
        }
      },
      features: {
        readEmails: graphResult.supported || imapResult.supported,
        deleteEmails: graphCanWriteMail || imapResult.supported,
        sendEmails: graphResult.supported,
        clearInbox: graphCanWriteMail || imapResult.supported,
        clearJunk: graphCanWriteMail || imapResult.supported
      },
      recommendations: []
    };

    if (graphResult.supported && imapResult.supported) {
      tokenInfo.recommendations.push('您的token同时支持Graph API和IMAP模式，建议优先使用Graph API以获得更好的性能');
    } else if (graphResult.supported) {
      tokenInfo.recommendations.push('您的token支持Graph API模式，可以使用所有高级功能');
    } else if (imapResult.supported) {
      tokenInfo.recommendations.push('您的token仅支持IMAP模式，功能相对有限但兼容性更好');
    } else {
      tokenInfo.recommendations.push('您的token似乎不支持邮件操作，请检查应用权限配置');
    }

    return res.status(200).json({
      success: true,
      tokenInfo,
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    return res.status(500).json({
      success: false,
      error: 'Token information detection failed',
      details: error.message
    });
  }
};
