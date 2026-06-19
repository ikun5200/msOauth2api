const API_KEY = process.env.AI_API_KEY;
const API_URL = process.env.AI_API_URL;
const MODEL = process.env.AI_MODEL;
const PASSWORD = process.env.PASSWORD;

function sendEvent(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
}

async function readJsonBody(req) {
  if (req.body && typeof req.body === 'object') {
    return req.body;
  }

  if (typeof req.body === 'string') {
    return JSON.parse(req.body);
  }

  const chunks = [];
  for await (const chunk of req) {
    chunks.push(Buffer.from(chunk));
  }

  const rawBody = Buffer.concat(chunks).toString('utf8');
  return rawBody ? JSON.parse(rawBody) : {};
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: '只支持 POST 请求' });
  }

  if (!API_KEY || !API_URL || !MODEL) {
    return res.status(500).json({ error: 'AI_API_KEY、AI_API_URL 或 AI_MODEL 未配置' });
  }

  let body;
  try {
    body = await readJsonBody(req);
  } catch {
    return res.status(400).json({ error: '无效的 JSON' });
  }

  const { messages, password } = body;

  if (!messages) {
    return res.status(400).json({ error: '缺少 messages 参数' });
  }

  if (PASSWORD && password !== PASSWORD) {
    return res.status(401).json({ error: '密码验证失败' });
  }

  res.writeHead(200, {
    'Content-Type': 'text/event-stream; charset=utf-8',
    'Cache-Control': 'no-cache, no-transform',
    Connection: 'keep-alive',
    'X-Accel-Buffering': 'no'
  });

  res.write(': connected\n\n');

  try {
    sendEvent(res, 'status', { message: '正在连接 AI...' });

    const response = await fetch(`${API_URL}/v1/chat/completions`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${API_KEY}`
      },
      body: JSON.stringify({
        model: MODEL,
        messages,
        stream: true
      })
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`AI 请求失败: ${response.status} - ${errorText}`);
    }

    const reader = response.body.getReader();

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      res.write(Buffer.from(value));
    }

    res.end();
  } catch (error) {
    sendEvent(res, 'error', { error: error.message });
    res.end();
  }
};
