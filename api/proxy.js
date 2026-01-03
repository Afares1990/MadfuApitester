// Simple proxy for Vercel serverless functions.
// Usage: /api/proxy?url=<encoded-target-url>
// For non-GET methods send the JSON body as usual and include the `url` query param.

module.exports = async (req, res) => {
  try {
    // Get target URL from query string ?url=<encoded>
    const target = (req.query && req.query.url) || (req.url && new URL(req.url, `http://${req.headers.host}`).searchParams.get('url'));
    if (!target) {
      res.status(400).json({ error: 'Missing `url` query parameter.' });
      return;
    }

    // Build fetch options
    const method = req.method || 'GET';
    const headers = Object.assign({}, req.headers);

    // Remove hop-by-hop headers and host-related headers that should not be forwarded
    delete headers['host'];
    delete headers['cookie'];
    delete headers['accept-encoding'];

    // Prepare body for non-GET requests
    let body;
    if (method !== 'GET' && method !== 'HEAD') {
      if (req.body && typeof req.body === 'object' && req.headers['content-type'] && req.headers['content-type'].includes('application/json')) {
        body = JSON.stringify(req.body);
      } else {
        // If Vercel didn't parse body, use rawBody (works in some setups) or stringified body fallback
        body = req.rawBody || (req.body ? JSON.stringify(req.body) : undefined);
      }
    }

    // Use global fetch (Node 18+ on Vercel)
    const fetchRes = await fetch(target, {
      method,
      headers,
      body,
      redirect: 'follow'
    });

    // Copy status
    res.status(fetchRes.status);

    // Copy response headers (but avoid hop-by-hop)
    fetchRes.headers.forEach((value, key) => {
      if (['transfer-encoding', 'content-encoding', 'connection'].includes(key.toLowerCase())) return;
      res.setHeader(key, value);
    });

    // Add permissive CORS for testing
    res.setHeader('access-control-allow-origin', '*');
    res.setHeader('access-control-allow-methods', 'GET,HEAD,PUT,PATCH,POST,DELETE,OPTIONS');
    res.setHeader('access-control-allow-headers', 'Content-Type, Authorization, APIKey, AppCode, PlatformTypeId, Token');

    // Handle OPTIONS preflight
    if (req.method === 'OPTIONS') {
      res.status(204).end();
      return;
    }

    // Stream the response body back to the client
    const buffer = await fetchRes.arrayBuffer();
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error('Proxy error', err);
    res.status(500).json({ error: 'Proxy error', details: String(err) });
  }
};
