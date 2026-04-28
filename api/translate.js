export default async function handler(req, res) {
  const { sl = 'auto', tl = 'en', q = '' } = req.query;
  if (!q) {
    res.status(400).json({ error: 'Missing q' });
    return;
  }
  const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=${encodeURIComponent(sl)}&tl=${encodeURIComponent(tl)}&dt=t&q=${encodeURIComponent(q)}`;
  try {
    const upstream = await fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0' },
    });
    const body = await upstream.text();
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Cache-Control', 'public, max-age=86400');
    res.status(upstream.status).send(body);
  } catch (e) {
    res.status(502).json({ error: 'Upstream fetch failed', detail: String(e) });
  }
}
