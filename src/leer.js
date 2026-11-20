// Función serverless de Vercel: lee imágenes/PDF (páginas como imagen) con Claude y devuelve JSON.
// Requiere la variable de entorno ANTHROPIC_API_KEY (opcional: ANTHROPIC_MODEL).
export const config = { maxDuration: 60 };

export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Usa POST' });
  const key = process.env.ANTHROPIC_API_KEY;
  if (!key) return res.status(500).json({ error: 'Falta ANTHROPIC_API_KEY en Vercel' });
  const { prompt = '', images = [] } = req.body || {};
  const content = [
    ...images.map(i => ({ type: 'image', source: { type: 'base64', media_type: i.type || 'image/jpeg', data: i.data } })),
    { type: 'text', text: prompt + '\n\nResponde únicamente con JSON válido, sin texto adicional ni bloques de código.' },
  ];
  try {
    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({
        model: process.env.ANTHROPIC_MODEL || 'claude-sonnet-5-5',
        max_tokens: 8000,
        messages: [{ role: 'user', content }],
      }),
    });
    const j = await r.json();
    if (!r.ok) return res.status(r.status).json({ error: j.error?.message || 'Error de la API' });
    const t = (j.content || []).map(c => c.text || '').join('').replace(/```json|```/g, '');
    const a = t.search(/[\[{]/), b = Math.max(t.lastIndexOf(']'), t.lastIndexOf('}'));
    res.status(200).json({ data: JSON.parse(t.slice(a, b + 1)) });
  } catch (e) {
    res.status(500).json({ error: String(e.message || e) });
  }
}
