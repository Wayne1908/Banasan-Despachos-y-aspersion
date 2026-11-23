export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  const { prompt, images = [] } = req.body || {};
  const geminiKey = (process.env.GEMINI_API_KEY || '').trim();
  const anthropicKey = (process.env.ANTHROPIC_API_KEY || '').trim();

  if (!geminiKey && !anthropicKey) {
    return res.status(500).json({ error: 'Falta GEMINI_API_KEY en Vercel' });
  }

  try {
    let rawText = '';

    if (geminiKey) {
      const parts = [{ text: prompt }];
      for (const img of images) {
        parts.push({
          inlineData: {
            mimeType: img.type || 'image/jpeg',
            data: img.data
          }
        });
      }

      const payload = JSON.stringify({
        contents: [{ role: 'user', parts }],
        generationConfig: {
          temperature: 0,
          responseMimeType: 'application/json'
        }
      });

      const endpoints = [
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(geminiKey)}`,
        `https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(geminiKey)}`
      ];

      let okRes = null;
      let lastErr = '';
      for (const url of endpoints) {
        const r = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: payload
        });
        if (r.ok) {
          okRes = await r.json();
          break;
        }
        lastErr = await r.text();
      }

      if (!okRes) {
        return res.status(502).json({ error: 'Error Gemini: ' + lastErr });
      }
      rawText = okRes.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
    } else {
      const content = [];
      for (const img of images) {
        content.push({
          type: 'image',
          source: {
            type: 'base64',
            media_type: img.type || 'image/jpeg',
            data: img.data
          }
        });
      }
      content.push({ type: 'text', text: prompt });

      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': anthropicKey,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 4096,
          temperature: 0,
          messages: [{ role: 'user', content }]
        })
      });

      if (!r.ok) {
        const errTxt = await r.text();
        return res.status(502).json({ error: 'Error Anthropic: ' + errTxt });
      }
      const j = await r.json();
      rawText = j.content?.[0]?.text || '[]';
    }

    const cleaned = rawText.replace(/```json|```/gi, '').trim();
    const data = JSON.parse(cleaned);
    return res.status(200).json({ data });
  } catch (e) {
    return res.status(500).json({ error: e.message || 'Error procesando IA' });
  }
}
