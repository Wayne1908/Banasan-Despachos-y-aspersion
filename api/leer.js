export default async function handler(req, res) {
  const geminiKey = (process.env.GEMINI_API_KEY || '').trim();
  const anthropicKey = (process.env.ANTHROPIC_API_KEY || '').trim();

  async function callGemini(parts, useJson = true) {
    const genCfg = { temperature: 0 };
    if (useJson) genCfg.responseMimeType = 'application/json';
    const payload = JSON.stringify({
      contents: [{ role: 'user', parts }],
      generationConfig: genCfg
    });

    const targets = [
      `https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(geminiKey)}`,
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${encodeURIComponent(geminiKey)}`,
      `https://aiplatform.googleapis.com/v1/publishers/google/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(geminiKey)}`,
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${encodeURIComponent(geminiKey)}`
    ];

    let lastErr = '';
    for (const url of targets) {
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: payload
      });
      if (r.ok) {
        const okRes = await r.json();
        return okRes.candidates?.[0]?.content?.parts?.[0]?.text || '[]';
      }
      lastErr = `(${r.status}) ` + await r.text();
    }
    throw new Error(lastErr);
  }

  if (req.method === 'GET') {
    if (!geminiKey && !anthropicKey) {
      return res.status(500).json({ ok: false, error: 'Falta GEMINI_API_KEY en Vercel (Environment Variables).' });
    }
    try {
      const txt = geminiKey ? await callGemini([{ text: 'Responde solo OK' }], false) : 'Anthropic OK';
      return res.status(200).json({ ok: true, keyPrefix: geminiKey.slice(0, 6) + '...', reply: txt });
    } catch (e) {
      return res.status(502).json({ ok: false, keyPrefix: geminiKey.slice(0, 6) + '...', error: e.message });
    }
  }

  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  if (!geminiKey && !anthropicKey) {
    return res.status(500).json({ error: 'Falta GEMINI_API_KEY en Vercel (Settings -> Environment Variables).' });
  }

  const { prompt, images = [] } = req.body || {};

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
      rawText = await callGemini(parts, true);
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
