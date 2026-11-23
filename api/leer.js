export default async function handler(req, res) {
  const geminiKey = (process.env.GEMINI_API_KEY || '').trim();
  const anthropicKey = (process.env.ANTHROPIC_API_KEY || '').trim();

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  async function tryModel(modelPath, parts, useJson) {
    const genCfg = { temperature: 0 };
    if (useJson) genCfg.responseMimeType = 'application/json';
    const cleanPath = modelPath.startsWith('models/') ? modelPath : `models/${modelPath}`;
    const url = `https://generativelanguage.googleapis.com/v1beta/${cleanPath}:generateContent?key=${encodeURIComponent(geminiKey)}`;

    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        contents: [{ role: 'user', parts }],
        generationConfig: genCfg
      })
    });

    if (!r.ok) {
      const txt = await r.text();
      const err = new Error(`[${cleanPath} -> ${r.status}] ${txt}`);
      err.status = r.status;
      throw err;
    }

    const okRes = await r.json();
    const resParts = okRes.candidates?.[0]?.content?.parts || [];
    const textPart = resParts.find((p) => p.text && !p.thought) || resParts[resParts.length - 1];
    return textPart?.text || '[]';
  }

  async function callGemini(parts, useJson = true) {
    const errs = [];
    const models = [
      'models/gemini-3.8-flash',
      'models/gemini-flash-latest',
      'models/gemini-flash-lite-latest',
      'models/gemini-3.8-pro',
      'models/gemini-3-flash-preview'
    ];

    for (let round = 0; round < 2; round++) {
      let saw503 = false;
      for (const m of models) {
        try {
          return await tryModel(m, parts, useJson);
        } catch (e) {
          if (e.status === 503) saw503 = true;
          errs.push(e.message);
        }
      }
      if (saw503 && round === 0) {
        await sleep(2000);
      } else {
        break;
      }
    }

    throw new Error(errs.slice(0, 5).join(' | '));
  }

  if (req.method === 'GET') {
    if (!geminiKey && !anthropicKey) {
      return res.status(500).json({ ok: false, error: 'Falta GEMINI_API_KEY en Vercel.' });
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
