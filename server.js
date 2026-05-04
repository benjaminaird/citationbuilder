const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/improve', async (req, res) => {
  try {
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: 'Anthropic API key is not configured on the server.' });
    }

    const { form, soa, citation } = req.body || {};
    if (!form || !soa || !citation) {
      return res.status(400).json({ error: 'Missing form, SOA, or citation text.' });
    }

    const award = form.award || 'NAM';
    const mustBeAllCaps = award === 'NAM' || award === 'NAVCOM';

    const systemPrompt = `You are an expert U.S. Marine Corps awards editor. Your job is controlled phrasing improvement only. You must improve clarity, action-impact-result language, and professional military tone without inventing facts. You must preserve required citation structure, opening phrase, closing phrase, pronouns, dates, billet, unit, command, and award level. Return JSON only. No markdown.`;

    const userPrompt = `Improve the wording of this draft Summary of Action and Proposed Citation.

STRICT RULES:
1. Do NOT invent facts, numbers, billets, units, awards, dates, or impacts.
2. Preserve the required citation opening sentence and closing sentence.
3. Citation must remain exactly one paragraph.
4. ${mustBeAllCaps ? 'The entire citation MUST be ALL CAPS.' : 'Use normal capitalization for the citation.'}
5. The citation must be derived only from the SOA and user achievements.
6. Keep the output concise and suitable for Marine Corps award routing.
7. Return ONLY valid JSON in this exact shape:
{"soa":"...","citation":"...","notes":["..."]}

MARINE / AWARD DATA:
${JSON.stringify(form, null, 2)}

CURRENT SUMMARY OF ACTION:
${soa}

CURRENT PROPOSED CITATION:
${citation}`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-5-20250929',
        max_tokens: 1800,
        temperature: 0.2,
        system: systemPrompt,
        messages: [{ role: 'user', content: userPrompt }]
      })
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Anthropic API error:', data);
      return res.status(response.status).json({
        error: data?.error?.message || 'Anthropic API request failed.'
      });
    }

    const text = data?.content?.find(block => block.type === 'text')?.text;
    if (!text) {
      return res.status(500).json({ error: 'Anthropic returned an empty response.' });
    }

    let parsed;
    try {
      const jsonText = text.trim().replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```$/i, '').trim();
      parsed = JSON.parse(jsonText);
    } catch (err) {
      console.error('Could not parse Anthropic response:', text);
      return res.status(500).json({ error: 'Could not parse AI response as JSON.' });
    }

    if (!parsed.soa || !parsed.citation) {
      return res.status(500).json({ error: 'AI response did not include both SOA and citation.' });
    }

    if (mustBeAllCaps) {
      parsed.citation = parsed.citation.toUpperCase();
    }

    parsed.citation = String(parsed.citation).replace(/\s+/g, ' ').trim();
    parsed.soa = String(parsed.soa).trim();
    parsed.notes = Array.isArray(parsed.notes) ? parsed.notes : [];

    res.json(parsed);
  } catch (err) {
    console.error('AI improve route error:', err);
    res.status(500).json({ error: 'Server error while improving wording.' });
  }
});

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`CitationBuilder running on port ${PORT}`);
});
