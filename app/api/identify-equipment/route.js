export const runtime = "nodejs";

const getOutputText = (data) => {
  if (data?.output_text) return data.output_text;

  return (data?.output || [])
    .flatMap((item) => item.content || [])
    .map((content) => content.text || content.output_text || "")
    .filter(Boolean)
    .join("\n");
};

const parseJsonFromText = (text) => {
  const raw = String(text || "").trim();

  try {
    return JSON.parse(raw);
  } catch {}

  const match = raw.match(/\{[\s\S]*\}/);
  if (!match) return null;

  try {
    return JSON.parse(match[0]);
  } catch {
    return null;
  }
};

const normalizeResult = (value = {}) => ({
  visualName: String(value.visualName || value.visual_name || "").trim(),
  visualDescription: String(value.visualDescription || value.visual_description || "").trim(),
  equipmentCategory: String(value.equipmentCategory || value.equipment_category || "").trim(),
  likelySearchTerms: Array.isArray(value.likelySearchTerms)
    ? value.likelySearchTerms.map((item) => String(item || "").trim()).filter(Boolean).slice(0, 8)
    : [],
  confidence: Number(value.confidence || 0),
  notes: String(value.notes || "").trim(),
});

export async function POST(request) {
  try {
    const { imageDataUrl } = await request.json();

    if (!imageDataUrl || !String(imageDataUrl).startsWith("data:image/")) {
      return Response.json({ error: "Missing image data." }, { status: 400 });
    }

    if (!process.env.OPENAI_API_KEY) {
      return Response.json(
        { error: "OPENAI_API_KEY is not configured in Vercel." },
        { status: 500 }
      );
    }

    const model = process.env.OPENAI_EQUIPMENT_MODEL || "gpt-4o-mini";

    const prompt = `
Identify the commercial kitchen equipment in this photo.

Return ONLY valid JSON:
{
  "visualName": "short equipment name",
  "visualDescription": "short visual description",
  "equipmentCategory": "pot | pan | lid | tray | container | utensil | machine_part | tool | rack | unknown",
  "likelySearchTerms": ["word1", "word2", "word3"],
  "confidence": number from 0 to 1,
  "notes": "short note if blurry or uncertain"
}

Rules:
- Do not invent an inventory code.
- Focus on what is visible.
- Use common kitchen equipment words.
- Keep the answer short.
`;

    const apiResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        max_output_tokens: 300,
        input: [
          {
            role: "user",
            content: [
              {
                type: "input_text",
                text: prompt,
              },
              {
                type: "input_image",
                image_url: imageDataUrl,
                detail: "low",
              },
            ],
          },
        ],
      }),
    });

    const data = await apiResponse.json();

    if (!apiResponse.ok) {
      return Response.json(
        {
          error: data?.error?.message || "Could not identify equipment.",
        },
        { status: 500 }
      );
    }

    const outputText = getOutputText(data);
    const parsed = parseJsonFromText(outputText);

    if (!parsed) {
      return Response.json(
        {
          error: "AI response could not be parsed.",
          raw: outputText,
        },
        { status: 500 }
      );
    }

    return Response.json({
      result: normalizeResult(parsed),
    });
  } catch (error) {
    return Response.json(
      {
        error: error?.message || "Could not identify equipment.",
      },
      { status: 500 }
    );
  }
}
