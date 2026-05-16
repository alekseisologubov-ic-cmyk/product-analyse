export const runtime = "nodejs";

const getTextFromGeminiResponse = (data) =>
  (data?.candidates || [])
    .flatMap((candidate) => candidate?.content?.parts || [])
    .map((part) => part.text || "")
    .filter(Boolean)
    .join("\n");

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

const getImagePartsFromDataUrl = (imageDataUrl) => {
  const text = String(imageDataUrl || "");
  const match = text.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);

  if (!match) {
    return null;
  }

  return {
    mimeType: match[1],
    base64Data: match[2],
  };
};

const normalizeResult = (value = {}) => ({
  visualName: String(value.visualName || value.visual_name || "").trim(),
  visualDescription: String(value.visualDescription || value.visual_description || "").trim(),
  equipmentCategory: String(value.equipmentCategory || value.equipment_category || "").trim(),
  likelySearchTerms: Array.isArray(value.likelySearchTerms)
    ? value.likelySearchTerms
        .map((item) => String(item || "").trim())
        .filter(Boolean)
        .slice(0, 8)
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

    if (!process.env.GEMINI_API_KEY) {
      return Response.json(
        { error: "GEMINI_API_KEY is not configured in Vercel." },
        { status: 500 }
      );
    }

    const imageParts = getImagePartsFromDataUrl(imageDataUrl);

    if (!imageParts) {
      return Response.json(
        { error: "Invalid image data." },
        { status: 400 }
      );
    }

    const model = process.env.GEMINI_MODEL || "gemini-2.0-flash";

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
- If it looks like a plastic lid, say lid.
- If it looks like a hotel pan, gastronorm pan, tray, stock pot, sauce pan, saute pan, container, rack, utensil, or machine part, name it clearly.
`;

    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${process.env.GEMINI_API_KEY}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          contents: [
            {
              role: "user",
              parts: [
                { text: prompt },
                {
                  inline_data: {
                    mime_type: imageParts.mimeType,
                    data: imageParts.base64Data,
                  },
                },
              ],
            },
          ],
          generationConfig: {
            temperature: 0.1,
            maxOutputTokens: 300,
          },
        }),
      }
    );

    const data = await response.json();

    if (!response.ok) {
      return Response.json(
        {
          error:
            data?.error?.message ||
            "Could not identify equipment with Gemini.",
        },
        { status: response.status }
      );
    }

    const outputText = getTextFromGeminiResponse(data);
    const parsed = parseJsonFromText(outputText);

    if (!parsed) {
      return Response.json(
        {
          error: "Gemini response could not be parsed.",
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
