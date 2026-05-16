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
  matchedCandidateIndex:
    value.matchedCandidateIndex === null || value.matchedCandidateIndex === undefined
      ? null
      : Number(value.matchedCandidateIndex),
  matchedCode: String(value.matchedCode || value.matched_code || "").trim(),
  matchedName: String(value.matchedName || value.matched_name || "").trim(),
  matchStatus: String(value.matchStatus || value.match_status || "not_found").trim(),
  confidence: Number(value.confidence || 0),
  possibleCandidateIndexes: Array.isArray(value.possibleCandidateIndexes)
    ? value.possibleCandidateIndexes.map((item) => Number(item)).filter((item) => Number.isFinite(item))
    : [],
  notes: String(value.notes || "").trim(),
});

export async function POST(request) {
  try {
    const { imageDataUrl, candidates = [] } = await request.json();

    if (!imageDataUrl || !String(imageDataUrl).startsWith("data:image/")) {
      return Response.json({ error: "Missing image data." }, { status: 400 });
    }

    if (!process.env.OPENAI_API_KEY) {
      return Response.json(
        { error: "OPENAI_API_KEY is not configured in Vercel." },
        { status: 500 }
      );
    }

    const safeCandidates = Array.isArray(candidates)
      ? candidates.slice(0, 900).map((item, index) => ({
          index,
          code: String(item.code || "").slice(0, 80),
          name: String(item.name || "").slice(0, 160),
          category: String(item.category || "").slice(0, 120),
          sheetName: String(item.sheetName || "").slice(0, 120),
        }))
      : [];

    const candidateText = safeCandidates
      .map(
        (item) =>
          `${item.index}. Code: ${item.code || "N/A"} | Name: ${item.name || "N/A"} | Category: ${item.category || "N/A"} | Sheet: ${item.sheetName || "N/A"}`
      )
      .join("\n");

    const model =
      process.env.OPENAI_EQUIPMENT_MODEL ||
      process.env.OPENAI_TEMPERATURE_MODEL ||
      "gpt-4o-mini";

    const prompt = `
You are helping an equipment inventory app.

Analyze the equipment in the photo. Then compare it with this uploaded master inventory list.

Return ONLY valid JSON with this schema:
{
  "visualName": "short name of the equipment visible in the photo",
  "visualDescription": "short visual description",
  "matchedCandidateIndex": number or null,
  "matchedCode": "code from candidate list or empty string",
  "matchedName": "name from candidate list or empty string",
  "matchStatus": "exact | possible | not_found",
  "confidence": number from 0 to 1,
  "possibleCandidateIndexes": [number, number, number],
  "notes": "short explanation"
}

Rules:
- Use the candidate list as the source of truth for matching.
- If the item clearly exists in the candidate list, return matchStatus "exact".
- If it may exist but you are not sure, return matchStatus "possible" and include up to 5 possibleCandidateIndexes.
- If it does not appear in the list, return matchStatus "not_found".
- Do not invent codes.
- If the image is blurry, say so in notes.
- Equipment examples: saucepan, stock pot, saute pan, baking tray, knife, cutting board, whisk, tong, ladle, blender jar, mixer bowl, gastronorm pan, hotel pan, tray, rack, container.

Candidate list:
${candidateText || "No candidates were provided."}
`;

    const apiResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
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
                detail: "high",
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
