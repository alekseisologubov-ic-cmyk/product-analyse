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
  productName: String(value.productName || value.product_name || "").trim(),
  foodCategory: String(value.foodCategory || value.food_category || "").trim(),
  temperatureText: String(value.temperatureText || value.temperature_text || "").trim(),
  temperatureValue:
    value.temperatureValue === null || value.temperatureValue === undefined
      ? ""
      : value.temperatureValue,
  temperatureUnit: String(value.temperatureUnit || value.temperature_unit || "F").trim() || "F",
  readerVisible: Boolean(value.readerVisible ?? value.reader_visible ?? false),
  isChickenOrPoultry: Boolean(value.isChickenOrPoultry ?? value.is_chicken_or_poultry ?? false),
  confidence: Number(value.confidence || 0),
  notes: String(value.notes || "").trim(),
});

export async function POST(request) {
  try {
    const { imageDataUrl } = await request.json();

    if (!imageDataUrl || !String(imageDataUrl).startsWith("data:image/")) {
      return Response.json(
        { error: "Missing image data." },
        { status: 400 }
      );
    }

    if (!process.env.OPENAI_API_KEY) {
      return Response.json(
        { error: "OPENAI_API_KEY is not configured in Vercel." },
        { status: 500 }
      );
    }

    const model = process.env.OPENAI_TEMPERATURE_MODEL || "gpt-4o-mini";

    const prompt = `
You are helping a food safety temperature log app.

Analyze the photo. It may show:
1. food or product on a plate, tray, pan, or container
2. a thermometer or temperature reader display

Return ONLY valid JSON with this schema:
{
  "productName": "short product name visible in photo, for example Chicken Breast, Chicken Thigh, Fish, Sauce, Soup, Unknown",
  "foodCategory": "poultry | beef | pork | fish | shellfish | sauce | soup | vegetable | prepared_food | unknown",
  "temperatureText": "exact visible temperature text, for example 41.2 F, 6 C, or empty string if not readable",
  "temperatureValue": number or null,
  "temperatureUnit": "F | C | unknown",
  "readerVisible": true or false,
  "isChickenOrPoultry": true or false,
  "confidence": number from 0 to 1,
  "notes": "short note, include if the display is blurry or food is uncertain"
}

Rules:
- Do not guess a temperature if the thermometer display is not readable.
- If any cut of chicken is visible, including breast, thigh, leg, wing, tender, diced chicken, cooked chicken, or raw chicken, set isChickenOrPoultry true and foodCategory poultry.
- If unsure, use Unknown and confidence below 0.5.
- Keep productName short.
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
          error:
            data?.error?.message ||
            "Could not analyze temperature photo.",
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
        error: error?.message || "Could not analyze photo.",
      },
      { status: 500 }
    );
  }
}
