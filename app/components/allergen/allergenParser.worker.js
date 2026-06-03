import { parseIngredientByLocationArrayBuffer } from "../../lib/allergenParser";

self.onmessage = (event) => {
  const { arrayBuffer } = event.data || {};

  try {
    if (!arrayBuffer) {
      throw new Error("No workbook data received by allergen parser worker.");
    }

    const startedAt = performance.now();
    const parsed = parseIngredientByLocationArrayBuffer(arrayBuffer);
    const elapsedMs = Math.round(performance.now() - startedAt);

    self.postMessage({
      ok: true,
      parsed,
      elapsedMs,
    });
  } catch (error) {
    self.postMessage({
      ok: false,
      error:
        error?.message ||
        "Could not parse Ingredient by Location workbook.",
    });
  }
};
