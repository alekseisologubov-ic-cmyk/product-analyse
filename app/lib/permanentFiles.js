export const APP_FILE_BUCKET = "app-files";
export const INGREDIENT_BY_LOCATION_PATH = "ingredient-by-location/latest.xlsx";

export const uploadIngredientByLocationFileToStorage = async ({ supabase, file }) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  if (!file) {
    throw new Error("No file selected.");
  }

  const { error } = await supabase.storage
    .from(APP_FILE_BUCKET)
    .upload(INGREDIENT_BY_LOCATION_PATH, file, {
      contentType:
        file.type ||
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      upsert: true,
      cacheControl: "60",
    });

  if (error) {
    throw error;
  }

  return true;
};

export const downloadIngredientByLocationFileFromStorage = async ({ supabase }) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  const { data, error } = await supabase.storage
    .from(APP_FILE_BUCKET)
    .download(INGREDIENT_BY_LOCATION_PATH);

  if (error) {
    throw error;
  }

  if (!data) {
    throw new Error("Permanent Ingredient by Location file was not found.");
  }

  return data.arrayBuffer();
};
