export const APP_FILE_BUCKET = "app-files";

export const INGREDIENT_BY_LOCATION_FOLDER = "ingredient-by-location";
export const INGREDIENT_BY_LOCATION_PATH = `${INGREDIENT_BY_LOCATION_FOLDER}/latest.xlsx`;
export const INGREDIENT_BY_LOCATION_MANIFEST_PATH = `${INGREDIENT_BY_LOCATION_FOLDER}/manifest.json`;

export const uploadIngredientByLocationFileToStorage = async ({
  supabase,
  file,
}) => {
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

export const downloadIngredientByLocationFileFromStorage = async ({
  supabase,
}) => {
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

const downloadStorageJson = async ({ supabase, path }) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  if (!path) {
    throw new Error("Storage JSON path is missing.");
  }

  const { data, error } = await supabase.storage
    .from(APP_FILE_BUCKET)
    .download(path);

  if (error) {
    throw error;
  }

  if (!data) {
    throw new Error(`JSON file was not found: ${path}`);
  }

  const text = await data.text();

  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Could not read JSON file: ${path}`);
  }
};

const uploadStorageJson = async ({
  supabase,
  path,
  value,
  cacheControl = "60",
  upsert = true,
}) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  if (!path) {
    throw new Error("Storage JSON path is missing.");
  }

  const blob = new Blob([JSON.stringify(value)], {
    type: "application/json",
  });

  const { error } = await supabase.storage
    .from(APP_FILE_BUCKET)
    .upload(path, blob, {
      cacheControl,
      contentType: "application/json",
      upsert,
    });

  if (error) {
    throw error;
  }

  return true;
};

export const uploadIngredientByLocationParsedDataToStorage = async ({
  supabase,
  parsed,
  fileName,
}) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  if (!parsed?.rows?.length) {
    throw new Error("No parsed allergen rows to save.");
  }

  if (!parsed?.venues?.length) {
    throw new Error("No parsed allergen venues to save.");
  }

  const version = new Date().toISOString().replace(/[:.]/g, "-");
  const parsedPath = `${INGREDIENT_BY_LOCATION_FOLDER}/parsed-${version}.json`;

  await uploadStorageJson({
    supabase,
    path: parsedPath,
    value: parsed,
    cacheControl: "86400",
    upsert: false,
  });

  const manifest = {
    version,
    parsedPath,
    fileName: fileName || "Permanent Ingredient by Location",
    rowCount: parsed.rows.length,
    venueCount: parsed.venues.length,
    sourceSheet: parsed.sourceSheet || "",
    parserVersion: parsed.parserVersion || 1,
    updatedAt: new Date().toISOString(),
  };

  await uploadStorageJson({
    supabase,
    path: INGREDIENT_BY_LOCATION_MANIFEST_PATH,
    value: manifest,
    cacheControl: "30",
    upsert: true,
  });

  return manifest;
};

export const downloadIngredientByLocationParsedDataFromStorage = async ({
  supabase,
}) => {
  if (!supabase) {
    throw new Error("Supabase is not connected.");
  }

  const manifest = await downloadStorageJson({
    supabase,
    path: INGREDIENT_BY_LOCATION_MANIFEST_PATH,
  });

  if (!manifest?.parsedPath) {
    throw new Error("Parsed allergen manifest is missing parsedPath.");
  }

  const parsed = await downloadStorageJson({
    supabase,
    path: manifest.parsedPath,
  });

  if (!parsed?.rows?.length || !parsed?.venues?.length) {
    throw new Error("Parsed allergen cache is empty or invalid.");
  }

  return {
    manifest,
    parsed,
  };
};
