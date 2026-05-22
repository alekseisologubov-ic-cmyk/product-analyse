import { google } from "googleapis";

export const runtime = "nodejs";

const PICTURE_FOLDER_ID = "1RqnkXgtOXREZTH4p2N3FJrbdEOe44Nka";
const SPREADSHEET_ID = "1KRPzBqi3Tu_w1WDObdzHyagW15E3xydt";
const TARGET_SHEET_ID = 1792313930;

const normalizeEmail = (value) => String(value || "").trim().toLowerCase();

const getAdminEmails = () =>
  String(process.env.ADMIN_EMAILS || "")
    .split(",")
    .map((email) => normalizeEmail(email))
    .filter(Boolean);

const isAllowedAdminEmail = (email) => {
  const normalizedEmail = normalizeEmail(email);
  return Boolean(normalizedEmail && getAdminEmails().includes(normalizedEmail));
};

const getGoogleCredentials = () => {
  const clientEmail =
    process.env.GOOGLE_CLIENT_EMAIL ||
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL ||
    "";

  const privateKey = String(
    process.env.GOOGLE_PRIVATE_KEY ||
      process.env.GOOGLE_SERVICE_ACCOUNT_PRIVATE_KEY ||
      ""
  ).replace(/\\n/g, "\n");

  if (!clientEmail || !privateKey) {
    throw new Error(
      "Google service account credentials are missing. Add GOOGLE_CLIENT_EMAIL and GOOGLE_PRIVATE_KEY in Vercel."
    );
  }

  return { clientEmail, privateKey };
};

const getGoogleAuth = () => {
  const { clientEmail, privateKey } = getGoogleCredentials();

  return new google.auth.JWT({
    email: clientEmail,
    key: privateKey,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  });
};

const normalizeEquipmentPictureCode = (value) => {
  const text = String(value || "")
    .trim()
    .replace(/\.0+$/g, "");

  const match = text.match(/\d{4,}/);
  return match ? match[0].replace(/^0+/, "") : "";
};

const isImageFile = (file) => {
  const name = String(file?.name || "").toLowerCase();
  const mimeType = String(file?.mimeType || "").toLowerCase();

  return (
    mimeType.startsWith("image/") ||
    /\.(jpg|jpeg|png|webp|gif)$/i.test(name)
  );
};

const getDrivePictureUrl = (file) =>
  file?.webViewLink ||
  (file?.id ? `https://drive.google.com/file/d/${file.id}/view?usp=drive_link` : "");

const quoteSheetTitle = (title) => `'${String(title || "").replace(/'/g, "''")}'`;

const listPictureFilesByCode = async (drive) => {
  const pictureByCode = new Map();
  const duplicateCodes = new Set();
  const files = [];
  let pageToken = undefined;

  do {
    const response = await drive.files.list({
      q: `'${PICTURE_FOLDER_ID}' in parents and trashed=false`,
      fields:
        "nextPageToken, files(id,name,mimeType,webViewLink,thumbnailLink)",
      pageSize: 100,
      pageToken,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });

    const pageFiles = response.data.files || [];
    files.push(...pageFiles);
    pageToken = response.data.nextPageToken || undefined;
  } while (pageToken);

  files.filter(isImageFile).forEach((file) => {
    const codeKey = normalizeEquipmentPictureCode(file.name);
    const pictureUrl = getDrivePictureUrl(file);

    if (!codeKey || !pictureUrl) return;

    if (pictureByCode.has(codeKey)) {
      duplicateCodes.add(codeKey);
      return;
    }

    pictureByCode.set(codeKey, {
      codeKey,
      fileName: file.name,
      pictureUrl,
      fileId: file.id,
    });
  });

  return {
    pictureByCode,
    totalDriveFiles: files.length,
    imageFileCount: files.filter(isImageFile).length,
    duplicateCodes: [...duplicateCodes],
  };
};

const getTargetSheetTitle = async (sheets) => {
  const response = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: "sheets(properties(sheetId,title))",
  });

  const targetSheet = (response.data.sheets || []).find(
    (sheet) => Number(sheet.properties?.sheetId) === Number(TARGET_SHEET_ID)
  );

  if (!targetSheet?.properties?.title) {
    throw new Error(
      `Target sheet gid ${TARGET_SHEET_ID} was not found in spreadsheet ${SPREADSHEET_ID}.`
    );
  }

  return targetSheet.properties.title;
};

const chunkArray = (items, size) => {
  const chunks = [];

  for (let index = 0; index < items.length; index += size) {
    chunks.push(items.slice(index, index + size));
  }

  return chunks;
};

export async function POST(request) {
  try {
    const adminUploadCode = process.env.ADMIN_UPLOAD_CODE || "";

    if (!adminUploadCode) {
      return Response.json(
        { error: "ADMIN_UPLOAD_CODE is not configured on the server." },
        { status: 500 }
      );
    }

    const body = await request.json().catch(() => ({}));
    const userEmail = normalizeEmail(body.userEmail);
    const adminCode = String(body.adminCode || "").trim();
    const dryRun = Boolean(body.dryRun);

    if (!isAllowedAdminEmail(userEmail)) {
      return Response.json(
        { error: "This email is not allowed to update picture links." },
        { status: 403 }
      );
    }

    if (adminCode !== adminUploadCode) {
      return Response.json(
        { error: "Admin upload code is incorrect." },
        { status: 403 }
      );
    }

    const auth = getGoogleAuth();
    const drive = google.drive({ version: "v3", auth });
    const sheets = google.sheets({ version: "v4", auth });

    const { pictureByCode, totalDriveFiles, imageFileCount, duplicateCodes } =
      await listPictureFilesByCode(drive);

    if (!pictureByCode.size) {
      return Response.json(
        {
          error:
            "No image files with equipment codes were found in the Drive picture folder.",
          totalDriveFiles,
          imageFileCount,
        },
        { status: 404 }
      );
    }

    const sheetTitle = await getTargetSheetTitle(sheets);
    const quotedSheetTitle = quoteSheetTitle(sheetTitle);

    const valuesResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${quotedSheetTitle}!A:I`,
      valueRenderOption: "UNFORMATTED_VALUE",
    });

    const rows = valuesResponse.data.values || [];
    const updates = [];
    const matchedCodes = [];
    const unmatchedCodes = [];

    rows.forEach((row, rowIndex) => {
      const rowNumber = rowIndex + 1;
      const productCode = row?.[3]; // Column D
      const codeKey = normalizeEquipmentPictureCode(productCode);

      if (!codeKey) return;

      const pictureMatch = pictureByCode.get(codeKey);

      if (!pictureMatch) {
        unmatchedCodes.push(codeKey);
        return;
      }

      updates.push({
        range: `${quotedSheetTitle}!H${rowNumber}`,
        values: [[pictureMatch.pictureUrl]],
      });

      matchedCodes.push(codeKey);
    });

    if (!updates.length) {
      return Response.json({
        ok: true,
        dryRun,
        message:
          "No matching product codes were found between column D and the Drive picture folder.",
        sheetTitle,
        totalDriveFiles,
        imageFileCount,
        matchedRows: 0,
        unmatchedRows: unmatchedCodes.length,
        duplicatePictureCodes: duplicateCodes.length,
        sampleUnmatchedCodes: [...new Set(unmatchedCodes)].slice(0, 30),
      });
    }

    if (!dryRun) {
      const updateChunks = chunkArray(updates, 400);

      for (const chunk of updateChunks) {
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            valueInputOption: "USER_ENTERED",
            data: chunk,
          },
        });
      }
    }

    return Response.json({
      ok: true,
      dryRun,
      sheetTitle,
      spreadsheetId: SPREADSHEET_ID,
      targetSheetId: TARGET_SHEET_ID,
      driveFolderId: PICTURE_FOLDER_ID,
      totalDriveFiles,
      imageFileCount,
      pictureCodesFound: pictureByCode.size,
      matchedRows: updates.length,
      unmatchedRows: unmatchedCodes.length,
      duplicatePictureCodes: duplicateCodes.length,
      sampleMatchedCodes: [...new Set(matchedCodes)].slice(0, 30),
      sampleUnmatchedCodes: [...new Set(unmatchedCodes)].slice(0, 30),
    });
  } catch (error) {
    return Response.json(
      {
        error: error?.message || "Could not write picture links to Google Sheet.",
      },
      { status: 500 }
    );
  }
}
