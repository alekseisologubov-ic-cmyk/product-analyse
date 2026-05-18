export const runtime = "nodejs";

const DEFAULT_FOLDER_ID = "19RH3TcKSZbMpQCh1DkGWa5Zl8e0nILs_";

const getNumbersFromText = (value) => {
  const matches = String(value || "").match(/\d{4,}/g) || [];

  return [
    ...new Set(
      matches
        .map((item) => item.replace(/^0+/, "") || "0")
        .filter(Boolean)
    ),
  ];
};

const getDriveFileViewUrl = (fileId) =>
  `https://drive.google.com/file/d/${fileId}/view`;

const getDriveThumbnailUrl = (fileId, size = "w800") =>
  `https://drive.google.com/thumbnail?id=${fileId}&sz=${size}`;
const getDriveProxyImageUrl = (fileId) =>
  `/api/drive-image?fileId=${encodeURIComponent(fileId)}`;

export async function GET(request) {
  try {
    const apiKey = process.env.GOOGLE_DRIVE_API_KEY;

    if (!apiKey) {
      return Response.json(
        { error: "GOOGLE_DRIVE_API_KEY is not configured in Vercel." },
        { status: 500 }
      );
    }

    const requestUrl = new URL(request.url);

    const folderId =
      requestUrl.searchParams.get("folderId") ||
      process.env.GOOGLE_DRIVE_EQUIPMENT_FOLDER_ID ||
      DEFAULT_FOLDER_ID;

    const files = [];
    let pageToken = "";

    do {
      const params = new URLSearchParams({
        key: apiKey,
        q: `'${folderId}' in parents and trashed = false and mimeType contains 'image/'`,
        pageSize: "1000",
        fields: "nextPageToken,files(id,name,mimeType,modifiedTime,webViewLink)",
        supportsAllDrives: "true",
        includeItemsFromAllDrives: "true",
      });

      if (pageToken) {
        params.set("pageToken", pageToken);
      }

      const response = await fetch(
        `https://www.googleapis.com/drive/v3/files?${params.toString()}`
      );

      const data = await response.json();

      if (!response.ok) {
        return Response.json(
          {
            error:
              data?.error?.message ||
              "Could not load Google Drive picture library.",
          },
          { status: response.status }
        );
      }

      (data.files || []).forEach((file) => {
        files.push({
  id: file.id,
  name: file.name || "",
  mimeType: file.mimeType || "",
  modifiedTime: file.modifiedTime || "",
  numbers: getNumbersFromText(file.name),
  thumbnailUrl: getDriveProxyImageUrl(file.id),
  imageUrl: getDriveProxyImageUrl(file.id),
  driveThumbnailUrl: getDriveThumbnailUrl(file.id, "w800"),
  webViewLink: file.webViewLink || getDriveFileViewUrl(file.id),
});
      });

      pageToken = data.nextPageToken || "";
    } while (pageToken);

    return Response.json({
      folderId,
      count: files.length,
      files,
    });
  } catch (error) {
    return Response.json(
      {
        error: error?.message || "Could not load picture library.",
      },
      { status: 500 }
    );
  }
}
