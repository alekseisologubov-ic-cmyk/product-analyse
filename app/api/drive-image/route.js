export const runtime = "nodejs";
export const dynamic = "force-dynamic";

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
    const fileId = requestUrl.searchParams.get("fileId");

    if (!fileId) {
      return Response.json(
        { error: "Missing Google Drive fileId." },
        { status: 400 }
      );
    }

    const driveResponse = await fetch(
      `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media&key=${apiKey}`
    );

    if (!driveResponse.ok) {
      const text = await driveResponse.text();

      return Response.json(
        {
          error: "Could not load Google Drive image.",
          details: text.slice(0, 300),
        },
        { status: driveResponse.status }
      );
    }

    const contentType =
      driveResponse.headers.get("content-type") || "image/jpeg";

    const imageBuffer = await driveResponse.arrayBuffer();

    return new Response(imageBuffer, {
      headers: {
        "Content-Type": contentType,
        "Cache-Control": "public, max-age=86400",
      },
    });
  } catch (error) {
    return Response.json(
      {
        error: error?.message || "Could not load Google Drive image.",
      },
      { status: 500 }
    );
  }
}
