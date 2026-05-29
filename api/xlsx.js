export default async function handler(request) {
  if (request.method !== "GET") {
    return new Response("Method not allowed", {
      status: 405,
      headers: { "allow": "GET" },
    });
  }

  return new Response(
    JSON.stringify(
      {
        ok: false,
        error: "Raw workbook downloads are not exposed.",
        message: "Use /api/data for the sanitized site dataset.",
      },
      null,
      2
    ),
    {
      status: 403,
      headers: {
        "content-type": "application/json; charset=utf-8",
        "cache-control": "no-store",
      },
    }
  );
}
