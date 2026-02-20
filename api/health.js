export default async function handler(request) {
  return new Response(
    JSON.stringify({ ok: true, now: new Date().toISOString() }, null, 2),
    {
      status: 200,
      headers: {
        "content-type": "application/json; charset=utf-8",
        "cache-control": "no-store",
      },
    }
  );
}
