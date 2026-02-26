// Vercel Serverless Function (Node.js)
//
// Current behavior: serve the preprocessed JSON bundled with the deployment.
// Next step: swap this logic to fetch the latest XLSX/ArcGIS JSON, merge, and return.

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

module.exports = (req, res) => {
  try {
    const filePath = path.join(process.cwd(), "hand_sharepoint_dashboard", "assets", "data", "sites.json");
    const body = fs.readFileSync(filePath, "utf8");

    const etag = crypto.createHash("sha1").update(body).digest("hex");
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    res.setHeader("Cache-Control", "public, max-age=300, stale-while-revalidate=86400");
    res.setHeader("ETag", etag);

    if (req.headers["if-none-match"] === etag) {
      res.statusCode = 304;
      return res.end();
    }

    res.statusCode = 200;
    return res.end(body);
  } catch (err) {
    res.statusCode = 500;
    res.setHeader("Content-Type", "application/json; charset=utf-8");
    return res.end(JSON.stringify({ error: "Failed to load bundled sites.json", message: String(err.message || err) }));
  }
};
