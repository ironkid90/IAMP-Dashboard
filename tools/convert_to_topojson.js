#!/usr/bin/env node
/**
 * Convert boundary GeoJSON files to TopoJSON for faster map loading.
 * Applies topology-aware simplification to reduce file size.
 *
 * Usage:  node tools/convert_to_topojson.js
 */

const fs = require("fs");
const path = require("path");
const topojson = require("topojson-server");
const simplify = require("topojson-simplify");

const DATA_DIR = path.join(__dirname, "..", "hand_sharepoint_dashboard", "assets", "data");

const FILES = [
  { input: "boundaries_admin1.geojson",        output: "boundaries_admin1.topojson",        name: "admin1", quantize: 1e5 },
  { input: "boundaries_admin2.geojson",        output: "boundaries_admin2.topojson",        name: "admin2", quantize: 1e5 },
  { input: "boundaries_admin3_subset.geojson", output: "boundaries_admin3_subset.topojson", name: "admin3", quantize: 1e5 },
  { input: "boundaries_admin3_full.geojson",   output: "boundaries_admin3_full.topojson",   name: "admin3_full", quantize: 1e5 },
];

for (const spec of FILES) {
  const inPath = path.join(DATA_DIR, spec.input);
  if (!fs.existsSync(inPath)) {
    console.warn(`⚠  Skipping ${spec.input} (not found)`);
    continue;
  }

  const geojson = JSON.parse(fs.readFileSync(inPath, "utf-8"));
  const inSize = fs.statSync(inPath).size;

  // Build topology
  const objects = {};
  objects[spec.name] = geojson;
  let topo = topojson.topology(objects, spec.quantize);

  // Simplify (retain ~15% of original detail — good for dashboard-level zoom)
  topo = simplify.presimplify(topo);
  topo = simplify.simplify(topo, 0.005);

  // Remove simplification weights to reduce output size
  topo = simplify.filter(topo, simplify.filterWeight(topo, 0.005));

  const outPath = path.join(DATA_DIR, spec.output);
  const json = JSON.stringify(topo);
  fs.writeFileSync(outPath, json, "utf-8");
  const outSize = Buffer.byteLength(json);

  const ratio = ((1 - outSize / inSize) * 100).toFixed(1);
  console.log(`✓  ${spec.input} → ${spec.output}  (${(inSize/1024).toFixed(0)} KB → ${(outSize/1024).toFixed(0)} KB, ${ratio}% smaller)`);
}

console.log("\nDone. TopoJSON files written to", DATA_DIR);
