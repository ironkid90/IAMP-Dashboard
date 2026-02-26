#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Preprocess IAMP dashboard inputs into fast, map-ready JSON.

Inputs (recommended):
  - Assessment export CSV (site-level operational fields, metrics, statuses)
  - Master sites XLSX (coordinates)
  - Lebanon boundaries GeoJSONs (admin1/admin2/admin3 with PCODEs)

Outputs:
  - sites.json (merged + redacted, ready for the dashboard)
  - boundaries_admin1.geojson (trimmed/rounded)
  - boundaries_admin2.geojson (trimmed/rounded)
  - boundaries_admin3_subset.geojson (trimmed/rounded; only PCODEs present in sites.json)
  - boundaries_admin3_full.geojson (trimmed/rounded; optional)

Notes:
  - This script avoids browser-side Excel parsing by producing a canonical JSON.
  - Spatial join (points -> admin3 polygons) is used to attach admin PCODEs.
  - PII is intentionally NOT exported in sites.json by default.
"""

import argparse
import json
import math
import os
import re
from datetime import datetime

import numpy as np
import pandas as pd
from shapely.geometry import Point, shape
from shapely.strtree import STRtree


BLANK_TOKENS = {"", "-", "—", "–", "na", "n/a", "null", "none"}


def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and np.isnan(v):
        return True
    s = str(v).strip().lower()
    return s in BLANK_TOKENS


def clean_str(v):
    if is_blank(v):
        return None
    return str(v).strip()


def clean_num(v):
    if v is None:
        return None
    try:
        if isinstance(v, float) and np.isnan(v):
            return None
        fx = float(v)
        if not math.isfinite(fx):
            return None
        if abs(fx - round(fx)) < 1e-9:
            return int(round(fx))
        return fx
    except Exception:
        return None


def round_coords(obj, ndigits=5):
    if isinstance(obj, list):
        return [round_coords(x, ndigits) for x in obj]
    if isinstance(obj, (float, int)):
        return round(float(obj), ndigits)
    return obj


def trim_geojson(in_path, out_path, keep_props, filter_key=None, filter_values=None, ndigits=5):
    gj = json.loads(open(in_path, "r", encoding="utf-8").read())
    feats_out = []
    for feat in gj.get("features", []):
        props = feat.get("properties", {}) or {}
        if filter_key and filter_values is not None:
            if props.get(filter_key) not in filter_values:
                continue

        new_props = {k: props.get(k) for k in keep_props if k in props}

        geom = feat.get("geometry")
        if geom:
            geom = dict(geom)
            geom["coordinates"] = round_coords(geom.get("coordinates"), ndigits=ndigits)

        feats_out.append({"type": "Feature", "properties": new_props, "geometry": geom})

    out_gj = {"type": "FeatureCollection", "features": feats_out}
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(json.dumps(out_gj, ensure_ascii=False, separators=(",", ":")))
    return len(feats_out)


def build_spatial_index(admin3_geojson_path):
    gj = json.loads(open(admin3_geojson_path, "r", encoding="utf-8").read())
    geoms = []
    props_by_id = {}
    for feat in gj.get("features", []):
        geom = shape(feat["geometry"])
        geoms.append(geom)
        props_by_id[id(geom)] = feat.get("properties", {}) or {}
    tree = STRtree(geoms)

    def lookup(lon, lat):
        pt = Point(lon, lat)
        for poly in tree.query(pt):
            # covers() includes boundary points
            if poly.covers(pt):
                return props_by_id.get(id(poly))
        return None

    return lookup


def main():
    ap = argparse.ArgumentParser(description="Preprocess IAMP dashboard data to canonical JSON.")
    ap.add_argument("--assessment_csv", required=True, help="Assessment CSV export (sites mapping).")
    ap.add_argument("--master_xlsx", required=True, help="Master sites list XLSX with coordinates.")
    ap.add_argument("--boundaries_dir", required=True, help="Folder containing lbn_admin1.geojson/lbn_admin2.geojson/lbn_admin3.geojson")
    ap.add_argument("--out_dir", required=True, help="Output directory for JSON files.")
    ap.add_argument("--subset_admin3", action="store_true", help="Output admin3 subset only (recommended for performance).")
    ap.add_argument("--write_full_admin3", action="store_true", help="Also output boundaries_admin3_full.geojson (trimmed).")
    args = ap.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    # 1) Read + clean assessment CSV
    df = pd.read_csv(args.assessment_csv)
    valid_pat = re.compile(r"^\\d{5}-\\d{2}-\\d{3}$")
    df = df[df["PCode"].astype(str).str.strip().apply(lambda x: bool(valid_pat.match(x)))].copy()

    # Numeric fields (coerce)
    num_cols = [
        "A- Number of Tents",
        "B- Number of Self-built Structures with Non-Concrete Roof",
        "C- Number of Prefab Structure",
        "D- Number of Self-built Structures with Concrete Roof",
        "Total number of Structures",
        "Total number of Households",
        "Total number of Individuals",
        "Number of Latrines",
    ]
    for c in num_cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # Assessed proxy
    df["assessed"] = ~df["Phone call status"].apply(is_blank)

    # 2) Read master coordinates XLSX (prefer New Sites if present)
    def read_sheet_safe(name):
        try:
            return pd.read_excel(args.master_xlsx, sheet_name=name)
        except Exception:
            return pd.DataFrame()

    df_all = read_sheet_safe("IAMP133 All Sites")
    df_new = read_sheet_safe("IAMP133 New Sites")

    def prep_coords(dfx):
        if dfx.empty:
            return dfx
        out = dfx.copy()
        out["PCode"] = out["PCode"].astype(str).str.strip()
        out["Latitude"] = pd.to_numeric(out.get("Latitude"), errors="coerce")
        out["Longitude"] = pd.to_numeric(out.get("Longitude"), errors="coerce")
        return out[["PCode", "Latitude", "Longitude"]]

    coords = pd.concat([prep_coords(df_new), prep_coords(df_all)], ignore_index=True)
    coords = coords.drop_duplicates(subset=["PCode"], keep="first")

    df = df.merge(coords, on="PCode", how="left")

    # 3) Spatial join to admin3 polygons (attach adm1/adm2/adm3 PCODEs)
    admin1_path = os.path.join(args.boundaries_dir, "lbn_admin1.geojson")
    admin2_path = os.path.join(args.boundaries_dir, "lbn_admin2.geojson")
    admin3_path = os.path.join(args.boundaries_dir, "lbn_admin3.geojson")

    lookup = build_spatial_index(admin3_path)

    adm1_pcodes, adm2_pcodes, adm3_pcodes = [], [], []
    for lon, lat in zip(df["Longitude"], df["Latitude"]):
        if pd.isna(lon) or pd.isna(lat):
            adm1_pcodes.append(None); adm2_pcodes.append(None); adm3_pcodes.append(None)
            continue
        props = lookup(float(lon), float(lat))
        if props:
            adm1_pcodes.append(props.get("adm1_pcode"))
            adm2_pcodes.append(props.get("adm2_pcode"))
            adm3_pcodes.append(props.get("adm3_pcode"))
        else:
            adm1_pcodes.append(None); adm2_pcodes.append(None); adm3_pcodes.append(None)

    df["adm1_pcode"] = adm1_pcodes
    df["adm2_pcode"] = adm2_pcodes
    df["adm3_pcode"] = adm3_pcodes

    # Fill missing admin codes by name-based mapping (from already joined rows)
    def norm(s):
        if is_blank(s):
            return ""
        return re.sub(r"\\s+", " ", re.sub(r"[^a-z0-9]+", " ", str(s).strip().lower())).strip()

    df["adm_key"] = list(zip(df["District"].map(norm), df["Cadaster"].map(norm)))
    mapping = (
        df.dropna(subset=["adm3_pcode"])
          .groupby("adm_key")[["adm1_pcode","adm2_pcode","adm3_pcode"]]
          .agg(lambda x: x.value_counts().index[0])
    )
    miss = df["adm3_pcode"].isna()
    if miss.any():
        for col in ["adm1_pcode","adm2_pcode","adm3_pcode"]:
            df.loc[miss, col] = df.loc[miss, "adm_key"].map(mapping[col])

    # 4) Lightweight QC flags
    df["qc_missing_coords"] = df["Latitude"].isna() | df["Longitude"].isna()
    df["qc_missing_site_status"] = df["assessed"] & df["Site Status"].apply(is_blank)
    df["qc_missing_totals_active"] = (df["Site Status"].astype(str).str.strip() == "Active") & df[["Total number of Individuals","Total number of Households","Total number of Structures"]].isna().any(axis=1)
    df["qc_inactive_missing_date"] = df["Site Status"].astype(str).str.contains("Inactive|Demolished", case=False, na=False) & df["Date of when site is Inactive or full demolish sites"].apply(is_blank)

    qc_cols = ["qc_missing_coords","qc_missing_site_status","qc_missing_totals_active","qc_inactive_missing_date"]
    df["qc_issue_count"] = df[qc_cols].sum(axis=1)
    df["qc_any_issue"] = df["qc_issue_count"] > 0

    # 5) Build redacted canonical JSON
    sites = []
    for _, r in df.iterrows():
        sites.append({
            "pcode": clean_str(r.get("PCode")),
            "name": clean_str(r.get("PCode Name")) or clean_str(r.get("Pcode_name")),
            "local_name": clean_str(r.get("Local Name")),
            "governorate": clean_str(r.get("Governorate")),
            "district": clean_str(r.get("District")),
            "cadaster": clean_str(r.get("Cadaster")),
            "site_status": clean_str(r.get("Site Status")),
            "phone_status": clean_str(r.get("Phone call status")),
            "assessed_date": clean_str(r.get("Date of phone assessment")),
            "lat": clean_num(r.get("Latitude")),
            "lon": clean_num(r.get("Longitude")),
            "adm1_pcode": clean_str(r.get("adm1_pcode")),
            "adm2_pcode": clean_str(r.get("adm2_pcode")),
            "adm3_pcode": clean_str(r.get("adm3_pcode")),
            "metrics": {
                "structures_total": clean_num(r.get("Total number of Structures")),
                "households_total": clean_num(r.get("Total number of Households")),
                "individuals_total": clean_num(r.get("Total number of Individuals")),
                "latrines_total": clean_num(r.get("Number of Latrines")),
                "tents": clean_num(r.get("A- Number of Tents")),
                "selfbuilt_nonconcrete": clean_num(r.get("B- Number of Self-built Structures with Non-Concrete Roof")),
                "prefab": clean_num(r.get("C- Number of Prefab Structure")),
                "selfbuilt_concrete": clean_num(r.get("D- Number of Self-built Structures with Concrete Roof")),
            },
            "qc": {
                "qc_any_issue": bool(r.get("qc_any_issue", False)),
                "qc_issue_count": int(r.get("qc_issue_count", 0)),
                "missing_coords": bool(r.get("qc_missing_coords", False)),
                "missing_site_status_when_assessed": bool(r.get("qc_missing_site_status", False)),
                "missing_totals_active": bool(r.get("qc_missing_totals_active", False)),
                "inactive_missing_date": bool(r.get("qc_inactive_missing_date", False)),
            },
            "record_status": clean_str(r.get("Record status")),
            "partner": clean_str(r.get("Partner Name")),
        })

    meta = {
        "generated_at_utc": datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source_files": {
            "assessment_csv": os.path.basename(args.assessment_csv),
            "master_xlsx": os.path.basename(args.master_xlsx),
        },
        "counts": {
            "records": int(len(sites)),
            "with_coords": int(sum(1 for s in sites if s["lat"] is not None and s["lon"] is not None)),
            "qc_any_issue": int(sum(1 for s in sites if s["qc"]["qc_any_issue"])),
        },
    }

    sites_path = os.path.join(args.out_dir, "sites.json")
    with open(sites_path, "w", encoding="utf-8") as f:
        f.write(json.dumps({"meta": meta, "sites": sites}, ensure_ascii=False, separators=(",", ":")))

    # 6) Boundaries (trim + round)
    keep1 = ["adm1_pcode","adm1_name","adm0_pcode","adm0_name"]
    keep2 = ["adm2_pcode","adm2_name","adm1_pcode","adm1_name","adm0_pcode","adm0_name"]
    keep3 = ["adm3_pcode","adm3_name","adm2_pcode","adm2_name","adm1_pcode","adm1_name","adm0_pcode","adm0_name"]

    trim_geojson(admin1_path, os.path.join(args.out_dir, "boundaries_admin1.geojson"), keep1, ndigits=5)
    trim_geojson(admin2_path, os.path.join(args.out_dir, "boundaries_admin2.geojson"), keep2, ndigits=5)

    adm3_codes = set(df["adm3_pcode"].dropna().astype(str).unique())

    if args.subset_admin3:
        trim_geojson(
            admin3_path,
            os.path.join(args.out_dir, "boundaries_admin3_subset.geojson"),
            keep3,
            filter_key="adm3_pcode",
            filter_values=adm3_codes,
            ndigits=5
        )
    else:
        trim_geojson(admin3_path, os.path.join(args.out_dir, "boundaries_admin3_subset.geojson"), keep3, ndigits=5)

    if args.write_full_admin3:
        trim_geojson(admin3_path, os.path.join(args.out_dir, "boundaries_admin3_full.geojson"), keep3, ndigits=5)

    print("Done.")
    print("  sites:", sites_path)
    print("  boundaries:", args.out_dir)


if __name__ == "__main__":
    main()
