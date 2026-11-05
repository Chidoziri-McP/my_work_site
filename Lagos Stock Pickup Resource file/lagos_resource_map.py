#!/usr/bin/env python3
"""
Lagos Resource Mapping -> interactive Folium map + LGA summary tables.

Usage:
    python lagos_resource_map.py [Book1.xlsx] [lagos_lga_boundaries.geojson] [out_dir]

Requires: pandas, openpyxl, geopandas, folium
"""
import difflib
import html
import json
import re
import sys
from collections import OrderedDict
from pathlib import Path

import folium
import geopandas as gpd
import numpy as np
import pandas as pd
from branca.element import Element
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

XLSX = Path(sys.argv[1] if len(sys.argv) > 1 else "Book1.xlsx")
GEOJSON = Path(sys.argv[2] if len(sys.argv) > 2 else "lagos_lga_boundaries.geojson")
OUT = Path(sys.argv[3] if len(sys.argv) > 3 else ".")
OUT.mkdir(parents=True, exist_ok=True)

HEADER_ROWS_TO_SKIP = 2          # title row(s) above the real header; blank row 1 is dropped by openpyxl range
SPLIT_PEOPLE = re.compile(r"\s*[,;\n]\s*")     # NOT "&" or "/" -> "Grace & Glory" stays whole
SPLIT_LOCATIONS = re.compile(r"\s*[,;\n]\s*")

# Known spelling variants -> official boundary-file name
LGA_ALIASES = {
    "ajeromi ifelodun": "Ajeromi/ Ifelodun",
    "ajeromiifelodun": "Ajeromi/ Ifelodun",
    "ifako ijaiye": "Ifako/Ijaye",
    "ifakoijaiye": "Ifako/Ijaye",
    "ibeju lekki": "Ibeju Lekki",
    "ibejulekki": "Ibeju Lekki",
    "etiosa": "Eti-Osa",
    "amuwoodofin": "Amuwo Odofin",
    "somolu": "Shomolu",
    "oshodiisolo": "Oshodi/Isolo",
}


# --------------------------------------------------------------------------- helpers
def squash(s):
    """Trim + collapse internal whitespace; return None for blanks."""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return None
    s = re.sub(r"\s+", " ", str(s).replace("\u00a0", " ")).strip()
    return s or None


def key(s):
    """Loose comparison key: lowercase letters/digits only."""
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def split_names(cell, splitter=SPLIT_PEOPLE):
    out = []
    s = squash(cell)
    if not s:
        return out
    for part in splitter.split(s):
        part = squash(part)
        if part:
            out.append(part)
    return out


def uniq(seq):
    """Order-preserving, case/punctuation-insensitive de-duplication."""
    seen, out = set(), []
    for x in seq:
        k = key(x)
        if k and k not in seen:
            seen.add(k)
            out.append(x)
    return out


# --------------------------------------------------------------------------- 1. read + clean
wb = load_workbook(XLSX, data_only=True)
ws = wb.worksheets[0]
sheet_name = ws.title

# Merged cells (RSM / Specialist span several rows): fill every cell of a merged block
# with the anchor value so each row knows who covers it.
merged_fill = {}
for rng in ws.merged_cells.ranges:
    anchor = ws.cell(rng.min_row, rng.min_col).value
    for r in range(rng.min_row, rng.max_row + 1):
        for c in range(rng.min_col, rng.max_col + 1):
            merged_fill[(r, c)] = anchor

rows = []
for r in range(1, ws.max_row + 1):
    vals = []
    for c in range(1, ws.max_column + 1):
        v = merged_fill.get((r, c), ws.cell(r, c).value)
        vals.append(v)
    rows.append(vals)

# find header row (contains "RSM" & "Local Govt") rather than trusting a fixed index
hdr_idx = next(
    i for i, row in enumerate(rows)
    if any(squash(v) == "RSM" for v in row) and any("local govt" in (squash(v) or "").lower() for v in row)
)
header = [squash(v) for v in rows[hdr_idx]]
df = pd.DataFrame(rows[hdr_idx + 1:], columns=header)
df = df.loc[:, [c for c in df.columns if c]]          # drop unnamed padding columns
df = df.rename(columns={"CDOs": "CDO", "TKAs": "TKA", "Local Govt. Area": "LGA_raw"})
df = df.dropna(how="all")

for col in df.columns:
    if col != "S/N":
        df[col] = df[col].map(squash)
df = df[df["LGA_raw"].notna() | df["Location"].notna()].copy()

log = []   # cleaning log

# --------------------------------------------------------------------------- 2. match LGAs to boundaries
gdf = gpd.read_file(GEOJSON).to_crs(4326)
gdf["LGA"] = gdf["LGA"].map(squash)
official = {key(n): n for n in gdf["LGA"]}


def match_lga(raw):
    """Return (official_name | None, method)."""
    if not raw:
        return None, "blank"
    k = key(raw)
    if k in official:
        return official[k], "exact" if raw == official[k] else "normalised"
    alias = LGA_ALIASES.get(raw.lower()) or LGA_ALIASES.get(k)
    if alias and key(alias) in official:
        return official[key(alias)], "alias"
    close = difflib.get_close_matches(k, official.keys(), n=1, cutoff=0.86)
    if close:
        return official[close[0]], "fuzzy"
    return None, "unmatched"


m = df["LGA_raw"].map(match_lga)
df["LGA"] = m.map(lambda t: t[0])
df["match_method"] = m.map(lambda t: t[1])
for raw, meth in df[["LGA_raw", "match_method"]].drop_duplicates().itertuples(index=False):
    if meth not in ("exact",):
        log.append(f"LGA '{raw}' -> {meth}")

# --------------------------------------------------------------------------- 3. explode multi-name cells
records = []   # one record per (row, location)
for _, row in df.iterrows():
    locs = split_names(row["Location"], SPLIT_LOCATIONS) or ["(location not given)"]
    if len(locs) > 1:
        log.append(f"S/N {row['S/N']}: Location '{row['Location']}' split into {len(locs)} locations: "
                   f"{', '.join(locs)} (TKA/DP/CDO of that row applied to each)")
    for cell in ("TKA", "DP", "CDO"):
        n = split_names(row[cell])
        if len(n) > 1:
            log.append(f"S/N {row['S/N']}: {cell} '{row[cell]}' split into {len(n)} entries")
    for loc in locs:
        records.append(dict(
            sn=row["S/N"], lga_raw=row["LGA_raw"], lga=row["LGA"], location=loc,
            tka=split_names(row["TKA"]), dp=split_names(row["DP"]), cdo=split_names(row["CDO"]),
            specialist=split_names(row["Specialist"]), rsm=split_names(row["RSM"]),
        ))
rec = pd.DataFrame(records)

matched = rec[rec["lga"].notna()]
unmatched = rec[rec["lga"].isna()]

# --------------------------------------------------------------------------- 4. aggregate per LGA -> per location
lga_data = OrderedDict()
same_name = []
for lga, sub in matched.groupby("lga", sort=False):
    locs = OrderedDict()
    for _, r in sub.iterrows():
        k = key(r["location"])
        d = locs.setdefault(k, dict(name=r["location"], tka=[], dp=[], cdo=[], sn=[]))
        d["tka"] += r["tka"]; d["dp"] += r["dp"]; d["cdo"] += r["cdo"]; d["sn"].append(r["sn"])
    for d in locs.values():
        for f in ("tka", "dp", "cdo"):
            d[f] = uniq(d[f])
        if key(d["name"]) == key(lga):
            same_name.append(lga)
    lga_data[lga] = dict(
        locations=list(locs.values()),
        specialist=uniq(sum(sub["specialist"].tolist(), [])),
        rsm=uniq(sum(sub["rsm"].tolist(), [])),
    )

if same_name:
    log.append("Location equals the LGA name (kept as given in the source): " + ", ".join(same_name))
log.append("Marker for 'Lagos Island' is pinned to the island (boundary polygon also covers lagoon water)")
# duplicate-row check (identical rows apart from S/N)
dups = df.duplicated(subset=["RSM", "Specialist", "CDO", "DP", "TKA", "Location", "LGA_raw"], keep="first").sum()
log.append(f"Exact duplicate rows removed: {int(dups)}")

# --------------------------------------------------------------------------- 5. summary table (every one of the 20 LGAs)
def joined(items, sep="; "):
    return sep.join(items)


summary_rows, detail_rows = [], []
for lga in sorted(gdf["LGA"]):
    d = lga_data.get(lga)
    if d:
        locs = d["locations"]
        tka = uniq(sum([l["tka"] for l in locs], []))
        dp = uniq(sum([l["dp"] for l in locs], []))
        cdo = uniq(sum([l["cdo"] for l in locs], []))
        spec, rsm = d["specialist"], d["rsm"]
        for l in locs:
            detail_rows.append(OrderedDict([
                ("LGA", lga), ("Location", l["name"]), ("TKAs", joined(l["tka"])),
                ("DPs", joined(l["dp"])), ("CDOs", joined(l["cdo"])),
                ("Specialists (LGA level)", joined(spec)), ("RSMs (LGA level)", joined(rsm)),
            ]))
    else:
        locs, tka, dp, cdo, spec, rsm = [], [], [], [], [], []
    total = len(uniq(tka + dp + cdo + spec + rsm))
    summary_rows.append(OrderedDict([
        ("LGA", lga),
        ("No. of Locations", len(locs)),
        ("Locations", joined([l["name"] for l in locs])),
        ("TKAs", joined(tka)), ("DPs", joined(dp)), ("CDOs", joined(cdo)),
        ("Specialists", joined(spec)), ("RSMs", joined(rsm)),
        ("No. of TKAs", len(tka)), ("No. of DPs", len(dp)), ("No. of CDOs", len(cdo)),
        ("No. of Specialists", len(spec)), ("No. of RSMs", len(rsm)),
        ("Total Unique Resources", total),
        ("Status", "Mapped" if locs else "No deployment in source data"),
    ]))
summary = pd.DataFrame(summary_rows)
detail = pd.DataFrame(detail_rows)

unm_rows = [OrderedDict([("S/N", r["sn"]), ("LGA in source file", r["lga_raw"] or "(blank)"),
                         ("Location", r["location"]),
                         ("TKAs", joined(r["tka"])), ("DPs", joined(r["dp"])), ("CDOs", joined(r["cdo"])),
                         ("Specialists", joined(r["specialist"])), ("RSMs", joined(r["rsm"])),
                         ("Reason", "No matching Lagos State LGA boundary")])
            for _, r in unmatched.iterrows()]
unmatched_df = pd.DataFrame(unm_rows, columns=["S/N", "LGA in source file", "Location", "TKAs", "DPs", "CDOs",
                                               "Specialists", "RSMs", "Reason"])

summary.to_csv(OUT / "lagos_lga_summary.csv", index=False, encoding="utf-8-sig")

# --------------------------------------------------------------------------- 6. colour scales
PALETTE = ["#deebf7", "#9ecae1", "#4292c6", "#2171b5", "#08306b"]   # light -> dark
NO_DATA = "#eeeeee"


def make_bins(values, n=5):
    vals = [v for v in values if v > 0]
    if not vals:
        return [1]
    lo, hi = min(vals), max(vals)
    if hi == lo:
        return [lo]
    edges = sorted(set(int(np.ceil(x)) for x in np.linspace(lo, hi, n + 1)[:-1]))
    return edges      # lower bounds of each class


def classify(v, edges):
    if v <= 0:
        return -1
    idx = 0
    for i, e in enumerate(edges):
        if v >= e:
            idx = i
    return idx


def class_labels(edges, maxv):
    labs = []
    for i, e in enumerate(edges):
        hi = (edges[i + 1] - 1) if i + 1 < len(edges) else maxv
        labs.append(f"{e}" if hi == e else f"{e}–{hi}")
    return labs


METRICS = OrderedDict([
    ("Number of deployment locations", "No. of Locations"),
    ("Total mapped resources (all roles)", "Total Unique Resources"),
])
scales = {}
for label, col in METRICS.items():
    vals = summary[col].tolist()
    edges = make_bins(vals, n=min(5, max(1, len(set(v for v in vals if v > 0)))))
    scales[label] = dict(edges=edges, labels=class_labels(edges, max(vals)))

# --------------------------------------------------------------------------- 7. popup / tooltip HTML
LOC_COLOR = "#d9480f"    # assigned DIRECTLY to a location
LGA_COLOR = "#5f3dc4"    # covering the WIDER LGA
e = html.escape


def people_list(items):
    if not items:
        return '<span class="none">— none listed —</span>'
    return "".join(f'<div class="p">• {e(i)}</div>' for i in items)


def popup_html(lga):
    d = lga_data.get(lga)
    css = f"""
    <style>
      .pp {{font-family:Arial,Helvetica,sans-serif;font-size:12.5px;color:#222;width:340px;
            max-height:300px;overflow-y:auto;padding-right:4px}}
      .pp h3 {{margin:0 0 2px 0;font-size:16px;color:#08306b}}
      .pp .sub {{color:#666;margin-bottom:8px}}
      .pp .sec {{border-left:5px solid;border-radius:3px;background:#fafafa;padding:6px 8px;margin:8px 0}}
      .pp .tag {{display:inline-block;font-size:10px;font-weight:bold;letter-spacing:.4px;color:#fff;
                 padding:1px 6px;border-radius:9px;margin-bottom:4px}}
      .pp .loc {{font-size:13.5px;font-weight:bold;margin:2px 0 4px 0}}
      .pp .role {{font-weight:bold;margin-top:4px;color:#444}}
      .pp .p {{margin-left:6px}}
      .pp .none {{color:#999;font-style:italic;margin-left:6px}}
    </style>"""
    if not d:
        return css + (f'<div class="pp"><h3>{e(lga)}</h3><div class="sub">Lagos State LGA</div>'
                      '<div class="none">No resources or deployment locations in the source file.</div></div>')
    locs = d["locations"]
    out = [css, f'<div class="pp"><h3>{e(lga)}</h3>',
           f'<div class="sub">{len(locs)} deployment location{"s" if len(locs) != 1 else ""}</div>',
           f'<span class="tag" style="background:{LOC_COLOR}">ASSIGNED TO LOCATION</span>']
    for l in locs:
        out.append(f'<div class="sec" style="border-color:{LOC_COLOR}"><div class="loc">📍 Location: {e(l["name"])}</div>')
        out.append('<div class="role">TKA(s)</div>' + people_list(l["tka"]))
        out.append('<div class="role">DP(s)</div>' + people_list(l["dp"]))
        out.append('<div class="role">CDO(s)</div>' + people_list(l["cdo"]))
        out.append("</div>")
    out.append(f'<span class="tag" style="background:{LGA_COLOR}">COVERING THE WHOLE LGA</span>')
    out.append(f'<div class="sec" style="border-color:{LGA_COLOR}">'
               '<div class="role" style="margin-top:0">Specialist(s)</div>' + people_list(d["specialist"]) +
               '<div class="role">RSM(s)</div>' + people_list(d["rsm"]) + "</div></div>")
    return "".join(out)


def tooltip_html(row):
    if row["No. of Locations"] == 0:
        return f"<b>{e(row['LGA'])}</b><br>No deployment in source data"
    return (f"<b>{e(row['LGA'])}</b><br>{row['No. of Locations']} location(s) · "
            f"{row['Total Unique Resources']} unique resources<br><i>Click for details</i>")


# --------------------------------------------------------------------------- 8. build map
gdf = gdf.merge(summary, on="LGA", how="left")
proj = gdf.to_crs(32631)
pts = proj.geometry.representative_point().to_crs(4326)
gdf["cx"], gdf["cy"] = pts.x, pts.y
# The boundary file's "Lagos Island" polygon also covers a large part of the lagoon, so its automatic
# interior point falls in open water. Pin that LGA's marker/label on the island itself (Marina/CMS area).
MARKER_OVERRIDES = {"Lagos Island": (6.4531, 3.3947)}      # (lat, lon)
for _lga, (_lat, _lon) in MARKER_OVERRIDES.items():
    gdf.loc[gdf["LGA"] == _lga, ["cy", "cx"]] = [_lat, _lon]

minx, miny, maxx, maxy = gdf.total_bounds
m = folium.Map(location=[(miny + maxy) / 2, (minx + maxx) / 2], zoom_start=10, tiles=None,
               control_scale=True, prefer_canvas=True)
folium.TileLayer("OpenStreetMap", name="OpenStreetMap", control=True, opacity=0.55).add_to(m)
m.fit_bounds([[miny, minx], [maxy, maxx]])
m.options["minZoom"] = 8

metric_group_names = {}
for label, col in METRICS.items():
    sc = scales[label]
    fg = folium.FeatureGroup(name=f"Colour by: {label}", overlay=False, show=(label == list(METRICS)[0]))
    metric_group_names[label] = fg.get_name()
    for _, row in gdf.iterrows():
        v = row[col]
        ci = classify(v, sc["edges"])
        fill = NO_DATA if ci < 0 else PALETTE[min(ci, len(PALETTE) - 1)] if len(sc["edges"]) == len(PALETTE) else \
            PALETTE[round(ci * (len(PALETTE) - 1) / max(1, len(sc["edges"]) - 1))]
        feat = json.loads(gpd.GeoSeries([row.geometry], crs=4326).to_json())["features"][0]
        gj = folium.GeoJson(
            feat,
            style_function=lambda _f, fill=fill: dict(fillColor=fill, color="#37474f", weight=1.2, fillOpacity=0.78),
            highlight_function=lambda _f: dict(weight=3, color="#000", fillOpacity=0.9),
            tooltip=folium.Tooltip(tooltip_html(row), sticky=True),
            popup=folium.Popup(popup_html(row["LGA"]), max_width=380, lazy=False, autoPanPaddingTopLeft=[20, 90], autoPanPaddingBottomRight=[20, 20]),
        )
        gj.add_to(fg)
    fg.add_to(m)

# numbered markers (number of deployment locations) — a second click target
mk = folium.FeatureGroup(name="Markers (no. of locations)", overlay=True, show=True)
for _, row in gdf.iterrows():
    n = int(row["No. of Locations"])
    if n == 0:
        bg, fg_c, txt, size = "#ffffff", "#777", "–", 22
    else:
        bg, fg_c, txt, size = LOC_COLOR, "#fff", str(n), 26 + min(n, 6) * 2
    icon = folium.DivIcon(
        html=(f'<div style="width:{size}px;height:{size}px;line-height:{size - 4}px;border-radius:50%;'
              f'background:{bg};color:{fg_c};border:2px solid #fff;box-shadow:0 0 4px rgba(0,0,0,.6);'
              f'text-align:center;font:bold 12px Arial;">{txt}</div>'),
        icon_size=(size, size), icon_anchor=(size // 2, size // 2))
    folium.Marker([row["cy"], row["cx"]], icon=icon,
                  tooltip=folium.Tooltip(tooltip_html(row), sticky=True),
                  popup=folium.Popup(popup_html(row["LGA"]), max_width=380, autoPanPaddingTopLeft=[20, 90], autoPanPaddingBottomRight=[20, 20])).add_to(mk)
mk.add_to(m)

# LGA name labels
lab = folium.FeatureGroup(name="LGA names", overlay=True, show=True)
for _, row in gdf.iterrows():
    folium.Marker([row["cy"], row["cx"]], interactive=False, icon=folium.DivIcon(
        html=(f'<div style="font:bold 10px Arial;color:#102a43;text-shadow:0 0 3px #fff,0 0 3px #fff,0 0 3px #fff;'
              f'white-space:nowrap;transform:translate(-50%,10px);text-align:center">{e(row["LGA"])}</div>'),
        icon_size=(0, 0))).add_to(lab)
lab.add_to(m)
folium.LayerControl(collapsed=False, position="topright").add_to(m)

# --------------------------------------------------------------------------- 9. legends, summary panel, unmatched table, CSV button
def legend_block(label):
    sc = scales[label]
    n = len(sc["edges"])
    items = []
    for i, lb in enumerate(sc["labels"]):
        col = PALETTE[round(i * (len(PALETTE) - 1) / max(1, n - 1))]
        items.append(f'<div><span class="sw" style="background:{col}"></span>{lb}</div>')
    items.append(f'<div><span class="sw" style="background:{NO_DATA}"></span>None mapped</div>')
    return f'<div class="lg" data-metric="{e(label)}"><b>{e(label)}</b>{"".join(items)}</div>'


def table_html(df_, cls):
    if df_.empty:
        return '<div class="none">None – every LGA in the source file matched a Lagos State boundary.</div>'
    head = "".join(f"<th>{e(c)}</th>" for c in df_.columns)
    body = "".join("<tr>" + "".join(f"<td>{e(str(v))}</td>" for v in r) + "</tr>" for r in df_.itertuples(index=False))
    return f'<table class="{cls}"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


panel_cols = ["LGA", "No. of Locations", "Locations", "TKAs", "DPs", "CDOs", "Specialists", "RSMs"]
csv_text = summary.to_csv(index=False)
unmatched_note = (f"{len(unmatched)} source row(s) could not be matched" if len(unmatched) else
                  "All LGAs in the source file matched official boundaries")
n_mapped = int((summary["No. of Locations"] > 0).sum())

ui = f"""
<style>
  .box {{position:fixed;z-index:9999;background:#fff;border:1px solid #9aa;border-radius:6px;
         box-shadow:0 2px 8px rgba(0,0,0,.3);font:12px Arial,Helvetica,sans-serif;color:#222}}
  #title {{top:10px;left:56px;padding:8px 12px;max-width:420px}}
  #title h2 {{margin:0;font-size:16px;color:#08306b}}
  #title .s {{color:#555;margin-top:2px}}
  #key {{bottom:28px;left:10px;padding:8px 10px;width:230px}}
  #key .lg {{display:none}} #key .lg.on {{display:block}}
  #key .sw {{display:inline-block;width:18px;height:12px;margin-right:6px;border:1px solid #888;vertical-align:middle}}
  #key div div {{margin:2px 0}}
  #key hr {{border:0;border-top:1px solid #ddd;margin:6px 0}}
  .dot {{display:inline-block;width:12px;height:12px;border-radius:50%;margin-right:6px;vertical-align:middle}}
  #sumbtn {{top:10px;right:260px;padding:0;border:0;background:none;box-shadow:none}}
  #sumbtn button {{font:bold 12px Arial;padding:7px 11px;margin-left:4px;border:1px solid #08306b;border-radius:5px;
                   background:#08306b;color:#fff;cursor:pointer}}
  #sumbtn button.alt {{background:#fff;color:#08306b}}
  #sumpanel {{display:none;top:52px;left:56px;right:260px;max-height:62vh;overflow:auto;padding:10px}}
  #sumpanel h4 {{margin:6px 0;color:#08306b}}
  #sumpanel table {{border-collapse:collapse;width:100%;font-size:11.5px}}
  #sumpanel th {{position:sticky;top:0;background:#08306b;color:#fff;text-align:left;padding:4px 6px}}
  #sumpanel td {{border-bottom:1px solid #e1e1e1;padding:3px 6px;vertical-align:top}}
  #sumpanel tr:nth-child(even) td {{background:#f6f9fc}}
  #sumpanel .none {{color:#777;font-style:italic;padding:4px 0}}
</style>
<div class="box" id="title"><h2>Lagos Region Resource Mapping 2026</h2>
  <div class="s">{n_mapped} of {len(summary)} LGAs have mapped deployments · {int(summary['No. of Locations'].sum())} locations ·
  click an LGA or numbered marker for details</div></div>
<div class="box" id="sumbtn">
  <button class="alt" onclick="var p=document.getElementById('sumpanel');p.style.display=p.style.display=='block'?'none':'block'">Summary &amp; unmatched LGAs</button>
  <button onclick="dlCSV()">⬇ Download summary (CSV)</button></div>
<div class="box" id="key">
  {''.join(legend_block(l) for l in METRICS)}
  <hr><b>Who is shown in the popup</b>
  <div><span class="dot" style="background:{LOC_COLOR}"></span>Assigned to a <b>location</b>: TKA, DP, CDO</div>
  <div><span class="dot" style="background:{LGA_COLOR}"></span>Covering the <b>whole LGA</b>: Specialist, RSM</div>
  <div style="color:#666;margin-top:4px">Numbered circle = no. of deployment locations</div></div>
<div class="box" id="sumpanel">
  <h4>LGA summary (all {len(summary)} Lagos State LGAs)</h4>{table_html(summary[panel_cols], "sum")}
  <h4>Unmatched LGAs — {unmatched_note}</h4>{table_html(unmatched_df, "unm")}</div>
<script>
  const CSV_TEXT = {json.dumps(csv_text)};
  function dlCSV() {{
    const b = new Blob(["\\ufeff" + CSV_TEXT], {{type: "text/csv;charset=utf-8;"}});
    const a = document.createElement("a"); a.href = URL.createObjectURL(b);
    a.download = "lagos_lga_summary.csv"; document.body.appendChild(a); a.click(); a.remove();
  }}
  // show the legend that matches the active colour layer
  window.addEventListener("load", function () {{
    const names = {json.dumps(list(METRICS))};
    const legends = document.querySelectorAll("#key .lg");
    function show(label) {{ legends.forEach(l => l.classList.toggle("on", l.dataset.metric === label)); }}
    show(names[0]);
    const map = window["{m.get_name()}"];
    map.on("baselayerchange", function (ev) {{
      const n = ev.name.replace("Colour by: ", ""); if (names.includes(n)) show(n);
    }});
  }});
</script>
"""
m.get_root().html.add_child(Element(ui))
m.save(OUT / "lagos_resource_map.html")

# --------------------------------------------------------------------------- 10. Excel summary workbook
xl = OUT / "lagos_lga_summary.xlsx"
cleaning = pd.DataFrame({"Cleaning / matching notes": log})
with pd.ExcelWriter(xl, engine="openpyxl") as w:
    summary.to_excel(w, sheet_name="LGA Summary", index=False)
    detail.to_excel(w, sheet_name="Location Detail", index=False)
    (unmatched_df if len(unmatched_df) else pd.DataFrame(
        {"Unmatched LGAs": ["None – all LGAs in the source file matched Lagos State boundaries"]})
     ).to_excel(w, sheet_name="Unmatched LGAs", index=False)
    cleaning.to_excel(w, sheet_name="Cleaning Log", index=False)

wb2 = load_workbook(xl)
for sh in wb2.worksheets:
    for c in sh[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="08306B")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in sh.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
    for i, col in enumerate(sh.columns, 1):
        longest = max(len(str(c.value or "")) for c in col)
        sh.column_dimensions[get_column_letter(i)].width = min(max(12, longest + 2), 55)
    sh.freeze_panes = "B2" if sh.title in ("LGA Summary", "Location Detail") else "A2"
wb2.save(xl)

# --------------------------------------------------------------------------- console report
print(f"Source sheet: '{sheet_name}' | data rows: {len(df)} | location records: {len(rec)}")
print(f"LGAs matched: {matched['lga'].nunique()} of {df['LGA_raw'].nunique()} distinct source LGAs; "
      f"unmatched rows: {len(unmatched)}")
print("Match methods:", df["match_method"].value_counts().to_dict())
print(summary[["LGA", "No. of Locations", "Total Unique Resources"]].to_string(index=False))
print("\n".join(["", "LOG:"] + log))
