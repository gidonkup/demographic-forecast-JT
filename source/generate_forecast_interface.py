#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generate Demographic Forecast Interface HTML
Reads TAZ shapefile + 3 Excel scenario files → self-contained HTML
"""

import geopandas as gpd
import pandas as pd
import json
import math
import urllib.request
from pathlib import Path
from shapely.ops import unary_union

BASE = Path(__file__).parent
OUT  = BASE.parent / 'index.html'

YEARS = [2020, 2025, 2030, 2035, 2040, 2045, 2050]
SCENARIOS = ['jtmt', 'bau', 'iplan']
EXCEL_FILES = {
    'jtmt':  ('241029_forecast_2020_till_2050_jtmt.xlsx',  0),
    'bau':   ('241029_forecast_2020_till_2050_bau.xlsx',   0),
    'iplan': ('241029_forecast_2020_till_2050_iplan.xlsx', 0),
}
YEAR_COLS = {
    2020: ('aprt_20',   'pop_without_dorms_yeshiva',      'total_emp'),
    2025: ('aprt_2025', 'pop_without_dorms_yeshiva_2025', 'total_emp_2025'),
    2030: ('aprt_2030', 'pop_without_dorms_yeshiva_2030', 'total_emp_2030'),
    2035: ('aprt_2035', 'pop_without_dorms_yeshiva_2035', 'total_emp_2035'),
    2040: ('aprt_2040', 'pop_without_dorms_yeshiva_2040', 'total_emp_2040'),
    2045: ('aprt_2045', 'pop_without_dorms_yeshiva_2045', 'total_emp_2045'),
    2050: ('aprt_2050', 'pop_without_dorms_yeshiva_2050', 'total_emp_2050'),
}

# ── helpers ───────────────────────────────────────────────────────────────

def rv(v):
    if v is None: return 0
    try:
        f = float(v)
        if math.isnan(f) or math.isinf(f): return 0
        return int(round(f))
    except Exception:
        return 0

def rnd_coords(c, d=5):
    if isinstance(c[0], (int, float)):
        return [round(c[0], d), round(c[1], d)]
    return [rnd_coords(x, d) for x in c]

def geom_to_dict(geom, d=5):
    gs = gpd.GeoSeries([geom])
    g = json.loads(gs.to_json())['features'][0]['geometry']
    g['coordinates'] = rnd_coords(g['coordinates'], d)
    return g

def jc(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(',', ':'))

# ── 1. Read shapefile ─────────────────────────────────────────────────────

print('Reading shapefile...')
gdf = gpd.read_file(str(BASE / 'TAZ_V4_241103_with_geo_info.shp'))
gdf = gdf.to_crs('EPSG:4326')
gdf['geometry'] = gdf.geometry.simplify(0.0002, preserve_topology=True)
gdf['geometry'] = gdf.geometry.buffer(0)  # fix invalid geoms

STR_COLS = ['Taz_name','Muni_Heb','SCHN_NAME','main_secto','in_jerusal','ENG_NAME_n','zonetype']
for c in STR_COLS:
    if c in gdf.columns:
        gdf[c] = gdf[c].fillna('').astype(str).str.strip()

gdf['Taz_num'] = pd.to_numeric(gdf['Taz_num'], errors='coerce').fillna(0).astype(int)

# district: use ENG_NAME_n, fallback to zonetype, fallback to 'אחר'
def pick_dist(row):
    v = str(row.get('ENG_NAME_n','')).strip()
    if v and v != 'nan': return v
    v = str(row.get('zonetype','')).strip()
    if v and v != 'nan': return v
    return 'אחר'
gdf['district'] = gdf.apply(pick_dist, axis=1)

print(f'  {len(gdf)} TAZ zones')

# ── 2. Read Excel files ───────────────────────────────────────────────────

print('Reading Excel files...')
excel_data = {}   # {scenario: {taz_num: {year: {a,p,e}}}}

for sc in SCENARIOS:
    fname, hdr = EXCEL_FILES[sc]
    print(f'  {sc} ...')
    df = pd.read_excel(str(BASE / fname), sheet_name=0, header=hdr, engine='openpyxl')
    df['Taz_num'] = pd.to_numeric(df.get('Taz_num', pd.Series(dtype=float)),
                                   errors='coerce').fillna(0).astype(int)
    df = df[df['Taz_num'] > 0].copy()

    sc_dict = {}
    for _, row in df.iterrows():
        taz = int(row['Taz_num'])
        yr_dict = {}
        for yr, (ac, pc, ec) in YEAR_COLS.items():
            yr_dict[yr] = {
                'a': rv(row.get(ac, 0)),
                'p': rv(row.get(pc, 0)),
                'e': rv(row.get(ec, 0)),
            }
        sc_dict[taz] = yr_dict
    excel_data[sc] = sc_dict

# ── 3. Build TAZ GeoJSON (geometry + all forecast data embedded) ──────────

print('Building TAZ GeoJSON...')
taz_features = []
for _, row in gdf.iterrows():
    geom = row.geometry
    if geom is None or geom.is_empty:
        continue
    taz = int(row['Taz_num'])

    d = {}
    for sc in SCENARIOS:
        sc_d = {}
        for yr in YEARS:
            yd = excel_data[sc].get(taz, {}).get(yr, {'a':0,'p':0,'e':0})
            sc_d[yr] = yd
        d[sc] = sc_d

    taz_features.append({
        'type': 'Feature',
        'geometry': geom_to_dict(geom),
        'properties': {
            'taz':  taz,
            'nm':   row['Taz_name'],
            'sec':  row['main_secto'],
            'schn': row['SCHN_NAME'],
            'muni': row['Muni_Heb'],
            'dist': row['district'],
            'jm':   rv(row.get('jeru_metro', 0)),
            'ij':   1 if row['in_jerusal'] == 'yes' else 0,
            'd':    d,
        }
    })
print(f'  {len(taz_features)} TAZ features')

# ── 4. Dissolved level helper ─────────────────────────────────────────────

def make_dissolved(source_gdf, col, tol=0.0003):
    features = []
    for name, grp in source_gdf.groupby(col):
        if not name or str(name) in ('', 'nan', 'אחר'): continue
        try:
            geom = unary_union(grp.geometry.values).simplify(tol, preserve_topology=True)
        except Exception:
            continue
        if geom is None or geom.is_empty: continue
        tazs = [int(t) for t in grp['Taz_num'].values if t > 0]
        vc = grp['main_secto'].value_counts()
        dom_sec = vc.index[0] if len(vc) else 'Jewish'
        features.append({
            'type': 'Feature',
            'geometry': geom_to_dict(geom),
            'properties': {'nm': str(name), 'sec': dom_sec, 'tazs': tazs}
        })
    return features

print('Building neighborhood GeoJSON (Jerusalem only)...')
gdf_jeru = gdf[gdf['in_jerusal'] == 'yes'].copy()
neigh_features = make_dissolved(gdf_jeru, 'SCHN_NAME', 0.0002)
print(f'  {len(neigh_features)} neighborhoods')

print('Building municipality GeoJSON...')
muni_features = make_dissolved(gdf, 'Muni_Heb', 0.0003)
print(f'  {len(muni_features)} municipalities')

print('Building district GeoJSON...')
dist_features = make_dissolved(gdf, 'district', 0.0005)
print(f'  {len(dist_features)} districts')

# ── 5. Aggregated data ────────────────────────────────────────────────────

def compute_agg(features):
    result = {}
    for f in features:
        nm = f['properties']['nm']
        tazs = f['properties']['tazs']
        sc_d = {}
        for sc in SCENARIOS:
            yr_d = {}
            for yr in YEARS:
                tot = {'a':0,'p':0,'e':0}
                for taz in tazs:
                    yd = excel_data[sc].get(taz, {}).get(yr, {})
                    for k in ('a','p','e'):
                        tot[k] += yd.get(k, 0)
                yr_d[yr] = tot
            sc_d[sc] = yr_d
        result[nm] = sc_d
    return result

print('Computing aggregated data...')
agg_data = {
    'neigh': compute_agg(neigh_features),
    'muni':  compute_agg(muni_features),
    'dist':  compute_agg(dist_features),
}

# strip tazs from props (save ~20% space)
for fl in (neigh_features, muni_features, dist_features):
    for f in fl:
        f['properties'].pop('tazs', None)

# ── 6. Serialize ──────────────────────────────────────────────────────────

print('Serializing...')
TAZ_JS   = jc({'type':'FeatureCollection','features':taz_features})
NEIGH_JS = jc({'type':'FeatureCollection','features':neigh_features})
MUNI_JS  = jc({'type':'FeatureCollection','features':muni_features})
DIST_JS  = jc({'type':'FeatureCollection','features':dist_features})
AGG_JS   = jc(agg_data)

print(f'  TAZ: {len(TAZ_JS)//1024} KB  Neigh: {len(NEIGH_JS)//1024} KB  Muni: {len(MUNI_JS)//1024} KB  Dist: {len(DIST_JS)//1024} KB  Agg: {len(AGG_JS)//1024} KB')

# ── 6b. Download/embed Leaflet ────────────────────────────────────────────

LEAFLET_CSS_PATH = Path(r'C:\Users\gidon\AppData\Local\Temp\leaflet.css')
LEAFLET_JS_PATH  = Path(r'C:\Users\gidon\AppData\Local\Temp\leaflet.js')

def get_leaflet():
    css_url = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.css'
    js_url  = 'https://unpkg.com/leaflet@1.9.4/dist/leaflet.js'
    if LEAFLET_CSS_PATH.exists() and LEAFLET_JS_PATH.exists():
        print('Using cached Leaflet...')
        css = LEAFLET_CSS_PATH.read_text(encoding='utf-8')
        js  = LEAFLET_JS_PATH.read_text(encoding='utf-8')
    else:
        print('Downloading Leaflet...')
        css = urllib.request.urlopen(css_url).read().decode('utf-8')
        js  = urllib.request.urlopen(js_url).read().decode('utf-8')
        LEAFLET_CSS_PATH.write_text(css, encoding='utf-8')
        LEAFLET_JS_PATH.write_text(js, encoding='utf-8')
    # Sanitize: </script inside a <script> block terminates HTML parsing prematurely
    import re as _re
    js  = _re.sub(r'<(/script)', r'<\\\1', js,  flags=_re.IGNORECASE)
    css = _re.sub(r'<(/style)',  r'<\\\1', css, flags=_re.IGNORECASE)
    print(f'  Leaflet CSS: {len(css)//1024} KB  JS: {len(js)//1024} KB')
    return css, js

LEAFLET_CSS, LEAFLET_JS = get_leaflet()

# ── 7. HTML ───────────────────────────────────────────────────────────────

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>ממשק תחזיות דמוגרפיות — תחבורה לירושלים</title>
<style>__LEAFLET_CSS__</style>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;font-family:"Segoe UI","Arial Hebrew",Arial,sans-serif;font-size:14px;color:#1e293b;direction:rtl}
button{font-family:inherit;cursor:pointer}

/* Header */
.app-header{display:flex;align-items:center;gap:12px;background:#1B3CC0;color:white;padding:8px 20px;height:52px;box-shadow:0 2px 8px rgba(0,0,0,.3);flex-shrink:0;position:relative;z-index:500}
.header-title{font-size:16px;font-weight:700;line-height:1.2}
.header-sub{font-size:11px;opacity:.8}

/* Layout */
.app-body{display:flex;height:calc(100vh - 52px);direction:ltr}
.map-wrap{flex:1;position:relative;min-width:0}
#map{width:100%;height:100%;background:#e8eaed}
.side-panel{width:33%;min-width:300px;max-width:480px;flex-shrink:0;background:#f0f4f8;border-left:1px solid #cbd5e1;overflow-y:auto;direction:rtl;display:flex;flex-direction:column;gap:0}

/* Panel cards */
.p-card{background:white;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.07);margin:8px 8px 0;flex-shrink:0}
.p-card:last-child{margin-bottom:8px}
.p-card-hdr{display:flex;align-items:center;justify-content:space-between;padding:8px 12px;border-bottom:1px solid #e2e8f0;background:#f8fafc;border-radius:8px 8px 0 0}
.p-card-title{font-size:11px;font-weight:700;color:#475569;text-transform:uppercase;letter-spacing:.06em}
.p-card-body{padding:10px 12px}

/* Segments */
.seg-row{display:flex;border:1px solid #cbd5e1;border-radius:7px;overflow:hidden}
.seg-btn{flex:1;border:none;padding:6px 4px;font-size:12px;font-family:inherit;cursor:pointer;background:white;color:#475569;transition:background .12s;line-height:1.3}
.seg-btn+.seg-btn{border-right:1px solid #cbd5e1}
.seg-btn.active{background:#1B3CC0;color:white;font-weight:600}

/* Pills */
.pill-row{display:flex;flex-wrap:wrap;gap:5px}
.pill{border:1px solid #cbd5e1;background:white;border-radius:20px;padding:4px 12px;font-size:12px;color:#475569;cursor:pointer;transition:all .12s;font-family:inherit}
.pill:hover{border-color:#1B3CC0;color:#1B3CC0}
.pill.active{background:#1B3CC0;border-color:#1B3CC0;color:white;font-weight:600}

/* Year */
.year-track{display:flex;gap:3px;flex-wrap:wrap}
.yr-pill{border:1px solid #cbd5e1;background:white;border-radius:6px;padding:4px 8px;font-size:12px;color:#475569;cursor:pointer;font-family:inherit;min-width:48px;text-align:center;transition:all .12s}
.yr-pill:hover{border-color:#1B3CC0}
.yr-pill.active{background:#1B3CC0;border-color:#1B3CC0;color:white;font-weight:700}

/* Sector chips */
.sec-chip{display:inline-flex;align-items:center;gap:5px;border-radius:20px;padding:3px 10px;font-size:12px;cursor:pointer;border:2px solid transparent;user-select:none;color:white;font-weight:600;transition:opacity .12s}
.sec-chip.off{opacity:.3}

/* Table */
.sum-wrap{overflow-x:auto;max-height:280px;overflow-y:auto}
.sum-table{width:100%;border-collapse:collapse;font-size:12px}
.sum-table th{background:#f1f5f9;padding:6px 8px;text-align:right;font-weight:700;font-size:11px;color:#475569;border-bottom:2px solid #e2e8f0;position:sticky;top:0;z-index:1;cursor:pointer;white-space:nowrap;user-select:none}
.sum-table th:hover{background:#e2e8f0}
.sum-table th.sort-active{color:#1B3CC0}
.sum-table td{padding:5px 8px;border-bottom:1px solid #f1f5f9;text-align:right;white-space:nowrap}
.sum-table tr:hover td{background:#f8fafc;cursor:pointer}
.sum-table tfoot td{background:#eff4ff;font-weight:700;color:#1B3CC0;border-top:2px solid #bfdbfe;position:sticky;bottom:0}
.td-name{max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600}
.td-dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-left:4px;vertical-align:middle}

/* Choropleth bar */
.choro-bar{height:14px;border-radius:3px;background:linear-gradient(to left,#b91c1c,#f59e0b,#fef08a);margin:4px 0}
.choro-scale{display:flex;justify-content:space-between;font-size:10px;color:#64748b}

/* Map tooltip */
#map-tip{position:absolute;background:rgba(15,23,42,.88);color:white;padding:6px 10px;border-radius:6px;font-size:12px;pointer-events:none;z-index:600;display:none;direction:rtl;max-width:220px;line-height:1.5;white-space:pre-wrap}

/* Map legend overlay */
#map-legend{position:absolute;bottom:28px;left:10px;z-index:500;background:rgba(255,255,255,.93);border-radius:8px;padding:8px 12px;font-size:12px;box-shadow:0 2px 8px rgba(0,0,0,.18);direction:rtl;pointer-events:none;min-width:130px}
#map-legend .leg-row{display:flex;align-items:center;gap:6px;padding:2px 0;line-height:1.3}
#map-legend .leg-sw{width:12px;height:12px;border-radius:3px;flex-shrink:0}
#map-legend .leg-title{font-weight:600;margin-bottom:4px;font-size:11px;color:#475569}
#taz-hint{display:none;align-items:center;justify-content:center;height:100%;color:#94a3b8;font-size:13px;text-align:center;padding:20px;direction:rtl}

/* Entity legend */
.ent-grid{display:grid;grid-template-columns:1fr 1fr;gap:2px 6px}
.ent-row{display:flex;align-items:center;gap:5px;padding:2px 0;min-width:0}
.ent-swatch{width:12px;height:12px;border-radius:3px;flex-shrink:0}
.ent-nm{font-size:11px;color:#374151;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;direction:rtl}

/* Detail panel */
#detail-pane{display:none;background:white;border-top:2px solid #1B3CC0;padding:10px 12px;flex-shrink:0;direction:rtl}
.dp-close{float:left;background:none;border:none;color:#94a3b8;font-size:16px;cursor:pointer;padding:0;line-height:1}
.dp-title{font-size:13px;font-weight:700;margin-bottom:5px;padding-left:24px}
.dp-sec{display:flex;align-items:center;gap:6px;font-size:12px;color:#475569;margin-bottom:8px}
.dp-table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:8px}
.dp-table td{padding:3px 6px;border-bottom:1px solid #f1f5f9}
.dp-table td:first-child{color:#64748b;width:50%}
.dp-table td:last-child{font-weight:700}
.dp-years{display:flex;gap:4px;flex-wrap:wrap}
.dp-yr{text-align:center;padding:4px 6px;border-radius:6px;border:1px solid #e2e8f0;font-size:11px;cursor:pointer;min-width:50px}
.dp-yr:hover:not(.dp-active){background:#eff4ff;border-color:#1B3CC0}
.dp-yr.dp-active{background:#1B3CC0;color:white;font-weight:700;border-color:#1B3CC0}
.dp-yr-val{font-size:10px;color:#94a3b8}
.dp-yr.dp-active .dp-yr-val{color:rgba(255,255,255,.8)}
/* Filter selects */
.flt-lbl{font-size:11px;color:#64748b;display:block;margin-bottom:3px}
.sf-select{width:100%;padding:5px 8px;border:1px solid #e2e8f0;border-radius:6px;font-size:12px;direction:rtl;font-family:inherit;background:#f8fafc;cursor:pointer;color:#0f172a;appearance:auto}
.sf-select:focus{outline:none;border-color:#1B3CC0}

/* Search */
#search-wrap{position:relative;padding:10px 12px 6px;border-bottom:1px solid #e2e8f0}
#search-box{display:flex;align-items:center;gap:6px;background:#f8fafc;border:1px solid #cbd5e1;border-radius:8px;padding:0 10px;transition:border-color .15s}
#search-box:focus-within{border-color:#1B3CC0;box-shadow:0 0 0 2px rgba(27,60,192,.12);background:#fff}
#search-icon{color:#94a3b8;font-size:14px;flex-shrink:0}
#search-inp{flex:1;border:none;background:transparent;padding:8px 4px;font-size:13px;direction:rtl;font-family:inherit;outline:none;color:#0f172a}
#search-inp::placeholder{color:#94a3b8}
#search-drop{position:absolute;top:calc(100% - 4px);right:12px;left:12px;background:#fff;border:1px solid #e2e8f0;border-radius:8px;box-shadow:0 4px 20px rgba(0,0,0,.13);z-index:2000;max-height:270px;overflow-y:auto;display:none}
.srch-item{padding:8px 12px;cursor:pointer;direction:rtl;display:flex;align-items:center;gap:8px;border-bottom:1px solid #f8fafc}
.srch-item:last-child{border-bottom:none}
.srch-item:hover{background:#eff6ff}
.srch-nm{flex:1;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.srch-tag{font-size:10px;color:#64748b;background:#f1f5f9;padding:2px 7px;border-radius:10px;white-space:nowrap;flex-shrink:0}
</style>
</head>
<body>

<header class="app-header">
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 110 80" width="34" height="24">
    <rect x="4" y="5" width="38" height="70" fill="white" rx="2"/>
    <rect x="68" y="5" width="38" height="70" fill="white" rx="2"/>
    <ellipse cx="55" cy="40" rx="31" ry="38" fill="#1B3CC0"/>
  </svg>
  <div>
    <div class="header-title">ממשק תחזיות דמוגרפיות</div>
    <div class="header-sub">תחבורה לירושלים — מטרופולין ירושלים 2020–2050</div>
  </div>
  <div style="margin-right:auto"></div>
  <span style="font-size:11px;opacity:.75;white-space:nowrap">נתונים גרסא 0.9</span>
  <a href="forecast-documentation.pdf" target="_blank"
     style="display:inline-flex;align-items:center;gap:5px;background:rgba(255,255,255,0.15);
            color:white;text-decoration:none;font-size:12px;padding:4px 10px;border-radius:6px;
            border:1px solid rgba(255,255,255,0.3);white-space:nowrap;transition:background .2s"
     onmouseover="this.style.background=\'rgba(255,255,255,0.25)\'"
     onmouseout="this.style.background=\'rgba(255,255,255,0.15)\'">
    &#128196; דוח תיעוד
  </a>
</header>

<div class="app-body">
  <div class="map-wrap">
    <div id="map"></div>
    <div id="map-tip"></div>
    <div id="map-legend"></div>
  </div>

  <div class="side-panel">

    <!-- חיפוש -->
    <div id="search-wrap">
      <div id="search-box">
        <span id="search-icon">&#128269;</span>
        <input id="search-inp" type="text" placeholder="חיפוש יישוב / שכונה / אזור תנועה..."
               oninput="doSearch(this.value)"
               onblur="setTimeout(hideSearchDrop,200)"
               autocomplete="off">
      </div>
      <div id="search-drop"></div>
    </div>

    <!-- תרחיש -->
    <div class="p-card">
      <div class="p-card-hdr"><span class="p-card-title">תרחיש</span></div>
      <div class="p-card-body">
        <div class="seg-row">
          <button class="seg-btn active" id="sc-jtmt"  onclick="setScenario(\'jtmt\')">תחבורה לירושלים</button>
          <button class="seg-btn"        id="sc-iplan" onclick="setScenario(\'iplan\')">דיור אסטרטגי</button>
          <button class="seg-btn"        id="sc-bau"   onclick="setScenario(\'bau\')">המשך מגמות</button>
        </div>
      </div>
    </div>

    <!-- שנה -->
    <div class="p-card">
      <div class="p-card-hdr"><span class="p-card-title">שנה</span></div>
      <div class="p-card-body">
        <div class="year-track" id="year-track"></div>
      </div>
    </div>

    <!-- רמת הצגה -->
    <div class="p-card">
      <div class="p-card-hdr"><span class="p-card-title">רמת הצגה</span></div>
      <div class="p-card-body">
        <div class="pill-row">
          <button class="pill active" id="lv-taz"   onclick="setLevel(\'taz\')">אזורי תנועה</button>
          <button class="pill"        id="lv-neigh" onclick="setLevel(\'neigh\')">שכונות ירושלים</button>
          <button class="pill"        id="lv-muni"  onclick="setLevel(\'muni\')">רשויות מקומיות</button>
          <button class="pill"        id="lv-dist"  onclick="setLevel(\'dist\')">מחוזות</button>
        </div>
      </div>
    </div>

    <!-- הצגה על מפה -->
    <div class="p-card">
      <div class="p-card-hdr"><span class="p-card-title">הצגה על המפה</span></div>
      <div class="p-card-body" style="display:flex;flex-direction:column;gap:8px">
        <div class="seg-row">
          <button class="seg-btn active" id="md-choro"  onclick="setMode(\'choro\')">כמותי</button>
          <button class="seg-btn"        id="md-sector" onclick="setMode(\'sector\')">מגזר</button>
          <button class="seg-btn"        id="md-entity" onclick="setMode(\'entity\')" style="display:none">ישות</button>
        </div>
        <!-- metric (choropleth mode) -->
        <div id="metric-row" class="pill-row">
          <button class="pill active" id="mt-p" onclick="setMetric(\'p\')">אוכלוסייה</button>
          <button class="pill"        id="mt-a" onclick="setMetric(\'a\')">יחידות דיור</button>
          <button class="pill"        id="mt-e" onclick="setMetric(\'e\')">מועסקים</button>
        </div>
        <!-- choropleth legend -->
        <div id="choro-leg">
          <div class="choro-bar"></div>
          <div class="choro-scale">
            <span id="ch-min">0</span><span id="ch-mid">—</span><span id="ch-max">0</span>
          </div>
        </div>
        <!-- entity legend -->
        <div id="entity-leg" style="display:none;max-height:170px;overflow-y:auto;direction:rtl"></div>
      </div>
    </div>

    <!-- סינון -->
    <div class="p-card">
      <div class="p-card-hdr">
        <span class="p-card-title">סינון</span>
        <button class="pill" onclick="clearAllFilters()" style="font-size:11px;padding:3px 10px;color:#dc2626;border-color:#fca5a5">נקה הכל</button>
      </div>
      <div class="p-card-body" style="display:flex;flex-direction:column;gap:8px">
        <div style="display:flex;align-items:center;justify-content:space-between">
          <span class="flt-lbl" style="margin:0">מגזר</span>
          <button class="pill" id="sec-all-btn" onclick="selectAllSectors()" style="font-size:11px;padding:3px 10px">בחר הכל</button>
        </div>
        <div id="sector-filter" class="pill-row" style="gap:4px"></div>
        <div>
          <span class="flt-lbl">רשות מקומית</span>
          <select id="sf-muni" onchange="setSFMuni(this.value)" class="sf-select">
            <option value="">הכל</option>
          </select>
        </div>
        <div id="sf-neigh-wrap" style="display:none">
          <span class="flt-lbl">שכונה</span>
          <select id="sf-neigh" onchange="setSFNeigh(this.value)" class="sf-select">
            <option value="">הכל</option>
          </select>
        </div>
      </div>
    </div>

    <!-- סיכום -->
    <div class="p-card" style="flex:1;display:flex;flex-direction:column;min-height:200px">
      <div class="p-card-hdr">
        <span class="p-card-title">סיכום</span>
        <span id="tbl-cap" style="font-size:11px;color:#94a3b8"></span>
      </div>
      <div id="sum-wrap" class="sum-wrap" style="flex:1">
        <table class="sum-table">
          <thead>
            <tr>
              <th onclick="sortTbl(\'nm\')" id="th-nm">שם</th>
              <th onclick="sortTbl(\'p\')"  id="th-p">אוכלוסייה</th>
              <th onclick="sortTbl(\'a\')"  id="th-a">יח"ד</th>
              <th onclick="sortTbl(\'e\')"  id="th-e">מועסקים</th>
            </tr>
          </thead>
          <tbody id="sum-tbody"></tbody>
          <tfoot><tr id="sum-foot"></tr></tfoot>
        </table>
      </div>
    </div>

    <!-- detail -->
    <div id="detail-pane">
      <button class="dp-close" onclick="closeDetail()">&#x2715;</button>
      <div id="detail-content"></div>
    </div>

  </div>
</div>

<script>__LEAFLET_JS__</script>
<script>
// ── DATA ─────────────────────────────────────────────────────────────────
const TAZ_GJ   = __TAZ_GJ__;
const NEIGH_GJ = __NEIGH_GJ__;
const MUNI_GJ  = __MUNI_GJ__;
const DIST_GJ  = __DIST_GJ__;
const AGG      = __AGG__;

const GJ = {taz:TAZ_GJ, neigh:NEIGH_GJ, muni:MUNI_GJ, dist:DIST_GJ};
const YEARS = [2020,2025,2030,2035,2040,2045,2050];

// ── SECTOR CONFIG ─────────────────────────────────────────────────────────
const SECTOR = {
  'Jewish':                        {label:'יהודי כללי',      color:'#78B858'},
  'U_Orthodox':                    {label:'חרדי',             color:'#3AAAC8'},
  'Arab':                          {label:'ערבי',             color:'#D4A020'},
  'arabs_behined_seperation_wall': {label:'ערבי מעבר לגדר',  color:'#D06030'},
  'Palestinian':                   {label:'פלסטינאי',         color:'#C04040'},
};
const DEFAULT_CLR = '#9AAABB';

function secClr(s) { return (SECTOR[s]||{}).color || DEFAULT_CLR; }
function secLbl(s) { return (SECTOR[s]||{}).label || s || 'אחר'; }

// ── ENTITY COLOR ──────────────────────────────────────────────────────────
// Golden-angle hue sequence for maximally distinct colors
const EMAP_PALETTE = Array.from({length:50}, (_,i) => {
  const h = Math.round((i * 137.508) % 360);
  const s = 58 + (i%3)*8;
  const l = 40 + (i%2)*12;
  return `hsl(${h},${s}%,${l}%)`;
});
let entityColors = {};
let hiddenEntities = new Set();

function buildEntityColors() {
  entityColors = {};
  let idx = 0;
  GJ[S.lv].features.forEach(f => {
    const nm = f.properties.nm || '';
    if (!(nm in entityColors)) { entityColors[nm] = EMAP_PALETTE[idx++ % EMAP_PALETTE.length]; }
  });
}

function toggleEntity(nm) {
  if (hiddenEntities.has(nm)) hiddenEntities.delete(nm); else hiddenEntities.add(nm);
  updateLegend();
  if (geoLayer) geoLayer.setStyle(styleF);
  updateTable();
}

function toggleAllEntities() {
  const allVis = Object.keys(entityColors).every(nm=>!hiddenEntities.has(nm));
  if (allVis) Object.keys(entityColors).forEach(nm=>hiddenEntities.add(nm));
  else hiddenEntities.clear();
  updateLegend();
  if (geoLayer) geoLayer.setStyle(styleF);
  updateTable();
}

// ── STATE ─────────────────────────────────────────────────────────────────
const S = {
  sc: 'jtmt', yr: 2025, lv: 'taz', mode: 'choro', met: 'p',
  filter: new Set(Object.keys(SECTOR)),
  sfMuni: '', sfNeigh: '',
  sortK: 'p', sortDir: -1,
};
let RANGE = {min:0, max:1};

// ── MAP ───────────────────────────────────────────────────────────────────
const map = L.map('map',{center:[31.80,35.20],zoom:10,preferCanvas:true});
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
  attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  maxZoom:18
}).addTo(map);

let geoLayer = null;
const tip = document.getElementById('map-tip');

// ── VALUE ACCESS ─────────────────────────────────────────────────────────
function getVal(props, met) {
  const m = met||S.met;
  if (S.lv === 'taz') return props.d?.[S.sc]?.[S.yr]?.[m] || 0;
  return AGG[S.lv]?.[props.nm]?.[S.sc]?.[S.yr]?.[m] || 0;
}

function isVisibleFeature(f) {
  const p = f.properties;
  if (S.lv === 'taz') return isVisibleTaz(p);
  if (!S.filter.has(p.sec || '')) return false;
  if (S.sfMuni  && S.lv === 'muni'  && p.nm !== S.sfMuni)  return false;
  if (S.sfNeigh && S.lv === 'neigh' && p.nm !== S.sfNeigh) return false;
  return true;
}

function computeRange() {
  const vals = GJ[S.lv].features.filter(isVisibleFeature).map(f=>getVal(f.properties)).filter(v=>v>0);
  return vals.length ? {min:Math.min(...vals), max:Math.max(...vals)} : {min:0,max:1};
}

// ── CHORO COLOR ───────────────────────────────────────────────────────────
function choroClr(v) {
  const {min,max} = RANGE;
  if (max<=min) return '#fef9c3';
  const t = Math.pow(Math.max(0,(v-min)/(max-min)),0.5);
  return `rgb(255,${Math.round(255-t*215)},${Math.round(80-t*80)})`;
}

// ── STYLE ─────────────────────────────────────────────────────────────────
// Returns true if a TAZ feature's properties pass all active filters
function isVisibleTaz(p) {
  if (!S.filter.has(p.sec || '')) return false;
  if (S.sfMuni  && p.muni !== S.sfMuni)  return false;
  if (S.sfNeigh && p.schn !== S.sfNeigh) return false;
  return true;
}

function styleF(feature) {
  const p   = feature.properties;
  const sec = p.sec || '';
  const nm  = p.nm  || '';
  const isTaz = S.lv === 'taz';
  const border = isTaz ? {color:'transparent', weight:0} : {color:'#888', weight:0.7};
  const hidden = {fillColor:'#e2e8f0', fillOpacity:.10, color:'transparent', weight:0};

  // ── Filters ───────────────────────────────────────────────────────────────
  if (isTaz) {
    if (!isVisibleTaz(p)) return hidden;
  } else {
    if (!S.filter.has(sec)) return hidden;
    if (S.sfMuni  && S.lv === 'muni'  && nm !== S.sfMuni)  return hidden;
    if (S.sfNeigh && S.lv === 'neigh' && nm !== S.sfNeigh) return hidden;
  }

  // ── Entity visibility ──────────────────────────────────────────────────────
  if (S.mode === 'entity' && hiddenEntities.has(nm)) return hidden;

  // ── Color ──────────────────────────────────────────────────────────────────
  let fill;
  if (S.mode === 'sector') fill = secClr(sec);
  else if (S.mode === 'entity') fill = entityColors[nm] || DEFAULT_CLR;
  else fill = choroClr(getVal(p));
  return {fillColor:fill, fillOpacity:.84, ...border, opacity:.8};
}

// ── MAP UPDATE ────────────────────────────────────────────────────────────
function updateMap() {
  if (S.mode==='choro') RANGE = computeRange();
  if (S.mode==='entity') buildEntityColors();   // must run before styleF is called
  if (geoLayer) map.removeLayer(geoLayer);

  geoLayer = L.geoJSON(GJ[S.lv], {
    style: styleF,
    onEachFeature(feature, layer) {
      layer.on({
        mouseover(e) {
          const l = e.target;
          l.setStyle({weight:2.5, color:'#1B3CC0', fillOpacity:.95});
          if (!L.Browser.ie && !L.Browser.opera && !L.Browser.edge) l.bringToFront();
          const p = feature.properties;
          const nm = S.lv==='taz' ? (p.nm||'TAZ '+p.taz) : p.nm;
          const ml = S.met==='p'?'אוכלוסייה':S.met==='a'?'יח"ד':'מועסקים';
          tip.innerHTML = '<strong>'+esc(nm)+'</strong>\\n'+secLbl(p.sec)+'\\n'+ml+': '+fmt(getVal(p));
          tip.style.display='block';
        },
        mousemove(e) {
          const r = map.getContainer().getBoundingClientRect();
          tip.style.left=(e.originalEvent.clientX-r.left+14)+'px';
          tip.style.top=(e.originalEvent.clientY-r.top-8)+'px';
        },
        mouseout(e) { geoLayer.resetStyle(e.target); tip.style.display='none'; },
        click() { showDetail(feature.properties); }
      });
    }
  }).addTo(map);

  updateLegend();
  updateTable();
}

// ── MAP LEGEND OVERLAY ────────────────────────────────────────────────────
function updateMapLegend() {
  const el = document.getElementById(\'map-legend\');
  if (!el) return;
  if (S.mode === \'sector\') {
    el.style.display = \'block\';
    el.innerHTML = \'<div class="leg-title">מגזר</div>\' +
      Object.entries(SECTOR).map(([k,v]) => {
        const dim = !S.filter.has(k);
        return \'<div class="leg-row" style="opacity:\' + (dim?0.3:1) + \'">\' +
               \'<span class="leg-sw" style="background:\' + v.color + \'"></span>\' +
               v.label + \'</div>\';
      }).join(\'\');
  } else if (S.mode === \'choro\') {
    const metLbl = S.met===\'p\'?\'אוכלוסייה\':S.met===\'a\'?\'יח"ד\':\'מועסקים\';
    const {min,max} = RANGE;
    el.style.display = \'block\';
    el.innerHTML = \'<div class="leg-title">\' + metLbl + \'</div>\' +
      \'<div class="choro-bar" style="height:12px;margin:4px 0"></div>\' +
      \'<div style="display:flex;justify-content:space-between;font-size:10px;color:#64748b">\' +
      \'<span>\' + fmt(max) + \'</span><span>\' + fmt(min) + \'</span></div>\';
  } else if (S.mode === \'entity\' && S.lv !== \'taz\') {
    const entries = Object.entries(entityColors).filter(([nm])=>!hiddenEntities.has(nm));
    el.style.display = \'block\';
    el.style.maxHeight = \'220px\';
    el.style.overflowY = \'auto\';
    el.innerHTML = \'<div class="leg-title">ישות</div>\' +
      entries.map(([nm,clr]) =>
        \'<div class="leg-row"><span class="leg-sw" style="background:\' + clr + \'"></span>\' +
        \'<span style="font-size:11px">\' + esc(nm) + \'</span></div>\'
      ).join(\'\');
  } else {
    el.style.display = \'none\';
  }
}

// ── LEGEND / FILTER ───────────────────────────────────────────────────────
function updateLegend() {
  const secEl  = document.getElementById('sector-filter');
  const allBtn = document.getElementById('sec-all-btn');
  const entLeg = document.getElementById('entity-leg');
  const choLeg = document.getElementById('choro-leg');
  const metRow = document.getElementById('metric-row');

  // Sector chips — always visible, global filter
  const allOn = Object.keys(SECTOR).every(k=>S.filter.has(k));
  allBtn.textContent = allOn ? 'נקה הכל' : 'בחר הכל';
  secEl.innerHTML = Object.entries(SECTOR).map(([k,v])=>{
    const on = S.filter.has(k);
    return `<button class="sec-chip${on?'':' off'}" style="background:${v.color}" onclick="toggleSec('${k}')">${v.label}</button>`;
  }).join('');

  // Sync dropdowns with state
  const muniSel = document.getElementById('sf-muni');
  if (muniSel) muniSel.value = S.sfMuni;
  const neighSel = document.getElementById('sf-neigh');
  if (neighSel) neighSel.value = S.sfNeigh;

  // Mode-specific controls
  metRow.style.display = S.mode==='choro' ? 'flex' : 'none';
  choLeg.style.display  = S.mode==='choro' ? 'block' : 'none';
  entLeg.style.display  = S.mode==='entity' ? 'block' : 'none';

  if (S.mode==='entity') {
    const entries = Object.entries(entityColors);
    const allVis = entries.every(([nm])=>!hiddenEntities.has(nm));
    entLeg.innerHTML =
      `<div style="display:flex;justify-content:flex-end;margin-bottom:4px"><button class="pill" onclick="toggleAllEntities()" style="font-size:11px;padding:3px 10px">${allVis?'נקה הכל':'בחר הכל'}</button></div>`+
      '<div class="ent-grid">'+entries.map(([nm,clr])=>{
        const off = hiddenEntities.has(nm);
        return `<div class="ent-row" onclick="toggleEntity('${esc(nm)}')" style="cursor:pointer;opacity:${off?.3:1}"><span class="ent-swatch" style="background:${clr}"></span><span class="ent-nm" title="${esc(nm)}">${esc(nm)}</span></div>`;
      }).join('')+'</div>';
  }

  if (S.mode==='choro') {
    const {min,max} = RANGE;
    document.getElementById('ch-min').textContent = fmt(min);
    document.getElementById('ch-mid').textContent = fmt(Math.round((min+max)/2));
    document.getElementById('ch-max').textContent = fmt(max);
  }
  updateMapLegend();
}

function setSFMuni(v) {
  S.sfMuni = v; S.sfNeigh = '';
  const ns = document.getElementById('sf-neigh');
  if (ns) ns.value = '';
  // Show neighborhood filter only when Jerusalem is selected
  const hasNeigh = v && GJ.taz.features.some(f=>f.properties.muni===v && f.properties.schn);
  const nw = document.getElementById('sf-neigh-wrap');
  if (nw) nw.style.display = hasNeigh ? '' : 'none';
  applyFilters();
}
function setSFNeigh(v) { S.sfNeigh = v; applyFilters(); }
function clearAllFilters() {
  Object.keys(SECTOR).forEach(k=>S.filter.add(k));
  S.sfMuni=''; S.sfNeigh='';
  const ms=document.getElementById('sf-muni'); if(ms) ms.value='';
  const ns=document.getElementById('sf-neigh'); if(ns) ns.value='';
  const nw=document.getElementById('sf-neigh-wrap'); if(nw) nw.style.display='none';
  applyFilters();
}
function applyFilters() {
  if (S.mode==='choro') RANGE = computeRange();
  updateLegend();
  if (geoLayer) geoLayer.setStyle(styleF);
  updateTable();
}

function toggleSec(k) {
  if (S.filter.has(k)) S.filter.delete(k); else S.filter.add(k);
  applyFilters();
}

function selectAllSectors() {
  const allOn = Object.keys(SECTOR).every(k=>S.filter.has(k));
  if (allOn) S.filter.clear(); else Object.keys(SECTOR).forEach(k=>S.filter.add(k));
  applyFilters();
}

// ── TABLE ─────────────────────────────────────────────────────────────────
function buildRows() {
  const sc=S.sc, yr=S.yr, m=S.sortK;
  let all;
  if (S.lv==='taz') {
    // Aggregate visible TAZ zones by sector
    const acc = {};
    Object.entries(SECTOR).forEach(([k,v])=>{ acc[k]={nm:v.label, sec:k, a:0, p:0, e:0}; });
    GJ.taz.features.forEach(f=>{
      const p=f.properties;
      if (!isVisibleTaz(p)) return;
      const d=p.d?.[sc]?.[yr]||{a:0,p:0,e:0};
      if (acc[p.sec]) { acc[p.sec].a+=d.a; acc[p.sec].p+=d.p; acc[p.sec].e+=d.e; }
    });
    all = Object.values(acc).filter(r=>r.a>0||r.p>0||r.e>0)
          .sort((a,b)=>m==='nm'?a.nm.localeCompare(b.nm,'he'):S.sortDir*(a[m]-b[m]));
    return {display: all, total: all, noData: 0, isSector: true};
  }
  let raw = Object.entries(AGG[S.lv]||{})
    .map(([nm,scD])=>{
      const d=scD?.[sc]?.[yr]||{a:0,p:0,e:0};
      return {nm, sec:'', a:d.a, p:d.p, e:d.e};
    });
  // Apply spatial filters at dissolved levels
  if (S.sfMuni  && S.lv==='muni')  raw = raw.filter(r=>r.nm===S.sfMuni);
  if (S.sfNeigh && S.lv==='neigh') raw = raw.filter(r=>r.nm===S.sfNeigh);
  if (S.mode==='entity') raw = raw.filter(r=>!hiddenEntities.has(r.nm));
  const withData = raw.filter(r=>r.a>0||r.p>0||r.e>0);
  const noData = raw.length - withData.length;
  all = withData.sort((a,b)=>m==='nm'?a.nm.localeCompare(b.nm,'he'):S.sortDir*(a[m]-b[m]));
  return {display: all, total: all, noData, isSector: false};
}

function updateTable() {
  const {display: rows, total: allRows, noData} = buildRows();
  const tot = allRows.reduce((acc,r)=>({a:acc.a+r.a,p:acc.p+r.p,e:acc.e+r.e}),{a:0,p:0,e:0});
  const lvLbl = {taz:'אזורי תנועה',neigh:'שכונות',muni:'רשויות מקומיות',dist:'מחוזות'};
  const capSuffix = rows.length < allRows.length ? ` | מוצגים ${rows.length}/${allRows.length}` : '';
  const noDataSuffix = (noData||0) > 0 ? ` | ${noData} ללא נתונים` : '';
  document.getElementById('tbl-cap').textContent = S.yr+' | '+lvLbl[S.lv]+capSuffix+noDataSuffix;

  ['nm','p','a','e'].forEach(k=>{
    const el=document.getElementById('th-'+k);
    if(!el) return;
    const base=k==='nm'?'שם':k==='p'?'אוכלוסייה':k==='a'?'יח"ד':'מועסקים';
    const arr=S.sortK===k?(S.sortDir===-1?' ▾':' ▴'):'';
    el.textContent=base+arr;
    el.classList.toggle('sort-active', S.sortK===k);
  });

  document.getElementById('sum-tbody').innerHTML = rows.map(r=>{
    const dotClr = S.mode==='entity' ? (entityColors[r.nm]||DEFAULT_CLR) : (r.sec ? secClr(r.sec) : '');
    const dot = dotClr ? `<span class="td-dot" style="background:${dotClr}"></span>` : '';
    return `<tr onclick="zoomToName('${esc(r.nm)}')">
      <td class="td-name" title="${esc(r.nm)}">${dot}${esc(r.nm)}</td>
      <td>${fmt(r.p)}</td><td>${fmt(r.a)}</td><td>${fmt(r.e)}</td>
    </tr>`;
  }).join('');

  document.getElementById('sum-foot').innerHTML =
    `<td><strong>סה"כ</strong></td><td><strong>${fmt(tot.p)}</strong></td><td><strong>${fmt(tot.a)}</strong></td><td><strong>${fmt(tot.e)}</strong></td>`;
}

function sortTbl(k) {
  if (S.sortK===k) S.sortDir*=-1; else { S.sortK=k; S.sortDir=-1; }
  updateTable();
}

function zoomToName(nm) {
  if (!geoLayer) return;
  geoLayer.eachLayer(layer=>{
    const p=layer.feature.properties;
    const n=S.lv==='taz'?(p.nm||'TAZ '+p.taz):p.nm;
    if (n===nm) {
      map.fitBounds(layer.getBounds(),{maxZoom:14,padding:[20,20]});
      showDetail(p);
    }
  });
}

// ── DETAIL ────────────────────────────────────────────────────────────────
function getD(props,sc,yr) {
  if (S.lv==='taz') return props.d?.[sc]?.[yr]||{a:0,p:0,e:0};
  return AGG[S.lv]?.[props.nm]?.[sc]?.[yr]||{a:0,p:0,e:0};
}

function showDetail(props) {
  const nm=S.lv==='taz'?(props.nm||'TAZ '+props.taz):props.nm;
  const sc=S.sc;
  const d=getD(props,sc,S.yr);

  let h=`<div class="dp-title">${esc(nm)}</div>`;
  h+=`<div class="dp-sec"><span style="width:12px;height:12px;border-radius:50%;background:${secClr(props.sec)};display:inline-block;flex-shrink:0"></span>${secLbl(props.sec)}</div>`;
  h+=`<table class="dp-table">
    <tr><td>אוכלוסייה</td><td>${fmt(d.p)}</td></tr>
    <tr><td>יחידות דיור</td><td>${fmt(d.a)}</td></tr>
    <tr><td>מועסקים</td><td>${fmt(d.e)}</td></tr>
  </table>`;

  h+='<div class="dp-years">';
  YEARS.forEach(y=>{
    const dY=getD(props,sc,y); const v=dY[S.met];
    const act=y===S.yr?' dp-active':'';
    h+=`<div class="dp-yr${act}" onclick="setYear(${y})"><div>${y}</div><div class="dp-yr-val">${fmt(v)}</div></div>`;
  });
  h+='</div>';

  document.getElementById('detail-content').innerHTML=h;
  document.getElementById('detail-pane').style.display='block';
}

function closeDetail() { document.getElementById('detail-pane').style.display='none'; }

// ── CONTROLS ──────────────────────────────────────────────────────────────
function setScenario(sc) {
  S.sc=sc;
  ['jtmt','bau','iplan'].forEach(s=>document.getElementById('sc-'+s).classList.toggle('active',s===sc));
  updateMap();
}

function setYear(yr) {
  S.yr=yr;
  document.querySelectorAll('.yr-pill').forEach(el=>el.classList.toggle('active',+el.dataset.yr===yr));
  updateMap();
}

function setLevel(lv) {
  S.lv=lv;
  hiddenEntities.clear();
  ['taz','neigh','muni','dist'].forEach(l=>document.getElementById('lv-'+l).classList.toggle('active',l===lv));
  const entBtn = document.getElementById('md-entity');
  if (lv==='taz') {
    entBtn.style.display='none';
    if (S.mode==='entity') { S.mode='choro'; ['sector','entity','choro'].forEach(m=>{const e=document.getElementById('md-'+m);if(e)e.classList.toggle('active',m==='choro');}); }
  } else {
    entBtn.style.display='';
  }
  closeDetail();
  updateMap();
}

function setMode(md) {
  S.mode=md;
  ['sector','entity','choro'].forEach(m=>{
    const el=document.getElementById('md-'+m);
    if(el) el.classList.toggle('active',m===md);
  });
  updateMap();
}

function setMetric(mt) {
  S.met=mt;
  ['p','a','e'].forEach(m=>document.getElementById('mt-'+m).classList.toggle('active',m===mt));
  updateMap();
}

// ── UTILS ─────────────────────────────────────────────────────────────────
function fmt(n) { return n ? Math.round(n).toLocaleString('he-IL') : '0'; }
function esc(s) { return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

// ── SEARCH ────────────────────────────────────────────────────────────────
const SEARCH_IDX = [];
(function buildIdx(){
  GJ.muni.features.forEach(f=>{ if(f.properties.nm) SEARCH_IDX.push({nm:f.properties.nm, type:'muni',  tag:'יישוב'}); });
  GJ.neigh.features.forEach(f=>{ if(f.properties.nm) SEARCH_IDX.push({nm:f.properties.nm, type:'neigh', tag:'שכונה'}); });
  GJ.taz.features.forEach(f=>{ if(f.properties.nm) SEARCH_IDX.push({nm:f.properties.nm,  type:'taz',   tag:'אזור תנועה'}); });
})();

let _srchResults = [];

function doSearch(q) {
  const drop = document.getElementById(\'search-drop\');
  if (!q.trim()) { drop.style.display=\'none\'; return; }
  const lq = q.trim();
  _srchResults = SEARCH_IDX.filter(r => r.nm.includes(lq)).slice(0, 14);
  if (!_srchResults.length) { drop.style.display=\'none\'; return; }
  drop.innerHTML = _srchResults.map((r,i) =>
    \'<div class="srch-item" onmousedown="selectSearch(\'+i+\',event)">\' +
    \'<span class="srch-nm">\' + esc(r.nm) + \'</span>\' +
    \'<span class="srch-tag">\' + r.tag + \'</span></div>\'
  ).join(\'\');
  drop.style.display = \'block\';
}

function selectSearch(idx, e) {
  if (e) e.preventDefault();
  const item = _srchResults[idx];
  if (!item) return;
  document.getElementById(\'search-inp\').value = item.nm;
  document.getElementById(\'search-drop\').style.display = \'none\';
  // Set spatial filter from search result
  if (item.type === \'muni\') { S.sfMuni = item.nm; S.sfNeigh = ''; applyFilters(); }
  else if (item.type === \'neigh\') { S.sfNeigh = item.nm; applyFilters(); }
  if (S.lv !== item.type) setLevel(item.type);
  zoomToName(item.nm);
}

function hideSearchDrop() {
  const el = document.getElementById(\'search-drop\');
  if (el) el.style.display = \'none\';
}

// ── INIT ──────────────────────────────────────────────────────────────────
(function init(){
  const track=document.getElementById('year-track');
  YEARS.forEach(y=>{
    const b=document.createElement('button');
    b.className='yr-pill'+(y===S.yr?' active':'');
    b.dataset.yr=y; b.textContent=y;
    b.onclick=()=>setYear(y);
    track.appendChild(b);
  });
  // Populate municipality dropdown
  const muniSel = document.getElementById('sf-muni');
  if (muniSel) {
    const names = GJ.muni.features.map(f=>f.properties.nm).filter(Boolean).sort((a,b)=>a.localeCompare(b,'he'));
    names.forEach(nm=>{ const o=document.createElement('option'); o.value=nm; o.textContent=nm; muniSel.appendChild(o); });
  }
  // Populate neighborhood dropdown
  const neighSel = document.getElementById('sf-neigh');
  if (neighSel) {
    const names = GJ.neigh.features.map(f=>f.properties.nm).filter(Boolean).sort((a,b)=>a.localeCompare(b,'he'));
    names.forEach(nm=>{ const o=document.createElement('option'); o.value=nm; o.textContent=nm; neighSel.appendChild(o); });
  }
  updateMap();
})();
</script>
</body>
</html>'''

# Inject data
print('Injecting data...')
html = HTML_TEMPLATE
html = html.replace('__LEAFLET_CSS__', LEAFLET_CSS)
html = html.replace('__LEAFLET_JS__',  LEAFLET_JS)
html = html.replace('__TAZ_GJ__',   TAZ_JS)
html = html.replace('__NEIGH_GJ__', NEIGH_JS)
html = html.replace('__MUNI_GJ__',  MUNI_JS)
html = html.replace('__DIST_GJ__',  DIST_JS)
html = html.replace('__AGG__',      AGG_JS)

OUT.write_text(html, encoding='utf-8')
sz = OUT.stat().st_size / 1024
print(f'\nDone! {sz:.0f} KB written to:\n{OUT}')
