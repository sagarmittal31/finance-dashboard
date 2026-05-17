#!/usr/bin/env python3
"""
Finance Dashboard Updater
─────────────────────────
Usage:
  python update_dashboard.py                          # uses default file paths
  python update_dashboard.py --primary "path/to/Primary data.xlsx" \\
                              --secondary "path/to/Secondary data.csv"
  python update_dashboard.py --categories "path/to/Expense Categories.docx"
  python update_dashboard.py --no-push               # generate HTML only, skip git push

All arguments are optional — defaults point to the data/ folder next to this script.
"""

import argparse
import csv
import json
import os
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime

# ── Paths relative to this script ────────────────────────────────────────────
ROOT        = os.path.dirname(os.path.abspath(__file__))
DATA_DIR    = os.path.join(ROOT, "data")
DOCS_DIR    = os.path.join(ROOT, "docs")
OUTPUT_HTML = os.path.join(DOCS_DIR, "index.html")

DEFAULT_PRIMARY    = os.path.join(DATA_DIR, "Primary data.xlsx")
DEFAULT_SECONDARY  = os.path.join(DATA_DIR, "Secondary data.csv")
DEFAULT_CATEGORIES = os.path.join(DATA_DIR, "Expense Categories.docx")


# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY RULES
# To update: edit the keyword lists below to match your Expense Categories.docx
# ══════════════════════════════════════════════════════════════════════════════
def infer_category(desc):
    if not desc:
        return "General"
    d = desc.lower().strip()

    # Clinic expenses — all salaries, rents, clinic-specific items
    if any(x in d for x in [
        "clinic", "medibliss", "wix sub", "n8n", "google my business",
        "homoeo board", "homoeopathic board", "social media manag",
        "stationery", "salary", "rent ",
    ]):
        return "Clinic expenses"

    # Investments
    if any(x in d for x in [
        "invest", "sip", "lic", "mutual fund", "mf ", "gold",
        "paytm shares", "esop", "dholera invest", "mtf invest",
        "yeida", "term life", "tata sip", "edelweiss",
    ]):
        return "Investments"

    # Travel
    if any(x in d for x in [
        "trip", "flight", "hotel", "train ticket", "thailand", "bhutan",
        "kerala", "e-sim", "chandigarh train", "varkala", "udaipur",
        "coron", "honeymoon", "vacation", "flights to", "ahmedabad flight",
    ]):
        return "Travel"

    # Petrol / Fuel
    if "petrol" in d or "fuel" in d:
        return "Petrol/Fuel"

    # Transport
    if any(x in d for x in [
        "uber", "ola ", "rapido", " cab", "cab ", "taxi",
        "auto ", "metro ", "porter", "driver booking", "parking",
    ]):
        return "Transport"

    # Groceries
    if any(x in d for x in [
        "grocer", "blinkit", "zepto", "instamart",
        "home supplies", "rosier food", "miscellaneous grocer",
    ]):
        return "Groceries"

    # Gifts — birthday cakes, donations, cash (before Eating Out)
    if any(x in d for x in [
        "gift", "farewell", "secret santa", "bhai dooj", "rakhi",
        "donation", "cash to ", "birthday cake", "cake ",
    ]):
        return "Gifts"

    # Eating Out — party food/drinks included
    if any(x in d for x in [
        "eating out", "eating", "food order", "dinner", "lunch", "breakfast",
        "cafe", "coffee", "swiggy", "zomato", "pizza", "burger", "shawarma",
        "sushi", "kulcha", "noodle", "kulfi", "momos", "chaap", "dimsum",
        "chips", "ice cream", "icecream", "spring roll", "springroll",
        "coconut water", "blue tokai", "midnight food", "midnight eating",
        "midnight munching", "midnight taco", "midnight swiggy", "laphing midnight",
        "taco bell", "kfc", "mcd", "subway", "gol gappe", "snowberry", "amaltas",
        "sandoz", "anardana", "youmee", "comorin", "garnita", "echoes", "xero",
        "ohashi", "bomba", "andhra", "lopera", "dohful", "salad days",
        "theobroma", "alcohol", "whiskey", "chakhna", "snack",
        "kitty lunch", "kitty party", "office dinner", "birthday party",
        "party food", "party drink", "leo lunch", "leo's lunch",
        "vimugdha", "lunch outing", "dessert", "cafe parmesan",
        "date peach", "date buho", "date with", "dinner with",
        "lunch with", "dinner date", "lunch date",
    ]):
        return "Eating Out"

    # Health and Grooming — health insurance included
    if any(x in d for x in [
        "medicine", "medic", "therapy", "skincare", "skin care", "salon",
        "haircut", "hair cut", "protein powder", "protein supplement",
        "vitamin", "supplement", "ashwagandha", "threptin", "cosmetic",
        "makeup", "sunscreen", "laser treatment", "dermat", "serum",
        "conditioner", "pilgrim", "nykaa", "doctor", "ent",
        "physiotherapy", "grooming", "fascial", "mehendi", "gym",
        "swimming session", "swim class", "blood test", "xray", "x-ray",
        "health insurance", "insurance",
    ]):
        return "Health and Grooming"

    # Entertainment
    if any(x in d for x in [
        "movie", "netflix", "youtube", "apple music", "apple subscription",
        "apple service", "spotify", "concert", "misfits", "board game",
        "jamming", "gaming", "azul game", "sony liv", "twitter subscription",
        "amazon subscription", "amazon prime", "comedy show", "zakir khan",
        "dandiya", "ram leela", "dussehra entertainment", "kitty game",
        "misfits drawing", "bhajan", "entertainment", "sunidhi", "disney",
    ]):
        return "Entertainment"

    # Personal Growth
    if any(x in d for x in [
        "swimming class", "music class", "bachata", "dance class", "audible",
        "book", "webinar", "lecture", "course", "workshop",
        "happiness program", "aol", "tanmay sir", "claude sub",
    ]):
        return "Personal Growth"

    # Shopping — mobile devices, electronics, amazon orders
    if any(x in d for x in [
        "shopping", "clothes", "clothing", "shoes", "myntra", "snitch",
        "decathlon", "sarojini", "loom", "earring", "saree", "suit",
        "kurte", "slippers", "footwear", "eyewear", "spectacles", "sunglass",
        "phone cover", "swimwear", "swim costume", "swimming costume",
        "jeans", "trouser", "dress", "blissclub", "legging", "uniqlo",
        "underwear", "phone tempered", "power bank", "watch battery",
        "mobile device", "electronic", "amazon order", "amazon home",
        "h&m", "zara",
    ]):
        return "Shopping"

    # Home expenses — pooja items, kitchen accessories
    if any(x in d for x in [
        "pillow", "washing machine", "wardrobe", "organiser", "plant",
        "dry cleaning", "house cleaning", "urban company", "room decor",
        "planter", "plant stand", "lunch box", "flower seeds", "ganesh",
        "decoration", "zazza cleaning", "clothes ironing", "toiletri",
        "appliance", "kitchen", "pooja", "puja", "moorti", "room freshner",
    ]):
        return "Home expenses"

    # Utilities
    if any(x in d for x in [
        "mobile bill", "internet", "recharge", "electricity",
        "water bill", "tv/phone",
    ]):
        return "Utilities"

    return "General"


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
def load_primary(path):
    try:
        import openpyxl
    except ImportError:
        print("Installing openpyxl..."); subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "-q"])
        import openpyxl

    print(f"  Loading primary: {os.path.basename(path)}")
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows_raw = list(ws.iter_rows(values_only=True))[1:]

    rows, cutoff = [], None
    for r in rows_raw:
        date, desc, cost, cat, *_ = r
        if not date:
            continue
        if cutoff is None or date > cutoff:
            cutoff = date
        rows.append({
            "date": date, "desc": desc or "",
            "cost": int(cost) if cost else 0,
            "cat": cat or "General",
            "year": date.year, "month": date.month,
            "ym": f"{date.year}-{str(date.month).zfill(2)}",
        })

    print(f"    {len(rows)} rows loaded | cutoff: {cutoff.strftime('%d %b %Y')}")
    return rows, cutoff


def load_secondary(path, cutoff):
    cutoff_str = cutoff.strftime("%Y-%m-%d")
    print(f"  Loading secondary: {os.path.basename(path)} (rows after {cutoff_str})")

    with open(path, encoding="utf-8-sig") as f:
        raw = list(csv.reader(f))

    rows, skipped = [], 0
    for row in raw[1:]:
        if len(row) < 4:
            continue
        date_str, desc, _, cost_str = row[0], row[1], row[2], row[3]
        if not date_str or desc in ("Total balance", "", None):
            continue
        if date_str <= cutoff_str:
            skipped += 1
            continue
        try:
            date = datetime.strptime(date_str, "%Y-%m-%d")
            cost = round(float(cost_str))
        except (ValueError, TypeError):
            continue
        desc = desc.strip()
        rows.append({
            "date": date, "desc": desc,
            "cost": cost, "cat": infer_category(desc),
            "year": date.year, "month": date.month,
            "ym": f"{date.year}-{str(date.month).zfill(2)}",
        })

    print(f"    {len(rows)} rows loaded | {skipped} rows before cutoff skipped")
    return rows


def parse_categories_docx(path):
    """Parse and display categories from the docx rulebook."""
    if not os.path.exists(path):
        print("  Categories docx not found — using built-in rules")
        return
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/document.xml") as f:
                tree = ET.parse(f)
        root = tree.getroot()
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        lines = []
        for para in root.iter(f"{{{ns}}}p"):
            line = "".join(t.text or "" for t in para.iter(f"{{{ns}}}t")).strip()
            if line:
                lines.append(line)
        cats_found = [l for l in lines if " - " in l and not l.startswith("These")]
        print(f"  Categories docx parsed — {len(cats_found)} categories found:")
        for c in cats_found:
            print(f"    • {c.split(' - ')[0]}")
        print()
        print("  NOTE: If you've changed category rules in the docx,")
        print("  update the keyword lists in infer_category() in this script.")
    except Exception as e:
        print(f"  Warning: Could not parse categories docx ({e})")


# ══════════════════════════════════════════════════════════════════════════════
# HTML GENERATION
# ══════════════════════════════════════════════════════════════════════════════
def build_chart_data(all_rows):
    months = sorted(set(d["ym"] for d in all_rows))
    cats   = sorted(set(d["cat"] for d in all_rows))
    lifestyle_cats = [c for c in cats if c not in ("Investments", "Clinic expenses")]

    pivot = defaultdict(lambda: defaultdict(int))
    for d in all_rows:
        pivot[d["cat"]][d["ym"]] += d["cost"]

    monthly_total = defaultdict(int)
    for d in all_rows:
        monthly_total[d["ym"]] += d["cost"]

    yoy = defaultdict(lambda: defaultdict(int))
    for d in all_rows:
        yoy[d["year"]][d["cat"]] += d["cost"]

    last6 = months[-6:]; prev6 = months[-12:-6]
    trend = {}
    for cat in cats:
        l6 = sum(pivot[cat][m] for m in last6)
        p6 = sum(pivot[cat][m] for m in prev6)
        trend[cat] = round((l6 - p6) / p6 * 100, 1) if p6 > 0 else (100.0 if l6 > 0 else 0.0)

    # Serialisable raw rows (date as string)
    raw_serial = [{**r, "date": r["date"].strftime("%Y-%m-%d") if hasattr(r["date"], "strftime") else r["date"]} for r in all_rows]

    D = {
        "months": months, "cats": cats, "lifestyle_cats": lifestyle_cats,
        "monthly_total": dict(monthly_total),
        "pivot": {cat: dict(pivot[cat]) for cat in cats},
        "yoy": {str(y): dict(yoy[y]) for y in sorted(yoy.keys())},
        "trend": trend,
        "generated": datetime.now().strftime("%d %b %Y %H:%M"),
        "date_range": f"{months[0]} to {months[-1]}",
    }
    return D, raw_serial


def generate_html(D, raw_serial):
    data_json = json.dumps(D)
    raw_json  = json.dumps(raw_serial)
    gen_date  = D["generated"]
    date_range_label = f"{D['months'][0]} – {D['months'][-1]}"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Personal Finance Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{{--bg:#0f1117;--sur:#1a1d27;--sur2:#22263a;--brd:#2e3350;--acc:#6c63ff;--grn:#00d4aa;--red:#ff6b6b;--yel:#ffa94d;--txt:#e8eaf6;--t2:#8892b0;--t3:#4a5580}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:var(--bg);color:var(--txt);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}
nav{{background:var(--sur);border-bottom:1px solid var(--brd);padding:0 24px;display:flex;align-items:center;justify-content:space-between;height:56px;position:sticky;top:0;z-index:200}}
.brand{{font-size:16px;font-weight:700;color:var(--acc)}}.brand span{{color:var(--t2);font-size:11px;font-weight:400;margin-left:8px}}
.tabs{{display:flex;gap:3px}}
.tab{{background:none;border:none;color:var(--t2);padding:7px 14px;border-radius:7px;cursor:pointer;font-size:13px;font-weight:500;transition:all .2s}}
.tab:hover{{background:var(--sur2);color:var(--txt)}}.tab.active{{background:var(--acc);color:#fff}}
.gfbar{{background:var(--sur);border-bottom:1px solid var(--brd);padding:10px 24px;display:flex;align-items:center;gap:14px;flex-wrap:wrap;position:sticky;top:56px;z-index:190}}
.gfl{{font-size:11px;font-weight:600;color:var(--t2);text-transform:uppercase;letter-spacing:.7px}}
.gfbar select{{background:var(--sur2);border:1px solid var(--brd);color:var(--txt);padding:6px 11px;border-radius:7px;font-size:12px;cursor:pointer;outline:none}}
.gfbar select:focus{{border-color:var(--acc)}}
.qbtn{{background:none;border:1px solid var(--brd);color:var(--t2);padding:5px 11px;border-radius:7px;cursor:pointer;font-size:11px;font-weight:500;transition:all .2s}}
.qbtn:hover,.qbtn.active{{background:var(--acc);border-color:var(--acc);color:#fff}}
.gfsep{{width:1px;height:20px;background:var(--brd);margin:0 4px}}
.page{{display:none;padding:24px;max-width:1440px;margin:0 auto}}.page.active{{display:block}}
.fbar{{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap;align-items:center}}
.fbar select,.fbar input{{background:var(--sur);border:1px solid var(--brd);color:var(--txt);padding:7px 12px;border-radius:7px;font-size:12px;cursor:pointer;outline:none}}
.freset{{background:none;border:1px solid var(--brd);color:var(--t2);padding:7px 12px;border-radius:7px;cursor:pointer;font-size:12px;transition:all .2s}}
.freset:hover{{border-color:var(--acc);color:var(--acc)}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:14px;margin-bottom:22px}}
.kpi{{background:var(--sur);border:1px solid var(--brd);border-radius:13px;padding:18px;position:relative;overflow:hidden;transition:border-color .2s}}
.kpi:hover{{border-color:var(--acc)}}
.kpi::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:13px 13px 0 0}}
.kpi.b::before{{background:var(--acc)}}.kpi.g::before{{background:var(--grn)}}.kpi.r::before{{background:var(--red)}}.kpi.y::before{{background:var(--yel)}}.kpi.p::before{{background:#b48eff}}.kpi.t::before{{background:#66d9e8}}
.kl{{font-size:10px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.8px;margin-bottom:7px}}
.kv{{font-size:24px;font-weight:700;letter-spacing:-1px}}.ks{{font-size:11px;color:var(--t2);margin-top:3px}}
.badge{{display:inline-flex;align-items:center;font-size:10px;font-weight:700;padding:2px 7px;border-radius:20px;margin-top:5px}}
.badge.up{{background:rgba(255,107,107,.15);color:var(--red)}}.badge.dn{{background:rgba(0,212,170,.15);color:var(--grn)}}
.grid{{display:grid;gap:18px;margin-bottom:18px}}
.g2{{grid-template-columns:1fr 1fr}}.g3{{grid-template-columns:1fr 1fr 1fr}}.g73{{grid-template-columns:7fr 3fr}}.g64{{grid-template-columns:6fr 4fr}}
@media(max-width:860px){{.g2,.g3,.g73,.g64{{grid-template-columns:1fr}}}}
.card{{background:var(--sur);border:1px solid var(--brd);border-radius:13px;padding:20px}}
.ct{{font-size:13px;font-weight:600;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between;gap:8px;flex-wrap:wrap}}
.sub{{font-size:10px;color:var(--t2);font-weight:400}}.cw{{position:relative}}
.tbl{{width:100%;border-collapse:collapse;font-size:12px}}
.tbl th{{text-align:left;padding:8px 11px;font-size:10px;font-weight:700;color:var(--t2);text-transform:uppercase;letter-spacing:.6px;border-bottom:1px solid var(--brd);white-space:nowrap}}
.tbl td{{padding:8px 11px;border-bottom:1px solid var(--brd);vertical-align:middle}}
.tbl tr:last-child td{{border-bottom:none}}.tbl tr:hover td{{background:var(--sur2)}}
.pill{{display:inline-block;padding:2px 9px;border-radius:20px;font-size:10px;font-weight:600;white-space:nowrap}}
.amt{{font-variant-numeric:tabular-nums;font-weight:600}}
.hmwrap{{overflow:auto;max-height:520px;position:relative}}
.hmtbl{{border-collapse:collapse;font-size:11px;white-space:nowrap}}
.hmtbl th{{padding:5px 9px;color:var(--t2);font-weight:500;text-align:right;position:sticky;top:0;background:var(--sur);z-index:3}}
.hmtbl td{{padding:4px 8px;text-align:right;border-radius:4px;min-width:66px;cursor:pointer;transition:opacity .15s}}
.hmtbl td:hover{{opacity:.75;outline:1px solid #fff3}}
.hmc{{text-align:left!important;color:var(--txt);font-weight:500;padding-right:14px;position:sticky;left:0;background:var(--sur);z-index:2;box-shadow:2px 0 8px rgba(0,0,0,.4)}}
.hmtbl thead .hmc{{z-index:4}}
.trlist{{display:flex;flex-direction:column;gap:11px}}
.tritem{{display:flex;flex-direction:column;gap:3px}}
.trh{{display:flex;justify-content:space-between;font-size:12px}}
.trbg{{height:5px;background:var(--sur2);border-radius:3px;overflow:hidden}}
.trfill{{height:100%;border-radius:3px;transition:width .8s ease}}
.tgrp{{display:flex;gap:3px;background:var(--sur2);border-radius:7px;padding:3px}}
.tbtn{{background:none;border:none;color:var(--t2);padding:5px 12px;border-radius:5px;cursor:pointer;font-size:11px;font-weight:500;transition:all .2s;white-space:nowrap}}
.tbtn.active{{background:var(--acc);color:#fff}}
.mswrap{{position:relative}}
.msbtn{{background:var(--sur2);border:1px solid var(--brd);color:var(--txt);padding:7px 12px;border-radius:7px;cursor:pointer;font-size:12px;display:flex;align-items:center;gap:6px;white-space:nowrap;transition:border-color .2s;min-width:160px;justify-content:space-between}}
.msbtn:hover,.msbtn.open{{border-color:var(--acc)}}
.mspanel{{position:absolute;top:calc(100% + 4px);left:0;background:var(--sur);border:1px solid var(--brd);border-radius:10px;padding:6px;min-width:200px;max-height:300px;overflow-y:auto;z-index:100;box-shadow:0 8px 24px rgba(0,0,0,.5);display:none}}
.mspanel.open{{display:block}}
.msact{{display:flex;gap:6px;padding:4px 4px 8px;border-bottom:1px solid var(--brd);margin-bottom:4px}}
.msact button{{background:none;border:1px solid var(--brd);color:var(--t2);padding:3px 10px;border-radius:5px;cursor:pointer;font-size:11px;transition:all .2s}}
.msact button:hover{{border-color:var(--acc);color:var(--acc)}}
.mspanel label{{display:flex;align-items:center;gap:8px;padding:5px 7px;border-radius:6px;cursor:pointer;font-size:12px;white-space:nowrap}}
.mspanel label:hover{{background:var(--sur2)}}
.mspanel label input{{accent-color:var(--acc);width:13px;height:13px;cursor:pointer;flex-shrink:0}}
.igrid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:14px}}
.icard{{background:var(--sur);border:1px solid var(--brd);border-radius:13px;padding:18px;display:flex;gap:13px}}
.icard.warn{{border-color:rgba(255,165,0,.35);background:rgba(255,165,0,.04)}}
.icard.good{{border-color:rgba(0,212,170,.35);background:rgba(0,212,170,.04)}}
.icard.info{{border-color:rgba(108,99,255,.35);background:rgba(108,99,255,.04)}}
.icard.alert{{border-color:rgba(255,107,107,.35);background:rgba(255,107,107,.04)}}
.iico{{font-size:24px;flex-shrink:0;margin-top:1px}}
.ititle{{font-size:13px;font-weight:600;margin-bottom:5px}}
.itext{{font-size:12px;color:var(--t2);line-height:1.65}}
.modal-ov{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:500;align-items:center;justify-content:center}}
.modal-ov.open{{display:flex}}
.modal-box{{background:var(--sur);border:1px solid var(--brd);border-radius:15px;width:min(720px,96vw);max-height:82vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 24px 64px rgba(0,0,0,.6)}}
.modal-hd{{padding:16px 20px;border-bottom:1px solid var(--brd);display:flex;justify-content:space-between;align-items:center;flex-shrink:0}}
.modal-hd span{{font-weight:600;font-size:14px}}
.modal-hd button{{background:none;border:none;color:var(--t2);font-size:22px;cursor:pointer;line-height:1;padding:0 2px}}
.modal-hd button:hover{{color:var(--txt)}}
.modal-bd{{padding:16px 20px;overflow-y:auto}}
.modal-ttl{{font-size:12px;color:var(--t2);margin-bottom:11px}}
::-webkit-scrollbar{{width:5px;height:5px}}::-webkit-scrollbar-track{{background:var(--sur)}}
::-webkit-scrollbar-thumb{{background:var(--brd);border-radius:3px}}::-webkit-scrollbar-thumb:hover{{background:var(--acc)}}
canvas{{cursor:pointer}}
.gen-badge{{font-size:10px;color:var(--t3);padding:4px 10px;background:var(--sur2);border-radius:20px;border:1px solid var(--brd)}}
</style>
</head>
<body>
<nav>
  <div class="brand">💰 Finance Dashboard <span>{date_range_label}</span></div>
  <div style="display:flex;align-items:center;gap:12px">
    <span class="gen-badge">Updated {gen_date}</span>
    <div class="tabs">
      <button class="tab active" onclick="go('overview',this)">Overview</button>
      <button class="tab" onclick="go('trends',this)">Trends</button>
      <button class="tab" onclick="go('categories',this)">Categories</button>
      <button class="tab" onclick="go('transactions',this)">Transactions</button>
      <button class="tab" onclick="go('insights',this)">Insights</button>
    </div>
  </div>
</nav>
<div class="gfbar">
  <span class="gfl">Period:</span>
  <select id="gf-from" onchange="onGF()"></select>
  <span style="color:var(--t2);font-size:12px">→</span>
  <select id="gf-to" onchange="onGF()"></select>
  <div class="gfsep"></div>
  <button class="qbtn" onclick="setQ(3,this)">3M</button>
  <button class="qbtn" onclick="setQ(6,this)">6M</button>
  <button class="qbtn" onclick="setQ(12,this)">1Y</button>
  <button class="qbtn" onclick="setQ(24,this)">2Y</button>
  <button class="qbtn active" onclick="setQ(0,this)">All</button>
  <div class="gfsep"></div>
  <span id="gf-summary" style="font-size:11px;color:var(--t2)"></span>
</div>

<!-- OVERVIEW -->
<div id="pg-overview" class="page active">
  <div class="kpis" id="ov-kpis"></div>
  <div class="grid g73">
    <div class="card">
      <div class="ct">Monthly Spending
        <div class="tgrp">
          <button class="tbtn active" onclick="ovMode='stacked';sa(this);rnOverview()">Stacked</button>
          <button class="tbtn" onclick="ovMode='lifestyle';sa(this);rnOverview()">Lifestyle</button>
          <button class="tbtn" onclick="ovMode='total';sa(this);rnOverview()">Total</button>
        </div>
      </div>
      <div class="cw" style="height:270px"><canvas id="c-monthly"></canvas></div>
    </div>
    <div class="card"><div class="ct">Spend Split <span class="sub">click segment</span></div><div class="cw" style="height:270px"><canvas id="c-donut"></canvas></div></div>
  </div>
  <div class="grid g2">
    <div class="card"><div class="ct">Category Breakdown <span class="sub">click to drill down</span></div><div class="cw" style="height:290px"><canvas id="c-catbar"></canvas></div></div>
    <div class="card"><div class="ct">Biggest Expenses <span class="sub">excl. investments</span></div><div style="overflow-y:auto;max-height:290px"><table class="tbl" id="t-top"></table></div></div>
  </div>
</div>

<!-- TRENDS -->
<div id="pg-trends" class="page">
  <div class="fbar">
    <span style="font-size:12px;color:var(--t2);font-weight:600">Category:</span>
    <select id="tr-cat" onchange="rnTrend()"><option value="all">All Lifestyle</option></select>
    <div class="tgrp" style="margin-left:auto">
      <button class="tbtn active" onclick="trView='abs';sa(this);rnTrend()">Amount</button>
      <button class="tbtn" onclick="trView='pct';sa(this);rnTrend()">% of Monthly</button>
    </div>
  </div>
  <div class="grid"><div class="card"><div class="ct" id="tr-title">Monthly Trend <span class="sub">click a point</span></div><div class="cw" style="height:300px"><canvas id="c-trend"></canvas></div></div></div>
  <div class="grid g2">
    <div class="card"><div class="ct">Year-on-Year <span class="sub">click a bar</span></div><div class="cw" style="height:280px"><canvas id="c-yoy"></canvas></div></div>
    <div class="card"><div class="ct">Category Trends <span class="sub">last 6 vs prior 6 months</span></div><div id="tr-bars" class="trlist" style="max-height:280px;overflow-y:auto;padding-right:4px"></div></div>
  </div>
  <div class="grid"><div class="card"><div class="ct">Stacked Monthly Breakdown <span class="sub">click a segment</span></div><div class="cw" style="height:320px"><canvas id="c-stack"></canvas></div></div></div>
</div>

<!-- CATEGORIES -->
<div id="pg-categories" class="page">
  <div class="card" style="margin-bottom:18px"><div class="ct">Heatmap — Category × Month <span class="sub">click any cell</span></div><div class="hmwrap" id="hm"></div></div>
  <div class="grid g2">
    <div class="card"><div class="ct">Monthly Average by Category <span class="sub">click a bar</span></div><div class="cw" style="height:320px"><canvas id="c-avg"></canvas></div></div>
    <div class="card"><div class="ct">Category Summary</div><div style="overflow-y:auto;max-height:320px"><table class="tbl" id="t-catsum"></table></div></div>
  </div>
</div>

<!-- TRANSACTIONS -->
<div id="pg-transactions" class="page">
  <div class="fbar">
    <span style="font-size:12px;color:var(--t2);font-weight:600">Category:</span>
    <div class="mswrap" id="ms-cat">
      <button class="msbtn" id="ms-cat-btn" onclick="toggleMs('ms-cat')"><span id="ms-cat-lbl">All Categories</span><span>▾</span></button>
      <div class="mspanel" id="ms-cat-panel">
        <div class="msact"><button onclick="msAll('ms-cat',true)">Select All</button><button onclick="msAll('ms-cat',false)">Clear</button></div>
        <div id="ms-cat-list"></div>
      </div>
    </div>
    <span style="font-size:12px;color:var(--t2);font-weight:600">Year:</span>
    <select id="tx-yr" onchange="rnTx()"><option value="all">All</option></select>
    <span style="font-size:12px;color:var(--t2);font-weight:600">Month:</span>
    <select id="tx-mo" onchange="rnTx()"><option value="all">All</option></select>
    <input type="text" id="tx-q" placeholder="Search…" oninput="rnTx()" style="min-width:170px">
    <button class="freset" onclick="resetTx()">Reset</button>
    <span id="tx-info" style="font-size:11px;color:var(--t2);margin-left:auto"></span>
  </div>
  <div class="card"><div style="overflow-y:auto;max-height:570px"><table class="tbl" id="t-tx"></table></div></div>
</div>

<!-- INSIGHTS -->
<div id="pg-insights" class="page">
  <div id="i-grid" class="igrid"></div>
  <div class="grid g2" style="margin-top:20px">
    <div class="card"><div class="ct">Lifestyle % of Total <span class="sub">click a point</span></div><div class="cw" style="height:260px"><canvas id="c-lifepct"></canvas></div></div>
    <div class="card"><div class="ct">Investment Consistency <span class="sub">click a bar</span></div><div class="cw" style="height:260px"><canvas id="c-invconsist"></canvas></div></div>
  </div>
</div>

<!-- MODAL -->
<div class="modal-ov" id="modal" onclick="closeModal()">
  <div class="modal-box" onclick="event.stopPropagation()">
    <div class="modal-hd"><span id="modal-title"></span><button onclick="closeModal()">×</button></div>
    <div class="modal-bd"><div class="modal-ttl" id="modal-sub"></div><table class="tbl" id="modal-tbl"></table></div>
  </div>
</div>

<script>
const D={data_json};
const RAW={raw_json};
const MONTHS=D.months,CATS=D.cats,PIVOT=D.pivot,YOY=D.yoy,TREND=D.trend,MTTL=D.monthly_total,LCATS=D.lifestyle_cats;
const CC={{'Investments':'#6c63ff','Clinic expenses':'#ff6b9d','General':'#8892b0','Travel':'#00d4aa','Gifts':'#ff6b6b','Shopping':'#ffa94d','Health and Grooming':'#69db7c','Eating Out':'#f06595','Petrol/Fuel':'#a9e34b','SCPH':'#74c0fc','Personal Growth':'#b197fc','Entertainment':'#f59f00','Home expenses':'#66d9e8','Groceries':'#2eb872','Transport':'#748ffc','Wedding':'#ff8787','Utilities':'#868e96','Loan':'#495057'}};
function cc(c){{return CC[c]||'#8892b0'}}
function fi(n){{return'₹'+Math.round(n).toLocaleString('en-IN')}}
function fs(n){{return n>=1000000?'₹'+(n/100000).toFixed(1)+'L':n>=1000?'₹'+(n/1000).toFixed(0)+'K':'₹'+Math.round(n)}}
function ml(ym){{const[y,m]=ym.split('-');return['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m]+'\\''+(y.slice(2))}}
function bucket(cat){{return cat==='Investments'?'inv':cat==='Clinic expenses'?'clinic':'life'}}
const CD={{color:'#8892b0',plugins:{{legend:{{labels:{{color:'#8892b0',font:{{size:11}}}}}},datalabels:{{display:false}}}},scales:{{x:{{ticks:{{color:'#8892b0',font:{{size:10}}}},grid:{{color:'#1e2235'}},border:{{color:'#2e3350'}}}},y:{{ticks:{{color:'#8892b0',font:{{size:10}},callback:v=>'₹'+(v>=100000?(v/100000).toFixed(1)+'L':v>=1000?(v/1000).toFixed(0)+'K':v)}},grid:{{color:'#1e2235'}},border:{{color:'#2e3350'}}}}}}}};
const CH={{}};
function dc(id){{if(CH[id]){{CH[id].destroy();delete CH[id]}}}}
function sa(btn){{const p=btn.closest('.tgrp');p.querySelectorAll('.tbtn').forEach(b=>b.classList.remove('active'));btn.classList.add('active')}}
function hover(e,els){{e.native.target.style.cursor=els[0]?'pointer':'default'}}
let gFrom=MONTHS[0],gTo=MONTHS[MONTHS.length-1],curPage='overview',ovMode='stacked',trView='abs';
function gMs(){{return MONTHS.filter(m=>m>=gFrom&&m<=gTo)}}
function gRows(){{return RAW.filter(r=>r.ym>=gFrom&&r.ym<=gTo)}}
function lifeMonthly(m){{return(MTTL[m]||0)-(PIVOT['Investments']?.[m]||0)-(PIVOT['Clinic expenses']?.[m]||0)}}

function initGF(){{
  const f=document.getElementById('gf-from'),t=document.getElementById('gf-to');
  MONTHS.forEach(m=>{{const l=ml(m);f.innerHTML+=`<option value="${{m}}">${{l}}</option>`;t.innerHTML+=`<option value="${{m}}">${{l}}</option>`;}});
  t.value=MONTHS[MONTHS.length-1];updGFSummary();
}}
function onGF(){{
  gFrom=document.getElementById('gf-from').value;gTo=document.getElementById('gf-to').value;
  if(gFrom>gTo){{gTo=gFrom;document.getElementById('gf-to').value=gFrom}}
  document.querySelectorAll('.qbtn').forEach(b=>b.classList.remove('active'));
  updGFSummary();rerender();
}}
function setQ(n,btn){{
  document.querySelectorAll('.qbtn').forEach(b=>b.classList.remove('active'));btn.classList.add('active');
  if(n===0){{gFrom=MONTHS[0];gTo=MONTHS[MONTHS.length-1];}}
  else{{gTo=MONTHS[MONTHS.length-1];gFrom=MONTHS[Math.max(0,MONTHS.length-n)];}}
  document.getElementById('gf-from').value=gFrom;document.getElementById('gf-to').value=gTo;
  updGFSummary();rerender();
}}
function updGFSummary(){{
  const ms=gMs(),rows=gRows(),total=rows.reduce((s,r)=>s+r.cost,0);
  document.getElementById('gf-summary').textContent=`${{ms.length}} months · ${{rows.length}} transactions · ${{fs(total)}} total`;
}}
function rerender(){{
  if(curPage==='overview') rnOverview();
  else if(curPage==='trends') rnTrend();
  else if(curPage==='categories') rnCats();
  else if(curPage==='transactions') rnTx();
  else if(curPage==='insights') rnInsights();
}}

function showModal(title,rows,sub){{
  document.getElementById('modal-title').textContent=title;
  const top=rows.slice(0,20),total=rows.reduce((s,r)=>s+r.cost,0);
  document.getElementById('modal-sub').textContent=(sub||'')+`${{rows.length}} transactions · ${{fi(total)}} total`+(rows.length>20?' (top 20 shown)':'');
  document.getElementById('modal-tbl').innerHTML='<thead><tr><th>Date</th><th>Description</th><th>Category</th><th style="text-align:right">Amount</th></tr></thead><tbody>'+
    top.map(r=>`<tr><td style="color:var(--t2);white-space:nowrap">${{r.date}}</td><td style="max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.desc}}">${{r.desc}}</td><td><span class="pill" style="background:${{cc(r.cat)}}22;color:${{cc(r.cat)}}">${{r.cat}}</span></td><td class="amt" style="text-align:right">${{fi(r.cost)}}</td></tr>`).join('')+'</tbody>';
  document.getElementById('modal').classList.add('open');
}}
function closeModal(){{document.getElementById('modal').classList.remove('open')}}
document.addEventListener('keydown',e=>{{if(e.key==='Escape')closeModal()}});

let msState={{}};
function initMs(){{
  const list=document.getElementById('ms-cat-list');
  CATS.forEach(c=>{{msState[c]=true;list.innerHTML+=`<label><input type="checkbox" value="${{c}}" checked onchange="msCh()"><span class="pill" style="background:${{cc(c)}}22;color:${{cc(c)}}">${{c}}</span></label>`;}});
}}
function toggleMs(id){{
  const panel=document.getElementById(id+'-panel'),btn=document.getElementById(id+'-btn');
  const isOpen=panel.classList.contains('open');panel.classList.toggle('open',!isOpen);btn.classList.toggle('open',!isOpen);
}}
function msAll(id,val){{document.querySelectorAll('#'+id+'-panel input').forEach(cb=>{{cb.checked=val;msState[cb.value]=val}});updMsLbl();rnTx();}}
function msCh(){{document.querySelectorAll('#ms-cat-panel input').forEach(cb=>msState[cb.value]=cb.checked);updMsLbl();rnTx();}}
function updMsLbl(){{const sel=Object.values(msState).filter(Boolean).length;document.getElementById('ms-cat-lbl').textContent=sel===CATS.length?'All Categories':sel===0?'None':sel+' selected';}}
function getMsSel(){{return CATS.filter(c=>msState[c])}}
document.addEventListener('click',e=>{{if(!e.target.closest('#ms-cat')){{document.getElementById('ms-cat-panel').classList.remove('open');document.getElementById('ms-cat-btn').classList.remove('open');}}}});

function rnOverview(){{
  const ms=gMs(),rows=gRows();
  const total=rows.reduce((s,r)=>s+r.cost,0);
  const inv=rows.filter(r=>r.cat==='Investments').reduce((s,r)=>s+r.cost,0);
  const cli=rows.filter(r=>r.cat==='Clinic expenses').reduce((s,r)=>s+r.cost,0);
  const life=total-inv-cli;
  const totalAvg=Math.round(total/ms.length),invAvg=Math.round(inv/ms.length),cliAvg=Math.round(cli/ms.length),lifeAvg=Math.round(life/ms.length);
  const catTtl={{}};rows.forEach(r=>{{catTtl[r.cat]=(catTtl[r.cat]||0)+r.cost}});
  const topLifeCat=LCATS.map(c=>{{return{{c,v:catTtl[c]||0}}}}).sort((a,b)=>b.v-a.v)[0]||{{c:'—',v:0}};
  document.getElementById('ov-kpis').innerHTML=`
    <div class="kpi b"><div class="kl">Total Spent</div><div class="kv">${{fs(total)}}</div><div class="ks">${{rows.length}} txns · ${{ms.length}} months</div></div>
    <div class="kpi b"><div class="kl">Overall Avg/Month</div><div class="kv">${{fs(totalAvg)}}</div><div class="ks">all spend per month</div></div>
    <div class="kpi g"><div class="kl">Investments</div><div class="kv">${{fs(inv)}}</div><div class="ks">${{total?Math.round(inv/total*100):0}}% of total</div></div>
    <div class="kpi g"><div class="kl">Investments Avg/Month</div><div class="kv">${{fs(invAvg)}}</div><div class="ks">per month</div></div>
    <div class="kpi t"><div class="kl">Clinic Expenses</div><div class="kv">${{fs(cli)}}</div><div class="ks">${{total?Math.round(cli/total*100):0}}% of total</div></div>
    <div class="kpi t"><div class="kl">Clinic Avg/Month</div><div class="kv">${{fs(cliAvg)}}</div><div class="ks">per month</div></div>
    <div class="kpi r"><div class="kl">Lifestyle Spend</div><div class="kv">${{fs(life)}}</div><div class="ks">${{total?Math.round(life/total*100):0}}% of total</div></div>
    <div class="kpi r"><div class="kl">Lifestyle Avg/Month</div><div class="kv">${{fs(lifeAvg)}}</div><div class="ks">per month</div></div>
    <div class="kpi p"><div class="kl">Top Lifestyle Cat.</div><div class="kv" style="font-size:16px">${{topLifeCat.c}}</div><div class="ks">${{fs(topLifeCat.v)}}</div></div>`;
  const invVals=ms.map(m=>PIVOT['Investments']?.[m]||0);
  const cliVals=ms.map(m=>PIVOT['Clinic expenses']?.[m]||0);
  const lifeVals=ms.map(m=>{{const t=MTTL[m]||0;return t-(PIVOT['Investments']?.[m]||0)-(PIVOT['Clinic expenses']?.[m]||0)}});
  const totalVals=ms.map(m=>MTTL[m]||0);
  const lAvg=Math.round(lifeVals.reduce((a,b)=>a+b,0)/lifeVals.length);
  let datasets;
  if(ovMode==='stacked'){{
    datasets=[
      {{label:'Lifestyle',data:lifeVals,backgroundColor:'rgba(108,99,255,.7)',borderColor:'#6c63ff',borderWidth:1,borderRadius:0,stack:'s'}},
      {{label:'Clinic',data:cliVals,backgroundColor:'rgba(255,107,157,.7)',borderColor:'#ff6b9d',borderWidth:1,borderRadius:0,stack:'s'}},
      {{label:'Investments',data:invVals,backgroundColor:'rgba(0,212,170,.7)',borderColor:'#00d4aa',borderWidth:1,borderRadius:0,stack:'s'}},
    ];
  }} else if(ovMode==='lifestyle'){{
    datasets=[
      {{label:'Lifestyle',data:lifeVals,backgroundColor:lifeVals.map(v=>v>lAvg*1.4?'rgba(255,107,107,.7)':'rgba(108,99,255,.65)'),borderColor:lifeVals.map(v=>v>lAvg*1.4?'#ff6b6b':'#6c63ff'),borderWidth:1,borderRadius:4}},
      {{label:'Avg',data:ms.map(()=>lAvg),type:'line',borderColor:'#ffa94d',borderDash:[5,4],borderWidth:1.5,pointRadius:0,tension:.4}},
    ];
  }} else {{
    datasets=[{{label:'Total',data:totalVals,backgroundColor:'rgba(108,99,255,.65)',borderColor:'#6c63ff',borderWidth:1,borderRadius:4}}];
  }}
  dc('monthly');
  CH['monthly']=new Chart(document.getElementById('c-monthly'),{{type:'bar',data:{{labels:ms.map(ml),datasets}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,legend:{{labels:{{color:'#8892b0',font:{{size:11}}}}}},tooltip:{{callbacks:{{label:ctx=>ctx.dataset.label+': '+fi(ctx.raw)}}}}}},
      scales:{{...CD.scales,x:{{...CD.scales.x,stacked:ovMode==='stacked'}},y:{{...CD.scales.y,stacked:ovMode==='stacked'}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const m=ms[els[0].index];const ds=els[0].datasetIndex;let cf=null;if(ovMode==='stacked')cf=[null,'Clinic expenses','Investments'][ds];const r=gRows().filter(x=>x.ym===m&&(cf?x.cat===cf:true)).sort((a,b)=>b.cost-a.cost);showModal(ml(m)+(cf?' · '+cf:''),r);}}
    }}
  }});
  const se=Object.entries(catTtl).sort((a,b)=>b[1]-a[1]);
  dc('donut');
  CH['donut']=new Chart(document.getElementById('c-donut'),{{type:'doughnut',
    data:{{labels:['Investments','Clinic','Lifestyle'],datasets:[{{data:[inv,cli,life],backgroundColor:['#6c63ff','#ff6b9d','#00d4aa'],borderWidth:2,borderColor:'#1a1d27'}}]}},
    options:{{responsive:true,maintainAspectRatio:false,cutout:'62%',onHover:hover,
      plugins:{{legend:{{display:true,position:'bottom',labels:{{color:'#8892b0',font:{{size:11}},padding:12}}}},datalabels:{{display:false}},
        tooltip:{{callbacks:{{label:ctx=>' '+ctx.label+': '+fi(ctx.raw)+' ('+Math.round(ctx.raw/total*100)+'%)'
      }}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const lbls=['Investments','Clinic expenses','Lifestyle'];const lbl=lbls[els[0].index];let r;if(lbl==='Investments')r=rows.filter(x=>x.cat==='Investments');else if(lbl==='Clinic expenses')r=rows.filter(x=>x.cat==='Clinic expenses');else r=rows.filter(x=>bucket(x.cat)==='life');showModal(lbl,r.sort((a,b)=>b.cost-a.cost));}}
    }}
  }});
  const se2=se.slice(0,14);
  dc('catbar');
  CH['catbar']=new Chart(document.getElementById('c-catbar'),{{type:'bar',
    data:{{labels:se2.map(([c])=>c),datasets:[{{data:se2.map(([,v])=>v),backgroundColor:se2.map(([c])=>cc(c)+'bb'),borderColor:se2.map(([c])=>cc(c)),borderWidth:1,borderRadius:4}}]}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,indexAxis:'y',onHover:hover,
      plugins:{{...CD.plugins,legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>fi(ctx.raw)}}}}}},
      scales:{{...CD.scales,x:{{...CD.scales.x,ticks:{{...CD.scales.x.ticks,callback:v=>fs(v)}}}},y:{{ticks:{{color:'#8892b0',font:{{size:10}}}},grid:{{display:false}},border:{{color:'#2e3350'}}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const cat=se2[els[0].index][0];showModal(cat,rows.filter(r=>r.cat===cat).sort((a,b)=>b.cost-a.cost));}}
    }}
  }});
  const top=rows.filter(r=>r.cat!=='Investments').sort((a,b)=>b.cost-a.cost).slice(0,12);
  document.getElementById('t-top').innerHTML='<thead><tr><th>Date</th><th>Description</th><th>Category</th><th style="text-align:right">Amount</th></tr></thead><tbody>'+
    top.map(r=>`<tr><td style="color:var(--t2);white-space:nowrap">${{r.date}}</td><td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.desc}}">${{r.desc}}</td><td><span class="pill" style="background:${{cc(r.cat)}}22;color:${{cc(r.cat)}}">${{r.cat}}</span></td><td class="amt" style="text-align:right;color:var(--red)">${{fi(r.cost)}}</td></tr>`).join('')+'</tbody>';
}}

function rnTrend(){{
  const ms=gMs(),rows=gRows();
  const cat=document.getElementById('tr-cat').value;
  let vals=cat==='all'?ms.map(m=>lifeMonthly(m)):ms.map(m=>PIVOT[cat]?.[m]||0);
  if(trView==='pct') vals=vals.map((v,i)=>MTTL[ms[i]]?+(v/MTTL[ms[i]]*100).toFixed(1):0);
  const avg=vals.reduce((a,b)=>a+b,0)/vals.length;
  document.getElementById('tr-title').textContent=(cat==='all'?'Lifestyle Spend':cat)+(trView==='pct'?' (% of Monthly)':'');
  dc('trend');
  CH['trend']=new Chart(document.getElementById('c-trend'),{{type:'line',
    data:{{labels:ms.map(ml),datasets:[
      {{label:cat==='all'?'Lifestyle':cat,data:vals,borderColor:cat==='all'?'#6c63ff':cc(cat),backgroundColor:(cat==='all'?'#6c63ff':cc(cat))+'20',fill:true,tension:.35,pointRadius:3,pointHoverRadius:6,borderWidth:2}},
      {{label:'Average',data:ms.map(()=>avg),borderColor:'#ffa94d',borderDash:[5,4],borderWidth:1.5,pointRadius:0,tension:0}}
    ]}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,tooltip:{{callbacks:{{label:ctx=>trView==='pct'?ctx.raw+'%':fi(ctx.raw)}}}}}},
      scales:trView==='pct'?{{...CD.scales,y:{{...CD.scales.y,ticks:{{...CD.scales.y.ticks,callback:v=>v+'%'}}}}}}:CD.scales,
      onClick:(evt,els)=>{{if(!els.length)return;const m=ms[els[0].index];const r=gRows().filter(x=>x.ym===m&&(cat==='all'?bucket(x.cat)==='life':x.cat===cat)).sort((a,b)=>b.cost-a.cost);showModal((cat==='all'?'Lifestyle':cat)+' · '+ml(m),r);}}
    }}
  }});
  const yrs=Object.keys(YOY).sort();
  const lc=LCATS.filter(c=>yrs.some(y=>YOY[y][c])).sort((a,b)=>(YOY[yrs[yrs.length-1]][b]||0)-(YOY[yrs[yrs.length-1]][a]||0)).slice(0,10);
  dc('yoy');
  CH['yoy']=new Chart(document.getElementById('c-yoy'),{{type:'bar',
    data:{{labels:lc,datasets:yrs.map((y,i)=>{{return{{label:y,data:lc.map(c=>YOY[y][c]||0),backgroundColor:['rgba(108,99,255,.7)','rgba(0,212,170,.7)','rgba(255,107,107,.7)','rgba(255,169,77,.7)'][i],borderWidth:0,borderRadius:3}}}})}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,tooltip:{{callbacks:{{label:ctx=>fi(ctx.raw)}}}}}},
      scales:{{...CD.scales,x:{{...CD.scales.x,ticks:{{...CD.scales.x.ticks,maxRotation:30,font:{{size:9}}}}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const cat=lc[els[0].index],yr=yrs[els[0].datasetIndex];const r=gRows().filter(x=>x.cat===cat&&x.year==yr).sort((a,b)=>b.cost-a.cost);showModal(cat+' · '+yr,r);}}
    }}
  }});
  const st=Object.entries(TREND).sort((a,b)=>Math.abs(b[1])-Math.abs(a[1])).slice(0,15);
  const mx=Math.max(...st.map(([,v])=>Math.abs(v)));
  document.getElementById('tr-bars').innerHTML=st.map(([c,p])=>{{const col=p>0?'#ff6b6b':'#00d4aa';return`<div class="tritem"><div class="trh"><span>${{c}}</span><span style="color:${{col}};font-weight:700">${{p>0?'↑':'↓'}}${{Math.abs(p)}}%</span></div><div class="trbg"><div class="trfill" style="width:${{Math.round(Math.abs(p)/mx*100)}}%;background:${{col}}"></div></div></div>`}}).join('');
  const sc=CATS.filter(c=>ms.some(m=>PIVOT[c]?.[m]>0));
  dc('stack');
  CH['stack']=new Chart(document.getElementById('c-stack'),{{type:'bar',
    data:{{labels:ms.map(ml),datasets:sc.map(c=>{{return{{label:c,data:ms.map(m=>PIVOT[c]?.[m]||0),backgroundColor:cc(c)+'cc',borderWidth:0,stack:'s'}}}})}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,tooltip:{{callbacks:{{label:ctx=>`${{ctx.dataset.label}}: ${{fi(ctx.raw)}}`}}}}}},
      scales:{{...CD.scales,x:{{...CD.scales.x,stacked:true}},y:{{...CD.scales.y,stacked:true}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const m=ms[els[0].index],c=sc[els[0].datasetIndex];const r=gRows().filter(x=>x.ym===m&&x.cat===c).sort((a,b)=>b.cost-a.cost);showModal(c+' · '+ml(m),r);}}
    }}
  }});
}}

function rnCats(){{
  const ms=gMs(),rows=gRows(),tot=rows.reduce((s,r)=>s+r.cost,0);
  const hc=CATS.filter(c=>ms.some(m=>PIVOT[c]?.[m]>0));
  const mx=Math.max(...hc.flatMap(c=>ms.map(m=>PIVOT[c]?.[m]||0)));
  function hcol(v){{if(!v)return'transparent';const i=Math.sqrt(v/mx);return i<0.2?`rgba(108,99,255,${{(i*4).toFixed(2)}})`:i<0.55?`rgba(255,169,77,${{(0.3+i*0.6).toFixed(2)}})`:`rgba(0,212,170,${{(0.4+i*0.6).toFixed(2)}})`}}
  document.getElementById('hm').innerHTML='<table class="hmtbl"><thead><tr><th class="hmc">Category</th>'+ms.map(m=>'<th>'+ml(m)+'</th>').join('')+'<th style="position:sticky;right:0;background:var(--sur);z-index:2">Total</th></tr></thead><tbody>'+
    hc.map(c=>{{const rt=ms.reduce((s,m)=>s+(PIVOT[c]?.[m]||0),0);if(!rt)return'';return`<tr><td class="hmc">${{c}}</td>${{ms.map(m=>{{const v=PIVOT[c]?.[m]||0;return`<td style="background:${{hcol(v)}};color:${{v>mx*0.3?'#fff':'#8892b0'}}" title="${{c}} · ${{ml(m)}}: ${{fi(v)}}" onclick="hmClick('${{c}}','${{m}}')">` +( v?fs(v):'')+`</td>`}}).join('')}}<td style="color:var(--txt);font-weight:600;position:sticky;right:0;background:var(--sur)">${{fs(rt)}}</td></tr>`}}).join('')+
    '<tr style="border-top:2px solid var(--brd)"><td class="hmc" style="font-weight:700">Total</td>'+ms.map(m=>`<td style="color:var(--t2);font-weight:600" onclick="hmClick(null,'${{m}}')">${{fs(MTTL[m]||0)}}</td>`).join('')+`<td style="color:var(--acc);font-weight:700;position:sticky;right:0;background:var(--sur)">${{fs(ms.reduce((s,m)=>s+(MTTL[m]||0),0))}}</td></tr></tbody></table>`;
  const ad=CATS.map(c=>{{return{{cat:c,avg:Math.round(ms.reduce((s,m)=>s+(PIVOT[c]?.[m]||0),0)/ms.length)}}}}).filter(x=>x.avg>0).sort((a,b)=>b.avg-a.avg);
  dc('avg');
  CH['avg']=new Chart(document.getElementById('c-avg'),{{type:'bar',
    data:{{labels:ad.map(x=>x.cat),datasets:[{{data:ad.map(x=>x.avg),backgroundColor:ad.map(x=>cc(x.cat)+'bb'),borderColor:ad.map(x=>cc(x.cat)),borderWidth:1,borderRadius:4}}]}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,indexAxis:'y',onHover:hover,
      plugins:{{...CD.plugins,legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>fi(ctx.raw)+'/month avg'}}}}}},
      scales:{{...CD.scales,x:{{...CD.scales.x,ticks:{{...CD.scales.x.ticks,callback:v=>fs(v)}}}},y:{{ticks:{{color:'#8892b0',font:{{size:10}}}},grid:{{display:false}},border:{{color:'#2e3350'}}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const cat=ad[els[0].index].cat;showModal(cat,rows.filter(r=>r.cat===cat).sort((a,b)=>b.cost-a.cost));}}
    }}
  }});
  const cd=CATS.map(c=>{{const t=ms.reduce((s,m)=>s+(PIVOT[c]?.[m]||0),0);return{{cat:c,t,avg:Math.round(t/ms.length),pct:tot?Math.round(t/tot*100):0}}}}).filter(x=>x.t>0).sort((a,b)=>b.t-a.t);
  document.getElementById('t-catsum').innerHTML='<thead><tr><th>Category</th><th style="text-align:right">Total</th><th style="text-align:right">Avg/mo</th><th style="text-align:right">Share</th></tr></thead><tbody>'+
    cd.map(x=>`<tr><td><span class="pill" style="background:${{cc(x.cat)}}22;color:${{cc(x.cat)}}">${{x.cat}}</span></td><td class="amt" style="text-align:right">${{fi(x.t)}}</td><td class="amt" style="text-align:right;color:var(--t2)">${{fi(x.avg)}}</td><td style="text-align:right"><div style="display:flex;align-items:center;justify-content:flex-end;gap:6px"><div style="width:50px;height:4px;background:var(--sur2);border-radius:2px"><div style="width:${{x.pct}}%;height:100%;background:${{cc(x.cat)}};border-radius:2px"></div></div><span style="font-size:10px;color:var(--t2)">${{x.pct}}%</span></div></td></tr>`).join('')+'</tbody>';
}}
function hmClick(cat,m){{const rows=gRows().filter(r=>(cat?r.cat===cat:true)&&r.ym===m).sort((a,b)=>b.cost-a.cost);showModal((cat||'All')+' · '+ml(m),rows);}}

function rnTx(){{
  const sel=getMsSel(),yr=document.getElementById('tx-yr').value,mo=document.getElementById('tx-mo').value,q=document.getElementById('tx-q').value.toLowerCase();
  let rows=gRows();
  rows=rows.filter(r=>sel.includes(r.cat));
  if(yr!=='all') rows=rows.filter(r=>r.year==yr);
  if(mo!=='all') rows=rows.filter(r=>r.month==mo);
  if(q) rows=rows.filter(r=>r.desc.toLowerCase().includes(q)||r.cat.toLowerCase().includes(q));
  rows.sort((a,b)=>b.date.localeCompare(a.date));
  document.getElementById('tx-info').textContent=rows.length+' transactions · '+fi(rows.reduce((s,r)=>s+r.cost,0));
  document.getElementById('t-tx').innerHTML='<thead><tr><th>Date</th><th>Description</th><th>Category</th><th style="text-align:right">Amount</th></tr></thead><tbody>'+
    rows.map(r=>`<tr><td style="color:var(--t2);white-space:nowrap">${{r.date}}</td><td style="max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.desc}}">${{r.desc}}</td><td><span class="pill" style="background:${{cc(r.cat)}}22;color:${{cc(r.cat)}}">${{r.cat}}</span></td><td class="amt" style="text-align:right">${{fi(r.cost)}}</td></tr>`).join('')+'</tbody>';
}}
function resetTx(){{msAll('ms-cat',true);['tx-yr','tx-mo'].forEach(id=>document.getElementById(id).value='all');document.getElementById('tx-q').value='';rnTx();}}

function rnInsights(){{
  const ms=gMs(),rows=gRows(),tot=rows.reduce((s,r)=>s+r.cost,0);
  const inv=rows.filter(r=>r.cat==='Investments').reduce((s,r)=>s+r.cost,0);
  const cli=rows.filter(r=>r.cat==='Clinic expenses').reduce((s,r)=>s+r.cost,0);
  const life=tot-inv-cli,invPct=tot?Math.round(inv/tot*100):0,lifePct=tot?Math.round(life/tot*100):0;
  const eo=rows.filter(r=>r.cat==='Eating Out').reduce((s,r)=>s+r.cost,0),eoAvg=Math.round(eo/ms.length),eoTr=TREND['Eating Out']||0;
  const travel=rows.filter(r=>r.cat==='Travel').reduce((s,r)=>s+r.cost,0);
  const gifts=rows.filter(r=>r.cat==='Gifts').reduce((s,r)=>s+r.cost,0);
  const shop=rows.filter(r=>r.cat==='Shopping').reduce((s,r)=>s+r.cost,0);
  const pg=rows.filter(r=>r.cat==='Personal Growth').reduce((s,r)=>s+r.cost,0),pgAvg=Math.round(pg/ms.length);
  const l3=ms.slice(-3).reduce((s,m)=>s+lifeMonthly(m),0)/3,p3=ms.slice(-6,-3).reduce((s,m)=>s+lifeMonthly(m),0)/3;
  const rTr=p3?Math.round((l3-p3)/p3*100):0;
  const ins=[
    {{t:invPct>=30?'good':'info',i:invPct>=30?'🏆':'📈',h:`Investment Rate: ${{invPct}}% of total`,b:invPct>=30?`Great discipline! You've invested ${{fi(inv)}} — ${{invPct}}% of all spending. Sustaining 30%+ builds serious wealth.`:`You've invested ${{fi(inv)}} (${{invPct}}%). Target 30%+ via automated SIPs.`}},
    {{t:'info',i:'🏥',h:`3-Bucket Split: Life ${{lifePct}}% · Clinic ${{tot?Math.round(cli/tot*100):0}}% · Inv ${{invPct}}%`,b:`Lifestyle (${{fi(life)}}), Clinic (${{fi(cli)}}), Investments (${{fi(inv)}}). Separating these gives a cleaner picture of personal vs professional spend.`}},
    {{t:eoTr>20?'alert':eoTr>8?'warn':'good',i:'🍽️',h:`Eating Out: ${{eoTr>0?'↑':'↓'}}${{Math.abs(eoTr)}}% trend`,b:`Avg ${{fi(eoAvg)}}/month. ${{eoTr>20?'Rising fast — cooking at home more could save ₹'+Math.round(eoAvg*0.35).toLocaleString()+'/month.':eoTr>8?'Slightly rising — set a weekly food budget.':'Stable or declining — great discipline!'}}`}},
    {{t:travel>800000?'warn':'info',i:'✈️',h:`Travel: ${{fi(travel)}} total`,b:`${{fi(travel)}} on travel (avg ${{fi(Math.round(travel/ms.length))}}/month). ${{travel>800000?'Book 60+ days ahead for 25-40% savings.':'Well-balanced spend — rich in experiences!'}}`}},
    {{t:gifts>600000?'warn':'info',i:'🎁',h:`Gifts: ${{fi(gifts)}} total`,b:`${{fi(gifts)}} on gifts (avg ${{fi(Math.round(gifts/ms.length))}}/month). ${{gifts>600000?'Set a per-occasion budget — thoughtfulness > price.':'Generous and well-balanced!'}}`}},
    {{t:shop>400000?'warn':'good',i:'🛍️',h:`Shopping: ${{fi(shop)}} total`,b:`${{fi(shop)}} on clothes & electronics. ${{shop>400000?'Try a 30-day wishlist rule to cut impulse buys.':'Shopping spend is well in control!'}}`}},
    {{t:pgAvg<3000?'warn':'good',i:'🎯',h:`Personal Growth: ${{fi(pgAvg)}}/month`,b:pgAvg<3000?`Only ${{fi(pgAvg)}}/month on skill development. Aim for 2-3% of income — ROI is compounding.`:`${{fi(pgAvg)}}/month on growth — excellent!`}},
    {{t:rTr>20?'alert':rTr>0?'warn':'good',i:rTr>0?'⚠️':'✅',h:`Recent Lifestyle Trend: ${{rTr>0?'+':''}}${{rTr}}%`,b:rTr>20?`Lifestyle spend up ${{rTr}}% recently. Set a monthly cap and review discretionary categories.`:rTr>0?'Slightly elevated recently. Monitor over next 2 months.':'Recent spend is lower than prior period — great!'}},
    {{t:'info',i:'💡',h:'Smart Spending Next Steps',b:'① Automate investments first (SIP before spending). ② Set category-wise monthly budgets. ③ Audit subscriptions quarterly. ④ Claim tax deductions on Clinic expenses.'}},
    {{t:'good',i:'📊',h:`${{rows.length}} transactions · ${{ms.length}} months tracked`,b:'Consistent tracking is the #1 financial habit. Visibility drives smarter decisions.'}},
  ];
  document.getElementById('i-grid').innerHTML=ins.map(x=>`<div class="icard ${{x.t}}"><div class="iico">${{x.i}}</div><div><div class="ititle">${{x.h}}</div><div class="itext">${{x.b}}</div></div></div>`).join('');
  const lp=ms.map(m=>{{const t=MTTL[m]||0;return t?+(lifeMonthly(m)/t*100).toFixed(1):0}});
  dc('lifepct');
  CH['lifepct']=new Chart(document.getElementById('c-lifepct'),{{type:'line',
    data:{{labels:ms.map(ml),datasets:[
      {{label:'Lifestyle %',data:lp,borderColor:'#6c63ff',backgroundColor:'#6c63ff20',fill:true,tension:.4,pointRadius:2,borderWidth:2}},
      {{label:'70%',data:ms.map(()=>70),borderColor:'#ff6b6b',borderDash:[4,4],borderWidth:1.5,pointRadius:0}}
    ]}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,tooltip:{{callbacks:{{label:ctx=>ctx.raw+'%'}}}}}},
      scales:{{...CD.scales,y:{{...CD.scales.y,min:0,max:110,ticks:{{...CD.scales.y.ticks,callback:v=>v+'%'}}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const m=ms[els[0].index];const r=gRows().filter(x=>x.ym===m&&bucket(x.cat)==='life').sort((a,b)=>b.cost-a.cost);showModal('Lifestyle · '+ml(m),r);}}
    }}
  }});
  const ib=ms.map(m=>PIVOT['Investments']?.[m]||0);
  dc('invconsist');
  CH['invconsist']=new Chart(document.getElementById('c-invconsist'),{{type:'bar',
    data:{{labels:ms.map(ml),datasets:[{{label:'Investments',data:ib,backgroundColor:ib.map(v=>v>0?'#6c63ffaa':'#2e3350'),borderColor:ib.map(v=>v>0?'#6c63ff':'#2e3350'),borderWidth:1,borderRadius:3}}]}},
    options:{{...CD,responsive:true,maintainAspectRatio:false,onHover:hover,
      plugins:{{...CD.plugins,legend:{{display:false}},tooltip:{{callbacks:{{label:ctx=>fi(ctx.raw)}}}}}},
      onClick:(evt,els)=>{{if(!els.length)return;const m=ms[els[0].index];const r=gRows().filter(x=>x.ym===m&&x.cat==='Investments').sort((a,b)=>b.cost-a.cost);showModal('Investments · '+ml(m),r);}}
    }}
  }});
}}

function go(id,btn){{
  curPage=id;
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(b=>b.classList.remove('active'));
  document.getElementById('pg-'+id).classList.add('active');
  if(btn) btn.classList.add('active');
  rerender();
}}

document.addEventListener('DOMContentLoaded',()=>{{
  initGF();initMs();
  LCATS.forEach(c=>{{const o=document.createElement('option');o.value=c;o.textContent=c;document.getElementById('tr-cat').appendChild(o)}});
  const yrs=[...new Set(RAW.map(r=>r.year))].sort();
  yrs.forEach(y=>{{const o=document.createElement('option');o.value=y;o.textContent=y;document.getElementById('tx-yr').appendChild(o)}});
  ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'].forEach((n,i)=>{{const o=document.createElement('option');o.value=i+1;o.textContent=n;document.getElementById('tx-mo').appendChild(o)}});
  rnOverview();
}});
</script>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# GIT
# ══════════════════════════════════════════════════════════════════════════════
def git_push(message):
    os.chdir(ROOT)
    try:
        subprocess.run(["git", "add", "docs/index.html"], check=True)
        result = subprocess.run(["git", "diff", "--cached", "--quiet"])
        if result.returncode == 0:
            print("  Nothing new to commit — dashboard is already up to date.")
            return False
        subprocess.run(["git", "commit", "-m", message], check=True)
        subprocess.run(["git", "push"], check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n  Git error: {e}")
        print("  The HTML was generated successfully at docs/index.html")
        print("  Push manually with: git push")
        return False


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="Regenerate and deploy the Finance Dashboard.")
    parser.add_argument("--primary",    default=DEFAULT_PRIMARY,    help="Path to Primary data.xlsx")
    parser.add_argument("--secondary",  default=DEFAULT_SECONDARY,  help="Path to Secondary data.csv")
    parser.add_argument("--categories", default=DEFAULT_CATEGORIES, help="Path to Expense Categories.docx")
    parser.add_argument("--no-push",    action="store_true",        help="Generate HTML only, skip git push")
    parser.add_argument("--message",    default="",                 help="Custom git commit message")
    args = parser.parse_args()

    print("\n══════════════════════════════════════════")
    print("  Finance Dashboard Updater")
    print("══════════════════════════════════════════")

    # Validate inputs
    for label, path in [("Primary", args.primary), ("Secondary", args.secondary)]:
        if not os.path.exists(path):
            print(f"\n  ERROR: {label} file not found: {path}")
            print(f"  Drop it in the data/ folder or pass --{label.lower()} <path>")
            sys.exit(1)

    # Parse categories docx if provided
    print("\n[1/4] Checking category rules...")
    parse_categories_docx(args.categories)

    # Load data
    print("\n[2/4] Loading data...")
    primary_rows, cutoff = load_primary(args.primary)
    secondary_rows       = load_secondary(args.secondary, cutoff)

    all_rows = sorted(primary_rows + secondary_rows, key=lambda x: x["date"])
    print(f"\n  Combined: {len(all_rows)} rows ({len(primary_rows)} primary + {len(secondary_rows)} secondary)")

    # Generate HTML
    print("\n[3/4] Generating dashboard HTML...")
    os.makedirs(DOCS_DIR, exist_ok=True)
    D, raw_serial = build_chart_data(all_rows)
    html          = generate_html(D, raw_serial)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  Written: docs/index.html ({len(html):,} chars)")
    print(f"  Date range: {D['date_range']}")

    # Git push
    if args.no_push:
        print("\n[4/4] Skipping git push (--no-push flag set)")
        print(f"\n  Open locally: file://{OUTPUT_HTML}")
    else:
        print("\n[4/4] Pushing to GitHub...")
        date_str = datetime.now().strftime("%d %b %Y %H:%M")
        msg = args.message or f"Update dashboard — {date_str}"
        pushed = git_push(msg)
        if pushed:
            print("\n  ✓ Dashboard updated successfully!")
            print("  It will be live at your GitHub Pages URL in ~60 seconds.")

    print("\n══════════════════════════════════════════\n")


if __name__ == "__main__":
    main()
