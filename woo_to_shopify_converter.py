"""
WooCommerce → Shopify CSV Converter  v2.0
==========================================
Converts WooCommerce product and customer export CSV files
into Shopify-compatible import CSVs.

USAGE:
  python woo_to_shopify_converter.py

REQUIREMENTS:
  pip install pandas openpyxl
"""

import os
import re
import csv
import sys
import itertools
import pandas as pd
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR  = os.path.join(SCRIPT_DIR, "shopify_output")

# ── Shopify official headers ──────────────────────────────────
PRODUCT_HEADERS = [
    "Handle","Title","Body (HTML)","Vendor","Type","Tags","Published",
    "Option1 Name","Option1 Value","Option2 Name","Option2 Value",
    "Option3 Name","Option3 Value",
    "Variant SKU","Variant Grams","Variant Inventory Tracker",
    "Variant Inventory Qty","Variant Inventory Policy",
    "Variant Fulfillment Service","Variant Price","Variant Compare At Price",
    "Variant Requires Shipping","Variant Taxable","Variant Barcode",
    "Image Src","Image Alt Text","SEO Title","SEO Description","Status",
]  # 29 columns

CUSTOMER_HEADERS = [
    "First Name","Last Name","Email","Accepts Email Marketing",
    "Default Address Company","Default Address Address1",
    "Default Address Address2","Default Address City",
    "Default Address Province Code","Default Address Country Code",
    "Default Address Zip","Default Address Phone","Phone",
    "Accepts SMS Marketing","Note","Tax Exempt","Tags",
]  # 17 columns


# ═══════════════════════════════════════════════════════════════
# UTILITIES
# ═══════════════════════════════════════════════════════════════

def log(msg, level="INFO"):
    tag = {"INFO":"i ","OK":"OK","WARN":"! ","ERROR":"X ","STEP":">>"}
    print("  [{}]  {}".format(tag.get(level,"  "), msg))


def ensure_output_dir():
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        return True
    except Exception as e:
        log("Cannot create output folder: {}".format(e), "ERROR")
        return False


def sanitize_cell(val):
    """
    Strip newlines from any cell value.
    Newlines inside quoted CSV fields cause Shopify to miscount line numbers,
    which breaks variant option validation.
    """
    if not isinstance(val, str):
        return val
    return " ".join(val.replace("\r\n"," ").replace("\r"," ").replace("\n"," ").split())


def write_csv(rows, headers, filename):
    if not ensure_output_dir():
        return
    clean = [[sanitize_cell(c) for c in row] for row in rows]
    path  = os.path.join(OUTPUT_DIR, filename)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f, quoting=csv.QUOTE_ALL).writerows([headers] + clean)
    log("Saved → {}  ({:,} bytes)".format(path, os.path.getsize(path)), "OK")


def read_file(path):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".csv":
            try:    return pd.read_csv(path, encoding="utf-8")
            except: return pd.read_csv(path, encoding="latin-1")
        elif ext in (".xlsx",".xlsm"):
            return pd.read_excel(path, engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(path, engine="xlrd")
        else:
            log("Unsupported file type: {}".format(ext), "ERROR"); return None
    except Exception as e:
        log("Cannot read {}: {}".format(os.path.basename(path), e), "ERROR"); return None


def sv(val, default=""):
    """Safe string — returns default for NaN / None / blank."""
    try:
        if val is None: return default
        if isinstance(val, float) and pd.isna(val): return default
        s = str(val).strip()
        return s if s.lower() not in ("nan","none","") else default
    except Exception:
        return default


def sfloat(val, default=""):
    """Safe 2-decimal float string. Empty for zero / missing."""
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)): return default
        f = float(val)
        return "{:.2f}".format(f) if f > 0 else default
    except Exception:
        return default


def sint(val, default=0):
    try:
        if val is None or (isinstance(val, float) and pd.isna(val)): return default
        return int(float(val))
    except Exception:
        return default


def find_col(df, candidates):
    for c in candidates:
        if c in df.columns: return c
    return None


def gv(row, col, default=""):
    if col is None: return default
    return sv(row.get(col, default), default)


# ═══════════════════════════════════════════════════════════════
# WOO-SPECIFIC PARSERS
# ═══════════════════════════════════════════════════════════════

def make_handle(title, sku=""):
    if not title:
        return sku.lower().replace(" ","-") if sku else "product"
    return re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")


def parse_woo_images(raw):
    """Comma-separated image URLs → list of https:// URLs."""
    if not raw or sv(raw) == "": return []
    result = []
    for url in str(raw).split(","):
        url = url.strip()
        if url.startswith("http://"):
            url = "https://" + url[7:]
        if url.startswith("https://"):
            result.append(url)
    return result


def parse_woo_categories(raw):
    """'Clothing > T-Shirts, Accessories' → 'T-Shirts, Accessories, Clothing'"""
    if not raw or sv(raw) == "": return ""
    all_names = []
    for cat in str(raw).split(","):
        for part in cat.split(">"):
            part = part.strip()
            if part: all_names.append(part)
    seen, unique = set(), []
    for n in all_names:
        if n.lower() not in seen:
            seen.add(n.lower()); unique.append(n)
    return ", ".join(unique)


def parse_woo_tags(raw):
    if not raw or sv(raw) == "": return ""
    return ", ".join([t.strip() for t in str(raw).split(",") if t.strip()])


def unescape_woo_attr_value(raw):
    """
    FIX 2: WooCommerce escapes commas inside attribute values with a backslash:
      '4\\,000 Lumens, 6\\,000 Lumens'  means two values: '4,000 Lumens' and '6,000 Lumens'

    Strategy:
      1. Replace  \\,  with a placeholder  \x00
      2. Split on  ,  to get individual values
      3. Replace  \x00  back to  ,  in each value
    Returns a list of clean individual values.
    """
    if not raw or sv(raw) == "": return []
    # Replace escaped comma (\,) with a placeholder
    placeholder = raw.replace("\\,", "\x00")
    # Split on unescaped commas
    parts = placeholder.split(",")
    # Restore the real commas and clean up
    return [p.replace("\x00", ",").strip() for p in parts if p.replace("\x00","").strip()]


def weight_to_grams(val, unit="kg"):
    """
    FIX 1: Auto-detect weight unit from column name.
    WooCommerce default is kg, but this store uses lbs.
    """
    try:
        w = float(val or 0)
        if unit == "lbs":
            return int(round(w * 453.592))
        else:
            return int(round(w * 1000))
    except Exception:
        return 0


def pick_price_and_compare(regular_price, sale_price):
    reg  = float(sfloat(regular_price, "0") or "0")
    sale = float(sfloat(sale_price,    "0") or "0")
    if sale > 0 and sale < reg:
        return "{:.2f}".format(sale), "{:.2f}".format(reg)
    elif reg > 0:
        return "{:.2f}".format(reg), ""
    return "", ""


def woo_published_to_shopify(val):
    try:    v = int(float(str(val).strip()))
    except: v = 1
    return ("TRUE","active") if v == 1 else ("FALSE","draft")


def build_body_html(description, short_description):
    desc  = sv(description)
    short = sv(short_description)
    if short and desc:
        return "<p><strong>{}</strong></p>\n\n{}".format(short, desc)
    return desc or short or ""


def get_seo_fields(row):
    """Check all major WP SEO plugin meta columns."""
    seo_title_keys = [
        "Meta: _yoast_wpseo_title","Meta: rank_math_title",
        "Meta: _aioseop_title","Meta: _aioseo_title",
    ]
    seo_desc_keys = [
        "Meta: _yoast_wpseo_metadesc","Meta: rank_math_description",
        "Meta: _aioseop_description","Meta: _aioseo_description",
    ]
    seo_title = next((sv(row.get(k,"")) for k in seo_title_keys if sv(row.get(k,""))), "")
    seo_desc  = next((sv(row.get(k,"")) for k in seo_desc_keys  if sv(row.get(k,""))), "")
    return seo_title, seo_desc


def get_woo_attributes(row, df_cols, single_value_only=False):
    """
    FIX 5: Support up to 9 attribute columns (WooCommerce allows up to 9).
    FIX 2: Unescape \\, in attribute values.
    Returns list of (name, value) tuples — single_value_only takes first value only.
    """
    attrs = []
    for i in range(1, 10):  # FIX 5: check up to 9 (Shopify uses only first 3)
        name_col  = "Attribute {} name".format(i)
        value_col = "Attribute {} value(s)".format(i)
        if name_col not in df_cols or value_col not in df_cols:
            break
        name  = sv(row.get(name_col, ""))
        value = sv(row.get(value_col, ""))
        if not name or not value:
            continue
        if single_value_only:
            # FIX 2: unescape then take first value
            values = unescape_woo_attr_value(value)
            value  = values[0] if values else value.split("|")[0].strip()
        attrs.append((name, value))
    return attrs


def expand_parent_variants(attrs):
    """
    When parent has multi-valued attributes and no variation rows,
    expand into one variant per combination using cartesian product.
    FIX 2: unescape_woo_attr_value used here for correct splitting.
    """
    option_lists = []
    for name, values_raw in attrs:
        # Try pipe-separated first (WC default), then unescape comma-separated
        if "|" in values_raw:
            values = [v.strip() for v in values_raw.split("|") if v.strip()]
        else:
            values = unescape_woo_attr_value(values_raw)
        if values:
            option_lists.append((name, values))

    if not option_lists:
        return []

    names  = [o[0] for o in option_lists]
    combos = list(itertools.product(*[o[1] for o in option_lists]))
    return [list(zip(names, combo)) for combo in combos]


def _blank_row(handle):
    row = [""] * len(PRODUCT_HEADERS)
    row[0] = handle
    return row


def is_valid_variation(row):
    """
    FIX 3+4: Skip variation rows that are invalid:
    - Blank/NaN attribute values (placeholder rows WooCommerce adds for out-of-stock)
    - Out-of-stock (In stock? = 0) AND no SKU AND no attribute values
    These rows produce empty option values which Shopify rejects.
    """
    attr1_val = sv(row.get("Attribute 1 value(s)", ""))
    attr2_val = sv(row.get("Attribute 2 value(s)", ""))
    sku       = sv(row.get("SKU", ""))
    in_stock  = str(row.get("In stock?", "1")).strip()

    # If ALL attribute values are blank → this is a placeholder row, skip it
    if not attr1_val and not attr2_val:
        return False

    return True


# ═══════════════════════════════════════════════════════════════
# PRODUCT CONVERTER
# ═══════════════════════════════════════════════════════════════

def convert_products(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None: return

    log("Loaded {} rows | {} columns".format(len(df), len(df.columns)))

    # FIX 1: Auto-detect weight unit from column name
    weight_col  = find_col(df, ["Weight (kg)","Weight (lbs)","Weight (g)","weight"])
    weight_unit = "kg"
    if weight_col:
        if "lbs" in weight_col.lower(): weight_unit = "lbs"
        elif "(g)" in weight_col.lower(): weight_unit = "g"
    log("Weight column: '{}' → unit={}".format(weight_col or "not found", weight_unit))

    # Print column check
    print()
    log("Key columns detected:", "INFO")
    key_checks = [
        ("Type",["Type"]),
        ("SKU",["SKU"]),
        ("Name",["Name"]),
        ("Published",["Published"]),
        ("Regular price",["Regular price"]),
        ("Sale price",["Sale price"]),
        ("Images",["Images"]),
        ("Categories",["Categories"]),
        ("Tags",["Tags"]),
        ("Weight",["Weight (kg)","Weight (lbs)","Weight (g)"]),
        ("Stock",["Stock"]),
        ("Parent",["Parent"]),
        ("Attribute 1 name",["Attribute 1 name"]),
        ("Brands",["Brands"]),
    ]
    for label, candidates in key_checks:
        col    = find_col(df, candidates)
        status = "✓ '{}'".format(col) if col else "✗ not found"
        print("    {:<25} {}".format(label, status))
    print()

    # FIX 6: Build variation lookup by BOTH SKU and id:ID
    parent_col = find_col(df, ["Parent","parent_id","parent"])
    variations_by_parent = {}

    if parent_col and "Type" in df.columns:
        for _, row in df.iterrows():
            if sv(row.get("Type","")).lower() != "variation":
                continue
            # FIX 3+4: skip invalid variation rows upfront
            if not is_valid_variation(row):
                log("Skipping invalid variation (blank attributes): ID={}".format(
                    sv(row.get("ID","?"))), "WARN")
                continue
            parent_ref = sv(row.get(parent_col, ""))
            if parent_ref:
                if parent_ref not in variations_by_parent:
                    variations_by_parent[parent_ref] = []
                variations_by_parent[parent_ref].append(row)

    # FIX 6: also index by  id:ID  for stores that export parent as numeric ID
    for _, row in df.iterrows():
        row_id = sv(row.get("ID",""))
        if row_id:
            key = "id:{}".format(row_id)
            existing = variations_by_parent.get(key, [])
            # merge any rows already keyed by id:N
            variations_by_parent[key] = existing

    all_rows  = []
    converted = 0
    skipped   = 0

    for _, row in df.iterrows():
        row_type = sv(row.get("Type","simple")).lower() if "Type" in df.columns else "simple"
        if row_type == "variation":
            continue

        title = sv(row.get("Name",""))
        if not title:
            skipped += 1
            continue

        sku    = sv(row.get("SKU",""))
        handle = make_handle(title, sku)
        body   = build_body_html(row.get("Description",""), row.get("Short description",""))

        # FIX 8: Brands column (WooCommerce brands plugin) → Vendor
        vendor = (
            sv(row.get("Brands",""))
            or sv(row.get("Brand",""))
            or sv(row.get("meta:_brand",""))
        )

        # Tags = categories + woo tags
        cat_tags = parse_woo_categories(sv(row.get("Categories","")))
        woo_tags = parse_woo_tags(sv(row.get("Tags","")))
        tags = ", ".join(filter(None, [cat_tags, woo_tags]))

        published, status = woo_published_to_shopify(row.get("Published", 1))
        seo_title, seo_desc = get_seo_fields(row)
        images    = parse_woo_images(sv(row.get("Images","")))
        first_img = images[0] if images else ""

        barcode = sv(row.get("GTIN, UPC, EAN, or ISBN","") or row.get("global_unique_id",""))
        if "." in barcode:
            barcode = barcode.rstrip("0").rstrip(".")

        # FIX 1: use detected weight unit
        raw_weight = row.get(weight_col) if weight_col else None
        weight_g   = weight_to_grams(raw_weight, weight_unit)

        qty        = sint(row.get("Stock", 0))
        in_stock   = sv(row.get("In stock?","1")) in ("1","true","yes","True","TRUE")
        inv_policy = "deny" if in_stock else "continue"

        # Get parent's attributes (may be pipe- or comma-escaped)
        attrs = get_woo_attributes(row, df.columns)

        # FIX 6: Look up variations by SKU and also by id:ID
        row_id     = sv(row.get("ID",""))
        variations = (
            variations_by_parent.get(sku, [])
            or variations_by_parent.get("id:{}".format(row_id), [])
            or variations_by_parent.get(row_id, [])
        )

        # Determine option names (up to 3 — Shopify limit)
        if variations:
            first_var  = variations[0]
            var_attrs  = get_woo_attributes(first_var, df.columns, single_value_only=True)
            opt_names  = [a[0] for a in var_attrs[:3]]
            first_vals = [a[1] for a in var_attrs[:3]]
        elif attrs:
            opt_names  = [a[0] for a in attrs[:3]]
            first_vals = [unescape_woo_attr_value(a[1])[0]
                          if unescape_woo_attr_value(a[1])
                          else a[1].split("|")[0].strip()
                          for a in attrs[:3]]
            # FIX 5: attributes beyond 3 → append to tags
            extra_attr_tags = []
            for name, val_raw in attrs[3:]:
                vals = unescape_woo_attr_value(val_raw) or [val_raw]
                extra_attr_tags.extend(vals)
            if extra_attr_tags:
                tags = ", ".join(filter(None, [tags] + extra_attr_tags))
        else:
            opt_names  = []
            first_vals = []

        if not opt_names:
            opt_names  = ["Title"]
            first_vals = ["Default Title"]

        # Pad to 3
        while len(opt_names)  < 3: opt_names.append("")
        while len(first_vals) < 3: first_vals.append("")

        def oname(i): return opt_names[i]  if i < len(opt_names)  else ""
        def oval(i):  return first_vals[i] if i < len(first_vals) else ""

        # Price for first variant
        if variations:
            fv        = variations[0]
            price, cp = pick_price_and_compare(fv.get("Regular price",""), fv.get("Sale price",""))
            var_sku   = sv(fv.get("SKU","")) or sku
            v_weight  = weight_to_grams(fv.get(weight_col) if weight_col else None, weight_unit) or weight_g
            v_qty     = sint(fv.get("Stock", qty))
            v_barc    = sv(fv.get("GTIN, UPC, EAN, or ISBN","")) or barcode
            v_in_stk  = sv(fv.get("In stock?","1")) in ("1","true","yes","True","TRUE")
            v_policy  = "deny" if v_in_stk else "continue"
        else:
            price, cp = pick_price_and_compare(row.get("Regular price",""), row.get("Sale price",""))
            var_sku   = sku
            v_weight  = weight_g
            v_qty     = qty
            v_barc    = barcode
            v_policy  = inv_policy

        # ── Main product row ──────────────────────────────────
        all_rows.append([
            handle, title, body, vendor, "", tags, published,
            oname(0), oval(0), oname(1), oval(1), oname(2), oval(2),
            var_sku, v_weight, "shopify", v_qty, v_policy, "manual",
            price, cp, "TRUE", "TRUE", v_barc,
            first_img, title if first_img else "",
            seo_title, seo_desc, status,
        ])  # 29 values

        # ── Additional variation rows ─────────────────────────
        # Track seen option combos to deduplicate.
        # WooCommerce sometimes has duplicate option combinations (same CRI+Lumens
        # on two different SKUs). Shopify requires each combination to be unique.
        seen_combos = set()
        if variations:
            first_va    = get_woo_attributes(variations[0], df.columns, single_value_only=True)
            first_combo = tuple(a[1] for a in first_va[:3])
            seen_combos.add(first_combo)

        for var_row in (variations[1:] if variations else []):
            va    = get_woo_attributes(var_row, df.columns, single_value_only=True)
            vvals = [a[1] for a in va[:3]]
            while len(vvals) < 3: vvals.append("")

            # Skip duplicate option combinations
            combo = tuple(vvals)
            if combo in seen_combos:
                log("Skipping duplicate variant combo {} for '{}' SKU={}".format(
                    combo, title[:35], sv(var_row.get("SKU","?"))), "WARN")
                continue
            seen_combos.add(combo)

            vp, vc  = pick_price_and_compare(var_row.get("Regular price",""), var_row.get("Sale price",""))
            vsku    = sv(var_row.get("SKU","")) or sku
            vwg     = weight_to_grams(var_row.get(weight_col) if weight_col else None, weight_unit) or weight_g
            vqty    = sint(var_row.get("Stock", qty))
            vbarc   = sv(var_row.get("GTIN, UPC, EAN, or ISBN","")) or barcode
            v_imgs  = parse_woo_images(sv(var_row.get("Images","")))
            v_img   = v_imgs[0] if v_imgs else ""
            vis_ok  = sv(var_row.get("In stock?","1")) in ("1","true","yes","True","TRUE")
            vpol    = "deny" if vis_ok else "continue"

            all_rows.append([
                handle,"","","","","","",
                oname(0), vvals[0], oname(1), vvals[1], oname(2), vvals[2],
                vsku, vwg, "shopify", vqty, vpol, "manual",
                vp, vc, "TRUE","TRUE", vbarc,
                v_img, title if v_img else "",
                "", "", status,
            ])  # 29 values

        # ── Synthetic variants (parent-only export) ───────────
        if row_type == "variable" and not variations and attrs:
            combos = expand_parent_variants(attrs[:3])
            for combo in combos[1:]:
                cvals = [v for _, v in combo]
                while len(cvals) < 3: cvals.append("")
                all_rows.append([
                    handle,"","","","","","",
                    oname(0), cvals[0], oname(1), cvals[1], oname(2), cvals[2],
                    sku, v_weight, "shopify", v_qty, v_policy, "manual",
                    price, cp, "TRUE","TRUE", barcode,
                    "","","","",status,
                ])

        # ── Extra image rows ──────────────────────────────────
        for img_url in images[1:]:
            ir = _blank_row(handle)
            ir[PRODUCT_HEADERS.index("Image Src")]      = img_url
            ir[PRODUCT_HEADERS.index("Image Alt Text")] = title
            all_rows.append(ir)

        converted += 1

    # Validate
    bad = [(i+2, len(r)) for i,r in enumerate(all_rows) if len(r) != len(PRODUCT_HEADERS)]
    if bad:
        log("BUG: {} rows with wrong column count!".format(len(bad)), "ERROR")
        for ri, rc in bad[:5]:
            log("  Row {}: {} cols (expected {})".format(ri,rc,len(PRODUCT_HEADERS)), "ERROR")
        return

    write_csv(all_rows, PRODUCT_HEADERS, "shopify_products.csv")
    print()
    log("Products converted  : {}".format(converted), "OK")
    log("Products skipped    : {}".format(skipped), "OK" if skipped==0 else "WARN")
    log("Total CSV rows      : {}".format(len(all_rows)), "OK")
    print()
    log("Import → Shopify Admin → Products → Import → shopify_products.csv", "INFO")
    print()


# ═══════════════════════════════════════════════════════════════
# CUSTOMER CONVERTER (unchanged)
# ═══════════════════════════════════════════════════════════════

def convert_customers(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None: return

    log("Loaded {} rows | {} columns".format(len(df), len(df.columns)))

    df.columns = [
        str(c).strip().lower().replace(" ","_").replace("-","_").replace("/","_")
        for c in df.columns
    ]

    c_first   = find_col(df,["first_name","billing_first_name","firstname","fname"])
    c_last    = find_col(df,["last_name","billing_last_name","lastname","lname"])
    c_email   = find_col(df,["email","billing_email","email_address","user_email"])
    c_phone   = find_col(df,["phone","billing_phone","phone_number","mobile"])
    c_company = find_col(df,["company","billing_company","company_name"])
    c_addr1   = find_col(df,["address_1","billing_address_1","address1","street"])
    c_addr2   = find_col(df,["address_2","billing_address_2","address2"])
    c_city    = find_col(df,["city","billing_city","town"])
    c_state   = find_col(df,["state","billing_state","province","county"])
    c_postcode= find_col(df,["postcode","billing_postcode","zip","postal_code"])
    c_country = find_col(df,["country","billing_country","country_code","billing_country_code"])
    c_note    = find_col(df,["note","customer_note","notes","order_notes"])

    COUNTRY_NAMES = {
        "united states":"US","usa":"US","united kingdom":"GB","uk":"GB",
        "canada":"CA","australia":"AU","india":"IN","germany":"DE",
        "france":"FR","italy":"IT","spain":"ES","netherlands":"NL",
        "new zealand":"NZ","singapore":"SG","ireland":"IE",
    }

    all_rows = []
    skipped  = 0

    for _, row in df.iterrows():
        email = gv(row, c_email)
        if not email or "@" not in email:
            skipped += 1
            continue

        prov = gv(row, c_state, "").upper().strip()
        if len(prov) > 4: prov = ""

        cc = gv(row, c_country, "US").strip()
        if len(cc) > 2:
            cc = COUNTRY_NAMES.get(cc.lower(), cc[:2].upper())
        else:
            cc = cc.upper()

        phone = gv(row, c_phone, "")
        if phone:
            d = re.sub(r"[^\d+]","",phone)
            if d and not d.startswith("+"):
                d = "+1"+d if len(d)==10 else "+"+d
            phone = d

        all_rows.append([
            gv(row,c_first), gv(row,c_last), email, "no",
            gv(row,c_company), gv(row,c_addr1), gv(row,c_addr2),
            gv(row,c_city), prov, cc,
            gv(row,c_postcode), phone, phone,
            "no", gv(row,c_note), "no", "",
        ])

    write_csv(all_rows, CUSTOMER_HEADERS, "shopify_customers.csv")
    print()
    log("Customers converted : {}".format(len(all_rows)), "OK")
    log("Customers skipped   : {} (no valid email)".format(skipped),
        "OK" if skipped==0 else "WARN")
    print()
    log("Import → Shopify Admin → Customers → Import → shopify_customers.csv", "INFO")
    print()


# ═══════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════
# ORDER CONVERTER  —  Matrixify format
# ═══════════════════════════════════════════════════════════════
#
# Shopify does NOT support native order CSV import.
# Output is in Matrixify format (free app from Shopify App Store).
# Import via: Shopify Admin → Apps → Matrixify → Import → shopify_orders.csv
#
# WooCommerce order export column names (official WooCommerce schema):
#   order_id / Order ID / order_number
#   order_date / date_created / order_created / Date
#   status / order_status / Status
#   billing_email / Billing Email
#   billing_first_name / Billing First Name
#   billing_last_name  / Billing Last Name
#   billing_company    / Billing Company
#   billing_address_1  / Billing Address 1
#   billing_address_2  / Billing Address 2
#   billing_city       / Billing City
#   billing_state      / Billing State
#   billing_postcode   / Billing Postcode
#   billing_country    / Billing Country
#   billing_phone      / Billing Phone
#   shipping_first_name ... (same pattern)
#   order_total / Total / order_subtotal
#   order_shipping / Shipping Total
#   order_tax / Tax Total
#   discount_total / Discount Total
#   payment_method / Payment Method
#   order_notes / Customer Note
#   line_items  (pipe-delimited: product_id:X|name:Y|quantity:Z|total:N)
#   OR separate columns: item_name, item_quantity, item_price, item_sku

ORDER_HEADERS = [
    "Name",                   # Order number e.g. #1001
    "Email",
    "Phone",
    "Currency",
    "Payment: Status",        # paid / pending / refunded / voided
    "Processed At",           # YYYY-MM-DD HH:MM:SS +0000
    "Send Receipt",           # FALSE
    "Inventory Behaviour",    # bypass
    "Note",
    "Tags",
    # Billing
    "Billing: First Name",
    "Billing: Last Name",
    "Billing: Company",
    "Billing: Phone",
    "Billing: Address 1",
    "Billing: Address 2",
    "Billing: City",
    "Billing: Province",
    "Billing: Province Code",
    "Billing: Country",
    "Billing: Country Code",
    "Billing: Zip",
    # Shipping
    "Shipping: First Name",
    "Shipping: Last Name",
    "Shipping: Company",
    "Shipping: Phone",
    "Shipping: Address 1",
    "Shipping: Address 2",
    "Shipping: City",
    "Shipping: Province",
    "Shipping: Province Code",
    "Shipping: Country",
    "Shipping: Country Code",
    "Shipping: Zip",
    # Line item fields (filled per row type)
    "Line: Type",             # Line Item / Shipping Line / Transaction
    "Line: Title",
    "Line: Quantity",
    "Line: Price",
    "Line: SKU",
    "Line: Grams",
    "Line: Requires Shipping",
    "Line: Taxable",
    "Line: Discount",
    # Tax
    "Tax 1: Title",
    "Tax 1: Price",
    "Tax 1: Rate",
    # Transaction
    "Transaction: Amount",
    "Transaction: Currency",
    "Transaction: Kind",
    "Transaction: Status",
    "Transaction: Gateway",
]  # 51 columns

# WooCommerce order status → Shopify / Matrixify payment status
WOO_ORDER_STATUS_MAP = {
    "completed":       "paid",
    "processing":      "paid",
    "wc-completed":    "paid",
    "wc-processing":   "paid",
    "pending":         "pending",
    "on-hold":         "pending",
    "wc-pending":      "pending",
    "wc-on-hold":      "pending",
    "refunded":        "refunded",
    "wc-refunded":     "refunded",
    "cancelled":       "voided",
    "canceled":        "voided",
    "wc-cancelled":    "voided",
    "failed":          "voided",
    "wc-failed":       "voided",
}


def fmt_order_date(raw):
    """Parse any WooCommerce date string → Matrixify ISO format."""
    try:
        return pd.to_datetime(raw, dayfirst=False).strftime("%Y-%m-%d %H:%M:%S +0000")
    except Exception:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S +0000")


def clean_phone_order(phone):
    """Return E.164 phone or empty string."""
    if not phone:
        return ""
    d = re.sub(r"[^\d+]", "", str(phone))
    if d and not d.startswith("+"):
        d = "+1" + d if len(d) == 10 else "+" + d
    return d


def parse_woo_line_items(raw):
    """
    WooCommerce exports line items in two formats:

    Format A — pipe-delimited single column (WooCommerce CSV Export plugin default):
      'product_id:99|name:Widget|sku:WID-01|quantity:2|total:29.99|subtotal:29.99'
    Multiple items separated by comma or newline.

    Format B — separate columns per item (some export plugins):
      item_name, item_quantity, item_price, item_sku
      item_name_2, item_quantity_2, item_price_2, item_sku_2

    Returns list of dicts: [{name, sku, qty, price}, ...]
    """
    if not raw or sv(raw) == "":
        return []

    items = []
    # Split multiple items — WooCommerce uses comma or newline between items
    # but values within items use pipe |
    # Strategy: split on ", product_id:" or "},{" (JSON) or newline
    raw_str = str(raw).strip()

    # Try pipe-delimited format first
    if "product_id:" in raw_str or "name:" in raw_str or "quantity:" in raw_str:
        # Split on occurrences of product_id: that are not at the start
        parts = re.split(r",\s*(?=product_id:|name:)", raw_str)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            item = {}
            for field in part.split("|"):
                field = field.strip()
                if ":" in field:
                    key, _, val = field.partition(":")
                    item[key.strip()] = val.strip()
            if item:
                items.append({
                    "name":  item.get("name",  item.get("product_name", "Item")),
                    "sku":   item.get("sku",   item.get("product_sku",  "")),
                    "qty":   item.get("quantity", "1"),
                    "price": item.get("total",    item.get("item_total", item.get("price", "0"))),
                })
        return items

    # Try JSON format: [{"product_id":99,"product_name":"Widget","quantity":1,"total":19.99}]
    if raw_str.startswith("[{"):
        try:
            import json
            parsed = json.loads(raw_str)
            for obj in parsed:
                items.append({
                    "name":  str(obj.get("product_name", obj.get("name", "Item"))),
                    "sku":   str(obj.get("sku",          "")),
                    "qty":   str(obj.get("quantity",     1)),
                    "price": str(obj.get("total",        obj.get("item_total", 0))),
                })
            return items
        except Exception:
            pass

    # Fallback: treat entire value as one item name
    return [{"name": raw_str[:200], "sku": "", "qty": "1", "price": "0"}]


def _order_base_row(name, email, phone, currency, pay_status, processed_at,
                    note, b_fname, b_lname, b_company, b_phone,
                    b_addr1, b_addr2, b_city, b_province, b_prov_code,
                    b_country, b_cc, b_zip,
                    s_fname, s_lname, s_company, s_phone,
                    s_addr1, s_addr2, s_city, s_province, s_prov_code,
                    s_country, s_cc, s_zip):
    """Return a 51-value base row with all line/tax/transaction cols blank."""
    return [
        name, email, phone, currency, pay_status, processed_at,
        "FALSE",   # Send Receipt
        "bypass",  # Inventory Behaviour
        note, "",  # Tags
        # Billing (12 fields)
        b_fname, b_lname, b_company, b_phone,
        b_addr1, b_addr2, b_city, b_province, b_prov_code,
        b_country, b_cc, b_zip,
        # Shipping (12 fields)
        s_fname, s_lname, s_company, s_phone,
        s_addr1, s_addr2, s_city, s_province, s_prov_code,
        s_country, s_cc, s_zip,
        # Line / Tax / Transaction cols — all blank, filled per row type
        "", "", "", "", "", "", "", "", "",
        "", "", "",
        "", "", "", "", "",
    ]   # 51 values — matches ORDER_HEADERS exactly


def convert_orders(input_path):
    log("Reading: {}".format(os.path.basename(input_path)), "STEP")
    df = read_file(input_path)
    if df is None:
        return

    log("Loaded {} rows | {} columns".format(len(df), len(df.columns)))

    # Normalize column names for matching
    df.columns = [
        str(c).strip().lower()
               .replace(" ", "_").replace("-", "_").replace("/", "_")
        for c in df.columns
    ]

    # ── Detect columns (WooCommerce has many export formats) ──
    c_id      = find_col(df, ["order_id","id","order_number","order_no","number"])
    c_date    = find_col(df, ["order_date","date_created","date","created_at",
                               "order_created","post_date"])
    c_status  = find_col(df, ["status","order_status","post_status"])
    c_currency= find_col(df, ["order_currency","currency","currency_code"])
    c_total   = find_col(df, ["order_total","total","grand_total","total_price"])
    c_subtotal= find_col(df, ["order_subtotal","subtotal","cart_total","subtotal_price"])
    c_shipping= find_col(df, ["order_shipping","shipping_total","shipping_cost",
                               "shipping","total_shipping"])
    c_tax     = find_col(df, ["order_tax","tax_total","total_tax","taxes","tax"])
    c_discount= find_col(df, ["discount_total","cart_discount","coupon_discount",
                               "total_discount","discount"])
    c_payment = find_col(df, ["payment_method","payment_method_title","payment"])
    c_note    = find_col(df, ["customer_note","order_notes","notes","note",
                               "customer_message"])
    # Billing
    c_b_email = find_col(df, ["billing_email","email","customer_email"])
    c_b_phone = find_col(df, ["billing_phone","phone","customer_phone"])
    c_b_fname = find_col(df, ["billing_first_name","billing_firstname","first_name"])
    c_b_lname = find_col(df, ["billing_last_name","billing_lastname","last_name"])
    c_b_comp  = find_col(df, ["billing_company","company"])
    c_b_addr1 = find_col(df, ["billing_address_1","billing_address1","billing_address",
                               "billing_street"])
    c_b_addr2 = find_col(df, ["billing_address_2","billing_address2"])
    c_b_city  = find_col(df, ["billing_city","city"])
    c_b_state = find_col(df, ["billing_state","billing_province","state","province"])
    c_b_zip   = find_col(df, ["billing_postcode","billing_zip","postcode","zip"])
    c_b_cntry = find_col(df, ["billing_country","country","billing_country_code"])
    # Shipping
    c_s_fname = find_col(df, ["shipping_first_name","ship_first_name"])
    c_s_lname = find_col(df, ["shipping_last_name","ship_last_name"])
    c_s_comp  = find_col(df, ["shipping_company","ship_company"])
    c_s_addr1 = find_col(df, ["shipping_address_1","shipping_address1","shipping_address",
                               "ship_address"])
    c_s_addr2 = find_col(df, ["shipping_address_2","shipping_address2","ship_address2"])
    c_s_city  = find_col(df, ["shipping_city","ship_city"])
    c_s_state = find_col(df, ["shipping_state","shipping_province","ship_state"])
    c_s_zip   = find_col(df, ["shipping_postcode","shipping_zip","ship_zip"])
    c_s_cntry = find_col(df, ["shipping_country","ship_country","shipping_country_code"])
    # Line items — may be one combined column or separate columns
    c_items   = find_col(df, ["line_items","order_items","items","products"])
    c_iname   = find_col(df, ["item_name","product_name","line_item_name"])
    c_iqty    = find_col(df, ["item_quantity","quantity","qty","line_qty"])
    c_iprice  = find_col(df, ["item_total","item_price","unit_price","line_price",
                               "product_price"])
    c_isku    = find_col(df, ["item_sku","product_sku","sku"])
    c_ship_method = find_col(df, ["shipping_method","shipping_method_title",
                                   "shipping_zone","ship_method"])

    # Print detection summary
    print()
    log("Detected order columns:", "INFO")
    for label, col in [
        ("order_id",        c_id),
        ("date",            c_date),
        ("status",          c_status),
        ("billing_email",   c_b_email),
        ("billing_name",    c_b_fname),
        ("billing_address", c_b_addr1),
        ("shipping_address",c_s_addr1),
        ("line_items",      c_items),
        ("item_name",       c_iname),
        ("order_total",     c_total),
        ("shipping_total",  c_shipping),
        ("tax_total",       c_tax),
    ]:
        found = "✓ '{}'".format(col) if col else "✗ not found"
        print("    {:<25} {}".format(label, found))
    print()

    all_rows = []
    skipped  = 0

    for _, row in df.iterrows():
        order_id = gv(row, c_id)
        if not order_id:
            skipped += 1
            continue

        order_name   = "#{}".format(order_id)
        email        = gv(row, c_b_email)
        phone        = clean_phone_order(gv(row, c_b_phone))
        currency     = gv(row, c_currency, "USD")
        processed_at = fmt_order_date(gv(row, c_date))
        note         = gv(row, c_note)

        raw_status   = gv(row, c_status, "completed").lower().strip()
        pay_status   = WOO_ORDER_STATUS_MAP.get(raw_status, "paid")

        # Billing address
        b_fname     = gv(row, c_b_fname)
        b_lname     = gv(row, c_b_lname)
        b_company   = gv(row, c_b_comp)
        b_phone     = clean_phone_order(gv(row, c_b_phone))
        b_addr1     = gv(row, c_b_addr1)
        b_addr2     = gv(row, c_b_addr2)
        b_city      = gv(row, c_b_city)
        b_state_raw = gv(row, c_b_state)
        b_prov_code = b_state_raw.upper() if len(b_state_raw) <= 4 else ""
        b_province  = b_state_raw if len(b_state_raw) > 2 else ""
        b_cntry_raw = gv(row, c_b_cntry, "US")
        b_cc        = b_cntry_raw.upper()[:2] if len(b_cntry_raw) <= 3 else "US"
        b_country   = b_cntry_raw if len(b_cntry_raw) > 2 else ""
        b_zip       = gv(row, c_b_zip)

        # Shipping — fall back to billing if empty
        s_fname     = gv(row, c_s_fname)  or b_fname
        s_lname     = gv(row, c_s_lname)  or b_lname
        s_company   = gv(row, c_s_comp)   or b_company
        s_addr1     = gv(row, c_s_addr1)  or b_addr1
        s_addr2     = gv(row, c_s_addr2)  or b_addr2
        s_city      = gv(row, c_s_city)   or b_city
        s_state_raw = gv(row, c_s_state)  or b_state_raw
        s_prov_code = s_state_raw.upper() if len(s_state_raw) <= 4 else b_prov_code
        s_province  = s_state_raw if len(s_state_raw) > 2 else b_province
        s_cntry_raw = gv(row, c_s_cntry, b_cntry_raw)
        s_cc        = s_cntry_raw.upper()[:2] if len(s_cntry_raw) <= 3 else b_cc
        s_country   = s_cntry_raw if len(s_cntry_raw) > 2 else b_country
        s_zip       = gv(row, c_s_zip)    or b_zip

        # Totals
        shipping_cost = sfloat(row.get(c_shipping) if c_shipping else None)
        tax_total     = sfloat(row.get(c_tax)      if c_tax      else None)
        order_total   = sfloat(row.get(c_total)    if c_total    else None)
        ship_method   = gv(row, c_ship_method, "Shipping")
        payment_gw    = gv(row, c_payment, "Custom Gateway")

        # Parse line items
        line_items = []
        if c_items:
            # Combined line_items column
            line_items = parse_woo_line_items(gv(row, c_items))
        elif c_iname:
            # Separate item columns
            line_items = [{
                "name":  gv(row, c_iname, "Item"),
                "sku":   gv(row, c_isku,  ""),
                "qty":   str(sint(row.get(c_iqty) if c_iqty else None) or 1),
                "price": sfloat(row.get(c_iprice) if c_iprice else None),
            }]

        if not line_items:
            # No line item data — create a placeholder row
            line_items = [{"name": "Order {}".format(order_id),
                           "sku": "", "qty": "1", "price": order_total or "0.00"}]

        # ── ROW 1+: one row per line item ─────────────────────
        for li_idx, item in enumerate(line_items):
            li_row = _order_base_row(
                order_name if li_idx == 0 else "",   # Name only on first row
                email      if li_idx == 0 else "",
                phone      if li_idx == 0 else "",
                currency, pay_status, processed_at,
                note       if li_idx == 0 else "",
                b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_cc, b_zip,
                s_fname, s_lname, s_company, "",
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_cc, s_zip,
            )
            li_row[ORDER_HEADERS.index("Line: Type")]              = "Line Item"
            li_row[ORDER_HEADERS.index("Line: Title")]             = item["name"]
            li_row[ORDER_HEADERS.index("Line: Quantity")]          = item["qty"]
            li_row[ORDER_HEADERS.index("Line: Price")]             = item["price"]
            li_row[ORDER_HEADERS.index("Line: SKU")]               = item["sku"]
            li_row[ORDER_HEADERS.index("Line: Requires Shipping")] = "TRUE"
            li_row[ORDER_HEADERS.index("Line: Taxable")]           = "TRUE"
            if tax_total and li_idx == 0:
                li_row[ORDER_HEADERS.index("Tax 1: Title")] = "Tax"
                li_row[ORDER_HEADERS.index("Tax 1: Price")] = tax_total
            all_rows.append(li_row)

        # ── Shipping Line row ──────────────────────────────────
        if shipping_cost:
            ship_row = _order_base_row(
                "", "", "", currency, pay_status, processed_at,
                "", b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_cc, b_zip,
                s_fname, s_lname, s_company, "",
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_cc, s_zip,
            )
            ship_row[ORDER_HEADERS.index("Line: Type")]  = "Shipping Line"
            ship_row[ORDER_HEADERS.index("Line: Title")] = ship_method
            ship_row[ORDER_HEADERS.index("Line: Price")] = shipping_cost
            all_rows.append(ship_row)

        # ── Transaction row (payment record) ──────────────────
        if order_total and pay_status == "paid":
            txn_row = _order_base_row(
                "", "", "", currency, pay_status, processed_at,
                "", b_fname, b_lname, b_company, b_phone,
                b_addr1, b_addr2, b_city, b_province, b_prov_code,
                b_country, b_cc, b_zip,
                s_fname, s_lname, s_company, "",
                s_addr1, s_addr2, s_city, s_province, s_prov_code,
                s_country, s_cc, s_zip,
            )
            txn_row[ORDER_HEADERS.index("Line: Type")]            = "Transaction"
            txn_row[ORDER_HEADERS.index("Transaction: Amount")]   = order_total
            txn_row[ORDER_HEADERS.index("Transaction: Currency")] = currency
            txn_row[ORDER_HEADERS.index("Transaction: Kind")]     = "sale"
            txn_row[ORDER_HEADERS.index("Transaction: Status")]   = "success"
            txn_row[ORDER_HEADERS.index("Transaction: Gateway")]  = payment_gw
            all_rows.append(txn_row)

    # Validate column count
    bad = [(i+2, len(r)) for i,r in enumerate(all_rows) if len(r) != len(ORDER_HEADERS)]
    if bad:
        log("BUG: {} rows with wrong column count!".format(len(bad)), "ERROR")
        for ri, rc in bad[:5]:
            log("  Row {}: {} cols (expected {})".format(
                ri, rc, len(ORDER_HEADERS)), "ERROR")
        return

    write_csv(all_rows, ORDER_HEADERS, "shopify_orders.csv")
    print()
    log("Orders converted   : {}".format(len(df) - skipped), "OK")
    log("Orders skipped     : {} (no order ID)".format(skipped),
        "OK" if skipped == 0 else "WARN")
    log("Total CSV rows     : {} (line + shipping + transaction rows)".format(
        len(all_rows)), "OK")
    print()
    log("IMPORTANT: Shopify does NOT support native order CSV import.", "WARN")
    log("Use the free Matrixify app:", "INFO")
    log("  1. Shopify Admin → Apps → Matrixify", "INFO")
    log("  2. Import → Add file → shopify_orders.csv", "INFO")
    log("  3. Review preview → click Import", "INFO")
    print()


def main():
    print()
    print("  +---------------------------------------------------------+")
    print("  |  WooCommerce → Shopify CSV Converter  v3.0             |")
    print("  |  Products (with Variants) + Customers + Orders         |")
    print("  +---------------------------------------------------------+")
    print()
    print("  Output folder: {}".format(OUTPUT_DIR))
    print()

    print("  =============================================================")
    print("  PRODUCTS")
    print("  =============================================================")
    print()
    print("  Export: WP Admin → Products → All Products → Export")
    print("          Select 'Export all columns' and tick 'Custom meta'")
    print()
    prod_path = input("  Path to WooCommerce products CSV (Enter to skip):\n  > ").strip().strip('"\'')
    if prod_path and os.path.exists(prod_path):
        print()
        convert_products(prod_path)
    elif prod_path:
        log("File not found: {}".format(prod_path), "ERROR")

    print("  =============================================================")
    print("  CUSTOMERS")
    print("  =============================================================")
    print()
    print("  Export: WP Admin → WooCommerce → Customers → Export")
    print()
    cust_path = input("  Path to WooCommerce customers CSV (Enter to skip):\n  > ").strip().strip('"\'')
    if cust_path and os.path.exists(cust_path):
        print()
        convert_customers(cust_path)
    elif cust_path:
        log("File not found: {}".format(cust_path), "ERROR")

    print("  =============================================================")
    print("  ORDERS")
    print("  =============================================================")
    print()
    print("  Export: WP Admin → WooCommerce → Orders")
    print("          Bulk Actions → Export to CSV")
    print("          (or use 'Advanced Order Export for WooCommerce' plugin)")
    print()
    ord_path = input("  Path to WooCommerce orders CSV (Enter to skip):\n  > ").strip().strip('"\'')
    if ord_path and os.path.exists(ord_path):
        print()
        convert_orders(ord_path)
    elif ord_path:
        log("File not found: {}".format(ord_path), "ERROR")

    print("  =============================================================")
    print("  OUTPUT → {}".format(OUTPUT_DIR))
    print("  =============================================================")
    if os.path.exists(OUTPUT_DIR):
        for f in sorted(os.listdir(OUTPUT_DIR)):
            fp = os.path.join(OUTPUT_DIR, f)
            if os.path.isfile(fp):
                print("    {:>12,} bytes  {}".format(os.path.getsize(fp), f))
    print()
    print("  HOW TO IMPORT:")
    print("  Products  → Shopify Admin → Products → Import → shopify_products.csv")
    print("  Customers → Shopify Admin → Customers → Import → shopify_customers.csv")
    print("  Orders    → Matrixify app → Import → shopify_orders.csv")
    print()

if __name__ == "__main__":
    main()
