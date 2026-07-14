#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 ET RFQ & PO Controller  --  Data Converter
================================================================================
Converts the raw RFQ-to-PO-to-shipment lifecycle Excel/CSV export into a clean,
browser-ready `data.json` consumed by the static dashboard (index.html).

Design goals
------------
* Map by *header text* (normalized), not by fixed column letters, so the script
  keeps working if a column is inserted or a heading is edited slightly.
* Never crash on blank optional columns, stray text in numeric fields, currency
  symbols, percentages, formulas that evaluated to errors, or invalid dates.
* Emit rich, self-describing metadata (date range, distinct entities, statuses,
  data-quality counters, reconciliation totals) so the dashboard needs zero
  server-side logic.
* Preserve original values *and* provide normalized / calculated helper fields.

Usage
-----
    python convert.py
    python convert.py "RFQ_Tracker_-_2024__New_.xlsx"
    python convert.py mydata.csv --sheet "RFQ Tracker" --out data.json

Requires: pandas, openpyxl  (see requirements.txt)
================================================================================
"""

import sys
import os
import re
import json
import glob
import argparse
import datetime as dt
from collections import defaultdict, Counter

try:
    import pandas as pd
except ImportError:
    sys.exit("ERROR: pandas is required.  Run:  pip install -r requirements.txt")


# ------------------------------------------------------------------ #
#  1.  CANONICAL FIELD MAP                                            #
# ------------------------------------------------------------------ #
# key = canonical field name used everywhere in the dashboard.
# value = list of accepted header spellings (compared after normalization).
FIELD_MAP = {
    "rfqSNo":            ["customer rfq s.no.", "customer rfq sno", "rfq s.no", "rfq sno", "s.no"],
    "etPOC":             ["et poc"],
    "custRfqDate":       ["customer rfq date"],
    "custRfqClosingDate":["customer rfq closing date", "rfq closing date"],
    "customer":          ["customer", "customer name"],
    "custPOC":           ["customer poc"],
    "custRfqNo":         ["customer rfq no.", "customer rfq no", "customer rfq number"],
    "lineItem":          ["customer line item", "customer line items", "line item"],
    "productCategory":   ["product category", "category"],
    "drawingNo":         ["drawing no", "drawing no.", "drawing number"],
    "oemPartNo":         ["oem part no", "oem part no.", "oem part number", "part no"],
    "itemDescription":   ["item description", "description"],
    "qty":               ["qty", "quantity"],
    "uom":               ["uom", "unit of measure"],
    "supplierQuoteDate": ["supplier quotation date", "supplier quote date"],
    "supplierName":      ["supplier name", "selected supplier", "supplier"],
    "sector":            ["sector"],
    "supplierQuoteRef":  ["supplier quote ref no.", "supplier quote reference no.",
                          "supplier quote ref no", "supplier quote reference"],
    "supplierRemarks":   ["supplier remarks / lead time", "supplier remarks",
                          "supplier remarks/lead time", "lead time"],
    "supplierUnitPrice": ["supplier unit price $", "supplier unit price usd", "supplier unit price"],
    "supplierTotalPrice":["total supplier price $", "total supplier price usd", "total supplier price"],
    "etQuoteDate":       ["et quotation date", "et quote date"],
    "etQuotedUnitPrice": ["et quoted unit price", "et quote unit price"],
    "etQuotedValue":     ["total et quoted value", "et quoted value", "total et quote value"],
    "etRfqStatus":       ["et rfq status", "rfq status"],
    "etQuoteStatus":     ["et quote status", "quote status"],
    "etQuoteRev":        ["et quotation rev", "et quote rev", "quotation rev", "revision"],
    "custPoDate":        ["customer po date"],
    "custPoNo":          ["customer po no", "customer po no.", "customer po number"],
    "etOaDate":          ["et oa date", "oa date"],
    "grossProfit":       ["gross profit $", "gross profit usd", "gross profit"],
    "grossProfitPct":    ["gross profit %", "gross profit percent", "gp %"],
    # ---- customer delivery timeline (AZ-BC) ----
    "custRequiredDate":  ["customer po required date", "customer required date", "customer po req date"],
    "etPromisedDate":    ["et po promised date", "et promised date"],
    "etRtsDate":         ["et po rts date", "et rts date"],
    "etActualShipDate":  ["et po actual ship date", "et actual ship date"],
    # ---- supplier PO / shipment (BD-BJ) ----
    "supplierPoNo":      ["supplier po#", "supplier po #", "supplier po no", "supplier po number"],
    "poDateToSupplier":  ["po date sent to suppler", "po date sent to supplier"],
    "etPoReqFromSupplier":["et po required date - supplier", "et po required date supplier",
                           "et required date from supplier"],
    "supplierPromisedDate":["supplier po promised date", "supplier promised date"],
    "supplierRtsDate":   ["supplier po rts date", "supplier rts date"],
    "supplierActualShipDate":["supplier po actual ship date", "supplier actual ship date"],
    "shipmentStatus":    ["shipment final status", "final shipment status", "shipment status"],
}

# Columns that hold *dates*
DATE_FIELDS = {
    "custRfqDate", "custRfqClosingDate", "supplierQuoteDate", "etQuoteDate",
    "custPoDate", "etOaDate", "custRequiredDate", "etPromisedDate", "etRtsDate",
    "etActualShipDate", "poDateToSupplier", "etPoReqFromSupplier",
    "supplierPromisedDate", "supplierRtsDate", "supplierActualShipDate",
}
# Columns that hold *numbers*
NUMERIC_FIELDS = {
    "qty", "supplierUnitPrice", "supplierTotalPrice", "etQuotedUnitPrice",
    "etQuotedValue", "grossProfit", "grossProfitPct",
}

# ------------------------------------------------------------------ #
#  2.  STATUS NORMALIZATION                                           #
# ------------------------------------------------------------------ #
RFQ_STATUS_NORM = {
    "quoted": "Quoted", "declined": "Declined", "acknowledged": "Acknowledged",
    "won": "Won", "lost": "Lost", "pending": "Pending", "open": "Open",
    "no quote": "No Quote", "cancelled": "Cancelled", "canceled": "Cancelled",
}
QUOTE_STATUS_NORM = {
    "won": "Won", "lost": "Lost", "declined": "Declined",
    "awaiting response": "Awaiting Response", "pending": "Pending", "open": "Open",
    "received supplier quote": "Received Supplier Quote",
    "awaiting supplier quote": "Awaiting Supplier Quote",
    "under clarification": "Under Clarification",
    "no quote": "No Quote", "cancelled": "Cancelled", "canceled": "Cancelled",
    "quoted": "Quoted", "acknowledged": "Acknowledged",
}
SHIPMENT_STATUS_NORM = {
    "open order": "Open Order", "open": "Open Order", "pending": "Pending",
    "in production": "In Production", "ready to ship": "Ready to Ship",
    "rts": "Ready to Ship", "in transit": "In Transit",
    "partially delivered": "Partially Delivered", "delivered": "Delivered",
    "cancelled": "Cancelled", "canceled": "Cancelled", "on hold": "On Hold",
    "delayed": "Delayed",
}


def norm_header(h):
    """Normalize a header for fuzzy matching."""
    if h is None:
        return ""
    s = str(h).lower().strip()
    s = s.replace("\u00a0", " ")
    s = re.sub(r"\s+", " ", s)
    return s


def clean_str(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    s = str(v).replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if s == "" or s.lower() in ("nan", "none", "null", "#n/a", "n/a", "na", "-", "\\"):
        return None
    return s


def clean_number(v):
    """Parse numbers that may carry $ , % spaces or Excel errors."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        if isinstance(v, float) and pd.isna(v):
            return None
        return float(v)
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null", "n/a", "na", "-"):
        return None
    if s.startswith("#") or "div/0" in s.lower() or "value!" in s.lower() or "ref!" in s.lower():
        return None
    is_pct = s.endswith("%")
    s = s.replace("$", "").replace(",", "").replace("%", "").replace("(", "-").replace(")", "").strip()
    try:
        num = float(s)
        return num / 100.0 if is_pct else num
    except ValueError:
        return None


def clean_date(v):
    """Return ISO YYYY-MM-DD string, or None. Returns ('bad', raw) sentinel via flag."""
    if v is None:
        return None, False
    if isinstance(v, float) and pd.isna(v):
        return None, False
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d"), False
    if isinstance(v, (int, float)):
        # Excel serial date
        try:
            base = dt.datetime(1899, 12, 30)
            return (base + dt.timedelta(days=float(v))).strftime("%Y-%m-%d"), False
        except Exception:
            return None, True
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "null", "n/a", "na", "-", "tbd", "tba"):
        return None, False
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=False)
    if pd.isna(parsed):
        parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        return None, True   # value present but unparseable -> invalid date
    return parsed.strftime("%Y-%m-%d"), False


def days_between(later_iso, earlier_iso):
    if not later_iso or not earlier_iso:
        return None
    try:
        a = dt.date.fromisoformat(later_iso)
        b = dt.date.fromisoformat(earlier_iso)
        return (a - b).days
    except Exception:
        return None


# ------------------------------------------------------------------ #
#  3.  LOAD SOURCE                                                    #
# ------------------------------------------------------------------ #
def find_source():
    for pat in ("*.xlsx", "*.xls", "*.csv"):
        hits = [f for f in glob.glob(pat) if not f.startswith("~$")]
        if hits:
            hits.sort(key=lambda f: os.path.getmtime(f), reverse=True)
            return hits[0]
    return None


def load_dataframe(path, sheet):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(path, dtype=object, encoding=enc, keep_default_na=False, na_values=[""])
            except Exception:
                continue
        return pd.read_csv(path, dtype=object)
    # Excel
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    xls = pd.ExcelFile(path, engine=engine)
    target = sheet
    if target is None or target not in xls.sheet_names:
        # prefer a sheet whose name mentions rfq / tracker, else the largest
        pref = [s for s in xls.sheet_names if re.search(r"rfq|tracker|data", s, re.I)]
        target = pref[0] if pref else xls.sheet_names[-1]
    return pd.read_excel(xls, sheet_name=target, header=0, dtype=object)


def resolve_columns(df):
    """Map canonical field -> actual dataframe column using normalized headers."""
    norm_to_actual = {}
    for col in df.columns:
        norm_to_actual.setdefault(norm_header(col), col)
    resolved, unmatched = {}, []
    for field, aliases in FIELD_MAP.items():
        found = None
        for a in aliases:
            if a in norm_to_actual:
                found = norm_to_actual[a]
                break
        if not found:  # loose contains-match fallback
            for a in aliases:
                for nh, actual in norm_to_actual.items():
                    if a == nh or (len(a) > 4 and a in nh):
                        found = actual
                        break
                if found:
                    break
        if found:
            resolved[field] = found
        else:
            unmatched.append(field)
    # Alt-supplier columns = every column NOT claimed above and not the REV col
    claimed = set(resolved.values())
    alt_cols = [c for c in df.columns if c not in claimed]
    return resolved, unmatched, alt_cols


# ------------------------------------------------------------------ #
#  4.  MAIN CONVERSION                                                #
# ------------------------------------------------------------------ #
def convert(path, sheet, out_path, html_out="index.html"):
    print(f"Reading: {path}")
    df = load_dataframe(path, sheet)
    total_source_rows = len(df)
    resolved, unmatched, alt_cols = resolve_columns(df)

    # The known "real" alternative-supplier headers are single tokens (supplier
    # short names). We keep every unclaimed column as a potential alt-supplier,
    # but skip an obvious REV helper column if it wasn't matched.
    alt_cols = [c for c in alt_cols if norm_header(c) not in ("", "et quotation rev")]

    print(f"Matched {len(resolved)} canonical fields; "
          f"{len(unmatched)} unmatched; {len(alt_cols)} alt-supplier columns.")
    if unmatched:
        print("  Unmatched (will be null in output):", ", ".join(unmatched))

    records = []
    dq = Counter()               # data-quality counters
    invalid_dates = 0
    missing_mandatory = 0

    def get(row, field):
        col = resolved.get(field)
        return row[col] if col is not None else None

    for idx, row in df.iterrows():
        rec = {"_row": int(idx) + 2}   # +2 => human Excel row (header = row 1)

        # ---- text fields ----
        for f in ["etPOC", "customer", "custPOC", "custRfqNo", "productCategory",
                  "drawingNo", "oemPartNo", "itemDescription", "uom", "supplierName",
                  "sector", "supplierQuoteRef", "supplierRemarks", "custPoNo",
                  "supplierPoNo", "rfqSNo", "etQuoteRev", "lineItem"]:
            rec[f] = clean_str(get(row, f))

        # ---- numeric fields ----
        for f in NUMERIC_FIELDS:
            rec[f] = clean_number(get(row, f))

        # ---- date fields ----
        for f in DATE_FIELDS:
            iso, bad = clean_date(get(row, f))
            rec[f] = iso
            if bad:
                invalid_dates += 1
                dq["invalidDate"] += 1

        # ---- normalized statuses (keep raw too) ----
        raw_rfq = clean_str(get(row, "etRfqStatus"))
        raw_q   = clean_str(get(row, "etQuoteStatus"))
        raw_ship = clean_str(get(row, "shipmentStatus"))
        rec["etRfqStatusRaw"] = raw_rfq
        rec["etQuoteStatusRaw"] = raw_q
        rec["shipmentStatusRaw"] = raw_ship
        rec["etRfqStatus"] = RFQ_STATUS_NORM.get(raw_rfq.lower(), raw_rfq) if raw_rfq else None
        rec["etQuoteStatus"] = QUOTE_STATUS_NORM.get(raw_q.lower(), raw_q) if raw_q else None
        rec["shipmentStatus"] = SHIPMENT_STATUS_NORM.get(raw_ship.lower(), raw_ship) if raw_ship else None

        # ---- normalize a couple of noisy categoricals (case only) ----
        if rec["productCategory"]:
            rec["productCategory"] = rec["productCategory"].title() if rec["productCategory"].islower() or rec["productCategory"].isupper() else rec["productCategory"]
        if rec["uom"]:
            u = rec["uom"].strip().lower()
            uom_map = {"ea": "Ea", "each": "Ea", "set": "Set", "kit": "Kit", "pcs": "Pcs",
                       "piece": "Pcs", "meter": "Meter", "mtr": "Meter", "m": "Meter",
                       "ft": "Foot", "foot": "Foot", "roll": "Roll", "lot": "Lot", "bag": "Bag"}
            rec["uom"] = uom_map.get(u, rec["uom"])

        # ---- gross profit % : source stores a fraction (0.26 = 26%). Normalize to percent number.
        if rec["grossProfitPct"] is not None and abs(rec["grossProfitPct"]) <= 5:
            rec["grossProfitPct"] = rec["grossProfitPct"] * 100.0

        # ---- alternative suppliers array ----
        alts = []
        for c in alt_cols:
            val = clean_str(row[c])
            if val is not None:
                alts.append({"supplier": str(c).strip(), "response": val})
        rec["alternativeSuppliers"] = alts
        rec["altSupplierCount"] = len(alts)

        # ---- calculated helper: gross profit (fallback) ----
        if rec["grossProfit"] is None and rec["etQuotedValue"] is not None and rec["supplierTotalPrice"] is not None:
            rec["grossProfitCalc"] = round(rec["etQuotedValue"] - rec["supplierTotalPrice"], 2)
        else:
            rec["grossProfitCalc"] = rec["grossProfit"]

        # ---- calculated response times (calendar days) ----
        rec["rfqResponseDays"] = days_between(rec["etQuoteDate"], rec["custRfqDate"])
        rec["closingVarianceDays"] = days_between(rec["etQuoteDate"], rec["custRfqClosingDate"])
        rec["supplierResponseDays"] = days_between(rec["supplierQuoteDate"], rec["custRfqDate"])
        rec["poToOaDays"] = days_between(rec["etOaDate"], rec["custPoDate"])
        rec["poToSupplierPoDays"] = days_between(rec["poDateToSupplier"], rec["custPoDate"])

        # ---- grouping keys ----
        cust = rec["customer"] or "?"
        if rec["custRfqNo"]:
            rec["rfqKey"] = f"{cust}||{rec['custRfqNo']}"
        elif rec["rfqSNo"]:
            rec["rfqKey"] = f"{cust}||{rec['rfqSNo']}||{rec['custRfqDate'] or ''}"
        else:
            rec["rfqKey"] = f"{cust}||row{rec['_row']}"
        rec["custPoKey"] = f"{cust}||{rec['custPoNo']}" if rec["custPoNo"] else None
        rec["supPoKey"] = f"{rec['supplierName']}||{rec['supplierPoNo']}" if (rec["supplierName"] and rec["supplierPoNo"]) else None
        rec["quoteKey"] = f"{rec['custRfqNo']}||{rec['etQuoteDate']}" if (rec["custRfqNo"] and rec["etQuoteDate"]) else None

        # ---- data-quality flags on the record ----
        if not rec["customer"]:
            dq["missingCustomer"] += 1
        if not rec["custRfqNo"]:
            dq["missingRfqNo"] += 1
        if not rec["etPOC"]:
            dq["missingEtPoc"] += 1
        if not rec["etQuoteStatus"]:
            dq["missingQuoteStatus"] += 1
        if rec["grossProfit"] is not None and rec["grossProfit"] < 0:
            dq["negativeGp"] += 1
        if rec["qty"] is not None and rec["qty"] <= 0:
            dq["nonPositiveQty"] += 1
        # illogical sequence: PO date before quotation date
        if rec["custPoDate"] and rec["etQuoteDate"] and rec["custPoDate"] < rec["etQuoteDate"]:
            dq["poBeforeQuote"] += 1
        # won without PO
        if (rec["etQuoteStatus"] == "Won") and not rec["custPoNo"]:
            dq["wonWithoutPo"] += 1

        # skip fully-empty trailing rows
        if any(rec.get(k) for k in ("customer", "etPOC", "custRfqNo", "etQuoteDate",
                                    "etQuotedValue", "etQuoteStatus", "rfqSNo")):
            records.append(rec)

    processed = len(records)

    # ------------------------------------------------------------------ #
    #  Reconciliation totals (dedup-aware) — matches dashboard JS logic  #
    # ------------------------------------------------------------------ #
    def dedup_total(rows, key, field):
        groups = defaultdict(list)
        for r in rows:
            v = r.get(field)
            if v is not None:
                groups[r[key]].append(v)
        tot = 0.0
        for _, vals in groups.items():
            rounded = {round(x, 2) for x in vals}
            tot += vals[0] if (len(vals) > 1 and len(rounded) == 1) else sum(vals)
        return round(tot, 2)

    won_rows = [r for r in records if r["etQuoteStatus"] == "Won"]
    recon = {
        "sourceRows": total_source_rows,
        "processedRows": processed,
        "uniqueRFQs": len({r["rfqKey"] for r in records}),
        "uniqueQuotedRFQs": len({r["rfqKey"] for r in records if r["etQuoteDate"]}),
        "uniqueWonRFQs": len({r["rfqKey"] for r in won_rows}),
        "uniqueCustomerPOs": len({r["custPoKey"] for r in records if r["custPoKey"]}),
        "uniqueSupplierPOs": len({r["supPoKey"] for r in records if r["supPoKey"]}),
        "totalEtQuotedValue": dedup_total(records, "rfqKey", "etQuotedValue"),
        "totalSupplierCost": dedup_total(records, "rfqKey", "supplierTotalPrice"),
        "totalGrossProfit": dedup_total(records, "rfqKey", "grossProfitCalc"),
        "wonEtQuotedValue": dedup_total(won_rows, "rfqKey", "etQuotedValue"),
        "wonGrossProfit": dedup_total(won_rows, "rfqKey", "grossProfitCalc"),
    }

    # distinct entity lists for slicers
    def distinct(field):
        return sorted({r[field] for r in records if r.get(field)})

    all_dates = [r["etQuoteDate"] for r in records if r["etQuoteDate"]] + \
                [r["custRfqDate"] for r in records if r["custRfqDate"]] + \
                [r["custPoDate"] for r in records if r["custPoDate"]]
    data_min = min(all_dates) if all_dates else None
    data_max = max(all_dates) if all_dates else None

    # field completeness matrix
    completeness = {}
    for f in ["customer", "etPOC", "custPOC", "custRfqNo", "custRfqDate", "etQuoteDate",
              "productCategory", "sector", "supplierName", "supplierTotalPrice",
              "etQuotedValue", "grossProfit", "etQuoteStatus", "custPoNo", "custPoDate",
              "custRequiredDate", "etPromisedDate", "etActualShipDate", "supplierPoNo",
              "supplierPromisedDate", "supplierActualShipDate", "shipmentStatus"]:
        filled = sum(1 for r in records if r.get(f) not in (None, ""))
        completeness[f] = {"filled": filled, "pct": round(100.0 * filled / processed, 1) if processed else 0}

    meta = {
        "generatedAt": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "sourceFile": os.path.basename(path),
        "dataDateMin": data_min,
        "dataDateMax": data_max,
        "reconciliation": recon,
        "distinct": {
            "customers": distinct("customer"),
            "etPOCs": distinct("etPOC"),
            "custPOCs": distinct("custPOC"),
            "suppliers": distinct("supplierName"),
            "sectors": distinct("sector"),
            "productCategories": distinct("productCategory"),
            "rfqStatuses": distinct("etRfqStatus"),
            "quoteStatuses": distinct("etQuoteStatus"),
            "shipmentStatuses": distinct("shipmentStatus"),
        },
        "dataQuality": dict(dq),
        "completeness": completeness,
        "unmatchedFields": unmatched,
        "altSupplierColumns": [str(c).strip() for c in alt_cols],
        "notes": [
            "Financial totals use dedup-aware aggregation: when a Total value repeats "
            "identically across every line of one RFQ/PO it is counted once; otherwise "
            "line values are summed.",
            "All day/delay metrics are CALENDAR days.",
            "Gross Profit % normalized to a percentage number (26.5 == 26.5%).",
            "Customer delivery timeline (AZ-BC) and Supplier PO/shipment (BD-BJ) fields "
            "are mapped but may be empty in this dataset; dashboard shows empty states.",
        ],
    }

    # Shrink: drop null / empty-string / empty-array / 0-count keys per record.
    # The dashboard treats any missing key as null. `_row` and `rfqKey` always kept.
    slim = []
    for r in records:
        o = {}
        for k, v in r.items():
            if k in ("_row", "rfqKey"):
                o[k] = v
                continue
            if v is None or v == "" or v == [] or (k == "altSupplierCount" and v == 0):
                continue
            o[k] = v
        slim.append(o)

    out = {"meta": meta, "records": slim}
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, separators=(",", ":"))

    # ---- build the self-contained, double-click-friendly index.html ----
    data_str = json.dumps(out, ensure_ascii=False, separators=(",", ":"))
    build_index_html(data_str, html_out)

    # ---- console summary ----
    size_mb = os.path.getsize(out_path) / 1024 / 1024
    print("\n================  CONVERSION SUMMARY  ================")
    print(f"  Source file            : {os.path.basename(path)}")
    print(f"  Total rows processed   : {processed:,} (of {total_source_rows:,} source rows)")
    print(f"  Total unique RFQs      : {recon['uniqueRFQs']:,}")
    print(f"  Unique customer RFQ #s : {len({r['custRfqNo'] for r in records if r['custRfqNo']}):,}")
    print(f"  Total customer POs     : {recon['uniqueCustomerPOs']:,}")
    print(f"  Total supplier POs     : {recon['uniqueSupplierPOs']:,}")
    print(f"  Total suppliers        : {len(distinct('supplierName')):,}")
    print(f"  Total customers        : {len(distinct('customer')):,}")
    print(f"  Invalid dates found    : {invalid_dates:,}")
    print(f"  Missing customer name  : {dq.get('missingCustomer', 0):,}")
    print(f"  Missing RFQ number     : {dq.get('missingRfqNo', 0):,}")
    print(f"  Total ET quoted value  : ${recon['totalEtQuotedValue']:,.2f}")
    print(f"  Total supplier cost    : ${recon['totalSupplierCost']:,.2f}")
    print(f"  Total gross profit     : ${recon['totalGrossProfit']:,.2f}")
    print(f"  Data date range        : {data_min}  ->  {data_max}")
    print(f"  Output file            : {os.path.abspath(out_path)}  ({size_mb:.2f} MB)")
    _hsz = os.path.getsize(html_out) / 1024 / 1024 if os.path.exists(html_out) else 0
    print(f"  Dashboard (self-cont.) : {os.path.abspath(html_out)}  ({_hsz:.2f} MB)")
    print("=====================================================\n")
    return out



# ===================================================================
# Embedded front-end assets (base64). Used to regenerate a fully
# self-contained index.html that opens by double-clicking (file://).
# ===================================================================
import base64 as _b64

_SHELL_B64 = (
    "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIiBkYXRhLXRoZW1lPSJsaWdodCI+CjxoZWFkPgogIDxtZXRhIGNoYXJzZXQ9IlVURi04IiAvPgogIDxt"
    "ZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MS4wIiAvPgogIDx0aXRsZT5FVCBSRlEgJmFt"
    "cDsgUE8gQ29udHJvbGxlcjwvdGl0bGU+CiAgPGxpbmsgcmVsPSJwcmVjb25uZWN0IiBocmVmPSJodHRwczovL2NkbmpzLmNsb3VkZmxhcmUuY29tIiAvPgog"
    "IDxsaW5rIHJlbD0ic3R5bGVzaGVldCIgaHJlZj0ic3R5bGVzLmNzcyIgLz4KPC9oZWFkPgo8Ym9keT4KICA8IS0tID09PT09PT09PT09PT09PT09PT09PSBM"
    "T0FESU5HIE9WRVJMQVkgPT09PT09PT09PT09PT09PT09PT09IC0tPgogIDxkaXYgaWQ9ImxvYWRlciIgY2xhc3M9ImxvYWRlciI+CiAgICA8ZGl2IGNsYXNz"
    "PSJsb2FkZXItYm94Ij4KICAgICAgPGRpdiBjbGFzcz0ic3Bpbm5lciI+PC9kaXY+CiAgICAgIDxkaXYgaWQ9ImxvYWRlck1zZyI+TG9hZGluZyBkYXNoYm9h"
    "cmQgZGF0YeKApjwvZGl2PgogICAgPC9kaXY+CiAgPC9kaXY+CgogIDwhLS0gPT09PT09PT09PT09PT09PT09PT09IEhFQURFUiA9PT09PT09PT09PT09PT09"
    "PT09PT0gLS0+CiAgPGhlYWRlciBjbGFzcz0iYXBwLWhlYWRlciI+CiAgICA8ZGl2IGNsYXNzPSJoZWFkZXItbGVmdCI+CiAgICAgIDxkaXYgY2xhc3M9ImJy"
    "YW5kIj4KICAgICAgICA8c3BhbiBjbGFzcz0iYnJhbmQtbWFyayI+RVQ8L3NwYW4+CiAgICAgICAgPGRpdiBjbGFzcz0iYnJhbmQtdGV4dCI+CiAgICAgICAg"
    "ICA8ZGl2IGNsYXNzPSJicmFuZC10aXRsZSI+UkZRICZhbXA7IFBPIENvbnRyb2xsZXI8L2Rpdj4KICAgICAgICAgIDxkaXYgY2xhc3M9ImJyYW5kLXN1YiIg"
    "aWQ9ImRhdGFSYW5nZUxhYmVsIj7igJQ8L2Rpdj4KICAgICAgICA8L2Rpdj4KICAgICAgPC9kaXY+CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJoZWFk"
    "ZXItY2VudGVyIj4KICAgICAgPGRpdiBjbGFzcz0ic2VhcmNoLXdyYXAiPgogICAgICAgIDxzdmcgY2xhc3M9InNlYXJjaC1pY28iIHZpZXdCb3g9IjAgMCAy"
    "NCAyNCIgd2lkdGg9IjE2IiBoZWlnaHQ9IjE2Ij48cGF0aCBmaWxsPSJjdXJyZW50Q29sb3IiIGQ9Ik0xNS41IDE0aC0uNzlsLS4yOC0uMjdhNi41IDYuNSAw"
    "IDEgMC0uNy43bC4yNy4yOHYuNzlsNSA0Ljk5TDIwLjQ5IDE5em0tNiAwQTQuNSA0LjUgMCAxIDEgMTQgOS41IDQuNDkgNC40OSAwIDAgMSA5LjUgMTQiLz48"
    "L3N2Zz4KICAgICAgICA8aW5wdXQgaWQ9Imdsb2JhbFNlYXJjaCIgdHlwZT0idGV4dCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBjdXN0b21lciwgUkZRLCBQTywg"
    "c3VwcGxpZXIsIHBhcnTigKYiIGF1dG9jb21wbGV0ZT0ib2ZmIiAvPgogICAgICAgIDxidXR0b24gaWQ9InNlYXJjaENsZWFyIiBjbGFzcz0ic2VhcmNoLWNs"
    "ZWFyIiB0aXRsZT0iQ2xlYXIgc2VhcmNoIiBoaWRkZW4+JnRpbWVzOzwvYnV0dG9uPgogICAgICA8L2Rpdj4KICAgIDwvZGl2PgoKICAgIDxkaXYgY2xhc3M9"
    "ImhlYWRlci1yaWdodCI+CiAgICAgIDxkaXYgY2xhc3M9ImhlYWRlci1tZXRhIj4KICAgICAgICA8ZGl2PjxzcGFuIGNsYXNzPSJobS1sYWJlbCI+UmVjb3Jk"
    "czwvc3Bhbj48c3BhbiBpZD0icmVjb3JkQ291bnQiIGNsYXNzPSJobS12YWwiPjA8L3NwYW4+PC9kaXY+CiAgICAgICAgPGRpdj48c3BhbiBjbGFzcz0iaG0t"
    "bGFiZWwiPkZpbHRlcnM8L3NwYW4+PHNwYW4gaWQ9ImZpbHRlckNvdW50IiBjbGFzcz0iaG0tdmFsIj4wPC9zcGFuPjwvZGl2PgogICAgICA8L2Rpdj4KICAg"
    "ICAgPGJ1dHRvbiBpZD0icmVzZXRCdG4iIGNsYXNzPSJidG4gYnRuLWdob3N0IiB0aXRsZT0iUmVzZXQgYWxsIGZpbHRlcnMiPlJlc2V0PC9idXR0b24+CiAg"
    "ICAgIDxkaXYgY2xhc3M9ImRyb3Bkb3duIj4KICAgICAgICA8YnV0dG9uIGlkPSJleHBvcnRCdG4iIGNsYXNzPSJidG4gYnRuLWdob3N0IiB0aXRsZT0iRXhw"
    "b3J0Ij5FeHBvcnQg4pa+PC9idXR0b24+CiAgICAgICAgPGRpdiBpZD0iZXhwb3J0TWVudSIgY2xhc3M9ImRyb3Bkb3duLW1lbnUiIGhpZGRlbj4KICAgICAg"
    "ICAgIDxidXR0b24gZGF0YS1leHBvcnQ9InJmcSI+UkZRIGRhdGEgKENTVik8L2J1dHRvbj4KICAgICAgICAgIDxidXR0b24gZGF0YS1leHBvcnQ9InBvIj5D"
    "dXN0b21lciBQTyBkYXRhIChDU1YpPC9idXR0b24+CiAgICAgICAgICA8YnV0dG9uIGRhdGEtZXhwb3J0PSJzdXBwbGllciI+U3VwcGxpZXIgZGF0YSAoQ1NW"
    "KTwvYnV0dG9uPgogICAgICAgICAgPGJ1dHRvbiBkYXRhLWV4cG9ydD0icmlzayI+UmlzayByZWdpc3RlciAoQ1NWKTwvYnV0dG9uPgogICAgICAgICAgPGJ1"
    "dHRvbiBkYXRhLWV4cG9ydD0iZHEiPkRhdGEtcXVhbGl0eSBpc3N1ZXMgKENTVik8L2J1dHRvbj4KICAgICAgICAgIDxidXR0b24gZGF0YS1leHBvcnQ9InN1"
    "bW1hcnkiPk1hbmFnZW1lbnQgc3VtbWFyeSAoSFRNTCk8L2J1dHRvbj4KICAgICAgICAgIDxidXR0b24gZGF0YS1leHBvcnQ9InByaW50Ij5QcmludCAvIFBE"
    "RiB2aWV3PC9idXR0b24+CiAgICAgICAgPC9kaXY+CiAgICAgIDwvZGl2PgogICAgICA8YnV0dG9uIGlkPSJ0aGVtZVRvZ2dsZSIgY2xhc3M9ImJ0biBidG4t"
    "aWNvbiIgdGl0bGU9IlRvZ2dsZSBsaWdodCAvIGRhcmsiPgogICAgICAgIDxzcGFuIGNsYXNzPSJ0aGVtZS1pY28iPuKXkDwvc3Bhbj4KICAgICAgPC9idXR0"
    "b24+CiAgICA8L2Rpdj4KICA8L2hlYWRlcj4KCiAgPCEtLSA9PT09PT09PT09PT09PT09PT09PT0gRklMVEVSIEJBUiA9PT09PT09PT09PT09PT09PT09PT0g"
    "LS0+CiAgPHNlY3Rpb24gY2xhc3M9ImZpbHRlci1iYXIiIGlkPSJmaWx0ZXJCYXIiPgogICAgPGRpdiBjbGFzcz0iZmlsdGVyLWdyaWQiIGlkPSJmaWx0ZXJH"
    "cmlkIj48IS0tIHNsaWNlcnMgaW5qZWN0ZWQgLS0+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjaGlwcy1yb3ciIGlkPSJjaGlwc1JvdyI+PCEtLSBhY3RpdmUg"
    "Y2hpcHMgaW5qZWN0ZWQgLS0+PC9kaXY+CiAgPC9zZWN0aW9uPgoKICA8IS0tID09PT09PT09PT09PT09PT09PT09PSBUQUIgTkFWID09PT09PT09PT09PT09"
    "PT09PT09PSAtLT4KICA8bmF2IGNsYXNzPSJ0YWItbmF2IiBpZD0idGFiTmF2Ij4KICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4gYWN0aXZlIiBkYXRhLXRh"
    "Yj0ib3ZlcnZpZXciPkV4ZWN1dGl2ZSBPdmVydmlldzwvYnV0dG9uPgogICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InJmcSI+UkZRIFRy"
    "YWNrZXI8L2J1dHRvbj4KICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJzdXBwbGllciI+U3VwcGxpZXIgJmFtcDsgUHJvY3VyZW1lbnQ8"
    "L2J1dHRvbj4KICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJwbyI+Q3VzdG9tZXIgUE8gJmFtcDsgRGVsaXZlcnk8L2J1dHRvbj4KICAg"
    "IDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJwb2MiPkVUIFBPQyBQZXJmb3JtYW5jZTwvYnV0dG9uPgogICAgPGJ1dHRvbiBjbGFzcz0idGFi"
    "LWJ0biIgZGF0YS10YWI9ImNvbXBhcmUiPlF1YXJ0ZXJseSAvIFllYXJseTwvYnV0dG9uPgogICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9"
    "ImN1c3RvbWVyIj5DdXN0b21lciBBbmFseXNpczwvYnV0dG9uPgogICAgPGJ1dHRvbiBjbGFzcz0idGFiLWJ0biIgZGF0YS10YWI9InJpc2siPlJpc2sgQW5h"
    "bHlzaXM8L2J1dHRvbj4KICAgIDxidXR0b24gY2xhc3M9InRhYi1idG4iIGRhdGEtdGFiPSJkcSI+RGF0YSBRdWFsaXR5ICZhbXA7IENvbnRyb2xzPC9idXR0"
    "b24+CiAgPC9uYXY+CgogIDwhLS0gPT09PT09PT09PT09PT09PT09PT09IE1BSU4gPT09PT09PT09PT09PT09PT09PT09IC0tPgogIDxtYWluIGNsYXNzPSJh"
    "cHAtbWFpbiI+CiAgICA8IS0tIHBhbmVscyBpbmplY3RlZC90b2dnbGVkIGJ5IEpTIC0tPgogICAgPHNlY3Rpb24gY2xhc3M9InRhYi1wYW5lbCBhY3RpdmUi"
    "IGlkPSJ0YWItb3ZlcnZpZXciPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGNsYXNzPSJ0YWItcGFuZWwiIGlkPSJ0YWItcmZxIj48L3NlY3Rpb24+CiAgICA8"
    "c2VjdGlvbiBjbGFzcz0idGFiLXBhbmVsIiBpZD0idGFiLXN1cHBsaWVyIj48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBjbGFzcz0idGFiLXBhbmVsIiBpZD0i"
    "dGFiLXBvIj48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBjbGFzcz0idGFiLXBhbmVsIiBpZD0idGFiLXBvYyI+PC9zZWN0aW9uPgogICAgPHNlY3Rpb24gY2xh"
    "c3M9InRhYi1wYW5lbCIgaWQ9InRhYi1jb21wYXJlIj48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBjbGFzcz0idGFiLXBhbmVsIiBpZD0idGFiLWN1c3RvbWVy"
    "Ij48L3NlY3Rpb24+CiAgICA8c2VjdGlvbiBjbGFzcz0idGFiLXBhbmVsIiBpZD0idGFiLXJpc2siPjwvc2VjdGlvbj4KICAgIDxzZWN0aW9uIGNsYXNzPSJ0"
    "YWItcGFuZWwiIGlkPSJ0YWItZHEiPjwvc2VjdGlvbj4KICA8L21haW4+CgogIDxmb290ZXIgY2xhc3M9ImFwcC1mb290ZXIiPgogICAgPHNwYW4gaWQ9ImZv"
    "b3RlclJlZnJlc2giPuKAlDwvc3Bhbj4KICAgIDxzcGFuIGNsYXNzPSJmb290LXNlcCI+4oCiPC9zcGFuPgogICAgPHNwYW4+QWxsIHZhbHVlcyBVU0Qgwrcg"
    "Y2FsZW5kYXItZGF5IG1ldHJpY3MgwrcgZGVkdXAtYXdhcmUgYWdncmVnYXRpb248L3NwYW4+CiAgICA8c3BhbiBjbGFzcz0iZm9vdC1zZXAiPuKAojwvc3Bh"
    "bj4KICAgIDxzcGFuPkVUIFJGUSAmYW1wOyBQTyBDb250cm9sbGVyPC9zcGFuPgogIDwvZm9vdGVyPgoKICA8IS0tID09PT09PT09PT09PT09PT09PT09PSBE"
    "UklMTC1ET1dOIE1PREFMID09PT09PT09PT09PT09PT09PT09PSAtLT4KICA8ZGl2IGlkPSJtb2RhbCIgY2xhc3M9Im1vZGFsIiBoaWRkZW4+CiAgICA8ZGl2"
    "IGNsYXNzPSJtb2RhbC1iYWNrZHJvcCIgZGF0YS1jbG9zZT48L2Rpdj4KICAgIDxkaXYgY2xhc3M9Im1vZGFsLWJveCI+CiAgICAgIDxkaXYgY2xhc3M9Im1v"
    "ZGFsLWhlYWQiPgogICAgICAgIDxoMyBpZD0ibW9kYWxUaXRsZSI+UmVjb3JkIGRldGFpbDwvaDM+CiAgICAgICAgPGJ1dHRvbiBjbGFzcz0iYnRuIGJ0bi1p"
    "Y29uIiBkYXRhLWNsb3NlPiZ0aW1lczs8L2J1dHRvbj4KICAgICAgPC9kaXY+CiAgICAgIDxkaXYgaWQ9Im1vZGFsQm9keSIgY2xhc3M9Im1vZGFsLWJvZHki"
    "PjwvZGl2PgogICAgPC9kaXY+CiAgPC9kaXY+CgogIDwhLS0gQ2hhcnQuanMgLS0+CiAgPHNjcmlwdCBzcmM9Imh0dHBzOi8vY2RuanMuY2xvdWRmbGFyZS5j"
    "b20vYWpheC9saWJzL0NoYXJ0LmpzLzQuNC4xL2NoYXJ0LnVtZC5taW4uanMiPjwvc2NyaXB0PgogIDxzY3JpcHQgc3JjPSJzY3JpcHQuanMiPjwvc2NyaXB0"
    "Pgo8L2JvZHk+CjwvaHRtbD4K"
)

_CSS_B64 = (
    "LyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBFVCBSRlEgJiBQTyBDb250cm9s"
    "bGVyIOKAlCBzdHlsZXMuY3NzCiAgIFByZW1pdW0gY29ycG9yYXRlIGRhc2hib2FyZCDCtyBsaWdodCArIGRhcmsgdGhlbWVzCiAgID09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KCjpyb290IHsKICAtLWJnOiAgICAgICAgI2VlZjFmNjsKICAt"
    "LXN1cmZhY2U6ICAgI2ZmZmZmZjsKICAtLXN1cmZhY2UtMjogI2Y3ZjlmYzsKICAtLXRleHQ6ICAgICAgIzE2MjAyZTsKICAtLW11dGVkOiAgICAgIzY0NzQ4"
    "YjsKICAtLWJvcmRlcjogICAgI2UyZThmMDsKICAtLWJvcmRlci0yOiAgI2NiZDVlMTsKICAtLXByaW1hcnk6ICAgIzJmNWJkNDsKICAtLXByaW1hcnktMjog"
    "IzRmNDZlNTsKICAtLWFjY2VudDogICAgIzBkOTQ4ODsKICAtLWdvb2Q6ICAgICAgIzE1ODg0YzsKICAtLWdvb2QtYmc6ICAgI2U2ZjRlYzsKICAtLXdhcm46"
    "ICAgICAgI2MwNzgwNjsKICAtLXdhcm4tYmc6ICAgI2ZkZjFkZTsKICAtLWJhZDogICAgICAgI2NmMzIzMjsKICAtLWJhZC1iZzogICAgI2ZiZTllOTsKICAt"
    "LWluZm86ICAgICAgIzJmNWJkNDsKICAtLXNoYWRvdzogICAgMCAxcHggMnB4IHJnYmEoMTYsMzIsNTQsLjA2KSwgMCA0cHggMTZweCByZ2JhKDE2LDMyLDU0"
    "LC4wNSk7CiAgLS1zaGFkb3ctbGc6IDAgOHB4IDQwcHggcmdiYSgxNiwzMiw1NCwuMTYpOwogIC0tcmFkaXVzOiAgICAxMnB4OwogIC0tcmFkaXVzLXNtOiA4"
    "cHg7CiAgLS1ncmlkOiAgICAgIHJnYmEoMTAwLDExNiwxMzksLjE2KTsKICAtLWZvbnQ6IC1hcHBsZS1zeXN0ZW0sIEJsaW5rTWFjU3lzdGVtRm9udCwgIlNl"
    "Z29lIFVJIiwgUm9ib3RvLCAiSGVsdmV0aWNhIE5ldWUiLCBBcmlhbCwgc2Fucy1zZXJpZjsKfQpbZGF0YS10aGVtZT0iZGFyayJdIHsKICAtLWJnOiAgICAg"
    "ICAgIzBkMTIxOTsKICAtLXN1cmZhY2U6ICAgIzE2MWQyOTsKICAtLXN1cmZhY2UtMjogIzFjMjUzMzsKICAtLXRleHQ6ICAgICAgI2U3ZWRmNTsKICAtLW11"
    "dGVkOiAgICAgIzkzYTFiNTsKICAtLWJvcmRlcjogICAgIzI3MzE0MDsKICAtLWJvcmRlci0yOiAgIzMzNDA0ZjsKICAtLXByaW1hcnk6ICAgIzViOGJmZjsK"
    "ICAtLXByaW1hcnktMjogIzhiOGJmZjsKICAtLWFjY2VudDogICAgIzJkZDRiZjsKICAtLWdvb2Q6ICAgICAgIzM1YzA3YzsKICAtLWdvb2QtYmc6ICAgIzEz"
    "MzIyNjsKICAtLXdhcm46ICAgICAgI2UwYTYzYzsKICAtLXdhcm4tYmc6ICAgIzNhMmMxMjsKICAtLWJhZDogICAgICAgI2YwNjE2YTsKICAtLWJhZC1iZzog"
    "ICAgIzNhMWMxZjsKICAtLWluZm86ICAgICAgIzViOGJmZjsKICAtLXNoYWRvdzogICAgMCAxcHggMnB4IHJnYmEoMCwwLDAsLjMpLCAwIDRweCAxNnB4IHJn"
    "YmEoMCwwLDAsLjI4KTsKICAtLXNoYWRvdy1sZzogMCA4cHggNDBweCByZ2JhKDAsMCwwLC41KTsKICAtLWdyaWQ6ICAgICAgcmdiYSgxNDcsMTYxLDE4MSwu"
    "MTQpOwp9CgoqIHsgYm94LXNpemluZzogYm9yZGVyLWJveDsgfQpodG1sLCBib2R5IHsgbWFyZ2luOiAwOyBwYWRkaW5nOiAwOyB9CmJvZHkgewogIGZvbnQt"
    "ZmFtaWx5OiB2YXIoLS1mb250KTsKICBiYWNrZ3JvdW5kOiB2YXIoLS1iZyk7CiAgY29sb3I6IHZhcigtLXRleHQpOwogIGZvbnQtc2l6ZTogMTRweDsKICBs"
    "aW5lLWhlaWdodDogMS40NTsKICAtd2Via2l0LWZvbnQtc21vb3RoaW5nOiBhbnRpYWxpYXNlZDsKfQpidXR0b24geyBmb250LWZhbWlseTogaW5oZXJpdDsg"
    "Y3Vyc29yOiBwb2ludGVyOyB9Cjo6LXdlYmtpdC1zY3JvbGxiYXIgeyBoZWlnaHQ6IDEwcHg7IHdpZHRoOiAxMHB4OyB9Cjo6LXdlYmtpdC1zY3JvbGxiYXIt"
    "dGh1bWIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1ib3JkZXItMik7IGJvcmRlci1yYWRpdXM6IDZweDsgfQo6Oi13ZWJraXQtc2Nyb2xsYmFyLXRyYWNrIHsgYmFj"
    "a2dyb3VuZDogdHJhbnNwYXJlbnQ7IH0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gbG9hZGVyIC0tLS0tLS0tLS0tLS0tLS0gKi8KLmxvYWRlciB7CiAgcG9zaXRp"
    "b246IGZpeGVkOyBpbnNldDogMDsgei1pbmRleDogOTk5OTsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6"
    "IGNlbnRlcjsKICBiYWNrZ3JvdW5kOiB2YXIoLS1iZyk7Cn0KLmxvYWRlci1ib3ggeyB0ZXh0LWFsaWduOiBjZW50ZXI7IGNvbG9yOiB2YXIoLS1tdXRlZCk7"
    "IH0KLnNwaW5uZXIgewogIHdpZHRoOiA0MnB4OyBoZWlnaHQ6IDQycHg7IG1hcmdpbjogMCBhdXRvIDE0cHg7CiAgYm9yZGVyOiAzcHggc29saWQgdmFyKC0t"
    "Ym9yZGVyKTsgYm9yZGVyLXRvcC1jb2xvcjogdmFyKC0tcHJpbWFyeSk7CiAgYm9yZGVyLXJhZGl1czogNTAlOyBhbmltYXRpb246IHNwaW4gLjhzIGxpbmVh"
    "ciBpbmZpbml0ZTsKfQpAa2V5ZnJhbWVzIHNwaW4geyB0byB7IHRyYW5zZm9ybTogcm90YXRlKDM2MGRlZyk7IH0gfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBo"
    "ZWFkZXIgLS0tLS0tLS0tLS0tLS0tLSAqLwouYXBwLWhlYWRlciB7CiAgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyB6LWluZGV4OiAyMDA7CiAgZGlzcGxh"
    "eTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxNnB4OwogIHBhZGRpbmc6IDEwcHggMThweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNl"
    "KTsKICBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3cpOwp9Ci5oZWFkZXItbGVmdCB7"
    "IGZsZXg6IDAgMCBhdXRvOyB9Ci5oZWFkZXItY2VudGVyIHsgZmxleDogMSAxIGF1dG87IGRpc3BsYXk6IGZsZXg7IGp1c3RpZnktY29udGVudDogY2VudGVy"
    "OyB9Ci5oZWFkZXItcmlnaHQgeyBmbGV4OiAwIDAgYXV0bzsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiA4cHg7IH0KLmJyYW5k"
    "IHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2FwOiAxMXB4OyB9Ci5icmFuZC1tYXJrIHsKICB3aWR0aDogMzhweDsgaGVpZ2h0OiAz"
    "OHB4OyBib3JkZXItcmFkaXVzOiA5cHg7CiAgYmFja2dyb3VuZDogbGluZWFyLWdyYWRpZW50KDEzNWRlZywgdmFyKC0tcHJpbWFyeSksIHZhcigtLXByaW1h"
    "cnktMikpOwogIGNvbG9yOiAjZmZmOyBmb250LXdlaWdodDogODAwOyBsZXR0ZXItc3BhY2luZzogLjVweDsKICBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVt"
    "czogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsgZm9udC1zaXplOiAxNXB4Owp9Ci5icmFuZC10aXRsZSB7IGZvbnQtd2VpZ2h0OiA3MDA7IGZv"
    "bnQtc2l6ZTogMTVweDsgbGV0dGVyLXNwYWNpbmc6IC0uMnB4OyB9Ci5icmFuZC1zdWIgeyBmb250LXNpemU6IDExLjVweDsgY29sb3I6IHZhcigtLW11dGVk"
    "KTsgfQoKLnNlYXJjaC13cmFwIHsKICBwb3NpdGlvbjogcmVsYXRpdmU7IHdpZHRoOiBtaW4oNDYwcHgsIDQ2dncpOwogIGRpc3BsYXk6IGZsZXg7IGFsaWdu"
    "LWl0ZW1zOiBjZW50ZXI7Cn0KLnNlYXJjaC1pY28geyBwb3NpdGlvbjogYWJzb2x1dGU7IGxlZnQ6IDExcHg7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IHBvaW50"
    "ZXItZXZlbnRzOiBub25lOyB9CiNnbG9iYWxTZWFyY2ggewogIHdpZHRoOiAxMDAlOyBwYWRkaW5nOiA5cHggMzBweCA5cHggMzRweDsKICBib3JkZXI6IDFw"
    "eCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOiA5OTlweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0t"
    "dGV4dCk7IGZvbnQtc2l6ZTogMTNweDsKICB0cmFuc2l0aW9uOiBib3JkZXItY29sb3IgLjE1cywgYm94LXNoYWRvdyAuMTVzOwp9CiNnbG9iYWxTZWFyY2g6"
    "Zm9jdXMgeyBvdXRsaW5lOiBub25lOyBib3JkZXItY29sb3I6IHZhcigtLXByaW1hcnkpOyBib3gtc2hhZG93OiAwIDAgMCAzcHggcmdiYSg0Nyw5MSwyMTIs"
    "LjE1KTsgfQouc2VhcmNoLWNsZWFyIHsKICBwb3NpdGlvbjogYWJzb2x1dGU7IHJpZ2h0OiA4cHg7IGJvcmRlcjogbm9uZTsgYmFja2dyb3VuZDogdmFyKC0t"
    "Ym9yZGVyKTsKICBjb2xvcjogdmFyKC0tdGV4dCk7IHdpZHRoOiAyMHB4OyBoZWlnaHQ6IDIwcHg7IGJvcmRlci1yYWRpdXM6IDUwJTsKICBmb250LXNpemU6"
    "IDE1cHg7IGxpbmUtaGVpZ2h0OiAxOyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6IGNlbnRlcjsKfQouaGVh"
    "ZGVyLW1ldGEgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDE0cHg7IG1hcmdpbi1yaWdodDogNHB4OyB9Ci5oZWFkZXItbWV0YSA+IGRpdiB7IGRpc3BsYXk6IGZs"
    "ZXg7IGZsZXgtZGlyZWN0aW9uOiBjb2x1bW47IGFsaWduLWl0ZW1zOiBmbGV4LWVuZDsgbGluZS1oZWlnaHQ6IDEuMTsgfQouaG0tbGFiZWwgeyBmb250LXNp"
    "emU6IDkuNXB4OyB0ZXh0LXRyYW5zZm9ybTogdXBwZXJjYXNlOyBsZXR0ZXItc3BhY2luZzogLjZweDsgY29sb3I6IHZhcigtLW11dGVkKTsgfQouaG0tdmFs"
    "IHsgZm9udC13ZWlnaHQ6IDcwMDsgZm9udC1zaXplOiAxNHB4OyB9CgouYnRuIHsKICBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBiYWNrZ3Jv"
    "dW5kOiB2YXIoLS1zdXJmYWNlLTIpOwogIGNvbG9yOiB2YXIoLS10ZXh0KTsgcGFkZGluZzogN3B4IDEycHg7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1"
    "cy1zbSk7CiAgZm9udC1zaXplOiAxMi41cHg7IGZvbnQtd2VpZ2h0OiA2MDA7IHRyYW5zaXRpb246IGJhY2tncm91bmQgLjE1cywgYm9yZGVyLWNvbG9yIC4x"
    "NXM7Cn0KLmJ0bjpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLWJvcmRlcik7IH0KLmJ0bi1naG9zdCB7IGJhY2tncm91bmQ6IHRyYW5zcGFyZW50OyB9Ci5i"
    "dG4taWNvbiB7IHBhZGRpbmc6IDdweCA5cHg7IGZvbnQtc2l6ZTogMTVweDsgfQoudGhlbWUtaWNvIHsgZGlzcGxheTogaW5saW5lLWJsb2NrOyB9CgouZHJv"
    "cGRvd24geyBwb3NpdGlvbjogcmVsYXRpdmU7IH0KLmRyb3Bkb3duLW1lbnUgewogIHBvc2l0aW9uOiBhYnNvbHV0ZTsgcmlnaHQ6IDA7IHRvcDogY2FsYygx"
    "MDAlICsgNnB4KTsgei1pbmRleDogMzAwOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwog"
    "IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cy1zbSk7IGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1sZyk7CiAgbWluLXdpZHRoOiAyMTBweDsgcGFkZGlu"
    "ZzogNnB4OyBkaXNwbGF5OiBmbGV4OyBmbGV4LWRpcmVjdGlvbjogY29sdW1uOwp9Ci5kcm9wZG93bi1tZW51IGJ1dHRvbiB7CiAgdGV4dC1hbGlnbjogbGVm"
    "dDsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZhcigtLXRleHQpOwogIHBhZGRpbmc6IDhweCAxMHB4OyBib3JkZXIt"
    "cmFkaXVzOiA2cHg7IGZvbnQtc2l6ZTogMTIuNXB4Owp9Ci5kcm9wZG93bi1tZW51IGJ1dHRvbjpob3ZlciB7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2Ut"
    "Mik7IH0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gZmlsdGVyIGJhciAtLS0tLS0tLS0tLS0tLS0tICovCi5maWx0ZXItYmFyIHsKICBwb3NpdGlvbjogc3RpY2t5"
    "OyB0b3A6IDU5cHg7IHotaW5kZXg6IDE1MDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyLWJvdHRvbTogMXB4IHNvbGlkIHZhcigtLWJv"
    "cmRlcik7CiAgcGFkZGluZzogMTBweCAxOHB4Owp9Ci5maWx0ZXItZ3JpZCB7CiAgZGlzcGxheTogZ3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOiByZXBl"
    "YXQoYXV0by1maWxsLCBtaW5tYXgoMTU4cHgsIDFmcikpOwogIGdhcDogOHB4Owp9Ci5zbGljZXIgeyBwb3NpdGlvbjogcmVsYXRpdmU7IH0KLnNsaWNlci1s"
    "YWJlbCB7CiAgZm9udC1zaXplOiA5LjVweDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IC41cHg7CiAgY29sb3I6IHZhcigt"
    "LW11dGVkKTsgbWFyZ2luLWJvdHRvbTogM3B4OyBmb250LXdlaWdodDogNzAwOwp9Ci5zbGljZXIgc2VsZWN0LCAuc2xpY2VyIGlucHV0IHsKICB3aWR0aDog"
    "MTAwJTsgcGFkZGluZzogNnB4IDhweDsgZm9udC1zaXplOiAxMnB4OwogIGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6"
    "IDdweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dCk7Cn0KLnNsaWNlciBzZWxlY3Q6Zm9jdXMsIC5zbGljZXIg"
    "aW5wdXQ6Zm9jdXMgeyBvdXRsaW5lOiBub25lOyBib3JkZXItY29sb3I6IHZhcigtLXByaW1hcnkpOyB9Ci5tcy10b2dnbGUgewogIHdpZHRoOiAxMDAlOyBw"
    "YWRkaW5nOiA2cHggOHB4OyBmb250LXNpemU6IDEycHg7IHRleHQtYWxpZ246IGxlZnQ7CiAgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9y"
    "ZGVyLXJhZGl1czogN3B4OyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOwogIGNvbG9yOiB2YXIoLS10ZXh0KTsgZGlzcGxheTogZmxleDsganVzdGlm"
    "eS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDRweDsKfQoubXMtdG9nZ2xlIC5jbnQgeyBjb2xvcjogdmFyKC0t"
    "bXV0ZWQpOyBmb250LXNpemU6IDExcHg7IH0KLm1zLXBhbmVsIHsKICBwb3NpdGlvbjogYWJzb2x1dGU7IHotaW5kZXg6IDQwMDsgdG9wOiBjYWxjKDEwMCUg"
    "KyAzcHgpOyBsZWZ0OiAwOyByaWdodDogMDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsg"
    "Ym9yZGVyLXJhZGl1czogOHB4OwogIGJveC1zaGFkb3c6IHZhcigtLXNoYWRvdy1sZyk7IG1heC1oZWlnaHQ6IDI2MHB4OyBvdmVyZmxvdzogYXV0bzsgcGFk"
    "ZGluZzogNnB4Owp9Ci5tcy1wYW5lbCBpbnB1dC5tcy1zZWFyY2ggeyB3aWR0aDogMTAwJTsgbWFyZ2luLWJvdHRvbTogNXB4OyB9Ci5tcy1vcHQgeyBkaXNw"
    "bGF5OiBmbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBnYXA6IDdweDsgcGFkZGluZzogNHB4IDZweDsgYm9yZGVyLXJhZGl1czogNXB4OyBmb250LXNpemU6"
    "IDEycHg7IGN1cnNvcjogcG9pbnRlcjsgfQoubXMtb3B0OmhvdmVyIHsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgfQoubXMtb3B0IGlucHV0IHsg"
    "d2lkdGg6IGF1dG87IH0KCi5jaGlwcy1yb3cgeyBkaXNwbGF5OiBmbGV4OyBmbGV4LXdyYXA6IHdyYXA7IGdhcDogNnB4OyBtYXJnaW4tdG9wOiA4cHg7IH0K"
    "LmNoaXBzLXJvdzplbXB0eSB7IGRpc3BsYXk6IG5vbmU7IH0KLmNoaXAgewogIGRpc3BsYXk6IGlubGluZS1mbGV4OyBhbGlnbi1pdGVtczogY2VudGVyOyBn"
    "YXA6IDZweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBhZGRpbmc6IDNweCA4"
    "cHg7IGJvcmRlci1yYWRpdXM6IDk5OXB4OyBmb250LXNpemU6IDExLjVweDsgY29sb3I6IHZhcigtLXRleHQpOwp9Ci5jaGlwIGIgeyBjb2xvcjogdmFyKC0t"
    "cHJpbWFyeSk7IGZvbnQtd2VpZ2h0OiA3MDA7IH0KLmNoaXAgYnV0dG9uIHsgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6"
    "IHZhcigtLW11dGVkKTsgZm9udC1zaXplOiAxNHB4OyBsaW5lLWhlaWdodDogMTsgcGFkZGluZzogMDsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSB0YWIgbmF2"
    "IC0tLS0tLS0tLS0tLS0tLS0gKi8KLnRhYi1uYXYgewogIGRpc3BsYXk6IGZsZXg7IGdhcDogMnB4OyBvdmVyZmxvdy14OiBhdXRvOwogIHBhZGRpbmc6IDAg"
    "MTJweDsgYmFja2dyb3VuZDogdmFyKC0tYmcpOwogIGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIHBvc2l0aW9uOiBzdGlja3k7"
    "IHRvcDogMTIycHg7IHotaW5kZXg6IDE0MDsKfQoudGFiLWJ0biB7CiAgYm9yZGVyOiBub25lOyBiYWNrZ3JvdW5kOiB0cmFuc3BhcmVudDsgY29sb3I6IHZh"
    "cigtLW11dGVkKTsKICBwYWRkaW5nOiAxMnB4IDE1cHg7IGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDYwMDsgd2hpdGUtc3BhY2U6IG5vd3JhcDsK"
    "ICBib3JkZXItYm90dG9tOiAyLjVweCBzb2xpZCB0cmFuc3BhcmVudDsgdHJhbnNpdGlvbjogY29sb3IgLjE1cywgYm9yZGVyLWNvbG9yIC4xNXM7Cn0KLnRh"
    "Yi1idG46aG92ZXIgeyBjb2xvcjogdmFyKC0tdGV4dCk7IH0KLnRhYi1idG4uYWN0aXZlIHsgY29sb3I6IHZhcigtLXByaW1hcnkpOyBib3JkZXItYm90dG9t"
    "LWNvbG9yOiB2YXIoLS1wcmltYXJ5KTsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBtYWluIC8gcGFuZWxzIC0tLS0tLS0tLS0tLS0tLS0gKi8KLmFwcC1tYWlu"
    "IHsgcGFkZGluZzogMThweDsgbWF4LXdpZHRoOiAxNjAwcHg7IG1hcmdpbjogMCBhdXRvOyB9Ci50YWItcGFuZWwgeyBkaXNwbGF5OiBub25lOyB9Ci50YWIt"
    "cGFuZWwuYWN0aXZlIHsgZGlzcGxheTogYmxvY2s7IGFuaW1hdGlvbjogZmFkZSAuMnMgZWFzZTsgfQpAa2V5ZnJhbWVzIGZhZGUgeyBmcm9tIHsgb3BhY2l0"
    "eTogMDsgdHJhbnNmb3JtOiB0cmFuc2xhdGVZKDRweCk7IH0gdG8geyBvcGFjaXR5OiAxOyB0cmFuc2Zvcm06IG5vbmU7IH0gfQoKLnNlY3Rpb24tdGl0bGUg"
    "ewogIGZvbnQtc2l6ZTogMTNweDsgZm9udC13ZWlnaHQ6IDcwMDsgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IC41cHg7CiAg"
    "Y29sb3I6IHZhcigtLW11dGVkKTsgbWFyZ2luOiAyMnB4IDJweCAxMHB4Owp9Ci5zZWN0aW9uLXRpdGxlOmZpcnN0LWNoaWxkIHsgbWFyZ2luLXRvcDogNHB4"
    "OyB9CgovKiAtLS0tLS0tLS0tLS0tLS0tIEtQSSBjYXJkcyAtLS0tLS0tLS0tLS0tLS0tICovCi5rcGktZ3JpZCB7CiAgZGlzcGxheTogZ3JpZDsgZ3JpZC10"
    "ZW1wbGF0ZS1jb2x1bW5zOiByZXBlYXQoYXV0by1maWxsLCBtaW5tYXgoMTgwcHgsIDFmcikpOyBnYXA6IDExcHg7Cn0KLmtwaSB7CiAgYmFja2dyb3VuZDog"
    "dmFyKC0tc3VyZmFjZSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzKTsgcGFkZGluZzog"
    "MTNweCAxNHB4OyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3cpOwogIHBvc2l0aW9uOiByZWxhdGl2ZTsgb3ZlcmZsb3c6IGhpZGRlbjsKfQoua3BpLWxhYmVs"
    "IHsgZm9udC1zaXplOiAxMXB4OyBjb2xvcjogdmFyKC0tbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyBtYXJnaW4tYm90dG9tOiA2cHg7IGxldHRlci1zcGFj"
    "aW5nOiAuMXB4OyB9Ci5rcGktdmFsdWUgeyBmb250LXNpemU6IDIycHg7IGZvbnQtd2VpZ2h0OiA3NTA7IGxldHRlci1zcGFjaW5nOiAtLjVweDsgbGluZS1o"
    "ZWlnaHQ6IDEuMTsgfQoua3BpLXN1YiB7IGZvbnQtc2l6ZTogMTFweDsgY29sb3I6IHZhcigtLW11dGVkKTsgbWFyZ2luLXRvcDogNHB4OyB9Ci5rcGkuZ29v"
    "ZCAgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLWdvb2QpOyB9Ci5rcGkud2FybiAgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLXdhcm4p"
    "OyB9Ci5rcGkuYmFkICAgeyBib3JkZXItbGVmdDogM3B4IHNvbGlkIHZhcigtLWJhZCk7IH0KLmtwaS5nb29kIC5rcGktdmFsdWUgIHsgY29sb3I6IHZhcigt"
    "LWdvb2QpOyB9Ci5rcGkud2FybiAua3BpLXZhbHVlICB7IGNvbG9yOiB2YXIoLS13YXJuKTsgfQoua3BpLmJhZCAgLmtwaS12YWx1ZSAgeyBjb2xvcjogdmFy"
    "KC0tYmFkKTsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBjaGFydCBjYXJkcyAtLS0tLS0tLS0tLS0tLS0tICovCi5jaGFydC1ncmlkIHsgZGlzcGxheTogZ3Jp"
    "ZDsgZ2FwOiAxM3B4OyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdCgxMiwgMWZyKTsgfQouY2FyZCB7CiAgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFj"
    "ZSk7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzKTsgcGFkZGluZzogMTRweCAxNXB4OyBi"
    "b3gtc2hhZG93OiB2YXIoLS1zaGFkb3cpOwp9Ci5jYXJkLmM2IHsgZ3JpZC1jb2x1bW46IHNwYW4gNjsgfQouY2FyZC5jNCB7IGdyaWQtY29sdW1uOiBzcGFu"
    "IDQ7IH0KLmNhcmQuYzggeyBncmlkLWNvbHVtbjogc3BhbiA4OyB9Ci5jYXJkLmMxMiB7IGdyaWQtY29sdW1uOiBzcGFuIDEyOyB9Ci5jYXJkLmMzIHsgZ3Jp"
    "ZC1jb2x1bW46IHNwYW4gMzsgfQouY2FyZC1oZWFkIHsgZGlzcGxheTogZmxleDsganVzdGlmeS1jb250ZW50OiBzcGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVt"
    "czogY2VudGVyOyBtYXJnaW4tYm90dG9tOiAxMHB4OyB9Ci5jYXJkLXRpdGxlIHsgZm9udC1zaXplOiAxM3B4OyBmb250LXdlaWdodDogNzAwOyB9Ci5jYXJk"
    "LWhpbnQgeyBmb250LXNpemU6IDEwLjVweDsgY29sb3I6IHZhcigtLW11dGVkKTsgfQouY2hhcnQtaG9sZGVyIHsgcG9zaXRpb246IHJlbGF0aXZlOyBoZWln"
    "aHQ6IDI2MHB4OyB9Ci5jaGFydC1ob2xkZXIudGFsbCB7IGhlaWdodDogMzIwcHg7IH0KLmNoYXJ0LWhvbGRlci5zaG9ydCB7IGhlaWdodDogMjEwcHg7IH0K"
    "Ci8qIC0tLS0tLS0tLS0tLS0tLS0gdGFibGVzIC0tLS0tLS0tLS0tLS0tLS0gKi8KLnRhYmxlLXRvb2xzIHsgZGlzcGxheTogZmxleDsgZmxleC13cmFwOiB3"
    "cmFwOyBnYXA6IDhweDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgbWFyZ2luLWJvdHRvbTogMTBweDsgfQoudGFibGUtdG9vbHMgaW5wdXRbdHlwZT10ZXh0XSB7"
    "CiAgcGFkZGluZzogN3B4IDEwcHg7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IDdweDsKICBiYWNrZ3JvdW5kOiB2"
    "YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dCk7IGZvbnQtc2l6ZTogMTIuNXB4OyBtaW4td2lkdGg6IDIwMHB4Owp9Ci50YWJsZS10b29scyBz"
    "ZWxlY3QgewogIHBhZGRpbmc6IDdweCA4cHg7IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IDdweDsKICBiYWNrZ3Jv"
    "dW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBjb2xvcjogdmFyKC0tdGV4dCk7IGZvbnQtc2l6ZTogMTJweDsKfQoudGFibGUtc2Nyb2xsIHsgb3ZlcmZsb3c6IGF1"
    "dG87IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cyk7IG1heC1oZWlnaHQ6IDYyMHB4OyB9CnRh"
    "YmxlLmRhdGEgeyB3aWR0aDogMTAwJTsgYm9yZGVyLWNvbGxhcHNlOiBjb2xsYXBzZTsgZm9udC1zaXplOiAxMnB4OyB3aGl0ZS1zcGFjZTogbm93cmFwOyB9"
    "CnRhYmxlLmRhdGEgdGhlYWQgdGggewogIHBvc2l0aW9uOiBzdGlja3k7IHRvcDogMDsgei1pbmRleDogMjsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNl"
    "LTIpOyBjb2xvcjogdmFyKC0tbXV0ZWQpOwogIHRleHQtYWxpZ246IGxlZnQ7IHBhZGRpbmc6IDlweCAxMXB4OyBmb250LXdlaWdodDogNzAwOyBmb250LXNp"
    "emU6IDExcHg7CiAgdGV4dC10cmFuc2Zvcm06IHVwcGVyY2FzZTsgbGV0dGVyLXNwYWNpbmc6IC4zcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIo"
    "LS1ib3JkZXIpOwogIGN1cnNvcjogcG9pbnRlcjsgdXNlci1zZWxlY3Q6IG5vbmU7Cn0KdGFibGUuZGF0YSB0aGVhZCB0aDpob3ZlciB7IGNvbG9yOiB2YXIo"
    "LS10ZXh0KTsgfQp0YWJsZS5kYXRhIHRoZWFkIHRoIC5hcnJvdyB7IG9wYWNpdHk6IC41OyBmb250LXNpemU6IDlweDsgfQp0YWJsZS5kYXRhIHRib2R5IHRk"
    "IHsgcGFkZGluZzogOHB4IDExcHg7IGJvcmRlci1ib3R0b206IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyB9CnRhYmxlLmRhdGEgdGJvZHkgdHI6aG92ZXIg"
    "eyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyB9CnRhYmxlLmRhdGEgdGJvZHkgdHIgeyBjdXJzb3I6IHBvaW50ZXI7IH0KdGFibGUuZGF0YSB0Zm9v"
    "dCB0ZCB7IHBhZGRpbmc6IDlweCAxMXB4OyBmb250LXdlaWdodDogNzAwOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlLTIpOyBib3JkZXItdG9wOiAycHgg"
    "c29saWQgdmFyKC0tYm9yZGVyKTsgfQoubnVtIHsgdGV4dC1hbGlnbjogcmlnaHQ7IGZvbnQtdmFyaWFudC1udW1lcmljOiB0YWJ1bGFyLW51bXM7IH0KLnBp"
    "bGwgewogIGRpc3BsYXk6IGlubGluZS1ibG9jazsgcGFkZGluZzogMnB4IDhweDsgYm9yZGVyLXJhZGl1czogOTk5cHg7IGZvbnQtc2l6ZTogMTAuNXB4OyBm"
    "b250LXdlaWdodDogNzAwOwp9Ci5waWxsLmcgeyBiYWNrZ3JvdW5kOiB2YXIoLS1nb29kLWJnKTsgY29sb3I6IHZhcigtLWdvb2QpOyB9Ci5waWxsLmEgeyBi"
    "YWNrZ3JvdW5kOiB2YXIoLS13YXJuLWJnKTsgY29sb3I6IHZhcigtLXdhcm4pOyB9Ci5waWxsLnIgeyBiYWNrZ3JvdW5kOiB2YXIoLS1iYWQtYmcpOyBjb2xv"
    "cjogdmFyKC0tYmFkKTsgfQoucGlsbC5uIHsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsgY29sb3I6IHZhcigtLW11dGVkKTsgfQoucGlsbC5iIHsg"
    "YmFja2dyb3VuZDogcmdiYSg0Nyw5MSwyMTIsLjEyKTsgY29sb3I6IHZhcigtLWluZm8pOyB9CgoucGFnZXIgeyBkaXNwbGF5OiBmbGV4OyBhbGlnbi1pdGVt"
    "czogY2VudGVyOyBnYXA6IDhweDsgbWFyZ2luLXRvcDogMTBweDsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tbXV0ZWQpOyB9Ci5wYWdlciBidXR0"
    "b24geyBwYWRkaW5nOiA1cHggMTBweDsgfQoucGFnZXIgYnV0dG9uOmRpc2FibGVkIHsgb3BhY2l0eTogLjQ7IGN1cnNvcjogZGVmYXVsdDsgfQouY29sLW1l"
    "bnUgeyBwb3NpdGlvbjogcmVsYXRpdmU7IH0KLmNvbC1tZW51LXBhbmVsIHsKICBwb3NpdGlvbjogYWJzb2x1dGU7IHJpZ2h0OiAwOyB0b3A6IGNhbGMoMTAw"
    "JSArIDRweCk7IHotaW5kZXg6IDUwOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3Jk"
    "ZXItcmFkaXVzOiA4cHg7CiAgYm94LXNoYWRvdzogdmFyKC0tc2hhZG93LWxnKTsgcGFkZGluZzogOHB4OyBtYXgtaGVpZ2h0OiAzMDBweDsgb3ZlcmZsb3c6"
    "IGF1dG87IG1pbi13aWR0aDogMTkwcHg7Cn0KLmNvbC1tZW51LXBhbmVsIGxhYmVsIHsgZGlzcGxheTogZmxleDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgZ2Fw"
    "OiA3cHg7IHBhZGRpbmc6IDRweCA2cHg7IGZvbnQtc2l6ZTogMTJweDsgYm9yZGVyLXJhZGl1czogNXB4OyB9Ci5jb2wtbWVudS1wYW5lbCBsYWJlbDpob3Zl"
    "ciB7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IH0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gYWxlcnRzIC8gaW5zaWdodHMgLS0tLS0tLS0tLS0tLS0t"
    "LSAqLwouYWxlcnRzLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IHJlcGVhdChhdXRvLWZpbGwsIG1pbm1heCgzMDBweCwg"
    "MWZyKSk7IGdhcDogMTBweDsgfQouYWxlcnQgewogIGRpc3BsYXk6IGZsZXg7IGdhcDogMTFweDsgcGFkZGluZzogMTJweCAxM3B4OyBib3JkZXItcmFkaXVz"
    "OiB2YXIoLS1yYWRpdXMpOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UpOyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3gtc2hhZG93"
    "OiB2YXIoLS1zaGFkb3cpOyBjdXJzb3I6IHBvaW50ZXI7Cn0KLmFsZXJ0OmhvdmVyIHsgYm9yZGVyLWNvbG9yOiB2YXIoLS1ib3JkZXItMik7IH0KLmFsZXJ0"
    "IC5zZXYgeyB3aWR0aDogNHB4OyBib3JkZXItcmFkaXVzOiA0cHg7IGZsZXg6IDAgMCA0cHg7IH0KLmFsZXJ0LmNyaXQgLnNldiB7IGJhY2tncm91bmQ6IHZh"
    "cigtLWJhZCk7IH0KLmFsZXJ0LmhpZ2ggLnNldiB7IGJhY2tncm91bmQ6IHZhcigtLXdhcm4pOyB9Ci5hbGVydC5tZWQgLnNldiAgeyBiYWNrZ3JvdW5kOiB2"
    "YXIoLS1pbmZvKTsgfQouYWxlcnQubG93IC5zZXYgIHsgYmFja2dyb3VuZDogdmFyKC0tbXV0ZWQpOyB9Ci5hbGVydC1ib2R5IHsgZmxleDogMTsgfQouYWxl"
    "cnQtdGl0bGUgeyBmb250LXdlaWdodDogNzAwOyBmb250LXNpemU6IDEzcHg7IG1hcmdpbi1ib3R0b206IDJweDsgfQouYWxlcnQtZGVzYyB7IGZvbnQtc2l6"
    "ZTogMTEuNXB4OyBjb2xvcjogdmFyKC0tbXV0ZWQpOyB9Ci5hbGVydC1tZXRhIHsgZm9udC1zaXplOiAxMXB4OyBtYXJnaW4tdG9wOiA1cHg7IGRpc3BsYXk6"
    "IGZsZXg7IGdhcDogMTJweDsgZmxleC13cmFwOiB3cmFwOyB9Ci5hbGVydC1tZXRhIGIgeyBjb2xvcjogdmFyKC0tdGV4dCk7IH0KCi5pbnNpZ2h0IHsKICBk"
    "aXNwbGF5OiBmbGV4OyBnYXA6IDEwcHg7IHBhZGRpbmc6IDEwcHggMTJweDsgYm9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzLXNtKTsKICBiYWNrZ3JvdW5k"
    "OiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgbWFyZ2luLWJvdHRvbTogOHB4OyBmb250LXNpemU6IDEyLjVweDsK"
    "fQouaW5zaWdodCAudGFnIHsKICBmbGV4OiAwIDAgYXV0bzsgZm9udC1zaXplOiA5LjVweDsgZm9udC13ZWlnaHQ6IDgwMDsgdGV4dC10cmFuc2Zvcm06IHVw"
    "cGVyY2FzZTsKICBwYWRkaW5nOiAycHggN3B4OyBib3JkZXItcmFkaXVzOiA1cHg7IGxldHRlci1zcGFjaW5nOiAuNHB4OyBoZWlnaHQ6IGZpdC1jb250ZW50"
    "Owp9Ci5pbnNpZ2h0LnBvcyAudGFnIHsgYmFja2dyb3VuZDogdmFyKC0tZ29vZC1iZyk7IGNvbG9yOiB2YXIoLS1nb29kKTsgfQouaW5zaWdodC53YXJuIC50"
    "YWcgeyBiYWNrZ3JvdW5kOiB2YXIoLS13YXJuLWJnKTsgY29sb3I6IHZhcigtLXdhcm4pOyB9Ci5pbnNpZ2h0LmNyaXQgLnRhZyB7IGJhY2tncm91bmQ6IHZh"
    "cigtLWJhZC1iZyk7IGNvbG9yOiB2YXIoLS1iYWQpOyB9Ci5pbnNpZ2h0Lm9wcCAudGFnIHsgYmFja2dyb3VuZDogcmdiYSg0Nyw5MSwyMTIsLjEyKTsgY29s"
    "b3I6IHZhcigtLWluZm8pOyB9Ci5pbnNpZ2h0LmRhdGEgLnRhZyB7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS1tdXRlZCk7"
    "IH0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gbWlzYyAtLS0tLS0tLS0tLS0tLS0tICovCi5lbXB0eS1zdGF0ZSB7CiAgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xv"
    "cjogdmFyKC0tbXV0ZWQpOyBwYWRkaW5nOiAzNHB4IDE2cHg7IGZvbnQtc2l6ZTogMTNweDsKICBib3JkZXI6IDFweCBkYXNoZWQgdmFyKC0tYm9yZGVyKTsg"
    "Ym9yZGVyLXJhZGl1czogdmFyKC0tcmFkaXVzKTsgYmFja2dyb3VuZDogdmFyKC0tc3VyZmFjZS0yKTsKfQouZW1wdHktc3RhdGUgYiB7IGRpc3BsYXk6IGJs"
    "b2NrOyBjb2xvcjogdmFyKC0tdGV4dCk7IG1hcmdpbi1ib3R0b206IDRweDsgZm9udC1zaXplOiAxNHB4OyB9Ci5ub3RlLWJhbm5lciB7CiAgZGlzcGxheTog"
    "ZmxleDsgZ2FwOiAxMHB4OyBhbGlnbi1pdGVtczogZmxleC1zdGFydDsKICBiYWNrZ3JvdW5kOiB2YXIoLS13YXJuLWJnKTsgYm9yZGVyOiAxcHggc29saWQg"
    "dmFyKC0td2Fybik7IGNvbG9yOiB2YXIoLS10ZXh0KTsKICBwYWRkaW5nOiAxMXB4IDE0cHg7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cyk7IGZvbnQt"
    "c2l6ZTogMTIuNXB4OyBtYXJnaW4tYm90dG9tOiAxNHB4Owp9Ci5kZWx0YS11cCB7IGNvbG9yOiB2YXIoLS1nb29kKTsgZm9udC13ZWlnaHQ6IDcwMDsgfQou"
    "ZGVsdGEtZG93biB7IGNvbG9yOiB2YXIoLS1iYWQpOyBmb250LXdlaWdodDogNzAwOyB9Ci5kZWx0YS1mbGF0IHsgY29sb3I6IHZhcigtLW11dGVkKTsgZm9u"
    "dC13ZWlnaHQ6IDcwMDsgfQouYWJiciB7IGJvcmRlci1ib3R0b206IDFweCBkb3R0ZWQgdmFyKC0tbXV0ZWQpOyBjdXJzb3I6IGhlbHA7IH0KCi5zZWxlY3Rv"
    "ci1yb3cgeyBkaXNwbGF5OiBmbGV4OyBnYXA6IDEwcHg7IGZsZXgtd3JhcDogd3JhcDsgYWxpZ24taXRlbXM6IGNlbnRlcjsgbWFyZ2luLWJvdHRvbTogMTRw"
    "eDsgfQouc2VsZWN0b3Itcm93IGxhYmVsIHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tbXV0ZWQpOyBmb250LXdlaWdodDogNjAwOyB9Ci5zZWxl"
    "Y3Rvci1yb3cgc2VsZWN0LCAuc2VsZWN0b3Itcm93IGlucHV0IHsKICBwYWRkaW5nOiA4cHggMTBweDsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0tYm9yZGVy"
    "KTsgYm9yZGVyLXJhZGl1czogOHB4OwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UpOyBjb2xvcjogdmFyKC0tdGV4dCk7IGZvbnQtc2l6ZTogMTNweDsg"
    "bWluLXdpZHRoOiAyMDBweDsKfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBtb2RhbCAtLS0tLS0tLS0tLS0tLS0tICovCi5tb2RhbCB7IHBvc2l0aW9uOiBmaXhl"
    "ZDsgaW5zZXQ6IDA7IHotaW5kZXg6IDEwMDA7IGRpc3BsYXk6IGZsZXg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGp1c3RpZnktY29udGVudDogY2VudGVyOyB9"
    "Ci5tb2RhbC1iYWNrZHJvcCB7IHBvc2l0aW9uOiBhYnNvbHV0ZTsgaW5zZXQ6IDA7IGJhY2tncm91bmQ6IHJnYmEoOCwxNCwyMiwuNSk7IH0KLm1vZGFsLWJv"
    "eCB7CiAgcG9zaXRpb246IHJlbGF0aXZlOyB6LWluZGV4OiAyOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgYm9yZGVyOiAxcHggc29saWQgdmFyKC0t"
    "Ym9yZGVyKTsKICBib3JkZXItcmFkaXVzOiB2YXIoLS1yYWRpdXMpOyBib3gtc2hhZG93OiB2YXIoLS1zaGFkb3ctbGcpOyB3aWR0aDogbWluKDY4MHB4LCA5"
    "NHZ3KTsKICBtYXgtaGVpZ2h0OiA4NnZoOyBvdmVyZmxvdzogYXV0bzsKfQoubW9kYWwtaGVhZCB7IGRpc3BsYXk6IGZsZXg7IGp1c3RpZnktY29udGVudDog"
    "c3BhY2UtYmV0d2VlbjsgYWxpZ24taXRlbXM6IGNlbnRlcjsgcGFkZGluZzogMTVweCAxOHB4OyBib3JkZXItYm90dG9tOiAxcHggc29saWQgdmFyKC0tYm9y"
    "ZGVyKTsgcG9zaXRpb246IHN0aWNreTsgdG9wOiAwOyBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgfQoubW9kYWwtaGVhZCBoMyB7IG1hcmdpbjogMDsg"
    "Zm9udC1zaXplOiAxNXB4OyB9Ci5tb2RhbC1ib2R5IHsgcGFkZGluZzogMTZweCAxOHB4OyB9Ci5rdiB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUt"
    "Y29sdW1uczogMTkwcHggMWZyOyBnYXA6IDRweCAxNHB4OyBmb250LXNpemU6IDEyLjVweDsgfQoua3YgZHQgeyBjb2xvcjogdmFyKC0tbXV0ZWQpOyB9Ci5r"
    "diBkZCB7IG1hcmdpbjogMDsgd29yZC1icmVhazogYnJlYWstd29yZDsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBmb290ZXIgLS0tLS0tLS0tLS0tLS0tLSAq"
    "LwouYXBwLWZvb3RlciB7CiAgdGV4dC1hbGlnbjogY2VudGVyOyBjb2xvcjogdmFyKC0tbXV0ZWQpOyBmb250LXNpemU6IDExLjVweDsKICBwYWRkaW5nOiAx"
    "OHB4OyBkaXNwbGF5OiBmbGV4OyBnYXA6IDEwcHg7IGp1c3RpZnktY29udGVudDogY2VudGVyOyBmbGV4LXdyYXA6IHdyYXA7Cn0KLmZvb3Qtc2VwIHsgb3Bh"
    "Y2l0eTogLjU7IH0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gcmVzcG9uc2l2ZSAtLS0tLS0tLS0tLS0tLS0tICovCkBtZWRpYSAobWF4LXdpZHRoOiAxMTAwcHgp"
    "IHsKICAuY2FyZC5jNiwgLmNhcmQuYzgsIC5jYXJkLmM0LCAuY2FyZC5jMyB7IGdyaWQtY29sdW1uOiBzcGFuIDEyOyB9Cn0KQG1lZGlhIChtYXgtd2lkdGg6"
    "IDg2MHB4KSB7CiAgLmhlYWRlci1jZW50ZXIgeyBkaXNwbGF5OiBub25lOyB9CiAgLmhlYWRlci1tZXRhIHsgZGlzcGxheTogbm9uZTsgfQogIC5maWx0ZXIt"
    "YmFyIHsgdG9wOiA1OXB4OyB9CiAgLnRhYi1uYXYgeyB0b3A6IDExNnB4OyB9Cn0KQG1lZGlhIChtYXgtd2lkdGg6IDU2MHB4KSB7CiAgLmtwaS12YWx1ZSB7"
    "IGZvbnQtc2l6ZTogMTlweDsgfQogIC5hcHAtbWFpbiB7IHBhZGRpbmc6IDEycHg7IH0KfQoKLyogPT09PT0gYWRkaXRpb25zOiBiYW5uZXJzLCBkZXRhaWwg"
    "bW9kYWwgZ3JpZCwgUE8gc2VhcmNoLCBjb21wYXJpc29uLCBEUSBtYXRyaXggPT09PT0gKi8KLmJhbm5lciB7CiAgbWFyZ2luOiA0cHggMCAxNHB4OyBwYWRk"
    "aW5nOiAxMXB4IDE0cHg7IGJvcmRlci1yYWRpdXM6IHZhcigtLXJhZGl1cywgMTBweCk7CiAgZm9udC1zaXplOiAxMi41cHg7IGxpbmUtaGVpZ2h0OiAxLjU1"
    "OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOwogIGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IGNvbG9yOiB2YXIoLS1tdXRlZCk7Cn0K"
    "LmJhbm5lci5pbmZvIHsgYm9yZGVyLWxlZnQ6IDNweCBzb2xpZCB2YXIoLS1pbmZvKTsgfQouYmFubmVyIGNvZGUgeyBiYWNrZ3JvdW5kOiByZ2JhKDEyNywx"
    "MjcsMTI3LC4xNCk7IHBhZGRpbmc6IDFweCA1cHg7IGJvcmRlci1yYWRpdXM6IDRweDsgZm9udC1zaXplOiAxMS41cHg7IH0KLmJhbm5lciBiIHsgY29sb3I6"
    "IHZhcigtLXRleHQpOyB9CgouZGV0YWlsLWdyaWQgeyBkaXNwbGF5OiBncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6IDFmciAxZnI7IGdhcDogNnB4IDE4"
    "cHg7IH0KQG1lZGlhIChtYXgtd2lkdGg6IDYyMHB4KSB7IC5kZXRhaWwtZ3JpZCB7IGdyaWQtdGVtcGxhdGUtY29sdW1uczogMWZyOyB9IH0KLmR0LXJvdyB7"
    "IGRpc3BsYXk6IGZsZXg7IGp1c3RpZnktY29udGVudDogc3BhY2UtYmV0d2VlbjsgZ2FwOiAxMnB4OyBwYWRkaW5nOiA2cHggMnB4OyBib3JkZXItYm90dG9t"
    "OiAxcHggZGFzaGVkIHZhcigtLWJvcmRlcik7IGZvbnQtc2l6ZTogMTNweDsgfQouZHQtayB7IGNvbG9yOiB2YXIoLS1tdXRlZCk7IH0KLmR0LXYgeyBjb2xv"
    "cjogdmFyKC0tdGV4dCk7IGZvbnQtd2VpZ2h0OiA2MDA7IHRleHQtYWxpZ246IHJpZ2h0OyB9CgoucG8tc2VhcmNoLXJvdyB7IGRpc3BsYXk6IGZsZXg7IGFs"
    "aWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogMTBweDsgbWFyZ2luLWJvdHRvbTogMTJweDsgZmxleC13cmFwOiB3cmFwOyB9Ci5wby1zZWFyY2gtcm93IGlucHV0"
    "IHsKICBmbGV4OiAwIDEgMzIwcHg7IHBhZGRpbmc6IDlweCAxMnB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXItMik7IGJvcmRlci1yYWRpdXM6"
    "IDlweDsKICBiYWNrZ3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgY29sb3I6IHZhcigtLXRleHQpOyBmb250LXNpemU6IDEzcHg7Cn0KLnBvLXNlYXJjaC1yb3cg"
    "aW5wdXQ6Zm9jdXMgeyBvdXRsaW5lOiAycHggc29saWQgdmFyKC0taW5mbyk7IGJvcmRlci1jb2xvcjogdmFyKC0taW5mbyk7IH0KLnBvLXNlYXJjaC1ub3Rl"
    "IHsgZm9udC1zaXplOiAxMnB4OyBjb2xvcjogdmFyKC0tbXV0ZWQpOyB9CgouY21wLWNvbnRyb2xzIHsgZGlzcGxheTogZmxleDsgZ2FwOiAxOHB4OyBhbGln"
    "bi1pdGVtczogY2VudGVyOyBtYXJnaW4tYm90dG9tOiAxNHB4OyBmbGV4LXdyYXA6IHdyYXA7IH0KLmNtcC1jb250cm9scyBsYWJlbCB7IGZvbnQtc2l6ZTog"
    "MTJweDsgY29sb3I6IHZhcigtLW11dGVkKTsgZGlzcGxheTogZmxleDsgZ2FwOiA3cHg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IH0KLmNtcC1jb250cm9scyBz"
    "ZWxlY3QgewogIHBhZGRpbmc6IDdweCAxMHB4OyBib3JkZXI6IDFweCBzb2xpZCB2YXIoLS1ib3JkZXItMik7IGJvcmRlci1yYWRpdXM6IDhweDsKICBiYWNr"
    "Z3JvdW5kOiB2YXIoLS1zdXJmYWNlKTsgY29sb3I6IHZhcigtLXRleHQpOyBmb250LXNpemU6IDEzcHg7Cn0KLmNtcC10YWJsZSB0aCwgLmNtcC10YWJsZSB0"
    "ZCB7IHdoaXRlLXNwYWNlOiBub3dyYXA7IH0KLmNtcC10YWJsZSB0ZC51cC1nb29kLCAuY21wLXRhYmxlIHRkLmRvd24tYmFkLCAuY21wLXRhYmxlIHRkLm5l"
    "dXRyYWwgeyBmb250LXdlaWdodDogNjAwOyB9CnRkLnVwLWdvb2QgeyBjb2xvcjogdmFyKC0tZ29vZCk7IH0KdGQuZG93bi1iYWQgeyBjb2xvcjogdmFyKC0t"
    "YmFkKTsgfQp0ZC5uZXV0cmFsIHsgY29sb3I6IHZhcigtLW11dGVkKTsgfQoKLmRxLW1hdHJpeCB7IGRpc3BsYXk6IGdyaWQ7IGdyaWQtdGVtcGxhdGUtY29s"
    "dW1uczogcmVwZWF0KGF1dG8tZmlsbCwgbWlubWF4KDI4MHB4LCAxZnIpKTsgZ2FwOiA4cHggMjBweDsgfQouZHEtY2VsbCB7IGRpc3BsYXk6IGdyaWQ7IGdy"
    "aWQtdGVtcGxhdGUtY29sdW1uczogMTUwcHggMWZyIDQycHg7IGFsaWduLWl0ZW1zOiBjZW50ZXI7IGdhcDogOHB4OyBmb250LXNpemU6IDEycHg7IH0KLmRx"
    "LWZpZWxkIHsgY29sb3I6IHZhcigtLW11dGVkKTsgb3ZlcmZsb3c6IGhpZGRlbjsgdGV4dC1vdmVyZmxvdzogZWxsaXBzaXM7IHdoaXRlLXNwYWNlOiBub3dy"
    "YXA7IH0KLmRxLWJhciB7IGhlaWdodDogOHB4OyBib3JkZXItcmFkaXVzOiA1cHg7IGJhY2tncm91bmQ6IHZhcigtLXN1cmZhY2UtMik7IG92ZXJmbG93OiBo"
    "aWRkZW47IGJvcmRlcjogMXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IH0KLmRxLWZpbGwgeyBkaXNwbGF5OiBibG9jazsgaGVpZ2h0OiAxMDAlOyBib3JkZXIt"
    "cmFkaXVzOiA1cHg7IH0KLmRxLWZpbGwuZyB7IGJhY2tncm91bmQ6IHZhcigtLWdvb2QpOyB9Ci5kcS1maWxsLmEgeyBiYWNrZ3JvdW5kOiB2YXIoLS13YXJu"
    "KTsgfQouZHEtZmlsbC5yIHsgYmFja2dyb3VuZDogdmFyKC0tYmFkKTsgfQouZHEtcGN0IHsgdGV4dC1hbGlnbjogcmlnaHQ7IGZvbnQtd2VpZ2h0OiA2MDA7"
    "IGNvbG9yOiB2YXIoLS10ZXh0KTsgfQouZHEtbm90ZSB7IG1hcmdpbi10b3A6IDE0cHg7IGZvbnQtc2l6ZTogMTJweDsgY29sb3I6IHZhcigtLW11dGVkKTsg"
    "bGluZS1oZWlnaHQ6IDEuNTU7IH0K"
)

_JS_B64 = (
    "LyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PQogICBFVCBSRlEgJiBQTyBDb250cm9s"
    "bGVyIOKAlCBzY3JpcHQuanMKICAgU3RhdGljLCBHaXRIdWItUGFnZXMtcmVhZHkgZGFzaGJvYXJkIGxvZ2ljICh2YW5pbGxhIEpTICsgQ2hhcnQuanMpCiAg"
    "ID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgU2VjdGlvbnM6CiAgICAgMS4gIFN0"
    "YXRlICYgY29uc3RhbnRzCiAgICAgMi4gIFV0aWxpdGllcyAoZm9ybWF0LCBkYXRlcywgbWF0aCkKICAgICAzLiAgRGVkdXAtYXdhcmUgYWdncmVnYXRpb24g"
    "KGNvcnJlY3QgdW5kZXIgYW55IGZpbHRlcikKICAgICA0LiAgVGhlbWUgaGFuZGxpbmcKICAgICA1LiAgRGF0YSBsb2FkaW5nCiAgICAgNi4gIEZpbHRlciBl"
    "bmdpbmUgIChzbGljZXJzLCBzZWFyY2gsIGNoYXJ0LWZpbHRlcnMsIGNoaXBzKQogICAgIDcuICBDaGFydCArIEtQSSBoZWxwZXJzCiAgICAgOC4gIEdlbmVy"
    "aWMgdGFibGUgY29tcG9uZW50IChzb3J0L3NlYXJjaC9wYWdlL2NvbHZpcy9leHBvcnQvZHJpbGwpCiAgICAgOS4gIFJGUSByb2xsLXVwICsgZGVsaXZlcnkt"
    "c3RhdHVzIGVuZ2luZQogICAgIDEwLiBSaXNrIGVuZ2luZQogICAgIDExLiBJbnNpZ2h0IGVuZ2luZQogICAgIDEyLiBUYWIgcmVuZGVyZXJzIChPdmVydmll"
    "dywgUkZRLCBTdXBwbGllciwgUE8sIFBPQywgQ29tcGFyZSwKICAgICAgICAgICAgICAgICAgICAgICAgQ3VzdG9tZXIsIFJpc2ssIERhdGEtUXVhbGl0eSkK"
    "ICAgICAxMy4gRXhwb3J0cwogICAgIDE0LiBCb290CiAgID09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09PT0gKi8KCi8qIC0tLS0tLS0tLS0tLS0tLS0gMS4gU1RBVEUgLS0tLS0tLS0tLS0tLS0tLSAqLwpsZXQgUkFXID0gW107CmxldCBNRVRBID0ge307"
    "CmNvbnN0IENIQVJUUyA9IHt9OwpsZXQgQ1VSUkVOVF9UQUIgPSAib3ZlcnZpZXciOwpsZXQgUkVOREVSRUQgPSB7fTsgICAgICAgICAgICAgICAgIC8vIHdo"
    "aWNoIHRhYnMgaGF2ZSBiZWVuIGRyYXduIGZvciBjdXJyZW50IGZpbHRlciBzdGF0ZQoKY29uc3QgRklMVEVSUyA9IHsKICBzZWFyY2g6ICIiLAogIGRhdGVG"
    "cm9tOiAiIiwgZGF0ZVRvOiAiIiwKICB5ZWFyOiAiQWxsIiwgcXVhcnRlcjogIkFsbCIsIG1vbnRoOiAiQWxsIiwKICBtdWx0aTogewogICAgZXRRdW90ZVN0"
    "YXR1czogW10sIGV0UmZxU3RhdHVzOiBbXSwgZXRQT0M6IFtdLCBjdXN0UE9DOiBbXSwKICAgIGN1c3RvbWVyOiBbXSwgc3VwcGxpZXJOYW1lOiBbXSwgc2Vj"
    "dG9yOiBbXSwgcHJvZHVjdENhdGVnb3J5OiBbXSwKICAgIHNoaXBtZW50U3RhdHVzOiBbXSwgcmZxUmVzdWx0OiBbXSwKICB9LAogIGNoYXJ0OiB7fSwgICAg"
    "ICAgICAgICAgICAgICAgICAgICAvLyBkaW1lbnNpb24gLT4gdmFsdWUgKGNsaWNrLXRvLWZpbHRlcikKfTsKCmNvbnN0IE1VTFRJX0RFRlMgPSBbCiAgWyJl"
    "dFF1b3RlU3RhdHVzIiwgIkVUIFF1b3RlIFN0YXR1cyIsICJxdW90ZVN0YXR1c2VzIl0sCiAgWyJldFJmcVN0YXR1cyIsICJFVCBSRlEgU3RhdHVzIiwgInJm"
    "cVN0YXR1c2VzIl0sCiAgWyJldFBPQyIsICJFVCBQT0MiLCAiZXRQT0NzIl0sCiAgWyJjdXN0UE9DIiwgIkN1c3RvbWVyIFBPQyIsICJjdXN0UE9DcyJdLAog"
    "IFsiY3VzdG9tZXIiLCAiQ3VzdG9tZXIiLCAiY3VzdG9tZXJzIl0sCiAgWyJzdXBwbGllck5hbWUiLCAiU3VwcGxpZXIiLCAic3VwcGxpZXJzIl0sCiAgWyJz"
    "ZWN0b3IiLCAiU2VjdG9yIiwgInNlY3RvcnMiXSwKICBbInByb2R1Y3RDYXRlZ29yeSIsICJQcm9kdWN0IENhdGVnb3J5IiwgInByb2R1Y3RDYXRlZ29yaWVz"
    "Il0sCl07CmNvbnN0IFJGUV9SRVNVTFRfT1BUUyA9IFsiV29uIiwgIkxvc3QiLCAiRGVjbGluZWQiLCAiT3BlbiIsICJQZW5kaW5nIl07CmNvbnN0IE1PTlRI"
    "UyA9IFsiSmFuIiwiRmViIiwiTWFyIiwiQXByIiwiTWF5IiwiSnVuIiwiSnVsIiwiQXVnIiwiU2VwIiwiT2N0IiwiTm92IiwiRGVjIl07CmNvbnN0IExPV19N"
    "QVJHSU4gPSA1OyAgICAgICAgICAgICAgLy8gJSB0aHJlc2hvbGQgKGNvbmZpZ3VyYWJsZSkKY29uc3QgSElHSF9NQVJHSU4gPSA2MDsgICAgICAgICAgIC8v"
    "ICUgdGhyZXNob2xkIChjb25maWd1cmFibGUpCgovKiAtLS0tLS0tLS0tLS0tLS0tIDIuIFVUSUxJVElFUyAtLS0tLS0tLS0tLS0tLS0tICovCmNvbnN0ICQg"
    "PSAocywgciA9IGRvY3VtZW50KSA9PiByLnF1ZXJ5U2VsZWN0b3Iocyk7CmNvbnN0ICQkID0gKHMsIHIgPSBkb2N1bWVudCkgPT4gQXJyYXkuZnJvbShyLnF1"
    "ZXJ5U2VsZWN0b3JBbGwocykpOwpjb25zdCBlc2MgPSAocykgPT4gU3RyaW5nKHMgPT0gbnVsbCA/ICIiIDogcykucmVwbGFjZSgvWyY8PiJdL2csIGMgPT4g"
    "KHsgIiYiOiImYW1wOyIsIjwiOiImbHQ7IiwiPiI6IiZndDsiLCciJzoiJnF1b3Q7IiB9W2NdKSk7CgpmdW5jdGlvbiBmbXRDdXIodikgewogIGlmICh2ID09"
    "IG51bGwgfHwgaXNOYU4odikpIHJldHVybiAi4oCUIjsKICByZXR1cm4gIiQiICsgTnVtYmVyKHYpLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIsIHsgbWluaW11"
    "bUZyYWN0aW9uRGlnaXRzOiAyLCBtYXhpbXVtRnJhY3Rpb25EaWdpdHM6IDIgfSk7Cn0KZnVuY3Rpb24gZm10Q29tcGFjdCh2KSB7CiAgaWYgKHYgPT0gbnVs"
    "bCB8fCBpc05hTih2KSkgcmV0dXJuICLigJQiOwogIGNvbnN0IGEgPSBNYXRoLmFicyh2KTsKICBpZiAoYSA+PSAxZTkpIHJldHVybiAiJCIgKyAodiAvIDFl"
    "OSkudG9GaXhlZCgyKSArICJCIjsKICBpZiAoYSA+PSAxZTYpIHJldHVybiAiJCIgKyAodiAvIDFlNikudG9GaXhlZCgyKSArICJNIjsKICBpZiAoYSA+PSAx"
    "ZTMpIHJldHVybiAiJCIgKyAodiAvIDFlMykudG9GaXhlZCgxKSArICJLIjsKICByZXR1cm4gIiQiICsgTnVtYmVyKHYpLnRvRml4ZWQoMCk7Cn0KZnVuY3Rp"
    "b24gZm10TnVtKHYpIHsgcmV0dXJuIHYgPT0gbnVsbCB8fCBpc05hTih2KSA/ICLigJQiIDogTnVtYmVyKHYpLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIpOyB9"
    "CmZ1bmN0aW9uIGZtdFBjdCh2LCBkID0gMSkgeyByZXR1cm4gdiA9PSBudWxsIHx8IGlzTmFOKHYpID8gIuKAlCIgOiBOdW1iZXIodikudG9GaXhlZChkKSAr"
    "ICIlIjsgfQpmdW5jdGlvbiBmbXREYXRlKGlzbykgewogIGlmICghaXNvKSByZXR1cm4gIuKAlCI7CiAgY29uc3QgcCA9IFN0cmluZyhpc28pLnNwbGl0KCIt"
    "Iik7CiAgaWYgKHAubGVuZ3RoICE9PSAzKSByZXR1cm4gaXNvOwogIHJldHVybiBgJHtwWzJdfSAke01PTlRIU1srcFsxXSAtIDFdfSAke3BbMF19YDsKfQpm"
    "dW5jdGlvbiBmbXREYXlzKHYpIHsgcmV0dXJuIHYgPT0gbnVsbCB8fCBpc05hTih2KSA/ICLigJQiIDogTWF0aC5yb3VuZCh2KSArICIgZCI7IH0KZnVuY3Rp"
    "b24gc2FmZURpdihhLCBiKSB7IHJldHVybiBiID8gYSAvIGIgOiAwOyB9CmZ1bmN0aW9uIGRheXNCZXR3ZWVuKGEsIGIpIHsKICBpZiAoIWEgfHwgIWIpIHJl"
    "dHVybiBudWxsOwogIHJldHVybiBNYXRoLnJvdW5kKChuZXcgRGF0ZShhKSAtIG5ldyBEYXRlKGIpKSAvIDg2NDAwMDAwKTsKfQpmdW5jdGlvbiB0b2RheUlT"
    "TygpIHsgcmV0dXJuIG5ldyBEYXRlKCkudG9JU09TdHJpbmcoKS5zbGljZSgwLCAxMCk7IH0KZnVuY3Rpb24gbW9udGhLZXkoaXNvKSB7IHJldHVybiBpc28g"
    "PyBpc28uc2xpY2UoMCwgNykgOiBudWxsOyB9CmZ1bmN0aW9uIHF1YXJ0ZXJPZihpc28pIHsgcmV0dXJuIGlzbyA/IE1hdGguZmxvb3IoKCtpc28uc2xpY2Uo"
    "NSwgNykgLSAxKSAvIDMpICsgMSA6IG51bGw7IH0KZnVuY3Rpb24gbWVkaWFuKGFycikgewogIGNvbnN0IGEgPSBhcnIuZmlsdGVyKHggPT4geCAhPSBudWxs"
    "KS5zb3J0KCh4LCB5KSA9PiB4IC0geSk7CiAgaWYgKCFhLmxlbmd0aCkgcmV0dXJuIG51bGw7CiAgY29uc3QgbSA9IE1hdGguZmxvb3IoYS5sZW5ndGggLyAy"
    "KTsKICByZXR1cm4gYS5sZW5ndGggJSAyID8gYVttXSA6IChhW20gLSAxXSArIGFbbV0pIC8gMjsKfQoKLyogLS0tLS0tLS0tLS0tLS0tLSAzLiBERURVUC1B"
    "V0FSRSBBR0dSRUdBVElPTiAtLS0tLS0tLS0tLS0tLS0tCiAgIFdoZW4gb25lICJUb3RhbCIgdmFsdWUgcmVwZWF0cyBpZGVudGljYWxseSBhY3Jvc3MgZXZl"
    "cnkgbGluZSBvZiBhCiAgIHNpbmdsZSBSRlEvUE8gaXQgaXMgY291bnRlZCBvbmNlOyBvdGhlcndpc2UgbGluZSB2YWx1ZXMgYXJlIHN1bW1lZC4KICAgQ29y"
    "cmVjdCB1bmRlciBBTlkgcm93IHN1YnNldCAoZmlsdGVycyksIGNvbXB1dGVkIGxpdmUuICAgICAgICAgICAgICAgKi8KZnVuY3Rpb24gZ3JvdXBEZWR1cChy"
    "b3dzLCBrZXlGaWVsZCwgdmFsdWVGaWVsZCkgewogIGNvbnN0IGdyb3VwcyA9IG5ldyBNYXAoKTsKICBmb3IgKGNvbnN0IHIgb2Ygcm93cykgewogICAgY29u"
    "c3QgdiA9IHJbdmFsdWVGaWVsZF07CiAgICBpZiAodiA9PSBudWxsIHx8IGlzTmFOKHYpKSBjb250aW51ZTsKICAgIGNvbnN0IGsgPSByW2tleUZpZWxkXTsK"
    "ICAgIGlmIChrID09IG51bGwpIGNvbnRpbnVlOwogICAgKGdyb3Vwcy5nZXQoaykgfHwgZ3JvdXBzLnNldChrLCBbXSkuZ2V0KGspKS5wdXNoKHYpOwogIH0K"
    "ICBjb25zdCBvdXQgPSBuZXcgTWFwKCk7CiAgZm9yIChjb25zdCBbaywgdmFsc10gb2YgZ3JvdXBzKSB7CiAgICBjb25zdCB1bmlxID0gbmV3IFNldCh2YWxz"
    "Lm1hcCh4ID0+IE1hdGgucm91bmQoeCAqIDEwMCkgLyAxMDApKTsKICAgIG91dC5zZXQoaywgKHZhbHMubGVuZ3RoID4gMSAmJiB1bmlxLnNpemUgPT09IDEp"
    "ID8gdmFsc1swXSA6IHZhbHMucmVkdWNlKChhLCBiKSA9PiBhICsgYiwgMCkpOwogIH0KICByZXR1cm4gb3V0Owp9CmZ1bmN0aW9uIGFnZ3JlZ2F0ZURlZHVw"
    "KHJvd3MsIGtleUZpZWxkLCB2YWx1ZUZpZWxkKSB7CiAgbGV0IHQgPSAwOwogIGZvciAoY29uc3QgdiBvZiBncm91cERlZHVwKHJvd3MsIGtleUZpZWxkLCB2"
    "YWx1ZUZpZWxkKS52YWx1ZXMoKSkgdCArPSB2OwogIHJldHVybiB0Owp9CmZ1bmN0aW9uIHVuaXF1ZUNvdW50KHJvd3MsIGtleUZpZWxkKSB7CiAgY29uc3Qg"
    "cyA9IG5ldyBTZXQoKTsKICBmb3IgKGNvbnN0IHIgb2Ygcm93cykgaWYgKHJba2V5RmllbGRdICE9IG51bGwpIHMuYWRkKHJba2V5RmllbGRdKTsKICByZXR1"
    "cm4gcy5zaXplOwp9CmZ1bmN0aW9uIHN1bUZpZWxkKHJvd3MsIGYpIHsKICBsZXQgdCA9IDA7IGZvciAoY29uc3QgciBvZiByb3dzKSBpZiAocltmXSAhPSBu"
    "dWxsICYmICFpc05hTihyW2ZdKSkgdCArPSByW2ZdOyByZXR1cm4gdDsKfQpmdW5jdGlvbiBhdmdGaWVsZChyb3dzLCBmKSB7CiAgY29uc3QgdiA9IHJvd3Mu"
    "bWFwKHIgPT4gcltmXSkuZmlsdGVyKHggPT4geCAhPSBudWxsICYmICFpc05hTih4KSk7CiAgcmV0dXJuIHYubGVuZ3RoID8gdi5yZWR1Y2UoKGEsIGIpID0+"
    "IGEgKyBiLCAwKSAvIHYubGVuZ3RoIDogbnVsbDsKfQoKLyogLS0tLS0tLS0tLS0tLS0tLSA0LiBUSEVNRSAtLS0tLS0tLS0tLS0tLS0tICovCmZ1bmN0aW9u"
    "IGluaXRUaGVtZSgpIHsKICBjb25zdCBzYXZlZCA9IGxvY2FsU3RvcmFnZS5nZXRJdGVtKCJldC10aGVtZSIpIHx8ICJsaWdodCI7CiAgZG9jdW1lbnQuZG9j"
    "dW1lbnRFbGVtZW50LnNldEF0dHJpYnV0ZSgiZGF0YS10aGVtZSIsIHNhdmVkKTsKfQpmdW5jdGlvbiB0b2dnbGVUaGVtZSgpIHsKICBjb25zdCBjdXIgPSBk"
    "b2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZ2V0QXR0cmlidXRlKCJkYXRhLXRoZW1lIik7CiAgY29uc3Qgbnh0ID0gY3VyID09PSAiZGFyayIgPyAibGlnaHQi"
    "IDogImRhcmsiOwogIGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5zZXRBdHRyaWJ1dGUoImRhdGEtdGhlbWUiLCBueHQpOwogIGxvY2FsU3RvcmFnZS5zZXRJ"
    "dGVtKCJldC10aGVtZSIsIG54dCk7CiAgUkVOREVSRUQgPSB7fTsKICByZWZyZXNoKHRydWUpOyAgICAgICAgICAgICAgICAgICAgLy8gcmVkcmF3IGNoYXJ0"
    "cyB3aXRoIG5ldyBwYWxldHRlCn0KZnVuY3Rpb24gY3NzVmFyKG5hbWUpIHsgcmV0dXJuIGdldENvbXB1dGVkU3R5bGUoZG9jdW1lbnQuZG9jdW1lbnRFbGVt"
    "ZW50KS5nZXRQcm9wZXJ0eVZhbHVlKG5hbWUpLnRyaW0oKTsgfQpmdW5jdGlvbiBwYWxldHRlKCkgewogIHJldHVybiB7CiAgICB0ZXh0OiBjc3NWYXIoIi0t"
    "dGV4dCIpLCBtdXRlZDogY3NzVmFyKCItLW11dGVkIiksIGdyaWQ6IGNzc1ZhcigiLS1ncmlkIiksCiAgICBwcmltYXJ5OiBjc3NWYXIoIi0tcHJpbWFyeSIp"
    "LCBhY2NlbnQ6IGNzc1ZhcigiLS1hY2NlbnQiKSwKICAgIGdvb2Q6IGNzc1ZhcigiLS1nb29kIiksIHdhcm46IGNzc1ZhcigiLS13YXJuIiksIGJhZDogY3Nz"
    "VmFyKCItLWJhZCIpLCBpbmZvOiBjc3NWYXIoIi0taW5mbyIpLAogICAgc3VyZmFjZTogY3NzVmFyKCItLXN1cmZhY2UiKSwKICAgIHNlcmllczogWyIjMmY1"
    "YmQ0IiwiIzBkOTQ4OCIsIiNjMDc4MDYiLCIjOGI1Y2Y2IiwiI2UwNTY3YyIsIiMxNTg4NGMiLAogICAgICAgICAgICAgIiMwODkxYjIiLCIjZDQ2MDJmIiwi"
    "IzYzNjZmMSIsIiNjYThhMDQiLCIjMGY3NjZlIiwiI2JlMTg1ZCJdLAogIH07Cn0KCi8qIC0tLS0tLS0tLS0tLS0tLS0gNS4gREFUQSBMT0FESU5HIC0tLS0t"
    "LS0tLS0tLS0tLS0gKi8KYXN5bmMgZnVuY3Rpb24gbG9hZERhdGEoKSB7CiAgY29uc3QgbXNnID0gJCgiI2xvYWRlck1zZyIpOwogIHRyeSB7CiAgICBsZXQg"
    "anNvbjsKICAgIC8vIFByZWZlcnJlZCBwYXRoOiBkYXRhIGVtYmVkZGVkIGRpcmVjdGx5IGluc2lkZSBpbmRleC5odG1sLgogICAgLy8gVGhpcyBpcyB3aGF0"
    "IGxldHMgdGhlIGRhc2hib2FyZCBvcGVuIGJ5IGRvdWJsZS1jbGlja2luZyB0aGUgZmlsZSAoZmlsZTovLyksCiAgICAvLyB3aXRoIG5vIHdlYiBzZXJ2ZXIg"
    "YW5kIG5vIGZldGNoKCkg4oCUIGJyb3dzZXJzIGJsb2NrIGZldGNoIG92ZXIgZmlsZTovLy4KICAgIGNvbnN0IGVtYmVkID0gZG9jdW1lbnQuZ2V0RWxlbWVu"
    "dEJ5SWQoImV0LWRhdGEiKTsKICAgIGlmIChlbWJlZCAmJiBlbWJlZC50ZXh0Q29udGVudCAmJiBlbWJlZC50ZXh0Q29udGVudC50cmltKCkubGVuZ3RoID4g"
    "MikgewogICAgICBtc2cudGV4dENvbnRlbnQgPSAiUGFyc2luZyByZWNvcmRz4oCmIjsKICAgICAganNvbiA9IEpTT04ucGFyc2UoZW1iZWQudGV4dENvbnRl"
    "bnQpOwogICAgfSBlbHNlIHsKICAgICAgLy8gRmFsbGJhY2s6IGxvYWQgZXh0ZXJuYWwgZGF0YS5qc29uICh1c2VkIG9ubHkgd2hlbiBzZXJ2ZWQgb3ZlciBI"
    "VFRQKS4KICAgICAgbXNnLnRleHRDb250ZW50ID0gIkZldGNoaW5nIGRhdGEuanNvbuKApiI7CiAgICAgIGNvbnN0IHJlcyA9IGF3YWl0IGZldGNoKCJkYXRh"
    "Lmpzb24iLCB7IGNhY2hlOiAibm8tc3RvcmUiIH0pOwogICAgICBpZiAoIXJlcy5vaykgdGhyb3cgbmV3IEVycm9yKCJIVFRQICIgKyByZXMuc3RhdHVzKTsK"
    "ICAgICAgbXNnLnRleHRDb250ZW50ID0gIlBhcnNpbmcgcmVjb3Jkc+KApiI7CiAgICAgIGpzb24gPSBhd2FpdCByZXMuanNvbigpOwogICAgfQogICAgUkFX"
    "ID0ganNvbi5yZWNvcmRzIHx8IFtdOwogICAgTUVUQSA9IGpzb24ubWV0YSB8fCB7fTsKICAgIC8vIHByZS1jb21wdXRlIGRlcml2ZWQgZGF0ZSBwYXJ0cyBv"
    "bmNlCiAgICBmb3IgKGNvbnN0IHIgb2YgUkFXKSB7CiAgICAgIHIuX3F5ID0gci5ldFF1b3RlRGF0ZSA/ICtyLmV0UXVvdGVEYXRlLnNsaWNlKDAsIDQpIDog"
    "bnVsbDsKICAgICAgci5fcXEgPSByLmV0UXVvdGVEYXRlID8gcXVhcnRlck9mKHIuZXRRdW90ZURhdGUpIDogbnVsbDsKICAgICAgci5fcW0gPSByLmV0UXVv"
    "dGVEYXRlID8gK3IuZXRRdW90ZURhdGUuc2xpY2UoNSwgNykgOiBudWxsOwogICAgfQogIH0gY2F0Y2ggKGUpIHsKICAgIG1zZy5pbm5lckhUTUwgPSBgPGIg"
    "c3R5bGU9ImNvbG9yOnZhcigtLWJhZCkiPkNvdWxkIG5vdCBsb2FkIGRhdGEuPC9iPjxicj4KICAgICAgUmUtcnVuIDxjb2RlPnB5dGhvbiBjb252ZXJ0LnB5"
    "PC9jb2RlPiB0byByZWJ1aWxkIGEgc2VsZi1jb250YWluZWQgaW5kZXguaHRtbC4KICAgICAgPGJyPjxzbWFsbD4ke2VzYyhlLm1lc3NhZ2UpfTwvc21hbGw+"
    "YDsKICAgIHRocm93IGU7CiAgfQp9CgovKiAtLS0tLS0tLS0tLS0tLS0tIDYuIEZJTFRFUiBFTkdJTkUgLS0tLS0tLS0tLS0tLS0tLSAqLwpmdW5jdGlvbiBi"
    "dWlsZFNsaWNlcnMoKSB7CiAgY29uc3QgZ3JpZCA9ICQoIiNmaWx0ZXJHcmlkIik7CiAgZ3JpZC5pbm5lckhUTUwgPSAiIjsKICAvLyBkYXRlIHNsaWNlcnMK"
    "ICBjb25zdCB5ZWFycyA9IEFycmF5LmZyb20obmV3IFNldChSQVcubWFwKHIgPT4gci5fcXkpLmZpbHRlcihCb29sZWFuKSkpLnNvcnQoKTsKICBncmlkLmFw"
    "cGVuZENoaWxkKG1ha2VTZWxlY3QoIlllYXIgKEVUIFF1b3RlKSIsICJ5ZWFyIiwgWyJBbGwiLCAuLi55ZWFyc10pKTsKICBncmlkLmFwcGVuZENoaWxkKG1h"
    "a2VTZWxlY3QoIlF1YXJ0ZXIiLCAicXVhcnRlciIsIFsiQWxsIiwgIlExIiwgIlEyIiwgIlEzIiwgIlE0Il0pKTsKICBncmlkLmFwcGVuZENoaWxkKG1ha2VT"
    "ZWxlY3QoIk1vbnRoIiwgIm1vbnRoIiwgWyJBbGwiLCAuLi5NT05USFNdKSk7CiAgLy8gY3VzdG9tIGRhdGUgcmFuZ2UKICBncmlkLmFwcGVuZENoaWxkKG1h"
    "a2VEYXRlKCJGcm9tIChFVCBRdW90ZSkiLCAiZGF0ZUZyb20iKSk7CiAgZ3JpZC5hcHBlbmRDaGlsZChtYWtlRGF0ZSgiVG8gKEVUIFF1b3RlKSIsICJkYXRl"
    "VG8iKSk7CiAgLy8gbXVsdGktc2VsZWN0cwogIGZvciAoY29uc3QgW2tleSwgbGFiZWwsIGRpc3RLZXldIG9mIE1VTFRJX0RFRlMpIHsKICAgIGNvbnN0IG9w"
    "dHMgPSAoTUVUQS5kaXN0aW5jdCAmJiBNRVRBLmRpc3RpbmN0W2Rpc3RLZXldKSB8fCBbXTsKICAgIGdyaWQuYXBwZW5kQ2hpbGQobWFrZU11bHRpKGxhYmVs"
    "LCBrZXksIG9wdHMpKTsKICB9CiAgZ3JpZC5hcHBlbmRDaGlsZChtYWtlTXVsdGkoIlJGUSBSZXN1bHQiLCAicmZxUmVzdWx0IiwgUkZRX1JFU1VMVF9PUFRT"
    "KSk7Cn0KZnVuY3Rpb24gbWFrZVNlbGVjdChsYWJlbCwga2V5LCBvcHRzKSB7CiAgY29uc3QgdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImRpdiIpOyB3"
    "LmNsYXNzTmFtZSA9ICJzbGljZXIiOwogIHcuaW5uZXJIVE1MID0gYDxkaXYgY2xhc3M9InNsaWNlci1sYWJlbCI+JHtsYWJlbH08L2Rpdj5gOwogIGNvbnN0"
    "IHMgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCJzZWxlY3QiKTsKICBzLmlubmVySFRNTCA9IG9wdHMubWFwKG8gPT4gYDxvcHRpb24gdmFsdWU9IiR7ZXNj"
    "KG8pfSI+JHtlc2Mobyl9PC9vcHRpb24+YCkuam9pbigiIik7CiAgcy52YWx1ZSA9IEZJTFRFUlNba2V5XTsKICBzLm9uY2hhbmdlID0gKCkgPT4geyBGSUxU"
    "RVJTW2tleV0gPSBzLnZhbHVlOyBvbkZpbHRlckNoYW5nZSgpOyB9OwogIHcuYXBwZW5kQ2hpbGQocyk7IHJldHVybiB3Owp9CmZ1bmN0aW9uIG1ha2VEYXRl"
    "KGxhYmVsLCBrZXkpIHsKICBjb25zdCB3ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiZGl2Iik7IHcuY2xhc3NOYW1lID0gInNsaWNlciI7CiAgdy5pbm5l"
    "ckhUTUwgPSBgPGRpdiBjbGFzcz0ic2xpY2VyLWxhYmVsIj4ke2xhYmVsfTwvZGl2PmA7CiAgY29uc3QgaSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImlu"
    "cHV0Iik7IGkudHlwZSA9ICJkYXRlIjsgaS52YWx1ZSA9IEZJTFRFUlNba2V5XTsKICBpLm9uY2hhbmdlID0gKCkgPT4geyBGSUxURVJTW2tleV0gPSBpLnZh"
    "bHVlOyBvbkZpbHRlckNoYW5nZSgpOyB9OwogIHcuYXBwZW5kQ2hpbGQoaSk7IHJldHVybiB3Owp9CmZ1bmN0aW9uIG1ha2VNdWx0aShsYWJlbCwga2V5LCBv"
    "cHRzKSB7CiAgY29uc3QgdyA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImRpdiIpOyB3LmNsYXNzTmFtZSA9ICJzbGljZXIiOwogIGNvbnN0IHNlbCA9IEZJ"
    "TFRFUlMubXVsdGlba2V5XTsKICB3LmlubmVySFRNTCA9IGA8ZGl2IGNsYXNzPSJzbGljZXItbGFiZWwiPiR7bGFiZWx9PC9kaXY+YDsKICBjb25zdCBidG4g"
    "PSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCJidXR0b24iKTsgYnRuLmNsYXNzTmFtZSA9ICJtcy10b2dnbGUiOwogIGNvbnN0IHJlZnJlc2hMYmwgPSAoKSA9"
    "PiB7CiAgICBidG4uaW5uZXJIVE1MID0gYDxzcGFuPiR7c2VsLmxlbmd0aCA/IHNlbC5sZW5ndGggKyAiIHNlbGVjdGVkIiA6ICJBbGwifTwvc3Bhbj48c3Bh"
    "biBjbGFzcz0iY250Ij7ilr48L3NwYW4+YDsKICB9OwogIHJlZnJlc2hMYmwoKTsKICBjb25zdCBwYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImRp"
    "diIpOyBwYW5lbC5jbGFzc05hbWUgPSAibXMtcGFuZWwiOyBwYW5lbC5oaWRkZW4gPSB0cnVlOwogIGNvbnN0IHNlYXJjaCA9IGRvY3VtZW50LmNyZWF0ZUVs"
    "ZW1lbnQoImlucHV0Iik7CiAgc2VhcmNoLmNsYXNzTmFtZSA9ICJtcy1zZWFyY2ggc2xpY2VyIjsgc2VhcmNoLnBsYWNlaG9sZGVyID0gImZpbHRlcuKApiI7"
    "CiAgc2VhcmNoLnN0eWxlLmNzc1RleHQgPSAicGFkZGluZzo1cHggN3B4O2ZvbnQtc2l6ZToxMXB4O2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTti"
    "b3JkZXItcmFkaXVzOjZweDtiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMik7Y29sb3I6dmFyKC0tdGV4dCk7d2lkdGg6MTAwJSI7CiAgcGFuZWwuYXBwZW5k"
    "Q2hpbGQoc2VhcmNoKTsKICBjb25zdCBsaXN0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiZGl2Iik7IHBhbmVsLmFwcGVuZENoaWxkKGxpc3QpOwogIGNv"
    "bnN0IGRyYXcgPSAoZmx0ID0gIiIpID0+IHsKICAgIGxpc3QuaW5uZXJIVE1MID0gIiI7CiAgICBvcHRzLmZpbHRlcihvID0+IG8udG9Mb3dlckNhc2UoKS5p"
    "bmNsdWRlcyhmbHQudG9Mb3dlckNhc2UoKSkpLnNsaWNlKDAsIDMwMCkuZm9yRWFjaChvID0+IHsKICAgICAgY29uc3Qgcm93ID0gZG9jdW1lbnQuY3JlYXRl"
    "RWxlbWVudCgibGFiZWwiKTsgcm93LmNsYXNzTmFtZSA9ICJtcy1vcHQiOwogICAgICBjb25zdCBjYiA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImlucHV0"
    "Iik7IGNiLnR5cGUgPSAiY2hlY2tib3giOyBjYi5jaGVja2VkID0gc2VsLmluY2x1ZGVzKG8pOwogICAgICBjYi5vbmNoYW5nZSA9ICgpID0+IHsKICAgICAg"
    "ICBpZiAoY2IuY2hlY2tlZCkgeyBpZiAoIXNlbC5pbmNsdWRlcyhvKSkgc2VsLnB1c2gobyk7IH0KICAgICAgICBlbHNlIHsgY29uc3QgaSA9IHNlbC5pbmRl"
    "eE9mKG8pOyBpZiAoaSA+IC0xKSBzZWwuc3BsaWNlKGksIDEpOyB9CiAgICAgICAgcmVmcmVzaExibCgpOyBvbkZpbHRlckNoYW5nZSgpOwogICAgICB9Owog"
    "ICAgICByb3cuYXBwZW5kQ2hpbGQoY2IpOyByb3cuYXBwZW5kQ2hpbGQoZG9jdW1lbnQuY3JlYXRlVGV4dE5vZGUoIiAiICsgbykpOwogICAgICBsaXN0LmFw"
    "cGVuZENoaWxkKHJvdyk7CiAgICB9KTsKICB9OwogIGRyYXcoKTsKICBzZWFyY2gub25pbnB1dCA9ICgpID0+IGRyYXcoc2VhcmNoLnZhbHVlKTsKICBidG4u"
    "b25jbGljayA9IChlKSA9PiB7CiAgICBlLnN0b3BQcm9wYWdhdGlvbigpOwogICAgJCQoIi5tcy1wYW5lbCIpLmZvckVhY2gocCA9PiB7IGlmIChwICE9PSBw"
    "YW5lbCkgcC5oaWRkZW4gPSB0cnVlOyB9KTsKICAgIHBhbmVsLmhpZGRlbiA9ICFwYW5lbC5oaWRkZW47CiAgfTsKICBkb2N1bWVudC5hZGRFdmVudExpc3Rl"
    "bmVyKCJjbGljayIsIChlKSA9PiB7IGlmICghdy5jb250YWlucyhlLnRhcmdldCkpIHBhbmVsLmhpZGRlbiA9IHRydWU7IH0pOwogIHcuYXBwZW5kQ2hpbGQo"
    "YnRuKTsgdy5hcHBlbmRDaGlsZChwYW5lbCk7IHJldHVybiB3Owp9CgpmdW5jdGlvbiBtYXRjaFNlYXJjaChyLCBxKSB7CiAgaWYgKCFxKSByZXR1cm4gdHJ1"
    "ZTsKICBjb25zdCBmaWVsZHMgPSBbci5jdXN0b21lciwgci5jdXN0UmZxTm8sIHIuY3VzdFBvTm8sIHIuc3VwcGxpZXJQb05vLCByLnN1cHBsaWVyTmFtZSwK"
    "ICAgIHIuY3VzdFBPQywgci5ldFBPQywgci5pdGVtRGVzY3JpcHRpb24sIHIub2VtUGFydE5vLCByLmRyYXdpbmdObywgci5wcm9kdWN0Q2F0ZWdvcnksCiAg"
    "ICByLnN1cHBsaWVyUXVvdGVSZWYsIHIuc2VjdG9yLCByLnNoaXBtZW50U3RhdHVzLCByLnN1cHBsaWVyUmVtYXJrcywgci5yZnFTTm9dOwogIHJldHVybiBm"
    "aWVsZHMuc29tZShmID0+IGYgJiYgU3RyaW5nKGYpLnRvTG93ZXJDYXNlKCkuaW5jbHVkZXMocSkpOwp9CmZ1bmN0aW9uIHJmcVJlc3VsdE9mKHIpIHsKICBj"
    "b25zdCBxID0gci5ldFF1b3RlU3RhdHVzOwogIGlmIChxID09PSAiV29uIikgcmV0dXJuICJXb24iOwogIGlmIChxID09PSAiTG9zdCIpIHJldHVybiAiTG9z"
    "dCI7CiAgaWYgKHEgPT09ICJEZWNsaW5lZCIgfHwgci5ldFJmcVN0YXR1cyA9PT0gIkRlY2xpbmVkIikgcmV0dXJuICJEZWNsaW5lZCI7CiAgaWYgKHEgPT09"
    "ICJQZW5kaW5nIiB8fCBxID09PSAiVW5kZXIgQ2xhcmlmaWNhdGlvbiIpIHJldHVybiAiUGVuZGluZyI7CiAgcmV0dXJuICJPcGVuIjsKfQpmdW5jdGlvbiBh"
    "cHBseUZpbHRlcnMoKSB7CiAgY29uc3QgZiA9IEZJTFRFUlMsIHEgPSBmLnNlYXJjaC50cmltKCkudG9Mb3dlckNhc2UoKTsKICBjb25zdCBxTWFwID0geyBR"
    "MTogMSwgUTI6IDIsIFEzOiAzLCBRNDogNCB9OwogIGNvbnN0IG1vbklkeCA9IE1PTlRIUy5pbmRleE9mKGYubW9udGgpICsgMTsKICByZXR1cm4gUkFXLmZp"
    "bHRlcihyID0+IHsKICAgIGlmIChmLnllYXIgIT09ICJBbGwiICYmIHIuX3F5ICE9PSArZi55ZWFyKSByZXR1cm4gZmFsc2U7CiAgICBpZiAoZi5xdWFydGVy"
    "ICE9PSAiQWxsIiAmJiByLl9xcSAhPT0gcU1hcFtmLnF1YXJ0ZXJdKSByZXR1cm4gZmFsc2U7CiAgICBpZiAoZi5tb250aCAhPT0gIkFsbCIgJiYgci5fcW0g"
    "IT09IG1vbklkeCkgcmV0dXJuIGZhbHNlOwogICAgaWYgKGYuZGF0ZUZyb20gJiYgKCFyLmV0UXVvdGVEYXRlIHx8IHIuZXRRdW90ZURhdGUgPCBmLmRhdGVG"
    "cm9tKSkgcmV0dXJuIGZhbHNlOwogICAgaWYgKGYuZGF0ZVRvICYmICghci5ldFF1b3RlRGF0ZSB8fCByLmV0UXVvdGVEYXRlID4gZi5kYXRlVG8pKSByZXR1"
    "cm4gZmFsc2U7CiAgICBmb3IgKGNvbnN0IFtrZXldIG9mIE1VTFRJX0RFRlMpIHsKICAgICAgY29uc3Qgc2VsID0gZi5tdWx0aVtrZXldOwogICAgICBpZiAo"
    "c2VsLmxlbmd0aCAmJiAhc2VsLmluY2x1ZGVzKHJba2V5XSkpIHJldHVybiBmYWxzZTsKICAgIH0KICAgIGlmIChmLm11bHRpLnJmcVJlc3VsdC5sZW5ndGgg"
    "JiYgIWYubXVsdGkucmZxUmVzdWx0LmluY2x1ZGVzKHJmcVJlc3VsdE9mKHIpKSkgcmV0dXJuIGZhbHNlOwogICAgZm9yIChjb25zdCBbZGltLCB2YWxdIG9m"
    "IE9iamVjdC5lbnRyaWVzKGYuY2hhcnQpKSB7CiAgICAgIGlmIChkaW0gPT09ICJtb250aEtleSIpIHsgaWYgKG1vbnRoS2V5KHIuZXRRdW90ZURhdGUpICE9"
    "PSB2YWwpIHJldHVybiBmYWxzZTsgfQogICAgICBlbHNlIGlmIChkaW0gPT09ICJfX3Jlc3VsdCIpIHsgaWYgKHJmcVJlc3VsdE9mKHIpICE9PSB2YWwpIHJl"
    "dHVybiBmYWxzZTsgfQogICAgICBlbHNlIGlmIChyW2RpbV0gIT09IHZhbCkgcmV0dXJuIGZhbHNlOwogICAgfQogICAgaWYgKHEgJiYgIW1hdGNoU2VhcmNo"
    "KHIsIHEpKSByZXR1cm4gZmFsc2U7CiAgICByZXR1cm4gdHJ1ZTsKICB9KTsKfQpmdW5jdGlvbiBhY3RpdmVGaWx0ZXJDb3VudCgpIHsKICBsZXQgbiA9IDA7"
    "CiAgaWYgKEZJTFRFUlMuc2VhcmNoLnRyaW0oKSkgbisrOwogIFsieWVhciIsICJxdWFydGVyIiwgIm1vbnRoIiwgImRhdGVGcm9tIiwgImRhdGVUbyJdLmZv"
    "ckVhY2goayA9PiB7CiAgICBpZiAoRklMVEVSU1trXSAmJiBGSUxURVJTW2tdICE9PSAiQWxsIikgbisrOwogIH0pOwogIGZvciAoY29uc3QgayBpbiBGSUxU"
    "RVJTLm11bHRpKSBpZiAoRklMVEVSUy5tdWx0aVtrXS5sZW5ndGgpIG4rKzsKICBuICs9IE9iamVjdC5rZXlzKEZJTFRFUlMuY2hhcnQpLmxlbmd0aDsKICBy"
    "ZXR1cm4gbjsKfQpmdW5jdGlvbiB1cGRhdGVDaGlwcygpIHsKICBjb25zdCByb3cgPSAkKCIjY2hpcHNSb3ciKTsgcm93LmlubmVySFRNTCA9ICIiOwogIGNv"
    "bnN0IGFkZCA9IChsYWJlbCwgdmFsLCBvbkNsZWFyKSA9PiB7CiAgICBjb25zdCBjID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiZGl2Iik7IGMuY2xhc3NO"
    "YW1lID0gImNoaXAiOwogICAgYy5pbm5lckhUTUwgPSBgJHtsYWJlbH06IDxiPiR7ZXNjKHZhbCl9PC9iPmA7CiAgICBjb25zdCBiID0gZG9jdW1lbnQuY3Jl"
    "YXRlRWxlbWVudCgiYnV0dG9uIik7IGIudGV4dENvbnRlbnQgPSAiw5ciOyBiLm9uY2xpY2sgPSBvbkNsZWFyOwogICAgYy5hcHBlbmRDaGlsZChiKTsgcm93"
    "LmFwcGVuZENoaWxkKGMpOwogIH07CiAgaWYgKEZJTFRFUlMuc2VhcmNoLnRyaW0oKSkgYWRkKCJTZWFyY2giLCBGSUxURVJTLnNlYXJjaCwgKCkgPT4geyBG"
    "SUxURVJTLnNlYXJjaCA9ICIiOyAkKCIjZ2xvYmFsU2VhcmNoIikudmFsdWUgPSAiIjsgJCgiI3NlYXJjaENsZWFyIikuaGlkZGVuID0gdHJ1ZTsgb25GaWx0"
    "ZXJDaGFuZ2UoKTsgfSk7CiAgWyJ5ZWFyIiwgInF1YXJ0ZXIiLCAibW9udGgiXS5mb3JFYWNoKGsgPT4geyBpZiAoRklMVEVSU1trXSAhPT0gIkFsbCIpIGFk"
    "ZChrWzBdLnRvVXBwZXJDYXNlKCkgKyBrLnNsaWNlKDEpLCBGSUxURVJTW2tdLCAoKSA9PiB7IEZJTFRFUlNba10gPSAiQWxsIjsgYnVpbGRTbGljZXJzKCk7"
    "IG9uRmlsdGVyQ2hhbmdlKCk7IH0pOyB9KTsKICBpZiAoRklMVEVSUy5kYXRlRnJvbSkgYWRkKCJGcm9tIiwgRklMVEVSUy5kYXRlRnJvbSwgKCkgPT4geyBG"
    "SUxURVJTLmRhdGVGcm9tID0gIiI7IGJ1aWxkU2xpY2VycygpOyBvbkZpbHRlckNoYW5nZSgpOyB9KTsKICBpZiAoRklMVEVSUy5kYXRlVG8pIGFkZCgiVG8i"
    "LCBGSUxURVJTLmRhdGVUbywgKCkgPT4geyBGSUxURVJTLmRhdGVUbyA9ICIiOyBidWlsZFNsaWNlcnMoKTsgb25GaWx0ZXJDaGFuZ2UoKTsgfSk7CiAgTVVM"
    "VElfREVGUy5jb25jYXQoW1sicmZxUmVzdWx0IiwgIlJGUSBSZXN1bHQiXV0pLmZvckVhY2goKFtrZXksIGxhYmVsXSkgPT4gewogICAgRklMVEVSUy5tdWx0"
    "aVtrZXldLmZvckVhY2godiA9PiBhZGQobGFiZWwsIHYsICgpID0+IHsKICAgICAgY29uc3QgaSA9IEZJTFRFUlMubXVsdGlba2V5XS5pbmRleE9mKHYpOyBG"
    "SUxURVJTLm11bHRpW2tleV0uc3BsaWNlKGksIDEpOyBidWlsZFNsaWNlcnMoKTsgb25GaWx0ZXJDaGFuZ2UoKTsKICAgIH0pKTsKICB9KTsKICBPYmplY3Qu"
    "ZW50cmllcyhGSUxURVJTLmNoYXJ0KS5mb3JFYWNoKChbZGltLCB2YWxdKSA9PiBhZGQoIkNoYXJ0IiwgdmFsLCAoKSA9PiB7IGRlbGV0ZSBGSUxURVJTLmNo"
    "YXJ0W2RpbV07IG9uRmlsdGVyQ2hhbmdlKCk7IH0pKTsKfQpmdW5jdGlvbiByZXNldEZpbHRlcnMoKSB7CiAgRklMVEVSUy5zZWFyY2ggPSAiIjsgRklMVEVS"
    "Uy5kYXRlRnJvbSA9ICIiOyBGSUxURVJTLmRhdGVUbyA9ICIiOwogIEZJTFRFUlMueWVhciA9ICJBbGwiOyBGSUxURVJTLnF1YXJ0ZXIgPSAiQWxsIjsgRklM"
    "VEVSUy5tb250aCA9ICJBbGwiOwogIGZvciAoY29uc3QgayBpbiBGSUxURVJTLm11bHRpKSBGSUxURVJTLm11bHRpW2tdID0gW107CiAgRklMVEVSUy5jaGFy"
    "dCA9IHt9OwogICQoIiNnbG9iYWxTZWFyY2giKS52YWx1ZSA9ICIiOyAkKCIjc2VhcmNoQ2xlYXIiKS5oaWRkZW4gPSB0cnVlOwogIGJ1aWxkU2xpY2Vycygp"
    "OyBvbkZpbHRlckNoYW5nZSgpOwp9CmZ1bmN0aW9uIGNoYXJ0RmlsdGVyKGRpbSwgdmFsKSB7CiAgaWYgKEZJTFRFUlMuY2hhcnRbZGltXSA9PT0gdmFsKSBk"
    "ZWxldGUgRklMVEVSUy5jaGFydFtkaW1dOwogIGVsc2UgRklMVEVSUy5jaGFydFtkaW1dID0gdmFsOwogIG9uRmlsdGVyQ2hhbmdlKCk7Cn0KZnVuY3Rpb24g"
    "b25GaWx0ZXJDaGFuZ2UoKSB7IFJFTkRFUkVEID0ge307IHJlZnJlc2goKTsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSA3LiBDSEFSVCArIEtQSSBIRUxQRVJT"
    "IC0tLS0tLS0tLS0tLS0tLS0gKi8KZnVuY3Rpb24gcmVuZGVyQ2hhcnQoaWQsIGNvbmZpZykgewogIGNvbnN0IGVsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5"
    "SWQoaWQpOwogIGlmICghZWwpIHJldHVybjsKICBpZiAoQ0hBUlRTW2lkXSkgeyBDSEFSVFNbaWRdLmRlc3Ryb3koKTsgZGVsZXRlIENIQVJUU1tpZF07IH0K"
    "ICBjb25zdCBwID0gcGFsZXR0ZSgpOwogIENoYXJ0LmRlZmF1bHRzLmNvbG9yID0gcC5tdXRlZDsKICBDaGFydC5kZWZhdWx0cy5mb250LmZhbWlseSA9IGdl"
    "dENvbXB1dGVkU3R5bGUoZG9jdW1lbnQuYm9keSkuZm9udEZhbWlseTsKICBjb25maWcub3B0aW9ucyA9IGNvbmZpZy5vcHRpb25zIHx8IHt9OwogIGNvbmZp"
    "Zy5vcHRpb25zLnJlc3BvbnNpdmUgPSB0cnVlOwogIGNvbmZpZy5vcHRpb25zLm1haW50YWluQXNwZWN0UmF0aW8gPSBmYWxzZTsKICBjb25maWcub3B0aW9u"
    "cy5hbmltYXRpb24gPSB7IGR1cmF0aW9uOiAzMDAgfTsKICBjb25maWcub3B0aW9ucy5wbHVnaW5zID0gY29uZmlnLm9wdGlvbnMucGx1Z2lucyB8fCB7fTsK"
    "ICBjb25maWcub3B0aW9ucy5wbHVnaW5zLmxlZ2VuZCA9IE9iamVjdC5hc3NpZ24oeyBsYWJlbHM6IHsgY29sb3I6IHAudGV4dCwgYm94V2lkdGg6IDEyLCBm"
    "b250OiB7IHNpemU6IDExIH0gfSB9LCBjb25maWcub3B0aW9ucy5wbHVnaW5zLmxlZ2VuZCB8fCB7fSk7CiAgaWYgKGNvbmZpZy5vcHRpb25zLnNjYWxlcykg"
    "ewogICAgZm9yIChjb25zdCBheCBvZiBPYmplY3QudmFsdWVzKGNvbmZpZy5vcHRpb25zLnNjYWxlcykpIHsKICAgICAgYXguZ3JpZCA9IE9iamVjdC5hc3Np"
    "Z24oeyBjb2xvcjogcC5ncmlkIH0sIGF4LmdyaWQgfHwge30pOwogICAgICBheC50aWNrcyA9IE9iamVjdC5hc3NpZ24oeyBjb2xvcjogcC5tdXRlZCwgZm9u"
    "dDogeyBzaXplOiAxMC41IH0gfSwgYXgudGlja3MgfHwge30pOwogICAgfQogIH0KICBDSEFSVFNbaWRdID0gbmV3IENoYXJ0KGVsLCBjb25maWcpOwogIHJl"
    "dHVybiBDSEFSVFNbaWRdOwp9CmZ1bmN0aW9uIGtwaShsYWJlbCwgdmFsdWUsIHN1YiwgY2xzKSB7CiAgcmV0dXJuIGA8ZGl2IGNsYXNzPSJrcGkgJHtjbHMg"
    "fHwgIiJ9Ij4KICAgIDxkaXYgY2xhc3M9ImtwaS1sYWJlbCI+JHtlc2MobGFiZWwpfTwvZGl2PgogICAgPGRpdiBjbGFzcz0ia3BpLXZhbHVlIj4ke3ZhbHVl"
    "fTwvZGl2PgogICAgJHtzdWIgPyBgPGRpdiBjbGFzcz0ia3BpLXN1YiI+JHtzdWJ9PC9kaXY+YCA6ICIifTwvZGl2PmA7Cn0KZnVuY3Rpb24gY2FyZFNoZWxs"
    "KHRpdGxlLCBpZCwgc3BhbiA9ICJjNiIsIGhpbnQgPSAiIikgewogIHJldHVybiBgPGRpdiBjbGFzcz0iY2FyZCAke3NwYW59Ij4KICAgIDxkaXYgY2xhc3M9"
    "ImNhcmQtaGVhZCI+PGRpdiBjbGFzcz0iY2FyZC10aXRsZSI+JHtlc2ModGl0bGUpfTwvZGl2PgogICAgJHtoaW50ID8gYDxkaXYgY2xhc3M9ImNhcmQtaGlu"
    "dCI+JHtlc2MoaGludCl9PC9kaXY+YCA6ICIifTwvZGl2PgogICAgPGRpdiBjbGFzcz0iY2hhcnQtaG9sZGVyIj48Y2FudmFzIGlkPSIke2lkfSI+PC9jYW52"
    "YXM+PC9kaXY+PC9kaXY+YDsKfQpmdW5jdGlvbiB0b3BOKG1hcCwgbikgewogIHJldHVybiBBcnJheS5mcm9tKG1hcC5lbnRyaWVzKCkpLnNvcnQoKGEsIGIp"
    "ID0+IGJbMV0gLSBhWzFdKS5zbGljZSgwLCBuKTsKfQoKLyogLS0tLS0tLS0tLS0tLS0tLSA4LiBHRU5FUklDIFRBQkxFIENPTVBPTkVOVCAtLS0tLS0tLS0t"
    "LS0tLS0tICovCi8qIGNmZzogeyBjb2x1bW5zOlt7a2V5LGxhYmVsLHR5cGUsZm10LHZpc31dLCByb3dzOltdLCBwZXJQYWdlLCBzZWFyY2hhYmxlLCBkcmls"
    "bCB9ICovCmZ1bmN0aW9uIG1ha2VUYWJsZShjb250YWluZXIsIGNmZykgewogIGNvbnN0IHN0YXRlID0geyBzb3J0OiBudWxsLCBkaXI6IDEsIHBhZ2U6IDEs"
    "IHBlcjogY2ZnLnBlclBhZ2UgfHwgMjUsIHE6ICIiLAogICAgdmlzOiBjZmcuY29sdW1ucy5tYXAoYyA9PiBjLnZpcyAhPT0gZmFsc2UpIH07CiAgY29udGFp"
    "bmVyLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IHRvb2xzID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiZGl2Iik7IHRvb2xzLmNsYXNzTmFtZSA9ICJ0YWJs"
    "ZS10b29scyI7CiAgY29uc3Qgc2VhcmNoID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiaW5wdXQiKTsgc2VhcmNoLnR5cGUgPSAidGV4dCI7IHNlYXJjaC5w"
    "bGFjZWhvbGRlciA9ICJTZWFyY2ggdGFibGXigKYiOwogIGNvbnN0IHBlclNlbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoInNlbGVjdCIpOwogIFsxMCwg"
    "MjUsIDUwLCAxMDBdLmZvckVhY2gobiA9PiBwZXJTZWwuYXBwZW5kQ2hpbGQobmV3IE9wdGlvbihuICsgIiAvIHBhZ2UiLCBuLCBuID09PSBzdGF0ZS5wZXIs"
    "IG4gPT09IHN0YXRlLnBlcikpKTsKICBjb25zdCBjb2xCdG4gPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCJidXR0b24iKTsgY29sQnRuLmNsYXNzTmFtZSA9"
    "ICJidG4gYnRuLWdob3N0IjsgY29sQnRuLnRleHRDb250ZW50ID0gIkNvbHVtbnMg4pa+IjsKICBjb25zdCBjb2xXcmFwID0gZG9jdW1lbnQuY3JlYXRlRWxl"
    "bWVudCgiZGl2Iik7IGNvbFdyYXAuY2xhc3NOYW1lID0gImNvbC1tZW51IjsKICBjb25zdCBjb2xQYW5lbCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImRp"
    "diIpOyBjb2xQYW5lbC5jbGFzc05hbWUgPSAiY29sLW1lbnUtcGFuZWwiOyBjb2xQYW5lbC5oaWRkZW4gPSB0cnVlOwogIGNmZy5jb2x1bW5zLmZvckVhY2go"
    "KGMsIGkpID0+IHsKICAgIGNvbnN0IGwgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCJsYWJlbCIpOwogICAgY29uc3QgY2IgPSBkb2N1bWVudC5jcmVhdGVF"
    "bGVtZW50KCJpbnB1dCIpOyBjYi50eXBlID0gImNoZWNrYm94IjsgY2IuY2hlY2tlZCA9IHN0YXRlLnZpc1tpXTsKICAgIGNiLm9uY2hhbmdlID0gKCkgPT4g"
    "eyBzdGF0ZS52aXNbaV0gPSBjYi5jaGVja2VkOyByZW5kZXIoKTsgfTsKICAgIGwuYXBwZW5kQ2hpbGQoY2IpOyBsLmFwcGVuZENoaWxkKGRvY3VtZW50LmNy"
    "ZWF0ZVRleHROb2RlKCIgIiArIGMubGFiZWwpKTsgY29sUGFuZWwuYXBwZW5kQ2hpbGQobCk7CiAgfSk7CiAgY29sQnRuLm9uY2xpY2sgPSAoZSkgPT4geyBl"
    "LnN0b3BQcm9wYWdhdGlvbigpOyBjb2xQYW5lbC5oaWRkZW4gPSAhY29sUGFuZWwuaGlkZGVuOyB9OwogIGRvY3VtZW50LmFkZEV2ZW50TGlzdGVuZXIoImNs"
    "aWNrIiwgKGUpID0+IHsgaWYgKCFjb2xXcmFwLmNvbnRhaW5zKGUudGFyZ2V0KSkgY29sUGFuZWwuaGlkZGVuID0gdHJ1ZTsgfSk7CiAgY29sV3JhcC5hcHBl"
    "bmRDaGlsZChjb2xCdG4pOyBjb2xXcmFwLmFwcGVuZENoaWxkKGNvbFBhbmVsKTsKICBjb25zdCBleHAgPSBkb2N1bWVudC5jcmVhdGVFbGVtZW50KCJidXR0"
    "b24iKTsgZXhwLmNsYXNzTmFtZSA9ICJidG4gYnRuLWdob3N0IjsgZXhwLnRleHRDb250ZW50ID0gIkNTViI7CiAgY29uc3QgY291bnQgPSBkb2N1bWVudC5j"
    "cmVhdGVFbGVtZW50KCJzcGFuIik7IGNvdW50LnN0eWxlLmNzc1RleHQgPSAiZm9udC1zaXplOjEycHg7Y29sb3I6dmFyKC0tbXV0ZWQpO21hcmdpbi1sZWZ0"
    "OmF1dG8iOwogIHRvb2xzLmFwcGVuZChzZWFyY2gsIHBlclNlbCwgY29sV3JhcCwgZXhwLCBjb3VudCk7CiAgY29udGFpbmVyLmFwcGVuZENoaWxkKHRvb2xz"
    "KTsKCiAgY29uc3Qgc2Nyb2xsID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiZGl2Iik7IHNjcm9sbC5jbGFzc05hbWUgPSAidGFibGUtc2Nyb2xsIjsKICBj"
    "b25zdCB0YWJsZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoInRhYmxlIik7IHRhYmxlLmNsYXNzTmFtZSA9ICJkYXRhIjsKICBzY3JvbGwuYXBwZW5kQ2hp"
    "bGQodGFibGUpOyBjb250YWluZXIuYXBwZW5kQ2hpbGQoc2Nyb2xsKTsKICBjb25zdCBwYWdlciA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImRpdiIpOyBw"
    "YWdlci5jbGFzc05hbWUgPSAicGFnZXIiOyBjb250YWluZXIuYXBwZW5kQ2hpbGQocGFnZXIpOwoKICBmdW5jdGlvbiBmaWx0ZXJlZCgpIHsKICAgIGxldCBy"
    "b3dzID0gY2ZnLnJvd3M7CiAgICBpZiAoc3RhdGUucSkgewogICAgICBjb25zdCBxID0gc3RhdGUucS50b0xvd2VyQ2FzZSgpOwogICAgICByb3dzID0gcm93"
    "cy5maWx0ZXIociA9PiBjZmcuY29sdW1ucy5zb21lKChjLCBpKSA9PiBzdGF0ZS52aXNbaV0gJiYgU3RyaW5nKHJbYy5rZXldID8/ICIiKS50b0xvd2VyQ2Fz"
    "ZSgpLmluY2x1ZGVzKHEpKSk7CiAgICB9CiAgICBpZiAoc3RhdGUuc29ydCkgewogICAgICBjb25zdCBjID0gY2ZnLmNvbHVtbnMuZmluZChjID0+IGMua2V5"
    "ID09PSBzdGF0ZS5zb3J0KTsKICAgICAgcm93cyA9IHJvd3Muc2xpY2UoKS5zb3J0KChhLCBiKSA9PiB7CiAgICAgICAgbGV0IHggPSBhW3N0YXRlLnNvcnRd"
    "LCB5ID0gYltzdGF0ZS5zb3J0XTsKICAgICAgICBpZiAoYyAmJiAoYy50eXBlID09PSAibnVtIiB8fCBjLnR5cGUgPT09ICJjdXIiIHx8IGMudHlwZSA9PT0g"
    "InBjdCIpKSB7IHggPSAreCB8fCAtSW5maW5pdHk7IHkgPSAreSB8fCAtSW5maW5pdHk7IH0KICAgICAgICBlbHNlIHsgeCA9IFN0cmluZyh4ID8/ICIiKTsg"
    "eSA9IFN0cmluZyh5ID8/ICIiKTsgfQogICAgICAgIHJldHVybiB4IDwgeSA/IC1zdGF0ZS5kaXIgOiB4ID4geSA/IHN0YXRlLmRpciA6IDA7CiAgICAgIH0p"
    "OwogICAgfQogICAgcmV0dXJuIHJvd3M7CiAgfQogIGZ1bmN0aW9uIGNlbGxIVE1MKGMsIHIpIHsKICAgIGNvbnN0IHYgPSByW2Mua2V5XTsKICAgIGlmIChj"
    "LnJlbmRlcikgcmV0dXJuIGMucmVuZGVyKHYsIHIpOwogICAgaWYgKHYgPT0gbnVsbCB8fCB2ID09PSAiIikgcmV0dXJuICLigJQiOwogICAgaWYgKGMudHlw"
    "ZSA9PT0gImN1ciIpIHJldHVybiBmbXRDdXIodik7CiAgICBpZiAoYy50eXBlID09PSAibnVtIikgcmV0dXJuIGZtdE51bSh2KTsKICAgIGlmIChjLnR5cGUg"
    "PT09ICJwY3QiKSByZXR1cm4gZm10UGN0KHYpOwogICAgaWYgKGMudHlwZSA9PT0gImRhdGUiKSByZXR1cm4gZm10RGF0ZSh2KTsKICAgIGlmIChjLnR5cGUg"
    "PT09ICJkYXlzIikgcmV0dXJuIGZtdERheXModik7CiAgICByZXR1cm4gZXNjKHYpOwogIH0KICBmdW5jdGlvbiByZW5kZXIoKSB7CiAgICBjb25zdCByb3dz"
    "ID0gZmlsdGVyZWQoKTsKICAgIGNvbnN0IHBhZ2VzID0gTWF0aC5tYXgoMSwgTWF0aC5jZWlsKHJvd3MubGVuZ3RoIC8gc3RhdGUucGVyKSk7CiAgICBzdGF0"
    "ZS5wYWdlID0gTWF0aC5taW4oc3RhdGUucGFnZSwgcGFnZXMpOwogICAgY29uc3Qgc2xpY2UgPSByb3dzLnNsaWNlKChzdGF0ZS5wYWdlIC0gMSkgKiBzdGF0"
    "ZS5wZXIsIHN0YXRlLnBhZ2UgKiBzdGF0ZS5wZXIpOwogICAgY29uc3QgY29scyA9IGNmZy5jb2x1bW5zLmZpbHRlcigoYywgaSkgPT4gc3RhdGUudmlzW2ld"
    "KTsKICAgIHRhYmxlLmlubmVySFRNTCA9CiAgICAgICI8dGhlYWQ+PHRyPiIgKyBjb2xzLm1hcChjID0+IHsKICAgICAgICBjb25zdCBhciA9IHN0YXRlLnNv"
    "cnQgPT09IGMua2V5ID8gKHN0YXRlLmRpciA9PT0gMSA/ICLilrIiIDogIuKWvCIpIDogIiI7CiAgICAgICAgcmV0dXJuIGA8dGggZGF0YS1rPSIke2Mua2V5"
    "fSIgY2xhc3M9IiR7Yy50eXBlID09PSAnY3VyJyB8fCBjLnR5cGUgPT09ICdudW0nIHx8IGMudHlwZSA9PT0gJ3BjdCcgfHwgYy50eXBlID09PSAnZGF5cycg"
    "PyAnbnVtJyA6ICcnfSI+JHtlc2MoYy5sYWJlbCl9IDxzcGFuIGNsYXNzPSJhcnJvdyI+JHthcn08L3NwYW4+PC90aD5gOwogICAgICB9KS5qb2luKCIiKSAr"
    "ICI8L3RyPjwvdGhlYWQ+PHRib2R5PiIgKwogICAgICAoc2xpY2UubGVuZ3RoID8gc2xpY2UubWFwKChyLCByaSkgPT4gIjx0ciBkYXRhLWk9JyIgKyByaSAr"
    "ICInPiIgKyBjb2xzLm1hcChjID0+CiAgICAgICAgYDx0ZCBjbGFzcz0iJHtjLnR5cGUgPT09ICdjdXInIHx8IGMudHlwZSA9PT0gJ251bScgfHwgYy50eXBl"
    "ID09PSAncGN0JyB8fCBjLnR5cGUgPT09ICdkYXlzJyA/ICdudW0nIDogJyd9Ij4ke2NlbGxIVE1MKGMsIHIpfTwvdGQ+YAogICAgICApLmpvaW4oIiIpICsg"
    "IjwvdHI+Iikuam9pbigiIikKICAgICAgICA6IGA8dHI+PHRkIGNvbHNwYW49IiR7Y29scy5sZW5ndGh9Ij48ZGl2IGNsYXNzPSJlbXB0eS1zdGF0ZSIgc3R5"
    "bGU9Im1hcmdpbjo4cHgiPk5vIHJlY29yZHMgbWF0Y2ggdGhlIGN1cnJlbnQgZmlsdGVycy48L2Rpdj48L3RkPjwvdHI+YCkgKwogICAgICAiPC90Ym9keT4i"
    "OwogICAgdGFibGUucXVlcnlTZWxlY3RvckFsbCgidGhlYWQgdGgiKS5mb3JFYWNoKHRoID0+IHRoLm9uY2xpY2sgPSAoKSA9PiB7CiAgICAgIGNvbnN0IGsg"
    "PSB0aC5kYXRhc2V0Lms7CiAgICAgIGlmIChzdGF0ZS5zb3J0ID09PSBrKSBzdGF0ZS5kaXIgKj0gLTE7IGVsc2UgeyBzdGF0ZS5zb3J0ID0gazsgc3RhdGUu"
    "ZGlyID0gMTsgfQogICAgICByZW5kZXIoKTsKICAgIH0pOwogICAgaWYgKGNmZy5kcmlsbCkgdGFibGUucXVlcnlTZWxlY3RvckFsbCgidGJvZHkgdHJbZGF0"
    "YS1pXSIpLmZvckVhY2godHIgPT4gdHIub25jbGljayA9ICgpID0+IGNmZy5kcmlsbChzbGljZVsrdHIuZGF0YXNldC5pXSkpOwogICAgY291bnQudGV4dENv"
    "bnRlbnQgPSByb3dzLmxlbmd0aC50b0xvY2FsZVN0cmluZygpICsgIiByb3dzIjsKICAgIHBhZ2VyLmlubmVySFRNTCA9ICIiOwogICAgY29uc3QgcHJldiA9"
    "IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImJ1dHRvbiIpOyBwcmV2LmNsYXNzTmFtZSA9ICJidG4gYnRuLWdob3N0IjsgcHJldi50ZXh0Q29udGVudCA9ICLi"
    "gLkgUHJldiI7IHByZXYuZGlzYWJsZWQgPSBzdGF0ZS5wYWdlIDw9IDE7CiAgICBjb25zdCBuZXh0ID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgiYnV0dG9u"
    "Iik7IG5leHQuY2xhc3NOYW1lID0gImJ0biBidG4tZ2hvc3QiOyBuZXh0LnRleHRDb250ZW50ID0gIk5leHQg4oC6IjsgbmV4dC5kaXNhYmxlZCA9IHN0YXRl"
    "LnBhZ2UgPj0gcGFnZXM7CiAgICBwcmV2Lm9uY2xpY2sgPSAoKSA9PiB7IHN0YXRlLnBhZ2UtLTsgcmVuZGVyKCk7IH07IG5leHQub25jbGljayA9ICgpID0+"
    "IHsgc3RhdGUucGFnZSsrOyByZW5kZXIoKTsgfTsKICAgIGNvbnN0IGxibCA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoInNwYW4iKTsgbGJsLnRleHRDb250"
    "ZW50ID0gYFBhZ2UgJHtzdGF0ZS5wYWdlfSBvZiAke3BhZ2VzfWA7CiAgICBwYWdlci5hcHBlbmQocHJldiwgbGJsLCBuZXh0KTsKICB9CiAgbGV0IGRlYjsK"
    "ICBzZWFyY2gub25pbnB1dCA9ICgpID0+IHsgY2xlYXJUaW1lb3V0KGRlYik7IGRlYiA9IHNldFRpbWVvdXQoKCkgPT4geyBzdGF0ZS5xID0gc2VhcmNoLnZh"
    "bHVlOyBzdGF0ZS5wYWdlID0gMTsgcmVuZGVyKCk7IH0sIDE4MCk7IH07CiAgcGVyU2VsLm9uY2hhbmdlID0gKCkgPT4geyBzdGF0ZS5wZXIgPSArcGVyU2Vs"
    "LnZhbHVlOyBzdGF0ZS5wYWdlID0gMTsgcmVuZGVyKCk7IH07CiAgZXhwLm9uY2xpY2sgPSAoKSA9PiBleHBvcnRSb3dzQ1NWKGNmZy5jb2x1bW5zLmZpbHRl"
    "cigoYywgaSkgPT4gc3RhdGUudmlzW2ldKSwgZmlsdGVyZWQoKSwgKGNmZy5uYW1lIHx8ICJ0YWJsZSIpICsgIi5jc3YiKTsKICByZW5kZXIoKTsKfQoKLyog"
    "LS0tLS0tLS0tLS0tLS0tLSA5LiBSRlEgUk9MTC1VUCBFTkdJTkUgLS0tLS0tLS0tLS0tLS0tLSAqLwpmdW5jdGlvbiBidWlsZFJvbGx1cChyb3dzKSB7CiAg"
    "Y29uc3QgcmV2ID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBjb2dzID0gZ3JvdXBEZWR1cChyb3dzLCAi"
    "cmZxS2V5IiwgInN1cHBsaWVyVG90YWxQcmljZSIpOwogIGNvbnN0IGdwID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIp"
    "OwogIGNvbnN0IGJ5S2V5ID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgciBvZiByb3dzKSB7CiAgICBjb25zdCBrID0gci5yZnFLZXk7CiAgICBsZXQgbyA9"
    "IGJ5S2V5LmdldChrKTsKICAgIGlmICghbykgewogICAgICBvID0geyBrZXk6IGssIGN1c3RvbWVyOiByLmN1c3RvbWVyLCBldFBPQzogci5ldFBPQywgY3Vz"
    "dFBPQzogci5jdXN0UE9DLAogICAgICAgIHByb2R1Y3RDYXRlZ29yeTogci5wcm9kdWN0Q2F0ZWdvcnksIHNlY3Rvcjogci5zZWN0b3IsIHN1cHBsaWVyOiBy"
    "LnN1cHBsaWVyTmFtZSwKICAgICAgICBjdXN0UmZxRGF0ZTogci5jdXN0UmZxRGF0ZSwgZXRRdW90ZURhdGU6IHIuZXRRdW90ZURhdGUsIGNsb3NpbmdEYXRl"
    "OiByLmN1c3RSZnFDbG9zaW5nRGF0ZSwKICAgICAgICBjdXN0UG9Obzogci5jdXN0UG9ObywgY3VzdFBvRGF0ZTogci5jdXN0UG9EYXRlLCBjdXN0UG9LZXk6"
    "IHIuY3VzdFBvS2V5LAogICAgICAgIHJldjogcmV2LmdldChrKSB8fCAwLCBjb2dzOiBjb2dzLmdldChrKSB8fCAwLCBncDogZ3AuZ2V0KGspIHx8IDAsCiAg"
    "ICAgICAgd29uOiBmYWxzZSwgbG9zdDogZmFsc2UsIGRlY2xpbmVkOiBmYWxzZSwgcXVvdGVkOiBmYWxzZSwgbGluZXM6IDAsCiAgICAgICAgcmVzcERheXM6"
    "IFtdLCBjbG9zaW5nVmFyOiBbXSB9OwogICAgICBieUtleS5zZXQoaywgbyk7CiAgICB9CiAgICBvLmxpbmVzKys7CiAgICBpZiAoci5ldFF1b3RlU3RhdHVz"
    "ID09PSAiV29uIikgby53b24gPSB0cnVlOwogICAgaWYgKHIuZXRRdW90ZVN0YXR1cyA9PT0gIkxvc3QiKSBvLmxvc3QgPSB0cnVlOwogICAgaWYgKHIuZXRR"
    "dW90ZVN0YXR1cyA9PT0gIkRlY2xpbmVkIiB8fCByLmV0UmZxU3RhdHVzID09PSAiRGVjbGluZWQiKSBvLmRlY2xpbmVkID0gdHJ1ZTsKICAgIGlmIChyLmV0"
    "UXVvdGVEYXRlKSBvLnF1b3RlZCA9IHRydWU7CiAgICBpZiAoci5jdXN0UG9ObyAmJiAhby5jdXN0UG9ObykgeyBvLmN1c3RQb05vID0gci5jdXN0UG9Obzsg"
    "by5jdXN0UG9LZXkgPSByLmN1c3RQb0tleTsgby5jdXN0UG9EYXRlID0gci5jdXN0UG9EYXRlOyB9CiAgICBpZiAoci5yZnFSZXNwb25zZURheXMgIT0gbnVs"
    "bCkgby5yZXNwRGF5cy5wdXNoKHIucmZxUmVzcG9uc2VEYXlzKTsKICAgIGlmIChyLmNsb3NpbmdWYXJpYW5jZURheXMgIT0gbnVsbCkgby5jbG9zaW5nVmFy"
    "LnB1c2goci5jbG9zaW5nVmFyaWFuY2VEYXlzKTsKICAgIGlmIChyLmV0UXVvdGVEYXRlICYmICghby5ldFF1b3RlRGF0ZSB8fCByLmV0UXVvdGVEYXRlIDwg"
    "by5ldFF1b3RlRGF0ZSkpIG8uZXRRdW90ZURhdGUgPSByLmV0UXVvdGVEYXRlOwogIH0KICBjb25zdCBsaXN0ID0gW107CiAgZm9yIChjb25zdCBvIG9mIGJ5"
    "S2V5LnZhbHVlcygpKSB7CiAgICBvLm1hcmdpbiA9IG8ucmV2ID8gKG8uZ3AgLyBvLnJldikgKiAxMDAgOiBudWxsOwogICAgby5yZXN1bHQgPSBvLndvbiA/"
    "ICJXb24iIDogby5sb3N0ID8gIkxvc3QiIDogby5kZWNsaW5lZCA/ICJEZWNsaW5lZCIgOiBvLnF1b3RlZCA/ICJPcGVuIiA6ICJQZW5kaW5nIjsKICAgIG8u"
    "cmVzcERheSA9IG8ucmVzcERheXMubGVuZ3RoID8gTWF0aC5taW4oLi4uby5yZXNwRGF5cykgOiBudWxsOwogICAgby5jbG9zaW5nVmFyRGF5ID0gby5jbG9z"
    "aW5nVmFyLmxlbmd0aCA/IE1hdGgubWF4KC4uLm8uY2xvc2luZ1ZhcikgOiBudWxsOwogICAgby5oYXNQTyA9ICEhby5jdXN0UG9ObzsKICAgIGxpc3QucHVz"
    "aChvKTsKICB9CiAgcmV0dXJuIHsgbGlzdCwgYnlLZXkgfTsKfQoKLyogLS0tLS0tLS0tLS0tLS0tLSAxMC4gUklTSyBFTkdJTkUgLS0tLS0tLS0tLS0tLS0t"
    "LSAqLwpmdW5jdGlvbiBjb21wdXRlUmlza3Mocm93cykgewogIGNvbnN0IFIgPSBidWlsZFJvbGx1cChyb3dzKTsKICBjb25zdCB0b2RheSA9IHRvZGF5SVNP"
    "KCk7CiAgY29uc3Qgcmlza3MgPSBbXTsKICBsZXQgcmlkID0gMDsKICBjb25zdCBiYW5kID0gcyA9PiBzID49IDc1ID8gIkNyaXRpY2FsIiA6IHMgPj0gNTAg"
    "PyAiSGlnaCIgOiBzID49IDI1ID8gIk1vZGVyYXRlIiA6ICJMb3ciOwogIGZvciAoY29uc3QgbyBvZiBSLmxpc3QpIHsKICAgIGxldCBzY29yZSA9IDA7IGNv"
    "bnN0IHJlYXNvbnMgPSBbXTsgbGV0IGV4cG9zdXJlID0gMDsgbGV0IGRheXMgPSBudWxsOwogICAgaWYgKG8ud29uICYmICFvLmhhc1BPKSB7IHNjb3JlICs9"
    "IDE1OyByZWFzb25zLnB1c2goIldvbiBSRlEgd2l0aG91dCBjdXN0b21lciBQTyIpOyBleHBvc3VyZSArPSBvLnJldjsgfQogICAgaWYgKG8uZ3AgPCAwKSB7"
    "IHNjb3JlICs9IDI1OyByZWFzb25zLnB1c2goIk5lZ2F0aXZlIGdyb3NzIHByb2ZpdCIpOyBleHBvc3VyZSArPSBNYXRoLmFicyhvLmdwKTsgfQogICAgZWxz"
    "ZSBpZiAoby5tYXJnaW4gIT0gbnVsbCAmJiBvLm1hcmdpbiA+IDAgJiYgby5tYXJnaW4gPCBMT1dfTUFSR0lOKSB7IHNjb3JlICs9IDE1OyByZWFzb25zLnB1"
    "c2goIlZlcnkgbG93IG1hcmdpbiAoPCIgKyBMT1dfTUFSR0lOICsgIiUpIik7IH0KICAgIGlmIChvLm1hcmdpbiAhPSBudWxsICYmIG8ubWFyZ2luID4gSElH"
    "SF9NQVJHSU4pIHsgc2NvcmUgKz0gODsgcmVhc29ucy5wdXNoKCJVbnVzdWFsbHkgaGlnaCBtYXJnaW4g4oCUIHZlcmlmeSIpOyB9CiAgICBpZiAoIW8uY3Vz"
    "dG9tZXIpIHsgc2NvcmUgKz0gMTA7IHJlYXNvbnMucHVzaCgiTWlzc2luZyBjdXN0b21lciBuYW1lIik7IH0KICAgIGlmIChvLnF1b3RlZCAmJiBvLnJldiA9"
    "PT0gMCkgeyBzY29yZSArPSAxMjsgcmVhc29ucy5wdXNoKCJNaXNzaW5nIEVUIHF1b3RlZCB2YWx1ZSIpOyB9CiAgICBpZiAoby5yZXYgPiAwICYmIG8uY29n"
    "cyA9PT0gMCAmJiBvLndvbikgeyBzY29yZSArPSAxMDsgcmVhc29ucy5wdXNoKCJNaXNzaW5nIHN1cHBsaWVyIGNvc3Qgb24gd29uIFJGUSIpOyB9CiAgICBp"
    "ZiAoby5jbG9zaW5nVmFyRGF5ICE9IG51bGwgJiYgby5jbG9zaW5nVmFyRGF5ID4gMCAmJiAhby53b24pIHsgc2NvcmUgKz0gODsgcmVhc29ucy5wdXNoKCJR"
    "dW90ZWQgYWZ0ZXIgY2xvc2luZyBkYXRlIik7IH0KICAgIGlmIChvLnJlc3VsdCA9PT0gIk9wZW4iICYmIG8uY3VzdFJmcURhdGUpIHsKICAgICAgY29uc3Qg"
    "YWdlID0gZGF5c0JldHdlZW4odG9kYXksIG8uY3VzdFJmcURhdGUpOwogICAgICBpZiAoYWdlICE9IG51bGwgJiYgYWdlID4gMzApIHsgc2NvcmUgKz0gMTI7"
    "IHJlYXNvbnMucHVzaCgiT3BlbiBSRlEgYWdlaW5nID4zMCBkYXlzIik7IGRheXMgPSBhZ2U7IH0KICAgIH0KICAgIGlmIChvLndvbiAmJiBvLmdwID4gMCAm"
    "JiBvLnJldiA+IDAgJiYgby5oYXNQTyAmJiBNYXRoLmFicygoby5yZXYgLSBvLmNvZ3MpIC0gby5ncCkgPiBNYXRoLm1heCgxLCBvLnJldiAqIDAuMDIpKSB7"
    "CiAgICAgIHNjb3JlICs9IDg7IHJlYXNvbnMucHVzaCgiR1AgZG9lcyBub3QgcmVjb25jaWxlIHRvIFJldiDiiJIgQ29zdCIpOwogICAgfQogICAgaWYgKCFy"
    "ZWFzb25zLmxlbmd0aCkgY29udGludWU7CiAgICBzY29yZSA9IE1hdGgubWluKDEwMCwgc2NvcmUpOwogICAgcmlza3MucHVzaCh7CiAgICAgIGlkOiAiUiIg"
    "KyAoKytyaWQpLnRvU3RyaW5nKCkucGFkU3RhcnQoNCwgIjAiKSwKICAgICAgY2F0ZWdvcnk6IG8uZ3AgPCAwIHx8IChvLm1hcmdpbiAhPSBudWxsICYmIG8u"
    "bWFyZ2luIDwgTE9XX01BUkdJTikgPyAiRmluYW5jaWFsIgogICAgICAgIDogby53b24gJiYgIW8uaGFzUE8gPyAiQ29tbWVyY2lhbCIgOiAhby5jdXN0b21l"
    "ciB8fCBvLnJldiA9PT0gMCA/ICJEYXRhIFF1YWxpdHkiIDogIlJGUSBFeGVjdXRpb24iLAogICAgICBzZXZlcml0eTogYmFuZChzY29yZSksIHNjb3JlLAog"
    "ICAgICBjdXN0b21lcjogby5jdXN0b21lciB8fCAi4oCUIiwgcmZxOiBvLmtleS5zcGxpdCgifHwiKVsxXSB8fCAi4oCUIiwKICAgICAgcG86IG8uY3VzdFBv"
    "Tm8gfHwgIuKAlCIsIGV0UE9DOiBvLmV0UE9DIHx8ICLigJQiLCBzdXBwbGllcjogby5zdXBwbGllciB8fCAi4oCUIiwKICAgICAgZGVzY3JpcHRpb246IHJl"
    "YXNvbnMuam9pbigiOyAiKSwKICAgICAgZGF5c092ZXJkdWU6IGRheXMsCiAgICAgIGV4cG9zdXJlOiBNYXRoLnJvdW5kKGV4cG9zdXJlKSwKICAgICAgYWN0"
    "aW9uOiByZWNvbW1lbmRBY3Rpb24ocmVhc29ucyksCiAgICAgIF9yZXY6IG8ucmV2LAogICAgfSk7CiAgfQogIHJpc2tzLnNvcnQoKGEsIGIpID0+IGIuc2Nv"
    "cmUgLSBhLnNjb3JlIHx8IGIuZXhwb3N1cmUgLSBhLmV4cG9zdXJlKTsKICByZXR1cm4gcmlza3M7Cn0KZnVuY3Rpb24gcmVjb21tZW5kQWN0aW9uKHJlYXNv"
    "bnMpIHsKICBpZiAocmVhc29ucy5zb21lKHIgPT4gci5pbmNsdWRlcygiV29uIFJGUSB3aXRob3V0IikpKSByZXR1cm4gIk9idGFpbiBjdXN0b21lciBQTyBh"
    "Z2FpbnN0IHdvbiBSRlEuIjsKICBpZiAocmVhc29ucy5zb21lKHIgPT4gci5pbmNsdWRlcygiTmVnYXRpdmUiKSkpIHJldHVybiAiVmFsaWRhdGUgbmVnYXRp"
    "dmUgZ3Jvc3MgcHJvZml0IGJlZm9yZSBwcm9jZXNzaW5nLiI7CiAgaWYgKHJlYXNvbnMuc29tZShyID0+IHIuaW5jbHVkZXMoImxvdyBtYXJnaW4iKSkpIHJl"
    "dHVybiAiUmV2aWV3IHByaWNpbmcgLyBjb3N0IGJlZm9yZSBjb21taXRtZW50LiI7CiAgaWYgKHJlYXNvbnMuc29tZShyID0+IHIuaW5jbHVkZXMoImNsb3Np"
    "bmcgZGF0ZSIpKSkgcmV0dXJuICJUaWdodGVuIFJGUSB0dXJuYXJvdW5kIHRvIG1lZXQgZGVhZGxpbmVzLiI7CiAgaWYgKHJlYXNvbnMuc29tZShyID0+IHIu"
    "aW5jbHVkZXMoImFnZWluZyIpKSkgcmV0dXJuICJGb2xsb3cgdXAgb3IgY2xvc2Ugc3RhbGUgb3BlbiBSRlEuIjsKICBpZiAocmVhc29ucy5zb21lKHIgPT4g"
    "ci5pbmNsdWRlcygicmVjb25jaWxlIikpKSByZXR1cm4gIlJlY29uY2lsZSBHUCBhZ2FpbnN0IFJldiDiiJIgQ29zdC4iOwogIGlmIChyZWFzb25zLnNvbWUo"
    "ciA9PiByLmluY2x1ZGVzKCJNaXNzaW5nIikpKSByZXR1cm4gIkNvbXBsZXRlIG1pc3NpbmcgbWFzdGVyIGRhdGEgZmllbGRzLiI7CiAgcmV0dXJuICJSZXZp"
    "ZXcgcmVjb3JkIGFuZCBjb25maXJtIHN0YXR1cy4iOwp9CgovKiAtLS0tLS0tLS0tLS0tLS0tIDExLiBJTlNJR0hUIEVOR0lORSAtLS0tLS0tLS0tLS0tLS0t"
    "ICovCmZ1bmN0aW9uIGdlbmVyYXRlSW5zaWdodHMocm93cykgewogIGNvbnN0IG91dCA9IFtdOwogIGNvbnN0IFIgPSBidWlsZFJvbGx1cChyb3dzKTsKICBj"
    "b25zdCBsaXN0ID0gUi5saXN0OwogIGNvbnN0IHdvbiA9IGxpc3QuZmlsdGVyKG8gPT4gby53b24pLCBsb3N0ID0gbGlzdC5maWx0ZXIobyA9PiBvLmxvc3Qp"
    "OwogIGNvbnN0IGZpbmFsaXplZCA9IHdvbi5sZW5ndGggKyBsb3N0Lmxlbmd0aDsKICBjb25zdCB3aW5SYXRlID0gc2FmZURpdih3b24ubGVuZ3RoLCBmaW5h"
    "bGl6ZWQpICogMTAwOwogIGNvbnN0IHJldiA9IGFnZ3JlZ2F0ZURlZHVwKHJvd3MsICJyZnFLZXkiLCAiZXRRdW90ZWRWYWx1ZSIpOwogIGNvbnN0IGdwID0g"
    "YWdncmVnYXRlRGVkdXAocm93cywgInJmcUtleSIsICJncm9zc1Byb2ZpdENhbGMiKTsKICBjb25zdCBtYXJnaW4gPSBzYWZlRGl2KGdwLCByZXYpICogMTAw"
    "OwogIGNvbnN0IHBvUm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5jdXN0UG9Obyk7CiAgY29uc3QgcG9WYWwgPSBhZ2dyZWdhdGVEZWR1cChwb1Jvd3MsICJj"
    "dXN0UG9LZXkiLCAiZXRRdW90ZWRWYWx1ZSIpOwoKICBpZiAoZmluYWxpemVkKSBvdXQucHVzaChbInBvcyIsICJXaW4gcmF0ZSIsIGBSRlEgd2luIHJhdGUg"
    "aXMgJHtmbXRQY3Qod2luUmF0ZSl9IGFjcm9zcyAke2ZpbmFsaXplZH0gZmluYWxpemVkIFJGUXMgKCR7d29uLmxlbmd0aH0gd29uIC8gJHtsb3N0Lmxlbmd0"
    "aH0gbG9zdCkuYF0pOwogIC8vIGN1c3RvbWVyIGNvbmNlbnRyYXRpb24KICBjb25zdCBjdXN0UE8gPSBuZXcgTWFwKCk7CiAgZ3JvdXBEZWR1cChwb1Jvd3Ms"
    "ICJjdXN0UG9LZXkiLCAiZXRRdW90ZWRWYWx1ZSIpLmZvckVhY2goKHYsIGspID0+IHsKICAgIGNvbnN0IGN1c3QgPSBwb1Jvd3MuZmluZChyID0+IHIuY3Vz"
    "dFBvS2V5ID09PSBrKT8uY3VzdG9tZXIgfHwgIj8iOwogICAgY3VzdFBPLnNldChjdXN0LCAoY3VzdFBPLmdldChjdXN0KSB8fCAwKSArIHYpOwogIH0pOwog"
    "IGNvbnN0IHRvdFBPID0gQXJyYXkuZnJvbShjdXN0UE8udmFsdWVzKCkpLnJlZHVjZSgoYSwgYikgPT4gYSArIGIsIDApOwogIGNvbnN0IHRvcDMgPSB0b3BO"
    "KGN1c3RQTywgMyk7CiAgY29uc3QgdG9wM3NoYXJlID0gc2FmZURpdih0b3AzLnJlZHVjZSgoYSwgWywgdl0pID0+IGEgKyB2LCAwKSwgdG90UE8pICogMTAw"
    "OwogIGlmICh0b3RQTyAmJiB0b3Azc2hhcmUgPiA1NSkgb3V0LnB1c2goWyJ3YXJuIiwgIkNvbmNlbnRyYXRpb24iLCBgVG9wIDMgY3VzdG9tZXJzIGNvbnRy"
    "aWJ1dGUgJHtmbXRQY3QodG9wM3NoYXJlKX0gb2YgY3VzdG9tZXIgUE8gdmFsdWUg4oCUIGNvbmNlbnRyYXRpb24gcmlzay5gXSk7CiAgaWYgKHRvcDMubGVu"
    "Z3RoKSBvdXQucHVzaChbIm9wcCIsICJUb3AgYWNjb3VudCIsIGAke2VzYyh0b3AzWzBdWzBdKX0gaXMgdGhlIGxhcmdlc3QgUE8gYWNjb3VudCBhdCAke2Zt"
    "dENvbXBhY3QodG9wM1swXVsxXSl9LmBdKTsKICAvLyB3b24gd2l0aG91dCBQTwogIGNvbnN0IHdvbk5vUG8gPSB3b24uZmlsdGVyKG8gPT4gIW8uaGFzUE8p"
    "OwogIGlmICh3b25Ob1BvLmxlbmd0aCkgb3V0LnB1c2goWyJjcml0IiwgIlJldmVudWUgYXQgcmlzayIsIGAke3dvbk5vUG8ubGVuZ3RofSB3b24gUkZRcyB3"
    "b3J0aCAke2ZtdENvbXBhY3Qod29uTm9Qby5yZWR1Y2UoKGEsIG8pID0+IGEgKyBvLnJldiwgMCkpfSBoYXZlIG5vIGN1c3RvbWVyIFBPIGNhcHR1cmVkLmBd"
    "KTsKICAvLyBuZWdhdGl2ZSBHUAogIGNvbnN0IG5lZ0dwID0gbGlzdC5maWx0ZXIobyA9PiBvLmdwIDwgMCk7CiAgaWYgKG5lZ0dwLmxlbmd0aCkgb3V0LnB1"
    "c2goWyJjcml0IiwgIk5lZ2F0aXZlIG1hcmdpbiIsIGAke25lZ0dwLmxlbmd0aH0gUkZRcyBzaG93IG5lZ2F0aXZlIGdyb3NzIHByb2ZpdCAoJHtmbXRDb21w"
    "YWN0KG5lZ0dwLnJlZHVjZSgoYSwgbykgPT4gYSArIG8uZ3AsIDApKX0pLmBdKTsKICAvLyByZXNwb25zZSB0aW1lCiAgY29uc3QgcnQgPSBsaXN0Lm1hcChv"
    "ID0+IG8ucmVzcERheSkuZmlsdGVyKHggPT4geCAhPSBudWxsKTsKICBpZiAocnQubGVuZ3RoKSBvdXQucHVzaChbInBvcyIsICJSZXNwb25zaXZlbmVzcyIs"
    "IGBNZWRpYW4gUkZRIHJlc3BvbnNlIHRpbWUgaXMgJHtmbXREYXlzKG1lZGlhbihydCkpfSAoYXZnICR7Zm10RGF5cyhydC5yZWR1Y2UoKGEsIGIpID0+IGEg"
    "KyBiLCAwKSAvIHJ0Lmxlbmd0aCl9KS5gXSk7CiAgLy8gbWFyZ2luIGhlYWx0aAogIGlmIChyZXYpIG91dC5wdXNoKFttYXJnaW4gPj0gMTUgPyAicG9zIiA6"
    "ICJ3YXJuIiwgIk1hcmdpbiIsIGBCbGVuZGVkIGdyb3NzIG1hcmdpbiBpcyAke2ZtdFBjdChtYXJnaW4pfSBvbiAke2ZtdENvbXBhY3QocmV2KX0gcXVvdGVk"
    "IHZhbHVlLmBdKTsKICAvLyBQT0MgbGVhZGVyCiAgY29uc3QgcG9jR3AgPSBuZXcgTWFwKCk7CiAgbGlzdC5mb3JFYWNoKG8gPT4geyBpZiAoby5ldFBPQykg"
    "cG9jR3Auc2V0KG8uZXRQT0MsIChwb2NHcC5nZXQoby5ldFBPQykgfHwgMCkgKyBvLmdwKTsgfSk7CiAgY29uc3QgdG9wUG9jID0gdG9wTihwb2NHcCwgMSlb"
    "MF07CiAgaWYgKHRvcFBvYykgb3V0LnB1c2goWyJwb3MiLCAiVG9wIHBlcmZvcm1lciIsIGAke2VzYyh0b3BQb2NbMF0pfSBsZWFkcyBvbiBncm9zcyBwcm9m"
    "aXQgY29udHJpYnV0aW9uICgke2ZtdENvbXBhY3QodG9wUG9jWzFdKX0pLmBdKTsKICAvLyBzdXBwbGllciBjb25jZW50cmF0aW9uCiAgY29uc3Qgc3VwU3Bl"
    "bmQgPSBuZXcgTWFwKCk7CiAgZ3JvdXBEZWR1cChyb3dzLmZpbHRlcihyID0+IHIuc3VwcGxpZXJOYW1lICYmIHIud29uICE9PSBmYWxzZSksICJyZnFLZXki"
    "LCAic3VwcGxpZXJUb3RhbFByaWNlIik7CiAgcm93cy5mb3JFYWNoKHIgPT4geyBpZiAoci5zdXBwbGllck5hbWUgJiYgci5zdXBwbGllclRvdGFsUHJpY2Up"
    "IHN1cFNwZW5kLnNldChyLnN1cHBsaWVyTmFtZSwgKHN1cFNwZW5kLmdldChyLnN1cHBsaWVyTmFtZSkgfHwgMCkgKyByLnN1cHBsaWVyVG90YWxQcmljZSk7"
    "IH0pOwogIGNvbnN0IHN1cFRvcCA9IHRvcE4oc3VwU3BlbmQsIDEpWzBdOwogIGNvbnN0IHN1cFRvdCA9IEFycmF5LmZyb20oc3VwU3BlbmQudmFsdWVzKCkp"
    "LnJlZHVjZSgoYSwgYikgPT4gYSArIGIsIDApOwogIGlmIChzdXBUb3AgJiYgc2FmZURpdihzdXBUb3BbMV0sIHN1cFRvdCkgPiAwLjQpIG91dC5wdXNoKFsi"
    "d2FybiIsICJTdXBwbGllciByZWxpYW5jZSIsIGAke2VzYyhzdXBUb3BbMF0pfSBhY2NvdW50cyBmb3IgJHtmbXRQY3Qoc2FmZURpdihzdXBUb3BbMV0sIHN1"
    "cFRvdCkgKiAxMDApfSBvZiBzdXBwbGllciBjb3N0IOKAlCBzaW5nbGUtc291cmNlIGV4cG9zdXJlLmBdKTsKICAvLyBkYXRhIHF1YWxpdHkKICBjb25zdCBt"
    "aXNzQ3VzdCA9IHJvd3MuZmlsdGVyKHIgPT4gIXIuY3VzdG9tZXIpLmxlbmd0aDsKICBpZiAobWlzc0N1c3QpIG91dC5wdXNoKFsiZGF0YSIsICJEYXRhIHF1"
    "YWxpdHkiLCBgJHttaXNzQ3VzdH0gbGluZSBpdGVtcyBhcmUgbWlzc2luZyBhIGN1c3RvbWVyIG5hbWUuYF0pOwogIC8vIGNvbnZlcnNpb24KICBjb25zdCBx"
    "dW90ZWQgPSBsaXN0LmZpbHRlcihvID0+IG8ucXVvdGVkKS5sZW5ndGg7CiAgY29uc3QgY29udiA9IHNhZmVEaXYodW5pcXVlQ291bnQocG9Sb3dzLCAiY3Vz"
    "dFBvS2V5IiksIHF1b3RlZCkgKiAxMDA7CiAgaWYgKHF1b3RlZCkgb3V0LnB1c2goW2NvbnYgPj0gMTUgPyAib3BwIiA6ICJ3YXJuIiwgIkNvbnZlcnNpb24i"
    "LCBgUXVvdGUtdG8tUE8gY29udmVyc2lvbiBpcyAke2ZtdFBjdChjb252KX0gKCR7dW5pcXVlQ291bnQocG9Sb3dzLCAiY3VzdFBvS2V5Iil9IFBPcyBmcm9t"
    "ICR7cXVvdGVkfSBxdW90ZWQgUkZRcykuYF0pOwogIHJldHVybiBvdXQ7Cn0KZnVuY3Rpb24gaW5zaWdodEhUTUwobGlzdCkgewogIGlmICghbGlzdC5sZW5n"
    "dGgpIHJldHVybiBgPGRpdiBjbGFzcz0iZW1wdHktc3RhdGUiPk5vIGluc2lnaHRzIGZvciB0aGUgY3VycmVudCBzZWxlY3Rpb24uPC9kaXY+YDsKICBjb25z"
    "dCBuYW1lcyA9IHsgcG9zOiAiUG9zaXRpdmUiLCB3YXJuOiAiV2FybmluZyIsIGNyaXQ6ICJDcml0aWNhbCIsIG9wcDogIk9wcG9ydHVuaXR5IiwgZGF0YTog"
    "IkRhdGEiIH07CiAgcmV0dXJuIGxpc3QubWFwKChbY2xzLCB0YWcsIHR4dF0pID0+CiAgICBgPGRpdiBjbGFzcz0iaW5zaWdodCAke2Nsc30iPjxzcGFuIGNs"
    "YXNzPSJ0YWciPiR7bmFtZXNbY2xzXSB8fCB0YWd9PC9zcGFuPjxkaXY+JHt0eHR9PC9kaXY+PC9kaXY+YCkuam9pbigiIik7Cn0KCi8qIC0tLS0tLS0tLS0t"
    "LS0tLS0gMTIuIFRBQiBSRU5ERVJFUlMgLS0tLS0tLS0tLS0tLS0tLSAqLwoKLyogLS0tLSBzaGFyZWQgbWV0cmljIGJsb2NrIC0tLS0gKi8KZnVuY3Rpb24g"
    "Y29yZU1ldHJpY3Mocm93cykgewogIGNvbnN0IFIgPSBidWlsZFJvbGx1cChyb3dzKTsKICBjb25zdCBsaXN0ID0gUi5saXN0OwogIGNvbnN0IHdvbiA9IGxp"
    "c3QuZmlsdGVyKG8gPT4gby53b24pLCBsb3N0ID0gbGlzdC5maWx0ZXIobyA9PiBvLmxvc3QpLCBkZWNsaW5lZCA9IGxpc3QuZmlsdGVyKG8gPT4gby5kZWNs"
    "aW5lZCAmJiAhby53b24gJiYgIW8ubG9zdCk7CiAgY29uc3QgcXVvdGVkID0gbGlzdC5maWx0ZXIobyA9PiBvLnF1b3RlZCk7CiAgY29uc3Qgb3BlbiA9IGxp"
    "c3QuZmlsdGVyKG8gPT4gby5yZXN1bHQgPT09ICJPcGVuIik7CiAgY29uc3QgZmluYWxpemVkID0gd29uLmxlbmd0aCArIGxvc3QubGVuZ3RoOwogIGNvbnN0"
    "IHBvUm93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5jdXN0UG9Obyk7CiAgY29uc3QgcmV2ID0gYWdncmVnYXRlRGVkdXAocm93cywgInJmcUtleSIsICJldFF1"
    "b3RlZFZhbHVlIik7CiAgY29uc3QgY29ncyA9IGFnZ3JlZ2F0ZURlZHVwKHJvd3MsICJyZnFLZXkiLCAic3VwcGxpZXJUb3RhbFByaWNlIik7CiAgY29uc3Qg"
    "Z3AgPSBhZ2dyZWdhdGVEZWR1cChyb3dzLCAicmZxS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpOwogIGNvbnN0IHdvblZhbCA9IGFnZ3JlZ2F0ZURlZHVwKHdv"
    "bi5sZW5ndGggPyByb3dzLmZpbHRlcihyID0+IHIuZXRRdW90ZVN0YXR1cyA9PT0gIldvbiIpIDogW10sICJyZnFLZXkiLCAiZXRRdW90ZWRWYWx1ZSIpOwog"
    "IGNvbnN0IHBvVmFsID0gYWdncmVnYXRlRGVkdXAocG9Sb3dzLCAiY3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBuUE8gPSB1bmlxdWVD"
    "b3VudChwb1Jvd3MsICJjdXN0UG9LZXkiKTsKICByZXR1cm4gewogICAgUiwgbGlzdCwgd29uLCBsb3N0LCBkZWNsaW5lZCwgcXVvdGVkLCBvcGVuLCBmaW5h"
    "bGl6ZWQsIHBvUm93cywgcmV2LCBjb2dzLCBncCwgd29uVmFsLCBwb1ZhbCwgblBPLAogICAgd2luUmF0ZTogc2FmZURpdih3b24ubGVuZ3RoLCBmaW5hbGl6"
    "ZWQpICogMTAwLAogICAgY29udjogc2FmZURpdihuUE8sIHF1b3RlZC5sZW5ndGgpICogMTAwLAogICAgbWFyZ2luOiBzYWZlRGl2KGdwLCByZXYpICogMTAw"
    "LAogICAgYXZnUE86IHNhZmVEaXYocG9WYWwsIG5QTyksCiAgICBhdmdSRlE6IHNhZmVEaXYocmV2LCBsaXN0Lmxlbmd0aCksCiAgICBhdmdHcFBPOiBzYWZl"
    "RGl2KGFnZ3JlZ2F0ZURlZHVwKHBvUm93cywgImN1c3RQb0tleSIsICJncm9zc1Byb2ZpdENhbGMiKSwgblBPKSwKICAgIHJlc3BBdmc6IGF2Z0ZpZWxkKHJv"
    "d3MsICJyZnFSZXNwb25zZURheXMiKSwKICAgIHN1cFJlc3BBdmc6IGF2Z0ZpZWxkKHJvd3MsICJzdXBwbGllclJlc3BvbnNlRGF5cyIpLAogICAgcG9PYUF2"
    "ZzogYXZnRmllbGQocm93cywgInBvVG9PYURheXMiKSwKICB9Owp9CgpmdW5jdGlvbiByZW5kZXJPdmVydmlldyhyb3dzKSB7CiAgY29uc3QgbSA9IGNvcmVN"
    "ZXRyaWNzKHJvd3MpOwogIGNvbnN0IGVsID0gJCgiI3RhYi1vdmVydmlldyIpOwogIGNvbnN0IG1hcmdpbkNscyA9IG0ubWFyZ2luID49IDE1ID8gImdvb2Qi"
    "IDogbS5tYXJnaW4gPj0gOCA/ICJ3YXJuIiA6ICJiYWQiOwogIGVsLmlubmVySFRNTCA9IGAKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkV4ZWN1"
    "dGl2ZSBLUElzPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCI+CiAgICAgICR7a3BpKCJVbmlxdWUgUkZRcyIsIGZtdE51bShtLmxpc3QubGVuZ3Ro"
    "KSwgZm10TnVtKHJvd3MubGVuZ3RoKSArICIgbGluZSBpdGVtcyIpfQogICAgICAke2twaSgiUXVvdGVkIFJGUXMiLCBmbXROdW0obS5xdW90ZWQubGVuZ3Ro"
    "KSl9CiAgICAgICR7a3BpKCJXb24gUkZRcyIsIGZtdE51bShtLndvbi5sZW5ndGgpLCAiIiwgImdvb2QiKX0KICAgICAgJHtrcGkoIkxvc3QgUkZRcyIsIGZt"
    "dE51bShtLmxvc3QubGVuZ3RoKSwgIiIsICJiYWQiKX0KICAgICAgJHtrcGkoIk9wZW4gUkZRcyIsIGZtdE51bShtLm9wZW4ubGVuZ3RoKSl9CiAgICAgICR7"
    "a3BpKCJSRlEgV2luIFJhdGUiLCBmbXRQY3QobS53aW5SYXRlKSwgIndvbiDDtyAod29uK2xvc3QpIiwgbS53aW5SYXRlID49IDQwID8gImdvb2QiIDogbS53"
    "aW5SYXRlID49IDIwID8gIndhcm4iIDogImJhZCIpfQogICAgICAke2twaSgiUXVvdGXihpJQTyBDb252ZXJzaW9uIiwgZm10UGN0KG0uY29udikpfQogICAg"
    "ICAke2twaSgiRVQgUXVvdGVkIFZhbHVlIiwgZm10Q29tcGFjdChtLnJldiksIGZtdEN1cihtLnJldikpfQogICAgICAke2twaSgiV29uIC8gUE8gVmFsdWUi"
    "LCBmbXRDb21wYWN0KG0ud29uVmFsKSl9CiAgICAgICR7a3BpKCJTdXBwbGllciBDb3N0IiwgZm10Q29tcGFjdChtLmNvZ3MpKX0KICAgICAgJHtrcGkoIkdy"
    "b3NzIFByb2ZpdCIsIGZtdENvbXBhY3QobS5ncCksIGZtdEN1cihtLmdwKSwgbS5ncCA+PSAwID8gImdvb2QiIDogImJhZCIpfQogICAgICAke2twaSgiR3Jv"
    "c3MgTWFyZ2luICUiLCBmbXRQY3QobS5tYXJnaW4pLCAiIiwgbWFyZ2luQ2xzKX0KICAgICAgJHtrcGkoIkN1c3RvbWVyIFBPcyIsIGZtdE51bShtLm5QTykp"
    "fQogICAgICAke2twaSgiQXZnIEN1c3RvbWVyIFBPIFZhbHVlIiwgZm10Q29tcGFjdChtLmF2Z1BPKSl9CiAgICAgICR7a3BpKCJBdmcgUkZRIFZhbHVlIiwg"
    "Zm10Q29tcGFjdChtLmF2Z1JGUSkpfQogICAgICAke2twaSgiQXZnIEdQIC8gUE8iLCBmbXRDb21wYWN0KG0uYXZnR3BQTykpfQogICAgICAke2twaSgiQXZn"
    "IFJGUSBSZXNwb25zZSIsIGZtdERheXMobS5yZXNwQXZnKSl9CiAgICAgICR7a3BpKCJBdmcgU3VwcGxpZXIgUXVvdGUgUmVzcC4iLCBmbXREYXlzKG0uc3Vw"
    "UmVzcEF2ZykpfQogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+TWFuYWdlbWVudCBJbnNpZ2h0czwvZGl2PgogICAgPGRpdiBp"
    "ZD0ib3ZJbnNpZ2h0cyI+PC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+VHJlbmRzICYgRGlzdHJpYnV0aW9uPC9kaXY+CiAgICA8ZGl2"
    "IGNsYXNzPSJjaGFydC1ncmlkIj4KICAgICAgJHtjYXJkU2hlbGwoIk1vbnRobHkgUkZRcyAoYnkgRVQgcXVvdGUgbW9udGgpIiwgIm92TW9udGhseVJGUSIs"
    "ICJjNiIpfQogICAgICAke2NhcmRTaGVsbCgiTW9udGhseSBFVCBRdW90ZWQgVmFsdWUgJiBHcm9zcyBQcm9maXQiLCAib3ZNb250aGx5VmFsIiwgImM2Iil9"
    "CiAgICAgICR7Y2FyZFNoZWxsKCJSRlEgUmVzdWx0IiwgIm92UmVzdWx0IiwgImM0IiwgImNsaWNrIHRvIGZpbHRlciIpfQogICAgICAke2NhcmRTaGVsbCgi"
    "RVQgUXVvdGUgU3RhdHVzIiwgIm92UVN0YXR1cyIsICJjNCIsICJjbGljayB0byBmaWx0ZXIiKX0KICAgICAgJHtjYXJkU2hlbGwoIldvbiB2cyBMb3N0IFJG"
    "UXMiLCAib3ZXaW5Mb3NzIiwgImM0Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJSRlEg4oaSIENhc2ggRnVubmVsIiwgIm92RnVubmVsIiwgImM2Iil9CiAgICAg"
    "ICR7Y2FyZFNoZWxsKCJHcm9zcyBNYXJnaW4gJSBUcmVuZCIsICJvdk1hcmdpblRyZW5kIiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJUb3AgMTAgQ3Vz"
    "dG9tZXJzIGJ5IFF1b3RlZCBWYWx1ZSIsICJvdlRvcEN1c3RRIiwgImM2IiwgImNsaWNrIHRvIGZpbHRlciIpfQogICAgICAke2NhcmRTaGVsbCgiVG9wIDEw"
    "IEN1c3RvbWVycyBieSBQTyBWYWx1ZSIsICJvdlRvcEN1c3RQTyIsICJjNiIsICJjbGljayB0byBmaWx0ZXIiKX0KICAgICAgJHtjYXJkU2hlbGwoIlRvcCAx"
    "MCBTdXBwbGllcnMgYnkgUHJvY3VyZW1lbnQgVmFsdWUiLCAib3ZUb3BTdXAiLCAiYzYiLCAiY2xpY2sgdG8gZmlsdGVyIil9CiAgICAgICR7Y2FyZFNoZWxs"
    "KCJUb3AgRVQgUE9DcyBieSBHcm9zcyBQcm9maXQiLCAib3ZUb3BQb2MiLCAiYzYiLCAiY2xpY2sgdG8gZmlsdGVyIil9CiAgICAgICR7Y2FyZFNoZWxsKCJP"
    "cGVuIFJGUSBBZ2VpbmcgKGRheXMgc2luY2UgUkZRIGRhdGUpIiwgIm92QWdlaW5nIiwgImM2Iil9CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0"
    "aW9uLXRpdGxlIj5NYW5hZ2VtZW50IEFsZXJ0czwvZGl2PgogICAgPGRpdiBjbGFzcz0iYWxlcnRzLWdyaWQiIGlkPSJvdkFsZXJ0cyI+PC9kaXY+YDsKCiAg"
    "JCgiI292SW5zaWdodHMiKS5pbm5lckhUTUwgPSBpbnNpZ2h0SFRNTChnZW5lcmF0ZUluc2lnaHRzKHJvd3MpKTsKCiAgLy8gbW9udGhseSBzZXJpZXMKICBj"
    "b25zdCBtb250aHMgPSBtb250aGx5U2VyaWVzKHJvd3MpOwogIHJlbmRlckNoYXJ0KCJvdk1vbnRobHlSRlEiLCB7CiAgICB0eXBlOiAiYmFyIiwKICAgIGRh"
    "dGE6IHsgbGFiZWxzOiBtb250aHMubGFiZWxzLCBkYXRhc2V0czogW3sgbGFiZWw6ICJVbmlxdWUgUkZRcyIsIGRhdGE6IG1vbnRocy5yZnFDb3VudCwgYmFj"
    "a2dyb3VuZENvbG9yOiBwYWxldHRlKCkucHJpbWFyeSwgYm9yZGVyUmFkaXVzOiAzIH1dIH0sCiAgICBvcHRpb25zOiB7IHNjYWxlczogeyB5OiB7IGJlZ2lu"
    "QXRaZXJvOiB0cnVlIH0gfSwgcGx1Z2luczogeyBsZWdlbmQ6IHsgZGlzcGxheTogZmFsc2UgfSB9IH0sCiAgfSk7CiAgcmVuZGVyQ2hhcnQoIm92TW9udGhs"
    "eVZhbCIsIHsKICAgIHR5cGU6ICJiYXIiLAogICAgZGF0YTogeyBsYWJlbHM6IG1vbnRocy5sYWJlbHMsIGRhdGFzZXRzOiBbCiAgICAgIHsgbGFiZWw6ICJR"
    "dW90ZWQgVmFsdWUiLCBkYXRhOiBtb250aHMucmV2LCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5wcmltYXJ5LCBib3JkZXJSYWRpdXM6IDMsIHlBeGlz"
    "SUQ6ICJ5IiB9LAogICAgICB7IGxhYmVsOiAiR3Jvc3MgUHJvZml0IiwgdHlwZTogImxpbmUiLCBkYXRhOiBtb250aHMuZ3AsIGJvcmRlckNvbG9yOiBwYWxl"
    "dHRlKCkuZ29vZCwgYmFja2dyb3VuZENvbG9yOiBwYWxldHRlKCkuZ29vZCwgdGVuc2lvbjogLjMsIHlBeGlzSUQ6ICJ5IiB9LAogICAgXSB9LAogICAgb3B0"
    "aW9uczogeyBzY2FsZXM6IHsgeTogeyBiZWdpbkF0WmVybzogdHJ1ZSwgdGlja3M6IHsgY2FsbGJhY2s6IHYgPT4gZm10Q29tcGFjdCh2KSB9IH0gfSB9LAog"
    "IH0pOwogIGRvbnV0KCJvdlJlc3VsdCIsIGNvdW50QnkobS5saXN0LCBvID0+IG8ucmVzdWx0KSwgKGxibCkgPT4gY2hhcnRGaWx0ZXIoIl9fcmVzdWx0Iiwg"
    "bGJsKSwgcmVzdWx0Q29sb3JzKTsKICBkb251dCgib3ZRU3RhdHVzIiwgY291bnRCeVJvd3Mocm93cywgImV0UXVvdGVTdGF0dXMiKSwgKGxibCkgPT4gY2hh"
    "cnRGaWx0ZXIoImV0UXVvdGVTdGF0dXMiLCBsYmwpKTsKICByZW5kZXJDaGFydCgib3ZXaW5Mb3NzIiwgewogICAgdHlwZTogImJhciIsIGRhdGE6IHsgbGFi"
    "ZWxzOiBbIldvbiIsICJMb3N0IiwgIkRlY2xpbmVkIiwgIk9wZW4iXSwKICAgICAgZGF0YXNldHM6IFt7IGRhdGE6IFttLndvbi5sZW5ndGgsIG0ubG9zdC5s"
    "ZW5ndGgsIG0uZGVjbGluZWQubGVuZ3RoLCBtLm9wZW4ubGVuZ3RoXSwKICAgICAgICBiYWNrZ3JvdW5kQ29sb3I6IFtwYWxldHRlKCkuZ29vZCwgcGFsZXR0"
    "ZSgpLmJhZCwgcGFsZXR0ZSgpLndhcm4sIHBhbGV0dGUoKS5tdXRlZF0sIGJvcmRlclJhZGl1czogMyB9XSB9LAogICAgb3B0aW9uczogeyBwbHVnaW5zOiB7"
    "IGxlZ2VuZDogeyBkaXNwbGF5OiBmYWxzZSB9IH0sIHNjYWxlczogeyB5OiB7IGJlZ2luQXRaZXJvOiB0cnVlIH0gfSB9LAogIH0pOwogIGZ1bm5lbENoYXJ0"
    "KCJvdkZ1bm5lbCIsIHJvd3MsIG0pOwogIHJlbmRlckNoYXJ0KCJvdk1hcmdpblRyZW5kIiwgewogICAgdHlwZTogImxpbmUiLCBkYXRhOiB7IGxhYmVsczog"
    "bW9udGhzLmxhYmVscywgZGF0YXNldHM6IFt7IGxhYmVsOiAiR3Jvc3MgTWFyZ2luICUiLCBkYXRhOiBtb250aHMubWFyZ2luLAogICAgICBib3JkZXJDb2xv"
    "cjogcGFsZXR0ZSgpLmFjY2VudCwgYmFja2dyb3VuZENvbG9yOiAidHJhbnNwYXJlbnQiLCB0ZW5zaW9uOiAuMywgc3BhbkdhcHM6IHRydWUgfV0gfSwKICAg"
    "IG9wdGlvbnM6IHsgc2NhbGVzOiB7IHk6IHsgdGlja3M6IHsgY2FsbGJhY2s6IHYgPT4gdiArICIlIiB9IH0gfSwgcGx1Z2luczogeyBsZWdlbmQ6IHsgZGlz"
    "cGxheTogZmFsc2UgfSB9IH0sCiAgfSk7CiAgaGJhcigib3ZUb3BDdXN0USIsIHRvcE4oZGVkdXBNYXBCeShyb3dzLCAicmZxS2V5IiwgImV0UXVvdGVkVmFs"
    "dWUiLCAiY3VzdG9tZXIiKSwgMTApLCBmbXRDb21wYWN0LCAobGJsKSA9PiBzZXRNdWx0aSgiY3VzdG9tZXIiLCBsYmwpKTsKICBoYmFyKCJvdlRvcEN1c3RQ"
    "TyIsIHRvcE4oZGVkdXBNYXBCeShtLnBvUm93cywgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIiwgImN1c3RvbWVyIiksIDEwKSwgZm10Q29tcGFjdCwg"
    "KGxibCkgPT4gc2V0TXVsdGkoImN1c3RvbWVyIiwgbGJsKSk7CiAgaGJhcigib3ZUb3BTdXAiLCB0b3BOKHN1bU1hcEJ5KHJvd3MsICJzdXBwbGllck5hbWUi"
    "LCAic3VwcGxpZXJUb3RhbFByaWNlIiksIDEwKSwgZm10Q29tcGFjdCwgKGxibCkgPT4gc2V0TXVsdGkoInN1cHBsaWVyTmFtZSIsIGxibCkpOwogIGhiYXIo"
    "Im92VG9wUG9jIiwgdG9wTihkZWR1cE1hcEJ5KHJvd3MsICJyZnFLZXkiLCAiZ3Jvc3NQcm9maXRDYWxjIiwgImV0UE9DIiksIDEwKSwgZm10Q29tcGFjdCwg"
    "KGxibCkgPT4gc2V0TXVsdGkoImV0UE9DIiwgbGJsKSk7CiAgY29uc3QgYWdlID0gYWdlaW5nQnVja2V0cyhtLm9wZW4pOwogIHJlbmRlckNoYXJ0KCJvdkFn"
    "ZWluZyIsIHsgdHlwZTogImJhciIsIGRhdGE6IHsgbGFiZWxzOiBhZ2UubGFiZWxzLCBkYXRhc2V0czogW3sgZGF0YTogYWdlLmRhdGEsIGJhY2tncm91bmRD"
    "b2xvcjogcGFsZXR0ZSgpLndhcm4sIGJvcmRlclJhZGl1czogMyB9XSB9LAogICAgb3B0aW9uczogeyBwbHVnaW5zOiB7IGxlZ2VuZDogeyBkaXNwbGF5OiBm"
    "YWxzZSB9IH0sIHNjYWxlczogeyB5OiB7IGJlZ2luQXRaZXJvOiB0cnVlIH0gfSB9IH0pOwoKICByZW5kZXJBbGVydHMoJCgiI292QWxlcnRzIiksIHJvd3Ms"
    "IG0pOwp9CgpmdW5jdGlvbiBtb250aGx5U2VyaWVzKHJvd3MpIHsKICBjb25zdCBtYXAgPSBuZXcgTWFwKCk7CiAgY29uc3QgcmV2TSA9IGdyb3VwRGVkdXAo"
    "cm93cywgInJmcUtleSIsICJldFF1b3RlZFZhbHVlIik7CiAgY29uc3QgZ3BNID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5IiwgImdyb3NzUHJvZml0Q2Fs"
    "YyIpOwogIC8vIHBlciBSRlEsIGFzc2lnbiB0byBpdHMgcXVvdGUgbW9udGgKICBjb25zdCBzZWVuUmZxID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgciBv"
    "ZiByb3dzKSB7CiAgICBjb25zdCBtayA9IG1vbnRoS2V5KHIuZXRRdW90ZURhdGUpOyBpZiAoIW1rKSBjb250aW51ZTsKICAgIGlmICghbWFwLmhhcyhtaykp"
    "IG1hcC5zZXQobWssIHsgcmZxOiBuZXcgU2V0KCksIHJldjogMCwgZ3A6IDAgfSk7CiAgICBjb25zdCBiID0gbWFwLmdldChtayk7IGIucmZxLmFkZChyLnJm"
    "cUtleSk7CiAgfQogIC8vIHJldmVudWUvZ3Agb25jZSBwZXIgUkZRLCBpbiB0aGUgUkZRJ3MgZWFybGllc3QgcXVvdGUgbW9udGgKICBjb25zdCByZnFNb250"
    "aCA9IG5ldyBNYXAoKTsKICBmb3IgKGNvbnN0IHIgb2Ygcm93cykgeyBjb25zdCBtayA9IG1vbnRoS2V5KHIuZXRRdW90ZURhdGUpOyBpZiAobWsgJiYgKCFy"
    "ZnFNb250aC5oYXMoci5yZnFLZXkpIHx8IG1rIDwgcmZxTW9udGguZ2V0KHIucmZxS2V5KSkpIHJmcU1vbnRoLnNldChyLnJmcUtleSwgbWspOyB9CiAgZm9y"
    "IChjb25zdCBbaywgbWtdIG9mIHJmcU1vbnRoKSB7IGlmICghbWFwLmhhcyhtaykpIG1hcC5zZXQobWssIHsgcmZxOiBuZXcgU2V0KCksIHJldjogMCwgZ3A6"
    "IDAgfSk7IGNvbnN0IGIgPSBtYXAuZ2V0KG1rKTsgYi5yZXYgKz0gcmV2TS5nZXQoaykgfHwgMDsgYi5ncCArPSBncE0uZ2V0KGspIHx8IDA7IH0KICBjb25z"
    "dCBsYWJlbHMgPSBBcnJheS5mcm9tKG1hcC5rZXlzKCkpLnNvcnQoKTsKICByZXR1cm4gewogICAgbGFiZWxzOiBsYWJlbHMubWFwKGwgPT4gTU9OVEhTWyts"
    "LnNsaWNlKDUpIC0gMV0gKyAiICIgKyBsLnNsaWNlKDIsIDQpKSwKICAgIHJmcUNvdW50OiBsYWJlbHMubWFwKGwgPT4gbWFwLmdldChsKS5yZnEuc2l6ZSks"
    "CiAgICByZXY6IGxhYmVscy5tYXAobCA9PiBtYXAuZ2V0KGwpLnJldiksCiAgICBncDogbGFiZWxzLm1hcChsID0+IG1hcC5nZXQobCkuZ3ApLAogICAgbWFy"
    "Z2luOiBsYWJlbHMubWFwKGwgPT4geyBjb25zdCBiID0gbWFwLmdldChsKTsgcmV0dXJuIGIucmV2ID8gKyhiLmdwIC8gYi5yZXYgKiAxMDApLnRvRml4ZWQo"
    "MSkgOiBudWxsOyB9KSwKICAgIGtleXM6IGxhYmVscywKICB9Owp9CmZ1bmN0aW9uIGNvdW50QnkobGlzdCwgZm4pIHsgY29uc3QgbSA9IG5ldyBNYXAoKTsg"
    "bGlzdC5mb3JFYWNoKG8gPT4geyBjb25zdCBrID0gZm4obykgfHwgIuKAlCI7IG0uc2V0KGssIChtLmdldChrKSB8fCAwKSArIDEpOyB9KTsgcmV0dXJuIG07"
    "IH0KZnVuY3Rpb24gY291bnRCeVJvd3Mocm93cywgZmllbGQpIHsgY29uc3QgbSA9IG5ldyBNYXAoKTsgcm93cy5mb3JFYWNoKHIgPT4geyBjb25zdCBrID0g"
    "cltmaWVsZF0gfHwgIuKAlCI7IG0uc2V0KGssIChtLmdldChrKSB8fCAwKSArIDEpOyB9KTsgcmV0dXJuIG07IH0KZnVuY3Rpb24gZGVkdXBNYXBCeShyb3dz"
    "LCBrZXlGaWVsZCwgdmFsRmllbGQsIGdyb3VwRmllbGQpIHsKICBjb25zdCBnID0gZ3JvdXBEZWR1cChyb3dzLCBrZXlGaWVsZCwgdmFsRmllbGQpOwogIGNv"
    "bnN0IG93bmVyID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgciBvZiByb3dzKSBpZiAoIW93bmVyLmhhcyhyW2tleUZpZWxkXSkpIG93bmVyLnNldChyW2tl"
    "eUZpZWxkXSwgcltncm91cEZpZWxkXSB8fCAi4oCUIik7CiAgY29uc3Qgb3V0ID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgW2ssIHZdIG9mIGcpIHsgY29u"
    "c3QgZ3JwID0gb3duZXIuZ2V0KGspIHx8ICLigJQiOyBvdXQuc2V0KGdycCwgKG91dC5nZXQoZ3JwKSB8fCAwKSArIHYpOyB9CiAgcmV0dXJuIG91dDsKfQpm"
    "dW5jdGlvbiBzdW1NYXBCeShyb3dzLCBncm91cEZpZWxkLCB2YWxGaWVsZCkgewogIGNvbnN0IG0gPSBuZXcgTWFwKCk7IHJvd3MuZm9yRWFjaChyID0+IHsg"
    "aWYgKHJbdmFsRmllbGRdICE9IG51bGwpIHsgY29uc3QgayA9IHJbZ3JvdXBGaWVsZF0gfHwgIuKAlCI7IG0uc2V0KGssIChtLmdldChrKSB8fCAwKSArIHJb"
    "dmFsRmllbGRdKTsgfSB9KTsgcmV0dXJuIG07Cn0KY29uc3QgcmVzdWx0Q29sb3JzID0geyBXb246ICItLWdvb2QiLCBMb3N0OiAiLS1iYWQiLCBEZWNsaW5l"
    "ZDogIi0td2FybiIsIE9wZW46ICItLW11dGVkIiwgUGVuZGluZzogIi0taW5mbyIgfTsKZnVuY3Rpb24gZG9udXQoaWQsIG1hcCwgb25DbGljaywgY29sb3JN"
    "YXApIHsKICBjb25zdCBsYWJlbHMgPSBBcnJheS5mcm9tKG1hcC5rZXlzKCkpLCBkYXRhID0gQXJyYXkuZnJvbShtYXAudmFsdWVzKCkpOwogIGNvbnN0IHAg"
    "PSBwYWxldHRlKCk7CiAgY29uc3QgY29sb3JzID0gbGFiZWxzLm1hcCgobCwgaSkgPT4gY29sb3JNYXAgJiYgY29sb3JNYXBbbF0gPyBjc3NWYXIoY29sb3JN"
    "YXBbbF0pIDogcC5zZXJpZXNbaSAlIHAuc2VyaWVzLmxlbmd0aF0pOwogIHJlbmRlckNoYXJ0KGlkLCB7CiAgICB0eXBlOiAiZG91Z2hudXQiLAogICAgZGF0"
    "YTogeyBsYWJlbHMsIGRhdGFzZXRzOiBbeyBkYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IGNvbG9ycywgYm9yZGVyQ29sb3I6IHAuc3VyZmFjZSwgYm9yZGVyV2lk"
    "dGg6IDIgfV0gfSwKICAgIG9wdGlvbnM6IHsgY3V0b3V0OiAiNTglIiwgcGx1Z2luczogeyBsZWdlbmQ6IHsgcG9zaXRpb246ICJyaWdodCIgfSB9LAogICAg"
    "ICBvbkNsaWNrOiAoZSwgZWxzKSA9PiB7IGlmIChvbkNsaWNrICYmIGVscy5sZW5ndGgpIG9uQ2xpY2sobGFiZWxzW2Vsc1swXS5pbmRleF0pOyB9IH0sCiAg"
    "fSk7Cn0KZnVuY3Rpb24gaGJhcihpZCwgZW50cmllcywgZm10LCBvbkNsaWNrKSB7CiAgY29uc3QgbGFiZWxzID0gZW50cmllcy5tYXAoZSA9PiBlWzBdKSwg"
    "ZGF0YSA9IGVudHJpZXMubWFwKGUgPT4gZVsxXSk7CiAgY29uc3QgcCA9IHBhbGV0dGUoKTsKICByZW5kZXJDaGFydChpZCwgewogICAgdHlwZTogImJhciIs"
    "CiAgICBkYXRhOiB7IGxhYmVscywgZGF0YXNldHM6IFt7IGRhdGEsIGJhY2tncm91bmRDb2xvcjogcC5wcmltYXJ5LCBib3JkZXJSYWRpdXM6IDMgfV0gfSwK"
    "ICAgIG9wdGlvbnM6IHsgaW5kZXhBeGlzOiAieSIsIHBsdWdpbnM6IHsgbGVnZW5kOiB7IGRpc3BsYXk6IGZhbHNlIH0sCiAgICAgIHRvb2x0aXA6IHsgY2Fs"
    "bGJhY2tzOiB7IGxhYmVsOiBjID0+IGZtdCA/IGZtdChjLnBhcnNlZC54KSA6IGMucGFyc2VkLnggfSB9IH0sCiAgICAgIHNjYWxlczogeyB4OiB7IGJlZ2lu"
    "QXRaZXJvOiB0cnVlLCB0aWNrczogeyBjYWxsYmFjazogdiA9PiBmbXQgPyBmbXQodikgOiB2IH0gfSB9LAogICAgICBvbkNsaWNrOiAoZSwgZWxzKSA9PiB7"
    "IGlmIChvbkNsaWNrICYmIGVscy5sZW5ndGgpIG9uQ2xpY2sobGFiZWxzW2Vsc1swXS5pbmRleF0pOyB9IH0sCiAgfSk7Cn0KZnVuY3Rpb24gZnVubmVsQ2hh"
    "cnQoaWQsIHJvd3MsIG0pIHsKICBjb25zdCBzdGFnZXMgPSBbCiAgICBbIlJGUXMgcmVjZWl2ZWQiLCBtLmxpc3QubGVuZ3RoXSwKICAgIFsiUXVvdGVkIiwg"
    "bS5xdW90ZWQubGVuZ3RoXSwKICAgIFsiU3VwcGxpZXIgcXVvdGVzIiwgdW5pcXVlQ291bnQocm93cy5maWx0ZXIociA9PiByLnN1cHBsaWVyUXVvdGVEYXRl"
    "KSwgInJmcUtleSIpXSwKICAgIFsiV29uIiwgbS53b24ubGVuZ3RoXSwKICAgIFsiQ3VzdG9tZXIgUE9zIiwgbS5uUE9dLAogIF07CiAgcmVuZGVyQ2hhcnQo"
    "aWQsIHsKICAgIHR5cGU6ICJiYXIiLAogICAgZGF0YTogeyBsYWJlbHM6IHN0YWdlcy5tYXAocyA9PiBzWzBdKSwgZGF0YXNldHM6IFt7IGRhdGE6IHN0YWdl"
    "cy5tYXAocyA9PiBzWzFdKSwgYmFja2dyb3VuZENvbG9yOiBwYWxldHRlKCkuc2VyaWVzLCBib3JkZXJSYWRpdXM6IDMgfV0gfSwKICAgIG9wdGlvbnM6IHsg"
    "aW5kZXhBeGlzOiAieSIsIHBsdWdpbnM6IHsgbGVnZW5kOiB7IGRpc3BsYXk6IGZhbHNlIH0gfSwgc2NhbGVzOiB7IHg6IHsgYmVnaW5BdFplcm86IHRydWUg"
    "fSB9IH0sCiAgfSk7Cn0KZnVuY3Rpb24gYWdlaW5nQnVja2V0cyhvcGVuTGlzdCkgewogIGNvbnN0IHRvZGF5ID0gdG9kYXlJU08oKTsKICBjb25zdCBiID0g"
    "eyAiMOKAkzIiOiAwLCAiM+KAkzUiOiAwLCAiNuKAkzEwIjogMCwgIjEx4oCTMjAiOiAwLCAiMjHigJMzMCI6IDAsICI+MzAiOiAwIH07CiAgb3Blbkxpc3Qu"
    "Zm9yRWFjaChvID0+IHsKICAgIGNvbnN0IGQgPSBkYXlzQmV0d2Vlbih0b2RheSwgby5jdXN0UmZxRGF0ZSk7CiAgICBpZiAoZCA9PSBudWxsKSByZXR1cm47"
    "CiAgICBpZiAoZCA8PSAyKSBiWyIw4oCTMiJdKys7IGVsc2UgaWYgKGQgPD0gNSkgYlsiM+KAkzUiXSsrOyBlbHNlIGlmIChkIDw9IDEwKSBiWyI24oCTMTAi"
    "XSsrOwogICAgZWxzZSBpZiAoZCA8PSAyMCkgYlsiMTHigJMyMCJdKys7IGVsc2UgaWYgKGQgPD0gMzApIGJbIjIx4oCTMzAiXSsrOyBlbHNlIGJbIj4zMCJd"
    "Kys7CiAgfSk7CiAgcmV0dXJuIHsgbGFiZWxzOiBPYmplY3Qua2V5cyhiKSwgZGF0YTogT2JqZWN0LnZhbHVlcyhiKSB9Owp9CmZ1bmN0aW9uIHNldE11bHRp"
    "KGtleSwgdmFsdWUpIHsKICBjb25zdCBhcnIgPSBGSUxURVJTLm11bHRpW2tleV07CiAgY29uc3QgaSA9IGFyci5pbmRleE9mKHZhbHVlKTsKICBpZiAoaSA+"
    "IC0xKSBhcnIuc3BsaWNlKGksIDEpOyBlbHNlIGFyci5wdXNoKHZhbHVlKTsKICBidWlsZFNsaWNlcnMoKTsgb25GaWx0ZXJDaGFuZ2UoKTsKfQoKZnVuY3Rp"
    "b24gcmVuZGVyQWxlcnRzKGNvbnRhaW5lciwgcm93cywgbSkgewogIGNvbnN0IHRvZGF5ID0gdG9kYXlJU08oKTsKICBjb25zdCBhbGVydHMgPSBbXTsKICBj"
    "b25zdCB3b25Ob1BvID0gbS53b24uZmlsdGVyKG8gPT4gIW8uaGFzUE8pOwogIGlmICh3b25Ob1BvLmxlbmd0aCkgYWxlcnRzLnB1c2goWyJjcml0IiwgIldv"
    "biBSRlFzIHdpdGhvdXQgY3VzdG9tZXIgUE8iLCB3b25Ob1BvLmxlbmd0aCwgd29uTm9Qby5yZWR1Y2UoKGEsIG8pID0+IGEgKyBvLnJldiwgMCksICgpID0+"
    "IGNoYXJ0RmlsdGVyKCJfX3Jlc3VsdCIsICJXb24iKV0pOwogIGNvbnN0IG5lZ0dwID0gbS5saXN0LmZpbHRlcihvID0+IG8uZ3AgPCAwKTsKICBpZiAobmVn"
    "R3AubGVuZ3RoKSBhbGVydHMucHVzaChbImNyaXQiLCAiTmVnYXRpdmUgZ3Jvc3MgcHJvZml0IiwgbmVnR3AubGVuZ3RoLCBuZWdHcC5yZWR1Y2UoKGEsIG8p"
    "ID0+IGEgKyBNYXRoLmFicyhvLmdwKSwgMCksIG51bGxdKTsKICBjb25zdCBsb3dNID0gbS5saXN0LmZpbHRlcihvID0+IG8ubWFyZ2luICE9IG51bGwgJiYg"
    "by5tYXJnaW4gPiAwICYmIG8ubWFyZ2luIDwgTE9XX01BUkdJTik7CiAgaWYgKGxvd00ubGVuZ3RoKSBhbGVydHMucHVzaChbImhpZ2giLCAiVW51c3VhbGx5"
    "IGxvdyBtYXJnaW4gKDwiICsgTE9XX01BUkdJTiArICIlKSIsIGxvd00ubGVuZ3RoLCBsb3dNLnJlZHVjZSgoYSwgbykgPT4gYSArIG8ucmV2LCAwKSwgbnVs"
    "bF0pOwogIGNvbnN0IGhpZ2hNID0gbS5saXN0LmZpbHRlcihvID0+IG8ubWFyZ2luICE9IG51bGwgJiYgby5tYXJnaW4gPiBISUdIX01BUkdJTik7CiAgaWYg"
    "KGhpZ2hNLmxlbmd0aCkgYWxlcnRzLnB1c2goWyJtZWQiLCAiVW51c3VhbGx5IGhpZ2ggbWFyZ2luICh2ZXJpZnkpIiwgaGlnaE0ubGVuZ3RoLCAwLCBudWxs"
    "XSk7CiAgY29uc3Qgc3RhbGUgPSBtLm9wZW4uZmlsdGVyKG8gPT4geyBjb25zdCBkID0gZGF5c0JldHdlZW4odG9kYXksIG8uY3VzdFJmcURhdGUpOyByZXR1"
    "cm4gZCAhPSBudWxsICYmIGQgPiAzMDsgfSk7CiAgaWYgKHN0YWxlLmxlbmd0aCkgYWxlcnRzLnB1c2goWyJoaWdoIiwgIk9wZW4gUkZRcyBhZ2VpbmcgYmV5"
    "b25kIDMwIGRheXMiLCBzdGFsZS5sZW5ndGgsIHN0YWxlLnJlZHVjZSgoYSwgbykgPT4gYSArIG8ucmV2LCAwKSwgbnVsbF0pOwogIGNvbnN0IG92ZXJkdWVR"
    "dW90ZSA9IG0ubGlzdC5maWx0ZXIobyA9PiBvLmNsb3NpbmdWYXJEYXkgIT0gbnVsbCAmJiBvLmNsb3NpbmdWYXJEYXkgPiAwICYmICFvLndvbik7CiAgaWYg"
    "KG92ZXJkdWVRdW90ZS5sZW5ndGgpIGFsZXJ0cy5wdXNoKFsibWVkIiwgIlJGUXMgcXVvdGVkIGFmdGVyIGNsb3NpbmcgZGF0ZSIsIG92ZXJkdWVRdW90ZS5s"
    "ZW5ndGgsIDAsIG51bGxdKTsKICBjb25zdCBtaXNzQ3VzdCA9IHVuaXF1ZUNvdW50KHJvd3MuZmlsdGVyKHIgPT4gIXIuY3VzdG9tZXIpLCAicmZxS2V5Iik7"
    "CiAgaWYgKG1pc3NDdXN0KSBhbGVydHMucHVzaChbIm1lZCIsICJSRlFzIG1pc3NpbmcgY3VzdG9tZXIgbmFtZSIsIG1pc3NDdXN0LCAwLCBudWxsXSk7CiAg"
    "Y29uc3Qgbm9RdW90ZVZhbCA9IG0ucXVvdGVkLmZpbHRlcihvID0+IG8ucmV2ID09PSAwKTsKICBpZiAobm9RdW90ZVZhbC5sZW5ndGgpIGFsZXJ0cy5wdXNo"
    "KFsibWVkIiwgIlF1b3RlZCBSRlFzIG1pc3NpbmcgcXVvdGVkIHZhbHVlIiwgbm9RdW90ZVZhbC5sZW5ndGgsIDAsIG51bGxdKTsKICAvLyBkdXBsaWNhdGUg"
    "UE9zIChzYW1lIGN1c3RQb0tleSBhY3Jvc3MgZGlmZmVyZW50IGN1c3RvbWVycyBpcyBpbXBvc3NpYmxlOyBkZXRlY3QgUE8gbnVtYmVyIHJldXNlZCBhY3Jv"
    "c3MgY3VzdG9tZXJzKQogIGNvbnN0IHBvQnlOdW0gPSBuZXcgTWFwKCk7CiAgcm93cy5mb3JFYWNoKHIgPT4geyBpZiAoci5jdXN0UG9ObykgeyBpZiAoIXBv"
    "QnlOdW0uaGFzKHIuY3VzdFBvTm8pKSBwb0J5TnVtLnNldChyLmN1c3RQb05vLCBuZXcgU2V0KCkpOyBwb0J5TnVtLmdldChyLmN1c3RQb05vKS5hZGQoci5j"
    "dXN0b21lcik7IH0gfSk7CiAgY29uc3QgZHVwUG8gPSBBcnJheS5mcm9tKHBvQnlOdW0udmFsdWVzKCkpLmZpbHRlcihzID0+IHMuc2l6ZSA+IDEpLmxlbmd0"
    "aDsKICBpZiAoZHVwUG8pIGFsZXJ0cy5wdXNoKFsibWVkIiwgIkN1c3RvbWVyIFBPIG51bWJlciB1c2VkIGJ5IG11bHRpcGxlIGN1c3RvbWVycyIsIGR1cFBv"
    "LCAwLCBudWxsXSk7CgogIGlmICghYWxlcnRzLmxlbmd0aCkgeyBjb250YWluZXIuaW5uZXJIVE1MID0gYDxkaXYgY2xhc3M9ImVtcHR5LXN0YXRlIj5ObyBh"
    "Y3RpdmUgYWxlcnRzIGZvciB0aGUgY3VycmVudCBzZWxlY3Rpb24uPC9kaXY+YDsgcmV0dXJuOyB9CiAgY29udGFpbmVyLmlubmVySFRNTCA9IGFsZXJ0cy5t"
    "YXAoKFtzZXYsIHRpdGxlLCBjb3VudCwgZXhwLCBmbl0pID0+IGAKICAgIDxkaXYgY2xhc3M9ImFsZXJ0ICR7c2V2fSIgJHtmbiA/ICdkYXRhLWNsaWNrPSIx"
    "IicgOiAiIn0+CiAgICAgIDxkaXYgY2xhc3M9InNldiI+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImFsZXJ0LWJvZHkiPgogICAgICAgIDxkaXYgY2xhc3M9"
    "ImFsZXJ0LXRpdGxlIj4ke2VzYyh0aXRsZSl9PC9kaXY+CiAgICAgICAgPGRpdiBjbGFzcz0iYWxlcnQtbWV0YSI+PHNwYW4+UmVjb3JkczogPGI+JHtmbXRO"
    "dW0oY291bnQpfTwvYj48L3NwYW4+CiAgICAgICAgJHtleHAgPyBgPHNwYW4+RXhwb3N1cmU6IDxiPiR7Zm10Q29tcGFjdChleHApfTwvYj48L3NwYW4+YCA6"
    "ICIifQogICAgICAgIDxzcGFuIGNsYXNzPSJwaWxsICR7c2V2ID09PSAnY3JpdCcgPyAncicgOiBzZXYgPT09ICdoaWdoJyA/ICdhJyA6ICdiJ30iPiR7c2V2"
    "LnRvVXBwZXJDYXNlKCl9PC9zcGFuPjwvZGl2PgogICAgICA8L2Rpdj48L2Rpdj5gKS5qb2luKCIiKTsKICBBcnJheS5mcm9tKGNvbnRhaW5lci5jaGlsZHJl"
    "bikuZm9yRWFjaCgoYywgaSkgPT4geyBjb25zdCBmbiA9IGFsZXJ0c1tpXVs0XTsgaWYgKGZuKSBjLm9uY2xpY2sgPSBmbjsgfSk7Cn0KCi8qID09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0KICAgUEFSVCAzIOKAlCBzaGFyZWQgcmVuZGVyIGhlbHBl"
    "cnMgKyBSRlEgLyBTdXBwbGllciAvIFBPIC8gUE9DIHRhYnMKICAgPT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PSAqLwoKLyogLS0tLSBzaGFyZWQgc21hbGwgaGVscGVycyAtLS0tICovCmZ1bmN0aW9uIGVtcHR5Q2FyZCh0aXRsZSwgc3BhbiwgbXNn"
    "KSB7CiAgcmV0dXJuIGA8ZGl2IGNsYXNzPSJjYXJkICR7c3Bhbn0iPgogICAgPGRpdiBjbGFzcz0iY2FyZC1oZWFkIj48ZGl2IGNsYXNzPSJjYXJkLXRpdGxl"
    "Ij4ke2VzYyh0aXRsZSl9PC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjaGFydC1ob2xkZXIiPjxkaXYgY2xhc3M9ImVtcHR5LXN0YXRlIj4ke2VzYyht"
    "c2cpfTwvZGl2PjwvZGl2PjwvZGl2PmA7Cn0KZnVuY3Rpb24gYmFyQ2hhcnQoaWQsIGVudHJpZXMsIGZtdCwgb25DbGljaywgY29sb3IpIHsKICBjb25zdCBs"
    "YWJlbHMgPSBlbnRyaWVzLm1hcChlID0+IGVbMF0pLCBkYXRhID0gZW50cmllcy5tYXAoZSA9PiBlWzFdKTsKICBjb25zdCBwID0gcGFsZXR0ZSgpOwogIHJl"
    "bmRlckNoYXJ0KGlkLCB7CiAgICB0eXBlOiAiYmFyIiwKICAgIGRhdGE6IHsgbGFiZWxzLCBkYXRhc2V0czogW3sgZGF0YSwgYmFja2dyb3VuZENvbG9yOiBj"
    "b2xvciB8fCBwLnByaW1hcnksIGJvcmRlclJhZGl1czogMyB9XSB9LAogICAgb3B0aW9uczogewogICAgICBwbHVnaW5zOiB7IGxlZ2VuZDogeyBkaXNwbGF5"
    "OiBmYWxzZSB9LCB0b29sdGlwOiB7IGNhbGxiYWNrczogeyBsYWJlbDogYyA9PiBmbXQgPyBmbXQoYy5wYXJzZWQueSkgOiBjLnBhcnNlZC55IH0gfSB9LAog"
    "ICAgICBzY2FsZXM6IHsgeTogeyBiZWdpbkF0WmVybzogdHJ1ZSwgdGlja3M6IHsgY2FsbGJhY2s6IHYgPT4gZm10ID8gZm10KHYpIDogdiB9IH0gfSwKICAg"
    "ICAgb25DbGljazogKGUsIGVscykgPT4geyBpZiAob25DbGljayAmJiBlbHMubGVuZ3RoKSBvbkNsaWNrKGxhYmVsc1tlbHNbMF0uaW5kZXhdKTsgfSwKICAg"
    "IH0sCiAgfSk7Cn0KZnVuY3Rpb24gd2luUmF0ZUJ5KGxpc3QsIGZpZWxkLCBtaW5OID0gMykgewogIGNvbnN0IGcgPSBuZXcgTWFwKCk7CiAgbGlzdC5mb3JF"
    "YWNoKG8gPT4gewogICAgaWYgKCFvLndvbiAmJiAhby5sb3N0KSByZXR1cm47CiAgICBjb25zdCBrID0gb1tmaWVsZF0gfHwgIuKAlCI7CiAgICBjb25zdCBi"
    "ID0gZy5nZXQoaykgfHwgeyB3OiAwLCBmOiAwIH07CiAgICBpZiAoby53b24pIGIudysrOyBpZiAoby53b24gfHwgby5sb3N0KSBiLmYrKzsKICAgIGcuc2V0"
    "KGssIGIpOwogIH0pOwogIHJldHVybiBBcnJheS5mcm9tKGcuZW50cmllcygpKQogICAgLmZpbHRlcigoWywgYl0pID0+IGIuZiA+PSBtaW5OKQogICAgLm1h"
    "cCgoW2ssIGJdKSA9PiBbaywgKyhiLncgLyBiLmYgKiAxMDApLnRvRml4ZWQoMSldKQogICAgLnNvcnQoKGEsIGIpID0+IGJbMV0gLSBhWzFdKS5zbGljZSgw"
    "LCAxMik7Cn0KZnVuY3Rpb24gcmVzcERpc3RyaWJ1dGlvbihsaXN0KSB7CiAgY29uc3QgYiA9IHsgIjAgZCI6IDAsICIx4oCTMiBkIjogMCwgIjPigJM1IGQi"
    "OiAwLCAiNuKAkzEwIGQiOiAwLCAiMTHigJMyMCBkIjogMCwgIj4yMCBkIjogMCB9OwogIGxpc3QuZm9yRWFjaChvID0+IHsKICAgIGNvbnN0IGQgPSBvLnJl"
    "c3BEYXk7IGlmIChkID09IG51bGwpIHJldHVybjsKICAgIGlmIChkIDw9IDApIGJbIjAgZCJdKys7IGVsc2UgaWYgKGQgPD0gMikgYlsiMeKAkzIgZCJdKys7"
    "IGVsc2UgaWYgKGQgPD0gNSkgYlsiM+KAkzUgZCJdKys7CiAgICBlbHNlIGlmIChkIDw9IDEwKSBiWyI24oCTMTAgZCJdKys7IGVsc2UgaWYgKGQgPD0gMjAp"
    "IGJbIjEx4oCTMjAgZCJdKys7IGVsc2UgYlsiPjIwIGQiXSsrOwogIH0pOwogIHJldHVybiBPYmplY3QuZW50cmllcyhiKTsKfQpmdW5jdGlvbiBsb3N0UmVh"
    "c29uKG8sIHJvd3NCeUtleSkgewogIGlmICghby5sb3N0ICYmICFvLmRlY2xpbmVkKSByZXR1cm4gbnVsbDsKICBjb25zdCBycyA9IHJvd3NCeUtleS5nZXQo"
    "by5rZXkpIHx8IFtdOwogIGNvbnN0IGhhc1N1cHBsaWVyID0gcnMuc29tZShyID0+IHIuc3VwcGxpZXJOYW1lICYmIHIuc3VwcGxpZXJUb3RhbFByaWNlKTsK"
    "ICBpZiAoby5kZWNsaW5lZCAmJiAhby5xdW90ZWQpIHJldHVybiB7IHJlYXNvbjogIk5vIGJpZCAvIGRlY2xpbmVkIiwgdHlwZTogIkNvbmZpcm1lZCIgfTsK"
    "ICBpZiAoIWhhc1N1cHBsaWVyKSByZXR1cm4geyByZWFzb246ICJObyBzdXBwbGllciBxdW90YXRpb24iLCB0eXBlOiAiSW5mZXJyZWQiIH07CiAgaWYgKG8u"
    "Y2xvc2luZ1ZhckRheSAhPSBudWxsICYmIG8uY2xvc2luZ1ZhckRheSA+IDApIHJldHVybiB7IHJlYXNvbjogIkxhdGUgcXVvdGF0aW9uIiwgdHlwZTogIklu"
    "ZmVycmVkIiB9OwogIGlmIChvLmRlY2xpbmVkKSByZXR1cm4geyByZWFzb246ICJEZWNsaW5lZCBieSBFVCIsIHR5cGU6ICJDb25maXJtZWQiIH07CiAgcmV0"
    "dXJuIHsgcmVhc29uOiAiVW5rbm93biAocHJpY2UvdGVjaG5pY2FsL2NvbW1lcmNpYWwpIiwgdHlwZTogIlVua25vd24iIH07Cn0KZnVuY3Rpb24gcm93c0J5"
    "S2V5TWFwKHJvd3MpIHsKICBjb25zdCBtID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgciBvZiByb3dzKSB7IGNvbnN0IGEgPSBtLmdldChyLnJmcUtleSkg"
    "fHwgW107IGEucHVzaChyKTsgbS5zZXQoci5yZnFLZXksIGEpOyB9CiAgcmV0dXJuIG07Cn0KZnVuY3Rpb24gb3Blbk1vZGFsKHRpdGxlLCBodG1sQm9keSkg"
    "ewogICQoIiNtb2RhbFRpdGxlIikudGV4dENvbnRlbnQgPSB0aXRsZTsKICAkKCIjbW9kYWxCb2R5IikuaW5uZXJIVE1MID0gaHRtbEJvZHk7CiAgJCgiI21v"
    "ZGFsIikuaGlkZGVuID0gZmFsc2U7Cn0KZnVuY3Rpb24gcm9sbHVwTW9kYWxCb2R5KG8pIHsKICBjb25zdCBmID0gWwogICAgWyJDdXN0b21lciIsIG8uY3Vz"
    "dG9tZXJdLCBbIkN1c3RvbWVyIFJGUSBOby4iLCBvLmtleS5zcGxpdCgifHwiKVsxXV0sCiAgICBbIkVUIFBPQyIsIG8uZXRQT0NdLCBbIkN1c3RvbWVyIFBP"
    "QyIsIG8uY3VzdFBPQ10sCiAgICBbIlByb2R1Y3QgQ2F0ZWdvcnkiLCBvLnByb2R1Y3RDYXRlZ29yeV0sIFsiU2VjdG9yIiwgby5zZWN0b3JdLCBbIlNlbGVj"
    "dGVkIFN1cHBsaWVyIiwgby5zdXBwbGllcl0sCiAgICBbIkN1c3RvbWVyIFJGUSBEYXRlIiwgZm10RGF0ZShvLmN1c3RSZnFEYXRlKV0sIFsiUkZRIENsb3Np"
    "bmcgRGF0ZSIsIGZtdERhdGUoby5jbG9zaW5nRGF0ZSldLAogICAgWyJFVCBRdW90YXRpb24gRGF0ZSIsIGZtdERhdGUoby5ldFF1b3RlRGF0ZSldLCBbIlJl"
    "c3BvbnNlIChkYXlzKSIsIGZtdERheXMoby5yZXNwRGF5KV0sCiAgICBbIkNsb3NpbmcgdmFyaWFuY2UgKGRheXMpIiwgby5jbG9zaW5nVmFyRGF5ID09IG51"
    "bGwgPyAi4oCUIiA6IG8uY2xvc2luZ1ZhckRheSArICIgZCJdLAogICAgWyJMaW5lIGl0ZW1zIiwgZm10TnVtKG8ubGluZXMpXSwKICAgIFsiRVQgUXVvdGVk"
    "IFZhbHVlIiwgZm10Q3VyKG8ucmV2KV0sIFsiU3VwcGxpZXIgQ29zdCIsIGZtdEN1cihvLmNvZ3MpXSwKICAgIFsiR3Jvc3MgUHJvZml0IiwgZm10Q3VyKG8u"
    "Z3ApXSwgWyJHcm9zcyBNYXJnaW4gJSIsIG8ubWFyZ2luID09IG51bGwgPyAi4oCUIiA6IGZtdFBjdChvLm1hcmdpbildLAogICAgWyJSZXN1bHQiLCBvLnJl"
    "c3VsdF0sIFsiQ3VzdG9tZXIgUE8gTm8uIiwgby5jdXN0UG9ObyB8fCAi4oCUIl0sIFsiQ3VzdG9tZXIgUE8gRGF0ZSIsIGZtdERhdGUoby5jdXN0UG9EYXRl"
    "KV0sCiAgXTsKICByZXR1cm4gYDxkaXYgY2xhc3M9ImRldGFpbC1ncmlkIj4ke2YubWFwKChbaywgdl0pID0+CiAgICBgPGRpdiBjbGFzcz0iZHQtcm93Ij48"
    "c3BhbiBjbGFzcz0iZHQtayI+JHtlc2Moayl9PC9zcGFuPjxzcGFuIGNsYXNzPSJkdC12Ij4ke3YgPT0gbnVsbCA/ICLigJQiIDogZXNjKHYpfTwvc3Bhbj48"
    "L2Rpdj5gKS5qb2luKCIiKX08L2Rpdj5gOwp9CmZ1bmN0aW9uIGRpc3RpbmN0Qnkocm93cywgcHJlZCwgZmllbGQpIHsKICBjb25zdCBzID0gbmV3IFNldCgp"
    "OwogIHJvd3MuZm9yRWFjaChyID0+IHsgaWYgKHByZWQocikgJiYgcltmaWVsZF0pIHMuYWRkKHJbZmllbGRdKTsgfSk7CiAgcmV0dXJuIHMuc2l6ZTsKfQoK"
    "LyogPT09PT09PT09PT09PT09PT0gVEFCIDIg4oCUIFJGUSBUUkFDS0VSID09PT09PT09PT09PT09PT09ICovCmZ1bmN0aW9uIHJlbmRlclJGUShyb3dzKSB7"
    "CiAgY29uc3QgbSA9IGNvcmVNZXRyaWNzKHJvd3MpOwogIGNvbnN0IGxpc3QgPSBtLmxpc3Q7CiAgY29uc3QgcmJrID0gcm93c0J5S2V5TWFwKHJvd3MpOwog"
    "IGNvbnN0IHJlc3BWYWxzID0gbGlzdC5tYXAobyA9PiBvLnJlc3BEYXkpLmZpbHRlcih4ID0+IHggIT0gbnVsbCk7CiAgY29uc3QgYmVmb3JlRGwgPSBsaXN0"
    "LmZpbHRlcihvID0+IG8uY2xvc2luZ1ZhckRheSAhPSBudWxsICYmIG8uY2xvc2luZ1ZhckRheSA8PSAwKS5sZW5ndGg7CiAgY29uc3QgYWZ0ZXJEbCA9IGxp"
    "c3QuZmlsdGVyKG8gPT4gby5jbG9zaW5nVmFyRGF5ICE9IG51bGwgJiYgby5jbG9zaW5nVmFyRGF5ID4gMCkubGVuZ3RoOwogIGNvbnN0IGRsS25vd24gPSBi"
    "ZWZvcmVEbCArIGFmdGVyRGw7CiAgY29uc3QgY2xvc2luZ1Nvb24gPSBsaXN0LmZpbHRlcihvID0+IHsKICAgIGlmIChvLnJlc3VsdCAhPT0gIk9wZW4iIHx8"
    "ICFvLmNsb3NpbmdEYXRlKSByZXR1cm4gZmFsc2U7CiAgICBjb25zdCBkID0gZGF5c0JldHdlZW4oby5jbG9zaW5nRGF0ZSwgdG9kYXlJU08oKSk7IHJldHVy"
    "biBkICE9IG51bGwgJiYgZCA+PSAwICYmIGQgPD0gMzsKICB9KS5sZW5ndGg7CiAgY29uc3Qgb3ZlcmR1ZVVucXVvdGVkID0gbGlzdC5maWx0ZXIobyA9PiAh"
    "by5xdW90ZWQgJiYgby5jbG9zaW5nRGF0ZSAmJiBkYXlzQmV0d2Vlbih0b2RheUlTTygpLCBvLmNsb3NpbmdEYXRlKSA+IDApLmxlbmd0aDsKICBjb25zdCBl"
    "bCA9ICQoIiN0YWItcmZxIik7CiAgZWwuaW5uZXJIVE1MID0gYAogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+UkZRIEtQSXM8L2Rpdj4KICAgIDxk"
    "aXYgY2xhc3M9ImtwaS1ncmlkIj4KICAgICAgJHtrcGkoIlJGUXMgUmVjZWl2ZWQiLCBmbXROdW0obGlzdC5sZW5ndGgpKX0KICAgICAgJHtrcGkoIlF1b3Rl"
    "ZCIsIGZtdE51bShtLnF1b3RlZC5sZW5ndGgpKX0KICAgICAgJHtrcGkoIldvbiIsIGZtdE51bShtLndvbi5sZW5ndGgpLCAiIiwgImdvb2QiKX0KICAgICAg"
    "JHtrcGkoIkxvc3QiLCBmbXROdW0obS5sb3N0Lmxlbmd0aCksICIiLCAiYmFkIil9CiAgICAgICR7a3BpKCJEZWNsaW5lZCIsIGZtdE51bShtLmRlY2xpbmVk"
    "Lmxlbmd0aCkpfQogICAgICAke2twaSgiT3BlbiIsIGZtdE51bShtLm9wZW4ubGVuZ3RoKSl9CiAgICAgICR7a3BpKCJXaW4gUmF0ZSIsIGZtdFBjdChtLndp"
    "blJhdGUpLCAid29uIMO3ICh3b24rbG9zdCkiLCBtLndpblJhdGUgPj0gNDAgPyAiZ29vZCIgOiBtLndpblJhdGUgPj0gMjAgPyAid2FybiIgOiAiYmFkIil9"
    "CiAgICAgICR7a3BpKCJMb3NzIFJhdGUiLCBmbXRQY3QoMTAwIC0gbS53aW5SYXRlKSl9CiAgICAgICR7a3BpKCJRdW90ZeKGklBPIENvbnZlcnNpb24iLCBm"
    "bXRQY3QobS5jb252KSl9CiAgICAgICR7a3BpKCJBdmcgUkZRIFJlc3BvbnNlIiwgZm10RGF5cyhtLnJlc3BBdmcpKX0KICAgICAgJHtrcGkoIk1lZGlhbiBS"
    "RlEgUmVzcG9uc2UiLCBmbXREYXlzKG1lZGlhbihyZXNwVmFscykpKX0KICAgICAgJHtrcGkoIlF1b3RlZCBCZWZvcmUgRGVhZGxpbmUiLCBmbXROdW0oYmVm"
    "b3JlRGwpKX0KICAgICAgJHtrcGkoIlF1b3RlZCBBZnRlciBEZWFkbGluZSIsIGZtdE51bShhZnRlckRsKSwgIiIsIGFmdGVyRGwgPyAid2FybiIgOiAiIil9"
    "CiAgICAgICR7a3BpKCJEZWFkbGluZSBDb21wbGlhbmNlIiwgZGxLbm93biA/IGZtdFBjdChiZWZvcmVEbCAvIGRsS25vd24gKiAxMDApIDogIuKAlCIpfQog"
    "ICAgICAke2twaSgiQ2xvc2luZyDiiaQgMyBEYXlzIChvcGVuKSIsIGZtdE51bShjbG9zaW5nU29vbiksICIiLCBjbG9zaW5nU29vbiA/ICJ3YXJuIiA6ICIi"
    "KX0KICAgICAgJHtrcGkoIk92ZXJkdWUgVW5xdW90ZWQiLCBmbXROdW0ob3ZlcmR1ZVVucXVvdGVkKSwgIiIsIG92ZXJkdWVVbnF1b3RlZCA/ICJiYWQiIDog"
    "IiIpfQogICAgICAke2twaSgiQXZnIExpbmUgSXRlbXMgLyBSRlEiLCAobGlzdC5yZWR1Y2UoKGEsIG8pID0+IGEgKyBvLmxpbmVzLCAwKSAvIChsaXN0Lmxl"
    "bmd0aCB8fCAxKSkudG9GaXhlZCgxKSl9CiAgICAgICR7a3BpKCJBdmcgUXVvdGVkIFZhbHVlIC8gUkZRIiwgZm10Q29tcGFjdChtLmF2Z1JGUSkpfQogICAg"
    "PC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+QW5hbHlzaXM8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXJ0LWdyaWQiPgogICAgICAk"
    "e2NhcmRTaGVsbCgiUkZRIFJlc3VsdCIsICJyZnFSZXN1bHQiLCAiYzQiLCAiY2xpY2sgdG8gZmlsdGVyIil9CiAgICAgICR7Y2FyZFNoZWxsKCJXaW4gLyBM"
    "b3NzIC8gRGVjbGluZWQgLyBPcGVuIiwgInJmcVdMRCIsICJjNCIpfQogICAgICAke2NhcmRTaGVsbCgiTG9zdCAmIERlY2xpbmVkIOKAlCByZWFzb24gKGlu"
    "ZmVycmVkKSIsICJyZnFMb3N0IiwgImM0Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJNb250aGx5IFJlY2VpdmVkIHZzIFF1b3RlZCB2cyBXb24iLCAicmZxTW9u"
    "dGhseSIsICJjNiIpfQogICAgICAke2NhcmRTaGVsbCgiUkZRIFJlc3BvbnNlLVRpbWUgRGlzdHJpYnV0aW9uIiwgInJmcVJlc3BEaXN0IiwgImM2Iil9CiAg"
    "ICAgICR7Y2FyZFNoZWxsKCJXaW4gUmF0ZSBieSBDdXN0b21lciAo4omlMyBmaW5hbGl6ZWQpIiwgInJmcVdpbkN1c3QiLCAiYzYiKX0KICAgICAgJHtjYXJk"
    "U2hlbGwoIldpbiBSYXRlIGJ5IEVUIFBPQyIsICJyZnFXaW5Qb2MiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIldpbiBSYXRlIGJ5IFByb2R1Y3QgQ2F0"
    "ZWdvcnkiLCAicmZxV2luQ2F0IiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJRdW90ZWQgdnMgV29uIFZhbHVlIGJ5IE1vbnRoIiwgInJmcVF2VyIsICJj"
    "NiIpfQogICAgICAke2NhcmRTaGVsbCgiUkZRIFZhbHVlIGJ5IEN1c3RvbWVyIChUb3AgMTApIiwgInJmcVZhbEN1c3QiLCAiYzYiLCAiY2xpY2sgdG8gZmls"
    "dGVyIil9CiAgICAgICR7Y2FyZFNoZWxsKCJPcGVuIFJGUSBBZ2VpbmciLCAicmZxQWdlaW5nIiwgImM2Iil9CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNz"
    "PSJzZWN0aW9uLXRpdGxlIj5SRlEgRGV0YWlsPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjYXJkIGMxMiI+PGRpdiBpZD0icmZxVGFibGUiPjwvZGl2PjwvZGl2"
    "PmA7CgogIGRvbnV0KCJyZnFSZXN1bHQiLCBjb3VudEJ5KGxpc3QsIG8gPT4gby5yZXN1bHQpLCAobGJsKSA9PiBjaGFydEZpbHRlcigiX19yZXN1bHQiLCBs"
    "YmwpLCByZXN1bHRDb2xvcnMpOwogIHJlbmRlckNoYXJ0KCJyZnFXTEQiLCB7CiAgICB0eXBlOiAiYmFyIiwgZGF0YTogeyBsYWJlbHM6IFsiV29uIiwgIkxv"
    "c3QiLCAiRGVjbGluZWQiLCAiT3BlbiJdLCBkYXRhc2V0czogW3sKICAgICAgZGF0YTogW20ud29uLmxlbmd0aCwgbS5sb3N0Lmxlbmd0aCwgbS5kZWNsaW5l"
    "ZC5sZW5ndGgsIG0ub3Blbi5sZW5ndGhdLAogICAgICBiYWNrZ3JvdW5kQ29sb3I6IFtwYWxldHRlKCkuZ29vZCwgcGFsZXR0ZSgpLmJhZCwgcGFsZXR0ZSgp"
    "Lndhcm4sIHBhbGV0dGUoKS5tdXRlZF0sIGJvcmRlclJhZGl1czogMyB9XSB9LAogICAgb3B0aW9uczogeyBwbHVnaW5zOiB7IGxlZ2VuZDogeyBkaXNwbGF5"
    "OiBmYWxzZSB9IH0sIHNjYWxlczogeyB5OiB7IGJlZ2luQXRaZXJvOiB0cnVlIH0gfSB9LAogIH0pOwogIC8vIGxvc3QgcmVhc29uCiAgY29uc3QgbHIgPSBu"
    "ZXcgTWFwKCk7CiAgbGlzdC5mb3JFYWNoKG8gPT4geyBjb25zdCByID0gbG9zdFJlYXNvbihvLCByYmspOyBpZiAocikgbHIuc2V0KHIucmVhc29uLCAobHIu"
    "Z2V0KHIucmVhc29uKSB8fCAwKSArIDEpOyB9KTsKICBpZiAobHIuc2l6ZSkgZG9udXQoInJmcUxvc3QiLCBsciwgbnVsbCk7CiAgZWxzZSByZW5kZXJDaGFy"
    "dCgicmZxTG9zdCIsIHsgdHlwZTogImRvdWdobnV0IiwgZGF0YTogeyBsYWJlbHM6IFsiTm8gbG9zdC9kZWNsaW5lZCBSRlFzIl0sIGRhdGFzZXRzOiBbeyBk"
    "YXRhOiBbMV0sIGJhY2tncm91bmRDb2xvcjogW3BhbGV0dGUoKS5tdXRlZF0gfV0gfSwgb3B0aW9uczogeyBjdXRvdXQ6ICI1OCUiIH0gfSk7CiAgLy8gbW9u"
    "dGhseSByZWNlaXZlZC9xdW90ZWQvd29uCiAgY29uc3QgbW0gPSByZnFNb250aGx5QnJlYWtkb3duKHJvd3MsIGxpc3QpOwogIHJlbmRlckNoYXJ0KCJyZnFN"
    "b250aGx5IiwgewogICAgdHlwZTogImJhciIsIGRhdGE6IHsgbGFiZWxzOiBtbS5sYWJlbHMsIGRhdGFzZXRzOiBbCiAgICAgIHsgbGFiZWw6ICJSZWNlaXZl"
    "ZCIsIGRhdGE6IG1tLnJlY2VpdmVkLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5tdXRlZCwgYm9yZGVyUmFkaXVzOiAzIH0sCiAgICAgIHsgbGFiZWw6"
    "ICJRdW90ZWQiLCBkYXRhOiBtbS5xdW90ZWQsIGJhY2tncm91bmRDb2xvcjogcGFsZXR0ZSgpLnByaW1hcnksIGJvcmRlclJhZGl1czogMyB9LAogICAgICB7"
    "IGxhYmVsOiAiV29uIiwgZGF0YTogbW0ud29uLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5nb29kLCBib3JkZXJSYWRpdXM6IDMgfSwKICAgIF0gfSwg"
    "b3B0aW9uczogeyBzY2FsZXM6IHsgeTogeyBiZWdpbkF0WmVybzogdHJ1ZSB9IH0gfSwKICB9KTsKICBiYXJDaGFydCgicmZxUmVzcERpc3QiLCByZXNwRGlz"
    "dHJpYnV0aW9uKGxpc3QpLCBudWxsLCBudWxsLCBwYWxldHRlKCkuaW5mbyk7CiAgaGJhcigicmZxV2luQ3VzdCIsIHdpblJhdGVCeShsaXN0LCAiY3VzdG9t"
    "ZXIiKSwgdiA9PiB2ICsgIiUiKTsKICBoYmFyKCJyZnFXaW5Qb2MiLCB3aW5SYXRlQnkobGlzdCwgImV0UE9DIiwgMSksIHYgPT4gdiArICIlIik7CiAgaGJh"
    "cigicmZxV2luQ2F0Iiwgd2luUmF0ZUJ5KGxpc3QsICJwcm9kdWN0Q2F0ZWdvcnkiKSwgdiA9PiB2ICsgIiUiKTsKICByZW5kZXJDaGFydCgicmZxUXZXIiwg"
    "ewogICAgdHlwZTogImJhciIsIGRhdGE6IHsgbGFiZWxzOiBtbS5sYWJlbHMsIGRhdGFzZXRzOiBbCiAgICAgIHsgbGFiZWw6ICJRdW90ZWQgVmFsdWUiLCBk"
    "YXRhOiBtbS5xdW90ZWRWYWwsIGJhY2tncm91bmRDb2xvcjogcGFsZXR0ZSgpLnByaW1hcnksIGJvcmRlclJhZGl1czogMyB9LAogICAgICB7IGxhYmVsOiAi"
    "V29uIFZhbHVlIiwgZGF0YTogbW0ud29uVmFsLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5nb29kLCBib3JkZXJSYWRpdXM6IDMgfSwKICAgIF0gfSwg"
    "b3B0aW9uczogeyBzY2FsZXM6IHsgeTogeyBiZWdpbkF0WmVybzogdHJ1ZSwgdGlja3M6IHsgY2FsbGJhY2s6IHYgPT4gZm10Q29tcGFjdCh2KSB9IH0gfSB9"
    "LAogIH0pOwogIGhiYXIoInJmcVZhbEN1c3QiLCB0b3BOKGRlZHVwTWFwQnkocm93cywgInJmcUtleSIsICJldFF1b3RlZFZhbHVlIiwgImN1c3RvbWVyIiks"
    "IDEwKSwgZm10Q29tcGFjdCwgKGxibCkgPT4gc2V0TXVsdGkoImN1c3RvbWVyIiwgbGJsKSk7CiAgY29uc3QgYWdlID0gYWdlaW5nQnVja2V0cyhtLm9wZW4p"
    "OwogIHJlbmRlckNoYXJ0KCJyZnFBZ2VpbmciLCB7IHR5cGU6ICJiYXIiLCBkYXRhOiB7IGxhYmVsczogYWdlLmxhYmVscywgZGF0YXNldHM6IFt7IGRhdGE6"
    "IGFnZS5kYXRhLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS53YXJuLCBib3JkZXJSYWRpdXM6IDMgfV0gfSwgb3B0aW9uczogeyBwbHVnaW5zOiB7IGxl"
    "Z2VuZDogeyBkaXNwbGF5OiBmYWxzZSB9IH0sIHNjYWxlczogeyB5OiB7IGJlZ2luQXRaZXJvOiB0cnVlIH0gfSB9IH0pOwoKICAvLyBkZXRhaWwgdGFibGUK"
    "ICBjb25zdCB0YWJsZVJvd3MgPSBsaXN0Lm1hcChvID0+IHsKICAgIGNvbnN0IGxyMiA9IGxvc3RSZWFzb24obywgcmJrKTsKICAgIHJldHVybiB7CiAgICAg"
    "IHJmcU5vOiBvLmtleS5zcGxpdCgifHwiKVsxXSB8fCBvLmtleS5zcGxpdCgifHwiKVswXSwgY3VzdG9tZXI6IG8uY3VzdG9tZXIsIGV0UE9DOiBvLmV0UE9D"
    "LAogICAgICBjdXN0UmZxRGF0ZTogby5jdXN0UmZxRGF0ZSwgZXRRdW90ZURhdGU6IG8uZXRRdW90ZURhdGUsIHJlc3BEYXk6IG8ucmVzcERheSwKICAgICAg"
    "b25UaW1lOiBvLmNsb3NpbmdWYXJEYXkgPT0gbnVsbCA/ICLigJQiIDogKG8uY2xvc2luZ1ZhckRheSA8PSAwID8gIk9uIHRpbWUiIDogIkxhdGUiKSwKICAg"
    "ICAgcHJvZHVjdENhdGVnb3J5OiBvLnByb2R1Y3RDYXRlZ29yeSwgc2VjdG9yOiBvLnNlY3RvciwgbGluZXM6IG8ubGluZXMsCiAgICAgIHN1cHBsaWVyOiBv"
    "LnN1cHBsaWVyLCByZXY6IG8ucmV2LCBjb2dzOiBvLmNvZ3MsIGdwOiBvLmdwLCBtYXJnaW46IG8ubWFyZ2luLAogICAgICBzdGF0dXM6IG8ucmVzdWx0LCBj"
    "dXN0UG9Obzogby5jdXN0UG9ObywgbG9zdFJlYXNvbjogbHIyID8gbHIyLnJlYXNvbiArICIgKCIgKyBscjIudHlwZSArICIpIiA6ICLigJQiLAogICAgICBf"
    "bzogbywKICAgIH07CiAgfSk7CiAgbWFrZVRhYmxlKCQoIiNyZnFUYWJsZSIpLCB7CiAgICByb3dzOiB0YWJsZVJvd3MsCiAgICBuYW1lOiAicmZxX2RldGFp"
    "bCIsCiAgICBwZXJQYWdlOiAyNSwgZHJpbGw6IChyKSA9PiBvcGVuTW9kYWwoIlJGUSDigJQgIiArIChyLnJmcU5vIHx8ICIiKSwgcm9sbHVwTW9kYWxCb2R5"
    "KHIuX28pKSwKICAgIGNvbHVtbnM6IFsKICAgICAgeyBrZXk6ICJyZnFObyIsIGxhYmVsOiAiUkZRIE5vLiIgfSwgeyBrZXk6ICJjdXN0b21lciIsIGxhYmVs"
    "OiAiQ3VzdG9tZXIiIH0sCiAgICAgIHsga2V5OiAiZXRQT0MiLCBsYWJlbDogIkVUIFBPQyIgfSwgeyBrZXk6ICJjdXN0UmZxRGF0ZSIsIGxhYmVsOiAiUkZR"
    "IERhdGUiLCB0eXBlOiAiZGF0ZSIgfSwKICAgICAgeyBrZXk6ICJldFF1b3RlRGF0ZSIsIGxhYmVsOiAiUXVvdGUgRGF0ZSIsIHR5cGU6ICJkYXRlIiB9LCB7"
    "IGtleTogInJlc3BEYXkiLCBsYWJlbDogIlJlc3AiLCB0eXBlOiAiZGF5cyIgfSwKICAgICAgeyBrZXk6ICJvblRpbWUiLCBsYWJlbDogIk9uIHRpbWUiIH0s"
    "IHsga2V5OiAicHJvZHVjdENhdGVnb3J5IiwgbGFiZWw6ICJDYXRlZ29yeSIsIHZpczogZmFsc2UgfSwKICAgICAgeyBrZXk6ICJzZWN0b3IiLCBsYWJlbDog"
    "IlNlY3RvciIsIHZpczogZmFsc2UgfSwgeyBrZXk6ICJsaW5lcyIsIGxhYmVsOiAiTGluZXMiLCB0eXBlOiAibnVtIiB9LAogICAgICB7IGtleTogInN1cHBs"
    "aWVyIiwgbGFiZWw6ICJTdXBwbGllciIsIHZpczogZmFsc2UgfSwgeyBrZXk6ICJyZXYiLCBsYWJlbDogIlF1b3RlZCIsIHR5cGU6ICJjdXIiIH0sCiAgICAg"
    "IHsga2V5OiAiY29ncyIsIGxhYmVsOiAiQ29zdCIsIHR5cGU6ICJjdXIiLCB2aXM6IGZhbHNlIH0sIHsga2V5OiAiZ3AiLCBsYWJlbDogIkdQIiwgdHlwZTog"
    "ImN1ciIgfSwKICAgICAgeyBrZXk6ICJtYXJnaW4iLCBsYWJlbDogIk1hcmdpbiIsIHR5cGU6ICJwY3QiLCByZW5kZXI6ICh2KSA9PiB2ID09IG51bGwgPyAi"
    "4oCUIiA6IGA8c3BhbiBjbGFzcz0icGlsbCAke3YgPCAwID8gJ3InIDogdiA8IDggPyAnYScgOiAnZyd9Ij4ke2ZtdFBjdCh2KX08L3NwYW4+YCB9LAogICAg"
    "ICB7IGtleTogInN0YXR1cyIsIGxhYmVsOiAiUmVzdWx0IiwgcmVuZGVyOiAodikgPT4gYDxzcGFuIGNsYXNzPSJwaWxsICR7diA9PT0gJ1dvbicgPyAnZycg"
    "OiB2ID09PSAnTG9zdCcgPyAncicgOiB2ID09PSAnRGVjbGluZWQnID8gJ2EnIDogJ2InfSI+JHtlc2Modil9PC9zcGFuPmAgfSwKICAgICAgeyBrZXk6ICJj"
    "dXN0UG9ObyIsIGxhYmVsOiAiQ3VzdCBQTyIgfSwgeyBrZXk6ICJsb3N0UmVhc29uIiwgbGFiZWw6ICJMb3N0IHJlYXNvbiIsIHZpczogZmFsc2UgfSwKICAg"
    "IF0sCiAgfSk7Cn0KZnVuY3Rpb24gcmZxTW9udGhseUJyZWFrZG93bihyb3dzLCBsaXN0KSB7CiAgY29uc3QgbWFwID0gbmV3IE1hcCgpOwogIGNvbnN0IGVu"
    "c3VyZSA9IGsgPT4geyBpZiAoIW1hcC5oYXMoaykpIG1hcC5zZXQoaywgeyByZWM6IG5ldyBTZXQoKSwgcTogbmV3IFNldCgpLCB3OiBuZXcgU2V0KCksIHF2"
    "OiAwLCB3djogMCB9KTsgcmV0dXJuIG1hcC5nZXQoayk7IH07CiAgY29uc3QgcmV2ID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5IiwgImV0UXVvdGVkVmFs"
    "dWUiKTsKICBjb25zdCBieUtleU1vbnRoID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgbyBvZiBsaXN0KSB7IGNvbnN0IG1rID0gbW9udGhLZXkoby5ldFF1"
    "b3RlRGF0ZSkgfHwgbW9udGhLZXkoby5jdXN0UmZxRGF0ZSk7IGlmICghbWspIGNvbnRpbnVlOyBieUtleU1vbnRoLnNldChvLmtleSwgbWspOwogICAgY29u"
    "c3QgYiA9IGVuc3VyZShtayk7IGIucmVjLmFkZChvLmtleSk7IGlmIChvLnF1b3RlZCkgeyBiLnEuYWRkKG8ua2V5KTsgYi5xdiArPSByZXYuZ2V0KG8ua2V5"
    "KSB8fCAwOyB9CiAgICBpZiAoby53b24pIHsgYi53LmFkZChvLmtleSk7IGIud3YgKz0gcmV2LmdldChvLmtleSkgfHwgMDsgfSB9CiAgY29uc3QgbGFiZWxz"
    "ID0gQXJyYXkuZnJvbShtYXAua2V5cygpKS5zb3J0KCk7CiAgcmV0dXJuIHsKICAgIGxhYmVsczogbGFiZWxzLm1hcChsID0+IE1PTlRIU1srbC5zbGljZSg1"
    "KSAtIDFdICsgIiAiICsgbC5zbGljZSgyLCA0KSksCiAgICByZWNlaXZlZDogbGFiZWxzLm1hcChsID0+IG1hcC5nZXQobCkucmVjLnNpemUpLCBxdW90ZWQ6"
    "IGxhYmVscy5tYXAobCA9PiBtYXAuZ2V0KGwpLnEuc2l6ZSksCiAgICB3b246IGxhYmVscy5tYXAobCA9PiBtYXAuZ2V0KGwpLncuc2l6ZSksIHF1b3RlZFZh"
    "bDogbGFiZWxzLm1hcChsID0+IG1hcC5nZXQobCkucXYpLCB3b25WYWw6IGxhYmVscy5tYXAobCA9PiBtYXAuZ2V0KGwpLnd2KSwKICB9Owp9CgovKiA9PT09"
    "PT09PT09PT09PT09PSBUQUIgMyDigJQgU1VQUExJRVIgJiBQUk9DVVJFTUVOVCA9PT09PT09PT09PT09PT09PSAqLwpmdW5jdGlvbiByZW5kZXJTdXBwbGll"
    "cihyb3dzKSB7CiAgY29uc3QgZWwgPSAkKCIjdGFiLXN1cHBsaWVyIik7CiAgY29uc3QgYWN0aXZlU3VwID0gZGlzdGluY3RCeShyb3dzLCAoKSA9PiB0cnVl"
    "LCAic3VwcGxpZXJOYW1lIik7CiAgY29uc3QgcXVvdGVkU3VwID0gZGlzdGluY3RCeShyb3dzLCByID0+IHIuc3VwcGxpZXJRdW90ZURhdGUsICJzdXBwbGll"
    "ck5hbWUiKTsKICBjb25zdCBzZWxlY3RlZFN1cCA9IGRpc3RpbmN0Qnkocm93cywgciA9PiByLmV0UXVvdGVTdGF0dXMgPT09ICJXb24iLCAic3VwcGxpZXJO"
    "YW1lIik7CiAgY29uc3QgYWx0Q29udGFjdGVkID0gcm93cy5yZWR1Y2UoKGEsIHIpID0+IGEgKyAoci5hbHRTdXBwbGllckNvdW50IHx8IDApLCAwKTsKICBj"
    "b25zdCBwcm9jVmFsdWUgPSBhZ2dyZWdhdGVEZWR1cChyb3dzLCAicmZxS2V5IiwgInN1cHBsaWVyVG90YWxQcmljZSIpOwogIGNvbnN0IHdvblJvd3MgPSBy"
    "b3dzLmZpbHRlcihyID0+IHIuZXRRdW90ZVN0YXR1cyA9PT0gIldvbiIpOwogIGNvbnN0IHByb2NXb24gPSBhZ2dyZWdhdGVEZWR1cCh3b25Sb3dzLCAicmZx"
    "S2V5IiwgInN1cHBsaWVyVG90YWxQcmljZSIpOwogIGNvbnN0IHN1cFJlc3BBdmcgPSBhdmdGaWVsZChyb3dzLCAic3VwcGxpZXJSZXNwb25zZURheXMiKTsK"
    "ICBjb25zdCBzdXBSZXNwTWVkID0gbWVkaWFuKHJvd3MubWFwKHIgPT4gci5zdXBwbGllclJlc3BvbnNlRGF5cykuZmlsdGVyKHggPT4geCAhPSBudWxsKSk7"
    "CgogIGVsLmlubmVySFRNTCA9IGAKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPlByb2N1cmVtZW50IEtQSXM8L2Rpdj4KICAgIDxkaXYgY2xhc3M9"
    "ImtwaS1ncmlkIj4KICAgICAgJHtrcGkoIkFjdGl2ZSBTdXBwbGllcnMiLCBmbXROdW0oYWN0aXZlU3VwKSl9CiAgICAgICR7a3BpKCJTdXBwbGllcnMgUXVv"
    "dGVkIiwgZm10TnVtKHF1b3RlZFN1cCkpfQogICAgICAke2twaSgiU3VwcGxpZXJzIFNlbGVjdGVkICh3b24pIiwgZm10TnVtKHNlbGVjdGVkU3VwKSl9CiAg"
    "ICAgICR7a3BpKCJBbHQgU3VwcGxpZXJzIENvbnRhY3RlZCIsIGZtdE51bShhbHRDb250YWN0ZWQpKX0KICAgICAgJHtrcGkoIlRvdGFsIFByb2N1cmVtZW50"
    "IFZhbHVlIiwgZm10Q29tcGFjdChwcm9jVmFsdWUpLCBmbXRDdXIocHJvY1ZhbHVlKSl9CiAgICAgICR7a3BpKCJQcm9jdXJlbWVudCBWYWx1ZSAod29uKSIs"
    "IGZtdENvbXBhY3QocHJvY1dvbikpfQogICAgICAke2twaSgiQXZnIFN1cHBsaWVyIFF1b3RlIFJlc3AuIiwgZm10RGF5cyhzdXBSZXNwQXZnKSl9CiAgICAg"
    "ICR7a3BpKCJNZWRpYW4gU3VwcGxpZXIgUXVvdGUgUmVzcC4iLCBmbXREYXlzKHN1cFJlc3BNZWQpKX0KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iYmFu"
    "bmVyIGluZm8iPgogICAgICBTdXBwbGllciBQTyBwbGFjZW1lbnQgJmFtcDsgc2hpcG1lbnQgdHJhY2tpbmcgKFN1cHBsaWVyIFBPIE5vLiwgUE8vUlRTL2Fj"
    "dHVhbC1zaGlwIGRhdGVzLCBmaW5hbCBzaGlwbWVudCBzdGF0dXMg4oCUCiAgICAgIHNvdXJjZSBjb2x1bW5zIEJE4oCTQkopIGFyZSA8Yj5lbXB0eSBpbiB0"
    "aGUgY3VycmVudCBzb3VyY2UgZmlsZTwvYj4sIHNvIHN1cHBsaWVyLWRlbGl2ZXJ5IEtQSXMsCiAgICAgIG9uLXRpbWUgJSwgbGVhZC10aW1lIGFuZCBkZWxh"
    "eSBtZXRyaWNzIGNhbm5vdCBiZSBjb21wdXRlZC4gVGhleSBhcmUgc2hvd24gYXMgZW1wdHkgc3RhdGVzIGJlbG93IGFuZCB3aWxsIHBvcHVsYXRlCiAgICAg"
    "IGF1dG9tYXRpY2FsbHkgb25jZSB0aG9zZSBjb2x1bW5zIGFyZSBmaWxsZWQgYW5kIDxjb2RlPmNvbnZlcnQucHk8L2NvZGU+IGlzIHJlLXJ1bi4KICAgIDwv"
    "ZGl2PgoKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPlByb2N1cmVtZW50IEFuYWx5c2lzPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjaGFydC1ncmlk"
    "Ij4KICAgICAgJHtjYXJkU2hlbGwoIlByb2N1cmVtZW50IFZhbHVlIGJ5IFN1cHBsaWVyIChUb3AgMTIpIiwgInN1cFZhbHVlIiwgImM2IiwgImNsaWNrIHRv"
    "IGZpbHRlciIpfQogICAgICAke2NhcmRTaGVsbCgiU3VwcGxpZXIgU2VsZWN0aW9uIEZyZXF1ZW5jeSAod29uIFJGUXMpIiwgInN1cEZyZXEiLCAiYzYiLCAi"
    "Y2xpY2sgdG8gZmlsdGVyIil9CiAgICAgICR7Y2FyZFNoZWxsKCJQcm9jdXJlbWVudCBWYWx1ZSBieSBTZWN0b3IiLCAic3VwU2VjdG9yIiwgImM2Iil9CiAg"
    "ICAgICR7Y2FyZFNoZWxsKCJQcm9jdXJlbWVudCBWYWx1ZSBieSBQcm9kdWN0IENhdGVnb3J5IiwgInN1cENhdCIsICJjNiIpfQogICAgICAke2NhcmRTaGVs"
    "bCgiQXZnIFN1cHBsaWVyIFF1b3RlIFJlc3BvbnNlIGJ5IFN1cHBsaWVyIiwgInN1cFJlc3AiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIkdyb3NzIFBy"
    "b2ZpdCBTdXBwb3J0ZWQgYnkgU3VwcGxpZXIgKFRvcCAxMikiLCAic3VwR3AiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlN1cHBsaWVyIFNjb3JlIChk"
    "YXRhLWF2YWlsYWJsZSBmYWN0b3JzKSIsICJzdXBTY29yZSIsICJjNiIsICJxdW90ZSByZXNwb25zaXZlbmVzcyArIHNlbGVjdGlvbiArIGNvbXBldGl0aXZl"
    "bmVzcyIpfQogICAgICAke2VtcHR5Q2FyZCgiU3VwcGxpZXIgT24tVGltZSBTaGlwbWVudCAlIiwgImM2IiwgIk5vIHN1cHBsaWVyIHNoaXBtZW50IGRhdGVz"
    "IGluIHNvdXJjZSAoY29sdW1ucyBCSOKAk0JJIGVtcHR5KS4iKX0KICAgIDwvZGl2PgoKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPlN1cHBsaWVy"
    "IERldGFpbDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY2FyZCBjMTIiPjxkaXYgaWQ9InN1cFRhYmxlIj48L2Rpdj48L2Rpdj5gOwoKICBoYmFyKCJzdXBWYWx1"
    "ZSIsIHRvcE4oc3VtTWFwQnkocm93cywgInN1cHBsaWVyTmFtZSIsICJzdXBwbGllclRvdGFsUHJpY2UiKSwgMTIpLCBmbXRDb21wYWN0LCAobGJsKSA9PiBz"
    "ZXRNdWx0aSgic3VwcGxpZXJOYW1lIiwgbGJsKSk7CiAgaGJhcigic3VwRnJlcSIsIHRvcE4oZGVkdXBDb3VudEJ5KHdvblJvd3MsICJyZnFLZXkiLCAic3Vw"
    "cGxpZXJOYW1lIiksIDEyKSwgbnVsbCwgKGxibCkgPT4gc2V0TXVsdGkoInN1cHBsaWVyTmFtZSIsIGxibCkpOwogIGhiYXIoInN1cFNlY3RvciIsIHRvcE4o"
    "c3VtTWFwQnkocm93cywgInNlY3RvciIsICJzdXBwbGllclRvdGFsUHJpY2UiKSwgMTIpLCBmbXRDb21wYWN0KTsKICBoYmFyKCJzdXBDYXQiLCB0b3BOKHN1"
    "bU1hcEJ5KHJvd3MsICJwcm9kdWN0Q2F0ZWdvcnkiLCAic3VwcGxpZXJUb3RhbFByaWNlIiksIDEyKSwgZm10Q29tcGFjdCk7CiAgaGJhcigic3VwUmVzcCIs"
    "IHRvcE4oYXZnTWFwQnkocm93cywgInN1cHBsaWVyTmFtZSIsICJzdXBwbGllclJlc3BvbnNlRGF5cyIpLCAxMiksIHYgPT4gTWF0aC5yb3VuZCh2KSArICIg"
    "ZCIpOwogIGhiYXIoInN1cEdwIiwgdG9wTihkZWR1cE1hcEJ5KHJvd3MsICJyZnFLZXkiLCAiZ3Jvc3NQcm9maXRDYWxjIiwgInN1cHBsaWVyTmFtZSIpLCAx"
    "MiksIGZtdENvbXBhY3QpOwoKICBjb25zdCBzY29yZXMgPSBzdXBwbGllclNjb3Jlcyhyb3dzKTsKICBoYmFyKCJzdXBTY29yZSIsIHNjb3Jlcy5zbGljZSgw"
    "LCAxMikubWFwKHMgPT4gW3MubmFtZSwgcy5zY29yZV0pLCB2ID0+IE1hdGgucm91bmQodikgKyAiLzEwMCIpOwoKICAvLyBzdXBwbGllciB0YWJsZQogIGNv"
    "bnN0IHRSb3dzID0gc2NvcmVzLm1hcChzID0+ICh7CiAgICBzdXBwbGllcjogcy5uYW1lLCBzZWN0b3I6IHMuc2VjdG9yLCBwb0NvdW50OiBzLnNlbENvdW50"
    "LCBwcm9jVmFsdWU6IHMuc3BlbmQsCiAgICBncFN1cHBvcnRlZDogcy5ncCwgcXVvdGVSZXNwOiBzLnJlc3BBdmcsIGNvbXBldGl0aXZlbmVzczogcy5tYXJn"
    "aW5BdmcsIHNjb3JlOiBzLnNjb3JlLCBiYW5kOiBzLmJhbmQsCiAgfSkpOwogIG1ha2VUYWJsZSgkKCIjc3VwVGFibGUiKSwgewogICAgcm93czogdFJvd3Ms"
    "IG5hbWU6ICJzdXBwbGllcl9kZXRhaWwiLCBwZXJQYWdlOiAyNSwKICAgIGNvbHVtbnM6IFsKICAgICAgeyBrZXk6ICJzdXBwbGllciIsIGxhYmVsOiAiU3Vw"
    "cGxpZXIiIH0sIHsga2V5OiAic2VjdG9yIiwgbGFiZWw6ICJTZWN0b3IiIH0sCiAgICAgIHsga2V5OiAicG9Db3VudCIsIGxhYmVsOiAiU2VsZWN0ZWQgKHdv"
    "biBSRlFzKSIsIHR5cGU6ICJudW0iIH0sCiAgICAgIHsga2V5OiAicHJvY1ZhbHVlIiwgbGFiZWw6ICJQcm9jdXJlbWVudCBWYWx1ZSIsIHR5cGU6ICJjdXIi"
    "IH0sCiAgICAgIHsga2V5OiAiZ3BTdXBwb3J0ZWQiLCBsYWJlbDogIkdQIFN1cHBvcnRlZCIsIHR5cGU6ICJjdXIiIH0sCiAgICAgIHsga2V5OiAicXVvdGVS"
    "ZXNwIiwgbGFiZWw6ICJBdmcgUXVvdGUgUmVzcCIsIHR5cGU6ICJkYXlzIiB9LAogICAgICB7IGtleTogImNvbXBldGl0aXZlbmVzcyIsIGxhYmVsOiAiQXZn"
    "IE1hcmdpbiAlIiwgdHlwZTogInBjdCIgfSwKICAgICAgeyBrZXk6ICJzY29yZSIsIGxhYmVsOiAiU2NvcmUiLCB0eXBlOiAibnVtIiwgcmVuZGVyOiAodiwg"
    "cikgPT4gYDxzcGFuIGNsYXNzPSJwaWxsICR7diA+PSA3NSA/ICdnJyA6IHYgPj0gNjAgPyAnYicgOiB2ID49IDQwID8gJ2EnIDogJ3InfSI+JHtNYXRoLnJv"
    "dW5kKHYpfTwvc3Bhbj5gIH0sCiAgICAgIHsga2V5OiAiYmFuZCIsIGxhYmVsOiAiUmF0aW5nIiB9LAogICAgXSwKICB9KTsKfQpmdW5jdGlvbiBkZWR1cENv"
    "dW50Qnkocm93cywga2V5RmllbGQsIGdyb3VwRmllbGQpIHsKICBjb25zdCBvd25lciA9IG5ldyBNYXAoKTsKICBmb3IgKGNvbnN0IHIgb2Ygcm93cykgaWYg"
    "KHJba2V5RmllbGRdICE9IG51bGwgJiYgIW93bmVyLmhhcyhyW2tleUZpZWxkXSkpIG93bmVyLnNldChyW2tleUZpZWxkXSwgcltncm91cEZpZWxkXSB8fCAi"
    "4oCUIik7CiAgY29uc3Qgb3V0ID0gbmV3IE1hcCgpOwogIGZvciAoY29uc3QgZyBvZiBvd25lci52YWx1ZXMoKSkgb3V0LnNldChnLCAob3V0LmdldChnKSB8"
    "fCAwKSArIDEpOwogIHJldHVybiBvdXQ7Cn0KZnVuY3Rpb24gYXZnTWFwQnkocm93cywgZ3JvdXBGaWVsZCwgdmFsRmllbGQpIHsKICBjb25zdCBtID0gbmV3"
    "IE1hcCgpOwogIHJvd3MuZm9yRWFjaChyID0+IHsgY29uc3QgdiA9IHJbdmFsRmllbGRdOyBpZiAodiAhPSBudWxsICYmICFpc05hTih2KSkgeyBjb25zdCBr"
    "ID0gcltncm91cEZpZWxkXSB8fCAi4oCUIjsgY29uc3QgYiA9IG0uZ2V0KGspIHx8IHsgczogMCwgbjogMCB9OyBiLnMgKz0gdjsgYi5uKys7IG0uc2V0KGss"
    "IGIpOyB9IH0pOwogIGNvbnN0IG91dCA9IG5ldyBNYXAoKTsgZm9yIChjb25zdCBbaywgYl0gb2YgbSkgb3V0LnNldChrLCBiLnMgLyBiLm4pOyByZXR1cm4g"
    "b3V0Owp9CmZ1bmN0aW9uIHN1cHBsaWVyU2NvcmVzKHJvd3MpIHsKICBjb25zdCBzcGVuZCA9IHN1bU1hcEJ5KHJvd3MsICJzdXBwbGllck5hbWUiLCAic3Vw"
    "cGxpZXJUb3RhbFByaWNlIik7CiAgY29uc3QgZ3AgPSBkZWR1cE1hcEJ5KHJvd3MsICJyZnFLZXkiLCAiZ3Jvc3NQcm9maXRDYWxjIiwgInN1cHBsaWVyTmFt"
    "ZSIpOwogIGNvbnN0IHJlc3AgPSBhdmdNYXBCeShyb3dzLCAic3VwcGxpZXJOYW1lIiwgInN1cHBsaWVyUmVzcG9uc2VEYXlzIik7CiAgY29uc3Qgd29uUm93"
    "cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5ldFF1b3RlU3RhdHVzID09PSAiV29uIik7CiAgY29uc3Qgc2VsID0gZGVkdXBDb3VudEJ5KHdvblJvd3MsICJyZnFL"
    "ZXkiLCAic3VwcGxpZXJOYW1lIik7CiAgY29uc3QgbWFyZ2luQnkgPSBuZXcgTWFwKCk7CiAgY29uc3QgZyA9IGdyb3VwRGVkdXAocm93cywgInJmcUtleSIs"
    "ICJldFF1b3RlZFZhbHVlIiksIGdjID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpOwogIGNvbnN0IG93bmVyID0gbmV3"
    "IE1hcCgpOyByb3dzLmZvckVhY2gociA9PiB7IGlmICghb3duZXIuaGFzKHIucmZxS2V5KSkgb3duZXIuc2V0KHIucmZxS2V5LCByLnN1cHBsaWVyTmFtZSB8"
    "fCAi4oCUIik7IH0pOwogIGZvciAoY29uc3QgW2ssIHJldl0gb2YgZykgeyBpZiAoIXJldikgY29udGludWU7IGNvbnN0IHMgPSBvd25lci5nZXQoayk7IGNv"
    "bnN0IG1hciA9IChnYy5nZXQoaykgfHwgMCkgLyByZXYgKiAxMDA7IGNvbnN0IGIgPSBtYXJnaW5CeS5nZXQocykgfHwgeyBzOiAwLCBuOiAwIH07IGIucyAr"
    "PSBtYXI7IGIubisrOyBtYXJnaW5CeS5zZXQocywgYik7IH0KICBjb25zdCBzZWN0b3JCeSA9IG5ldyBNYXAoKTsgcm93cy5mb3JFYWNoKHIgPT4geyBpZiAo"
    "ci5zdXBwbGllck5hbWUgJiYgIXNlY3RvckJ5LmhhcyhyLnN1cHBsaWVyTmFtZSkpIHNlY3RvckJ5LnNldChyLnN1cHBsaWVyTmFtZSwgci5zZWN0b3IpOyB9"
    "KTsKCiAgY29uc3QgbmFtZXMgPSBBcnJheS5mcm9tKG5ldyBTZXQocm93cy5tYXAociA9PiByLnN1cHBsaWVyTmFtZSkuZmlsdGVyKEJvb2xlYW4pKSk7CiAg"
    "Y29uc3QgbWF4U3BlbmQgPSBNYXRoLm1heCgxLCAuLi5BcnJheS5mcm9tKHNwZW5kLnZhbHVlcygpKSk7CiAgY29uc3QgbWF4U2VsID0gTWF0aC5tYXgoMSwg"
    "Li4uQXJyYXkuZnJvbShzZWwudmFsdWVzKCkpKTsKICBjb25zdCByZXNwQXJyID0gQXJyYXkuZnJvbShyZXNwLnZhbHVlcygpKTsKICBjb25zdCByZXNwTWlu"
    "ID0gTWF0aC5taW4oLi4ucmVzcEFyciwgMCksIHJlc3BNYXggPSBNYXRoLm1heCguLi5yZXNwQXJyLCAxKTsKICBjb25zdCBvdXQgPSBuYW1lcy5tYXAobiA9"
    "PiB7CiAgICBjb25zdCBzcCA9IHNwZW5kLmdldChuKSB8fCAwLCBzYyA9IHNlbC5nZXQobikgfHwgMDsKICAgIGNvbnN0IHJBdmcgPSByZXNwLmdldChuKTsK"
    "ICAgIGNvbnN0IG1CID0gbWFyZ2luQnkuZ2V0KG4pOyBjb25zdCBtYXJnaW5BdmcgPSBtQiA/IG1CLnMgLyBtQi5uIDogbnVsbDsKICAgIC8vIGF2YWlsYWJs"
    "ZS1mYWN0b3Igd2VpZ2h0ZWQgc2NvcmUgKHJld2VpZ2h0ZWQ7IHNoaXBtZW50IGZhY3RvcnMgYWJzZW50KQogICAgbGV0IG51bSA9IDAsIGRlbiA9IDA7CiAg"
    "ICAvLyBzZWxlY3Rpb24gZnJlcXVlbmN5ICgzMCkKICAgIG51bSArPSAoc2MgLyBtYXhTZWwpICogMzA7IGRlbiArPSAzMDsKICAgIC8vIGNvbXBldGl0aXZl"
    "bmVzcyB2aWEgbWFyZ2luIHF1YWxpdHkgKDM1KSDigJQgaGlnaGVyIG1hcmdpbiA9IG1vcmUgY29tcGV0aXRpdmUgYnV5CiAgICBpZiAobWFyZ2luQXZnICE9"
    "IG51bGwpIHsgbnVtICs9IE1hdGgubWF4KDAsIE1hdGgubWluKDEsIG1hcmdpbkF2ZyAvIDMwKSkgKiAzNTsgZGVuICs9IDM1OyB9CiAgICAvLyBxdW90ZSBy"
    "ZXNwb25zaXZlbmVzcyAoMzUpIOKAlCBmYXN0ZXIgaXMgYmV0dGVyCiAgICBpZiAockF2ZyAhPSBudWxsICYmIHJlc3BNYXggPiByZXNwTWluKSB7IG51bSAr"
    "PSAoMSAtIChyQXZnIC0gcmVzcE1pbikgLyAocmVzcE1heCAtIHJlc3BNaW4pKSAqIDM1OyBkZW4gKz0gMzU7IH0KICAgIGNvbnN0IHNjb3JlID0gZGVuID8g"
    "KG51bSAvIGRlbikgKiAxMDAgOiAwOwogICAgY29uc3QgYmFuZCA9IHNjb3JlID49IDc1ID8gIkdvb2QrIiA6IHNjb3JlID49IDYwID8gIldhdGNoIiA6IHNj"
    "b3JlID49IDQwID8gIlJldmlldyIgOiAiTGltaXRlZCBkYXRhIjsKICAgIHJldHVybiB7IG5hbWU6IG4sIHNwZW5kOiBzcCwgZ3A6IGdwLmdldChuKSB8fCAw"
    "LCByZXNwQXZnOiByQXZnLCBzZWxDb3VudDogc2MsIG1hcmdpbkF2Zywgc2NvcmUsIGJhbmQsIHNlY3Rvcjogc2VjdG9yQnkuZ2V0KG4pIH07CiAgfSk7CiAg"
    "b3V0LnNvcnQoKGEsIGIpID0+IGIuc3BlbmQgLSBhLnNwZW5kKTsKICByZXR1cm4gb3V0Owp9CgovKiA9PT09PT09PT09PT09PT09PSBUQUIgNCDigJQgQ1VT"
    "VE9NRVIgUE8gJiBERUxJVkVSWSA9PT09PT09PT09PT09PT09PSAqLwpsZXQgUE9fU0VBUkNIID0gIiI7CmZ1bmN0aW9uIHJlbmRlclBPKHJvd3MpIHsKICBj"
    "b25zdCBlbCA9ICQoIiN0YWItcG8iKTsKICBjb25zdCBwb1Jvd3MwID0gcm93cy5maWx0ZXIociA9PiByLmN1c3RQb05vKTsKICBjb25zdCBwb1Jvd3MgPSBQ"
    "T19TRUFSQ0ggPyBwb1Jvd3MwLmZpbHRlcihyID0+IFN0cmluZyhyLmN1c3RQb05vKS50b0xvd2VyQ2FzZSgpLmluY2x1ZGVzKFBPX1NFQVJDSC50b0xvd2Vy"
    "Q2FzZSgpKSkgOiBwb1Jvd3MwOwogIGNvbnN0IG5QTyA9IHVuaXF1ZUNvdW50KHBvUm93cywgImN1c3RQb0tleSIpOwogIGNvbnN0IHBvVmFsID0gYWdncmVn"
    "YXRlRGVkdXAocG9Sb3dzLCAiY3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBwb0dwID0gYWdncmVnYXRlRGVkdXAocG9Sb3dzLCAiY3Vz"
    "dFBvS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpOwogIGNvbnN0IGF2Z1BPID0gc2FmZURpdihwb1ZhbCwgblBPKTsKICBjb25zdCBwb09hQXZnID0gYXZnRmll"
    "bGQocG9Sb3dzLCAicG9Ub09hRGF5cyIpOwogIGNvbnN0IHdpdGhPYSA9IHVuaXF1ZUNvdW50KHBvUm93cy5maWx0ZXIociA9PiByLmV0T2FEYXRlKSwgImN1"
    "c3RQb0tleSIpOwoKICBlbC5pbm5lckhUTUwgPSBgCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5DdXN0b21lciBQTyBLUElzPC9kaXY+CiAgICA8"
    "ZGl2IGNsYXNzPSJwby1zZWFyY2gtcm93Ij4KICAgICAgPGlucHV0IGlkPSJwb1NlYXJjaCIgdHlwZT0idGV4dCIgcGxhY2Vob2xkZXI9IlNlYXJjaCBhIGN1"
    "c3RvbWVyIFBPIG51bWJlcuKApiIgdmFsdWU9IiR7ZXNjKFBPX1NFQVJDSCl9IiAvPgogICAgICAke1BPX1NFQVJDSCA/IGA8YnV0dG9uIGlkPSJwb1NlYXJj"
    "aENsZWFyIiBjbGFzcz0iYnRuIGJ0bi1naG9zdCI+Q2xlYXI8L2J1dHRvbj5gIDogIiJ9CiAgICAgIDxzcGFuIGNsYXNzPSJwby1zZWFyY2gtbm90ZSI+JHtQ"
    "T19TRUFSQ0ggPyBmbXROdW0oblBPKSArICIgUE8ocykgbWF0Y2giIDogIiJ9PC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCI+"
    "CiAgICAgICR7a3BpKCJUb3RhbCBDdXN0b21lciBQT3MiLCBmbXROdW0oblBPKSl9CiAgICAgICR7a3BpKCJDdXN0b21lciBQTyBWYWx1ZSIsIGZtdENvbXBh"
    "Y3QocG9WYWwpLCBmbXRDdXIocG9WYWwpKX0KICAgICAgJHtrcGkoIkF2ZyBDdXN0b21lciBQTyBWYWx1ZSIsIGZtdENvbXBhY3QoYXZnUE8pKX0KICAgICAg"
    "JHtrcGkoIkdyb3NzIFByb2ZpdCBvbiBQT3MiLCBmbXRDb21wYWN0KHBvR3ApKX0KICAgICAgJHtrcGkoIlBPcyB3aXRoIEVUIE9BIiwgZm10TnVtKHdpdGhP"
    "YSksIG5QTyA/IGZtdFBjdCh3aXRoT2EgLyBuUE8gKiAxMDApICsgIiBvZiBQT3MiIDogIiIpfQogICAgICAke2twaSgiQXZnIFBP4oaST0EgVGltZSIsIGZt"
    "dERheXMocG9PYUF2ZykpfQogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJiYW5uZXIgaW5mbyI+CiAgICAgIERlbGl2ZXJ5LXRpbWVsaW5lIGNvbHVtbnMg"
    "4oCUIEN1c3RvbWVyIFJlcXVpcmVkIERhdGUsIEVUIFByb21pc2VkIC8gUlRTIC8gQWN0dWFsLVNoaXAgRGF0ZSwgYW5kIFNoaXBtZW50IEZpbmFsCiAgICAg"
    "IFN0YXR1cyAoc291cmNlIGNvbHVtbnMgQVrigJNCQywgQkopIOKAlCBhcmUgPGI+ZW1wdHkgaW4gdGhlIGN1cnJlbnQgc291cmNlIGZpbGU8L2I+LiBPcmRl"
    "ciBmdWxmaWxtZW50IC8gZGVsaXZlcnktc3RhdHVzLAogICAgICBvbi10aW1lICUsIGFuZCBkZWxheSBtZXRyaWNzIHRoZXJlZm9yZSBjYW5ub3QgYmUgY2Fs"
    "Y3VsYXRlZCBhbmQgYXJlIHNob3duIGFzIGVtcHR5IHN0YXRlcy4gUE8gdmFsdWUsIE9BIHByb2Nlc3NpbmcKICAgICAgYW5kIHByb2ZpdGFiaWxpdHkgKHdo"
    "aWNoIHRoZSBkYXRhIHN1cHBvcnRzKSBhcmUgZnVsbHkgcG9wdWxhdGVkLgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+Q3Vz"
    "dG9tZXIgUE8gQW5hbHlzaXM8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXJ0LWdyaWQiPgogICAgICAke2NhcmRTaGVsbCgiQ3VzdG9tZXIgUE8gVmFsdWUg"
    "YnkgTW9udGgiLCAicG9UcmVuZCIsICJjNiIpfQogICAgICAke2NhcmRTaGVsbCgiUE8gVmFsdWUgYnkgQ3VzdG9tZXIgKFRvcCAxMikiLCAicG9CeUN1c3Qi"
    "LCAiYzYiLCAiY2xpY2sgdG8gZmlsdGVyIil9CiAgICAgICR7Y2FyZFNoZWxsKCJQTyBWYWx1ZSBieSBFVCBQT0MiLCAicG9CeVBvYyIsICJjNiIpfQogICAg"
    "ICAke2NhcmRTaGVsbCgiUE8gVmFsdWUgYnkgU2VjdG9yIiwgInBvQnlTZWN0b3IiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlBPIFZhbHVlIGJ5IFBy"
    "b2R1Y3QgQ2F0ZWdvcnkiLCAicG9CeUNhdCIsICJjNiIpfQogICAgICAke2NhcmRTaGVsbCgiUE8g4oaSIE9BIFByb2Nlc3NpbmcgRGF5cyAoZGlzdHJpYnV0"
    "aW9uKSIsICJwb09hRGlzdCIsICJjNiIpfQogICAgICAke2VtcHR5Q2FyZCgiT3JkZXJzIGJ5IERlbGl2ZXJ5IFN0YXR1cyIsICJjNiIsICJObyBkZWxpdmVy"
    "eS9zaGlwbWVudCBzdGF0dXMgaW4gc291cmNlIChjb2x1bW5zIEFa4oCTQkMsIEJKIGVtcHR5KS4iKX0KICAgICAgJHtlbXB0eUNhcmQoIk9uLVRpbWUgQ3Vz"
    "dG9tZXIgU2hpcG1lbnQgJSIsICJjNiIsICJObyBFVCBwcm9taXNlZCAvIGFjdHVhbC1zaGlwIGRhdGVzIGluIHNvdXJjZS4iKX0KICAgIDwvZGl2PgoKICAg"
    "IDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkN1c3RvbWVyIFBPIERldGFpbDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY2FyZCBjMTIiPjxkaXYgaWQ9InBv"
    "VGFibGUiPjwvZGl2PjwvZGl2PmA7CgogIC8vIHdpcmUgUE8gc2VhcmNoCiAgY29uc3QgcHNpID0gJCgiI3BvU2VhcmNoIik7CiAgbGV0IGRlYjsKICBwc2ku"
    "b25pbnB1dCA9ICgpID0+IHsgY2xlYXJUaW1lb3V0KGRlYik7IGRlYiA9IHNldFRpbWVvdXQoKCkgPT4geyBQT19TRUFSQ0ggPSBwc2kudmFsdWUudHJpbSgp"
    "OyBSRU5ERVJFRC5wbyA9IGZhbHNlOyByZW5kZXJQTyhhcHBseUZpbHRlcnMoKSk7IH0sIDIwMCk7IH07CiAgcHNpLm9ua2V5ZG93biA9IChlKSA9PiB7IGlm"
    "IChlLmtleSA9PT0gIkVudGVyIikgeyBQT19TRUFSQ0ggPSBwc2kudmFsdWUudHJpbSgpOyBSRU5ERVJFRC5wbyA9IGZhbHNlOyByZW5kZXJQTyhhcHBseUZp"
    "bHRlcnMoKSk7IH0gfTsKICBpZiAoJCgiI3BvU2VhcmNoQ2xlYXIiKSkgJCgiI3BvU2VhcmNoQ2xlYXIiKS5vbmNsaWNrID0gKCkgPT4geyBQT19TRUFSQ0gg"
    "PSAiIjsgUkVOREVSRUQucG8gPSBmYWxzZTsgcmVuZGVyUE8oYXBwbHlGaWx0ZXJzKCkpOyB9OwoKICBpZiAoIW5QTykgeyByZXR1cm47IH0KICAvLyBtb250"
    "aGx5IFBPIHZhbHVlIChieSBjdXN0UG9EYXRlKQogIGNvbnN0IHBtID0gcG9Nb250aGx5KHBvUm93cyk7CiAgcmVuZGVyQ2hhcnQoInBvVHJlbmQiLCB7CiAg"
    "ICB0eXBlOiAiYmFyIiwgZGF0YTogeyBsYWJlbHM6IHBtLmxhYmVscywgZGF0YXNldHM6IFt7IGxhYmVsOiAiUE8gVmFsdWUiLCBkYXRhOiBwbS52YWwsIGJh"
    "Y2tncm91bmRDb2xvcjogcGFsZXR0ZSgpLnByaW1hcnksIGJvcmRlclJhZGl1czogMyB9XSB9LAogICAgb3B0aW9uczogeyBwbHVnaW5zOiB7IGxlZ2VuZDog"
    "eyBkaXNwbGF5OiBmYWxzZSB9IH0sIHNjYWxlczogeyB5OiB7IGJlZ2luQXRaZXJvOiB0cnVlLCB0aWNrczogeyBjYWxsYmFjazogdiA9PiBmbXRDb21wYWN0"
    "KHYpIH0gfSB9IH0sCiAgfSk7CiAgaGJhcigicG9CeUN1c3QiLCB0b3BOKGRlZHVwTWFwQnkocG9Sb3dzLCAiY3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUi"
    "LCAiY3VzdG9tZXIiKSwgMTIpLCBmbXRDb21wYWN0LCAobGJsKSA9PiBzZXRNdWx0aSgiY3VzdG9tZXIiLCBsYmwpKTsKICBoYmFyKCJwb0J5UG9jIiwgdG9w"
    "TihkZWR1cE1hcEJ5KHBvUm93cywgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIiwgImV0UE9DIiksIDEyKSwgZm10Q29tcGFjdCk7CiAgaGJhcigicG9C"
    "eVNlY3RvciIsIHRvcE4oZGVkdXBNYXBCeShwb1Jvd3MsICJjdXN0UG9LZXkiLCAiZXRRdW90ZWRWYWx1ZSIsICJzZWN0b3IiKSwgMTIpLCBmbXRDb21wYWN0"
    "KTsKICBoYmFyKCJwb0J5Q2F0IiwgdG9wTihkZWR1cE1hcEJ5KHBvUm93cywgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIiwgInByb2R1Y3RDYXRlZ29y"
    "eSIpLCAxMiksIGZtdENvbXBhY3QpOwogIGJhckNoYXJ0KCJwb09hRGlzdCIsIHBvT2FEaXN0cmlidXRpb24ocG9Sb3dzKSwgbnVsbCwgbnVsbCwgcGFsZXR0"
    "ZSgpLmluZm8pOwoKICAvLyBQTyBkZXRhaWwgdGFibGUgKG9uZSByb3cgcGVyIFBPKQogIGNvbnN0IHBvTWFwID0gbmV3IE1hcCgpOwogIHBvUm93cy5mb3JF"
    "YWNoKHIgPT4gewogICAgbGV0IG8gPSBwb01hcC5nZXQoci5jdXN0UG9LZXkpOwogICAgaWYgKCFvKSB7IG8gPSB7IHBvTm86IHIuY3VzdFBvTm8sIGN1c3Rv"
    "bWVyOiByLmN1c3RvbWVyLCBldFBPQzogci5ldFBPQywgcG9EYXRlOiByLmN1c3RQb0RhdGUsIG9hRGF0ZTogci5ldE9hRGF0ZSwgc2VjdG9yOiByLnNlY3Rv"
    "ciwgbGluZXM6IDAsIGtleXM6IG5ldyBTZXQoKSB9OyBwb01hcC5zZXQoci5jdXN0UG9LZXksIG8pOyB9CiAgICBvLmxpbmVzKys7CiAgfSk7CiAgY29uc3Qg"
    "cG9WYWxNYXAgPSBncm91cERlZHVwKHBvUm93cywgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIik7CiAgY29uc3QgcG9HcE1hcCA9IGdyb3VwRGVkdXAo"
    "cG9Sb3dzLCAiY3VzdFBvS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpOwogIGNvbnN0IHBvQ29zdE1hcCA9IGdyb3VwRGVkdXAocG9Sb3dzLCAiY3VzdFBvS2V5"
    "IiwgInN1cHBsaWVyVG90YWxQcmljZSIpOwogIGNvbnN0IGtleUJ5SWQgPSBuZXcgTWFwKCk7IHBvUm93cy5mb3JFYWNoKHIgPT4ga2V5QnlJZC5zZXQoci5j"
    "dXN0UG9LZXksIHIpKTsKICBjb25zdCB0Um93cyA9IEFycmF5LmZyb20ocG9NYXAuZW50cmllcygpKS5tYXAoKFtrLCBvXSkgPT4gewogICAgY29uc3QgcmV2"
    "ID0gcG9WYWxNYXAuZ2V0KGspIHx8IDAsIGdwID0gcG9HcE1hcC5nZXQoaykgfHwgMDsKICAgIHJldHVybiB7IHBvTm86IG8ucG9ObywgY3VzdG9tZXI6IG8u"
    "Y3VzdG9tZXIsIGV0UE9DOiBvLmV0UE9DLCBwb0RhdGU6IG8ucG9EYXRlLCBvYURhdGU6IG8ub2FEYXRlLAogICAgICBwb1RvT2E6IG8ucG9EYXRlICYmIG8u"
    "b2FEYXRlID8gZGF5c0JldHdlZW4oby5vYURhdGUsIG8ucG9EYXRlKSA6IG51bGwsCiAgICAgIHNlY3Rvcjogby5zZWN0b3IsIGxpbmVzOiBvLmxpbmVzLCB2"
    "YWx1ZTogcmV2LCBjb3N0OiBwb0Nvc3RNYXAuZ2V0KGspIHx8IDAsIGdwLAogICAgICBtYXJnaW46IHJldiA/IGdwIC8gcmV2ICogMTAwIDogbnVsbCwgZGVs"
    "aXZlcnk6ICJObyBkYXRlIGluIHNvdXJjZSIgfTsKICB9KTsKICBtYWtlVGFibGUoJCgiI3BvVGFibGUiKSwgewogICAgcm93czogdFJvd3MsIG5hbWU6ICJj"
    "dXN0b21lcl9wb19kZXRhaWwiLCBwZXJQYWdlOiAyNSwKICAgIGNvbHVtbnM6IFsKICAgICAgeyBrZXk6ICJwb05vIiwgbGFiZWw6ICJDdXN0b21lciBQTyBO"
    "by4iIH0sIHsga2V5OiAiY3VzdG9tZXIiLCBsYWJlbDogIkN1c3RvbWVyIiB9LAogICAgICB7IGtleTogImV0UE9DIiwgbGFiZWw6ICJFVCBQT0MiIH0sIHsg"
    "a2V5OiAicG9EYXRlIiwgbGFiZWw6ICJQTyBEYXRlIiwgdHlwZTogImRhdGUiIH0sCiAgICAgIHsga2V5OiAib2FEYXRlIiwgbGFiZWw6ICJPQSBEYXRlIiwg"
    "dHlwZTogImRhdGUiIH0sIHsga2V5OiAicG9Ub09hIiwgbGFiZWw6ICJQT+KGkk9BIiwgdHlwZTogImRheXMiIH0sCiAgICAgIHsga2V5OiAic2VjdG9yIiwg"
    "bGFiZWw6ICJTZWN0b3IiLCB2aXM6IGZhbHNlIH0sIHsga2V5OiAibGluZXMiLCBsYWJlbDogIkxpbmVzIiwgdHlwZTogIm51bSIgfSwKICAgICAgeyBrZXk6"
    "ICJ2YWx1ZSIsIGxhYmVsOiAiUE8gVmFsdWUiLCB0eXBlOiAiY3VyIiB9LCB7IGtleTogImNvc3QiLCBsYWJlbDogIkNvc3QiLCB0eXBlOiAiY3VyIiwgdmlz"
    "OiBmYWxzZSB9LAogICAgICB7IGtleTogImdwIiwgbGFiZWw6ICJHUCIsIHR5cGU6ICJjdXIiIH0sCiAgICAgIHsga2V5OiAibWFyZ2luIiwgbGFiZWw6ICJN"
    "YXJnaW4iLCB0eXBlOiAicGN0IiwgcmVuZGVyOiB2ID0+IHYgPT0gbnVsbCA/ICLigJQiIDogYDxzcGFuIGNsYXNzPSJwaWxsICR7diA8IDAgPyAncicgOiB2"
    "IDwgOCA/ICdhJyA6ICdnJ30iPiR7Zm10UGN0KHYpfTwvc3Bhbj5gIH0sCiAgICAgIHsga2V5OiAiZGVsaXZlcnkiLCBsYWJlbDogIkRlbGl2ZXJ5Iiwgdmlz"
    "OiBmYWxzZSB9LAogICAgXSwKICB9KTsKfQpmdW5jdGlvbiBwb01vbnRobHkocG9Sb3dzKSB7CiAgY29uc3QgbWFwID0gbmV3IE1hcCgpOwogIGNvbnN0IHZh"
    "bCA9IGdyb3VwRGVkdXAocG9Sb3dzLCAiY3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBvd25lciA9IG5ldyBNYXAoKTsgcG9Sb3dzLmZv"
    "ckVhY2gociA9PiB7IGlmICghb3duZXIuaGFzKHIuY3VzdFBvS2V5KSkgb3duZXIuc2V0KHIuY3VzdFBvS2V5LCBtb250aEtleShyLmN1c3RQb0RhdGUpKTsg"
    "fSk7CiAgZm9yIChjb25zdCBbaywgdl0gb2YgdmFsKSB7IGNvbnN0IG1rID0gb3duZXIuZ2V0KGspOyBpZiAoIW1rKSBjb250aW51ZTsgbWFwLnNldChtaywg"
    "KG1hcC5nZXQobWspIHx8IDApICsgdik7IH0KICBjb25zdCBsYWJlbHMgPSBBcnJheS5mcm9tKG1hcC5rZXlzKCkpLnNvcnQoKTsKICByZXR1cm4geyBsYWJl"
    "bHM6IGxhYmVscy5tYXAobCA9PiBNT05USFNbK2wuc2xpY2UoNSkgLSAxXSArICIgIiArIGwuc2xpY2UoMiwgNCkpLCB2YWw6IGxhYmVscy5tYXAobCA9PiBt"
    "YXAuZ2V0KGwpKSB9Owp9CmZ1bmN0aW9uIHBvT2FEaXN0cmlidXRpb24ocG9Sb3dzKSB7CiAgY29uc3QgYiA9IHsgIuKJpDAgZCI6IDAsICIx4oCTMyBkIjog"
    "MCwgIjTigJM3IGQiOiAwLCAiOOKAkzE0IGQiOiAwLCAiPjE0IGQiOiAwIH07CiAgY29uc3Qgc2VlbiA9IG5ldyBTZXQoKTsKICBwb1Jvd3MuZm9yRWFjaChy"
    "ID0+IHsKICAgIGlmIChzZWVuLmhhcyhyLmN1c3RQb0tleSkpIHJldHVybjsgc2Vlbi5hZGQoci5jdXN0UG9LZXkpOwogICAgY29uc3QgZCA9IHIucG9Ub09h"
    "RGF5czsgaWYgKGQgPT0gbnVsbCkgcmV0dXJuOwogICAgaWYgKGQgPD0gMCkgYlsi4omkMCBkIl0rKzsgZWxzZSBpZiAoZCA8PSAzKSBiWyIx4oCTMyBkIl0r"
    "KzsgZWxzZSBpZiAoZCA8PSA3KSBiWyI04oCTNyBkIl0rKzsKICAgIGVsc2UgaWYgKGQgPD0gMTQpIGJbIjjigJMxNCBkIl0rKzsgZWxzZSBiWyI+MTQgZCJd"
    "Kys7CiAgfSk7CiAgcmV0dXJuIE9iamVjdC5lbnRyaWVzKGIpOwp9CgovKiA9PT09PT09PT09PT09PT09PSBUQUIgNSDigJQgRVQgUE9DIFBFUkZPUk1BTkNF"
    "ID09PT09PT09PT09PT09PT09ICovCmZ1bmN0aW9uIHJlbmRlclBPQyhyb3dzKSB7CiAgY29uc3QgZWwgPSAkKCIjdGFiLXBvYyIpOwogIGNvbnN0IFIgPSBi"
    "dWlsZFJvbGx1cChyb3dzKTsKICBjb25zdCBieVBvYyA9IG5ldyBNYXAoKTsKICBSLmxpc3QuZm9yRWFjaChvID0+IHsKICAgIGNvbnN0IGsgPSBvLmV0UE9D"
    "IHx8ICLigJQiOwogICAgY29uc3QgYiA9IGJ5UG9jLmdldChrKSB8fCB7IHBvYzogaywgcmZxczogMCwgcXVvdGVkOiAwLCB3b246IDAsIGxvc3Q6IDAsIHJl"
    "djogMCwgZ3A6IDAsIGNvZ3M6IDAsIHJlc3A6IFtdLCBwb0tleXM6IG5ldyBTZXQoKSB9OwogICAgYi5yZnFzKys7IGlmIChvLnF1b3RlZCkgYi5xdW90ZWQr"
    "KzsgaWYgKG8ud29uKSBiLndvbisrOyBpZiAoby5sb3N0KSBiLmxvc3QrKzsKICAgIGIucmV2ICs9IG8ucmV2OyBiLmdwICs9IG8uZ3A7IGIuY29ncyArPSBv"
    "LmNvZ3M7IGlmIChvLnJlc3BEYXkgIT0gbnVsbCkgYi5yZXNwLnB1c2goby5yZXNwRGF5KTsKICAgIGlmIChvLmN1c3RQb0tleSkgYi5wb0tleXMuYWRkKG8u"
    "Y3VzdFBvS2V5KTsKICAgIGJ5UG9jLnNldChrLCBiKTsKICB9KTsKICBjb25zdCBhcnIgPSBBcnJheS5mcm9tKGJ5UG9jLnZhbHVlcygpKS5tYXAoYiA9PiB7"
    "CiAgICBjb25zdCBmaW4gPSBiLndvbiArIGIubG9zdDsKICAgIGIud2luUmF0ZSA9IGZpbiA/IGIud29uIC8gZmluICogMTAwIDogMDsKICAgIGIuY29udiA9"
    "IGIucXVvdGVkID8gYi5wb0tleXMuc2l6ZSAvIGIucXVvdGVkICogMTAwIDogMDsKICAgIGIubWFyZ2luID0gYi5yZXYgPyBiLmdwIC8gYi5yZXYgKiAxMDAg"
    "OiAwOwogICAgYi5yZXNwQXZnID0gYi5yZXNwLmxlbmd0aCA/IGIucmVzcC5yZWR1Y2UoKGEsIGMpID0+IGEgKyBjLCAwKSAvIGIucmVzcC5sZW5ndGggOiBu"
    "dWxsOwogICAgYi5uUE8gPSBiLnBvS2V5cy5zaXplOwogICAgcmV0dXJuIGI7CiAgfSkuZmlsdGVyKGIgPT4gYi5wb2MgIT09ICLigJQiIHx8IGIucmZxcyA+"
    "IDApOwogIC8vIGJhbGFuY2VkIHNjb3JlIChhdmFpbGFibGUgZmFjdG9yczogR1AsIHdpbiByYXRlLCBjb252ZXJzaW9uLCByZXNwb25zZSwgbWFyZ2luKQog"
    "IGNvbnN0IG1heEdwID0gTWF0aC5tYXgoMSwgLi4uYXJyLm1hcChiID0+IGIuZ3ApKTsKICBjb25zdCByZXNwVmFscyA9IGFyci5tYXAoYiA9PiBiLnJlc3BB"
    "dmcpLmZpbHRlcih4ID0+IHggIT0gbnVsbCk7CiAgY29uc3Qgck1pbiA9IE1hdGgubWluKC4uLnJlc3BWYWxzLCAwKSwgck1heCA9IE1hdGgubWF4KC4uLnJl"
    "c3BWYWxzLCAxKTsKICBhcnIuZm9yRWFjaChiID0+IHsKICAgIGxldCBudW0gPSAwLCBkZW4gPSAwOwogICAgbnVtICs9IE1hdGgubWF4KDAsIGIuZ3AgLyBt"
    "YXhHcCkgKiAzMDsgZGVuICs9IDMwOyAgICAgICAgICAgICAgICAgICAgICAgLy8gR1AgYWNoaWV2ZW1lbnQKICAgIG51bSArPSBNYXRoLm1pbigxLCBiLndp"
    "blJhdGUgLyAxMDApICogMjI7IGRlbiArPSAyMjsgICAgICAgICAgICAgICAgICAgIC8vIHdpbiByYXRlCiAgICBudW0gKz0gTWF0aC5taW4oMSwgYi5jb252"
    "IC8gNTApICogMTY7IGRlbiArPSAxNjsgICAgICAgICAgICAgICAgICAgICAgICAvLyBjb252ZXJzaW9uCiAgICBpZiAoYi5yZXNwQXZnICE9IG51bGwgJiYg"
    "ck1heCA+IHJNaW4pIHsgbnVtICs9ICgxIC0gKGIucmVzcEF2ZyAtIHJNaW4pIC8gKHJNYXggLSByTWluKSkgKiAxNjsgZGVuICs9IDE2OyB9CiAgICBudW0g"
    "Kz0gTWF0aC5tYXgoMCwgTWF0aC5taW4oMSwgYi5tYXJnaW4gLyAzMCkpICogMTY7IGRlbiArPSAxNjsgICAgICAgICAvLyBtYXJnaW4gcXVhbGl0eQogICAg"
    "Yi5zY29yZSA9IGRlbiA/IG51bSAvIGRlbiAqIDEwMCA6IDA7CiAgfSk7CiAgYXJyLnNvcnQoKGEsIGIpID0+IGIuc2NvcmUgLSBhLnNjb3JlKTsKICBhcnIu"
    "Zm9yRWFjaCgoYiwgaSkgPT4geyBiLnJhbmsgPSBpICsgMTsgYi5xdWFydGlsZSA9IE1hdGguY2VpbCgoaSArIDEpIC8gTWF0aC5tYXgoMSwgYXJyLmxlbmd0"
    "aCkgKiA0KTsgfSk7CgogIGVsLmlubmVySFRNTCA9IGAKICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkVUIFBPQyBQZXJmb3JtYW5jZTwvZGl2Pgog"
    "ICAgPGRpdiBjbGFzcz0iYmFubmVyIGluZm8iPlNjb3JlcyBhcmUgYW4gaW50ZXJuYWwgbWFuYWdlbWVudCBpbmRpY2F0b3IgKGdyb3NzLXByb2ZpdCwgd2lu"
    "IHJhdGUsIGNvbnZlcnNpb24sIHJlc3BvbnNlIHRpbWUKICAgICAgYW5kIG1hcmdpbiBxdWFsaXR5IOKAlCByZXdlaWdodGVkIGZvciBhdmFpbGFibGUgZGF0"
    "YSksIDxiPm5vdCBhbiBhdWRpdGVkIGVtcGxveWVlIGFwcHJhaXNhbDwvYj4uIFNoaXBtZW50LWJhc2VkIGZhY3RvcnMKICAgICAgYXJlIGV4Y2x1ZGVkIGJl"
    "Y2F1c2UgZGVsaXZlcnkgY29sdW1ucyBhcmUgZW1wdHkgaW4gdGhlIHNvdXJjZS48L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXJ0LWdyaWQiPgogICAgICAk"
    "e2NhcmRTaGVsbCgiQmFsYW5jZWQgU2NvcmUgYnkgRVQgUE9DIiwgInBvY1Njb3JlIiwgImM2IiwgIjDigJMxMDAsIGhpZ2hlciBpcyBiZXR0ZXIiKX0KICAg"
    "ICAgJHtjYXJkU2hlbGwoIkdyb3NzIFByb2ZpdCBieSBFVCBQT0MiLCAicG9jR3AiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlJGUXMgdnMgV29uIGJ5"
    "IEVUIFBPQyIsICJwb2NSZnFXb24iLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIldpbiBSYXRlIGJ5IEVUIFBPQyIsICJwb2NXaW4iLCAiYzYiKX0KICAg"
    "ICAgJHtjYXJkU2hlbGwoIkVUIFF1b3RlZCBWYWx1ZSBieSBFVCBQT0MiLCAicG9jUmV2IiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJBdmcgUkZRIFJl"
    "c3BvbnNlIGJ5IEVUIFBPQyIsICJwb2NSZXNwIiwgImM2Iil9CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPlBPQyBTY29yZWNh"
    "cmQ8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNhcmQgYzEyIj48ZGl2IGlkPSJwb2NUYWJsZSI+PC9kaXY+PC9kaXY+YDsKCiAgaGJhcigicG9jU2NvcmUiLCBh"
    "cnIubWFwKGIgPT4gW2IucG9jLCArYi5zY29yZS50b0ZpeGVkKDEpXSksIHYgPT4gTWF0aC5yb3VuZCh2KSArICIvMTAwIik7CiAgaGJhcigicG9jR3AiLCB0"
    "b3BOKG5ldyBNYXAoYXJyLm1hcChiID0+IFtiLnBvYywgYi5ncF0pKSwgMjApLCBmbXRDb21wYWN0KTsKICByZW5kZXJDaGFydCgicG9jUmZxV29uIiwgewog"
    "ICAgdHlwZTogImJhciIsIGRhdGE6IHsgbGFiZWxzOiBhcnIubWFwKGIgPT4gYi5wb2MpLCBkYXRhc2V0czogWwogICAgICB7IGxhYmVsOiAiUkZRcyIsIGRh"
    "dGE6IGFyci5tYXAoYiA9PiBiLnJmcXMpLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5tdXRlZCwgYm9yZGVyUmFkaXVzOiAzIH0sCiAgICAgIHsgbGFi"
    "ZWw6ICJXb24iLCBkYXRhOiBhcnIubWFwKGIgPT4gYi53b24pLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5nb29kLCBib3JkZXJSYWRpdXM6IDMgfSwK"
    "ICAgIF0gfSwgb3B0aW9uczogeyBzY2FsZXM6IHsgeTogeyBiZWdpbkF0WmVybzogdHJ1ZSB9IH0gfSwKICB9KTsKICBoYmFyKCJwb2NXaW4iLCBhcnIubWFw"
    "KGIgPT4gW2IucG9jLCArYi53aW5SYXRlLnRvRml4ZWQoMSldKS5zb3J0KChhLCBiKSA9PiBiWzFdIC0gYVsxXSksIHYgPT4gdiArICIlIik7CiAgaGJhcigi"
    "cG9jUmV2IiwgdG9wTihuZXcgTWFwKGFyci5tYXAoYiA9PiBbYi5wb2MsIGIucmV2XSkpLCAyMCksIGZtdENvbXBhY3QpOwogIGhiYXIoInBvY1Jlc3AiLCBh"
    "cnIuZmlsdGVyKGIgPT4gYi5yZXNwQXZnICE9IG51bGwpLm1hcChiID0+IFtiLnBvYywgK2IucmVzcEF2Zy50b0ZpeGVkKDEpXSkuc29ydCgoYSwgYikgPT4g"
    "YVsxXSAtIGJbMV0pLCB2ID0+IE1hdGgucm91bmQodikgKyAiIGQiKTsKCiAgbWFrZVRhYmxlKCQoIiNwb2NUYWJsZSIpLCB7CiAgICByb3dzOiBhcnIsIG5h"
    "bWU6ICJwb2Nfc2NvcmVjYXJkIiwgcGVyUGFnZTogMjUsCiAgICBjb2x1bW5zOiBbCiAgICAgIHsga2V5OiAicmFuayIsIGxhYmVsOiAiIyIsIHR5cGU6ICJu"
    "dW0iIH0sIHsga2V5OiAicG9jIiwgbGFiZWw6ICJFVCBQT0MiIH0sCiAgICAgIHsga2V5OiAicmZxcyIsIGxhYmVsOiAiUkZRcyIsIHR5cGU6ICJudW0iIH0s"
    "IHsga2V5OiAicXVvdGVkIiwgbGFiZWw6ICJRdW90ZWQiLCB0eXBlOiAibnVtIiB9LAogICAgICB7IGtleTogIndvbiIsIGxhYmVsOiAiV29uIiwgdHlwZTog"
    "Im51bSIgfSwgeyBrZXk6ICJsb3N0IiwgbGFiZWw6ICJMb3N0IiwgdHlwZTogIm51bSIgfSwKICAgICAgeyBrZXk6ICJ3aW5SYXRlIiwgbGFiZWw6ICJXaW4g"
    "JSIsIHR5cGU6ICJwY3QiIH0sIHsga2V5OiAiY29udiIsIGxhYmVsOiAiQ29udiAlIiwgdHlwZTogInBjdCIgfSwKICAgICAgeyBrZXk6ICJuUE8iLCBsYWJl"
    "bDogIlBPcyIsIHR5cGU6ICJudW0iIH0sIHsga2V5OiAicmV2IiwgbGFiZWw6ICJRdW90ZWQgVmFsIiwgdHlwZTogImN1ciIgfSwKICAgICAgeyBrZXk6ICJn"
    "cCIsIGxhYmVsOiAiR3Jvc3MgUHJvZml0IiwgdHlwZTogImN1ciIgfSwgeyBrZXk6ICJtYXJnaW4iLCBsYWJlbDogIk1hcmdpbiIsIHR5cGU6ICJwY3QiIH0s"
    "CiAgICAgIHsga2V5OiAicmVzcEF2ZyIsIGxhYmVsOiAiQXZnIFJlc3AiLCB0eXBlOiAiZGF5cyIgfSwKICAgICAgeyBrZXk6ICJzY29yZSIsIGxhYmVsOiAi"
    "U2NvcmUiLCB0eXBlOiAibnVtIiwgcmVuZGVyOiB2ID0+IGA8c3BhbiBjbGFzcz0icGlsbCAke3YgPj0gNzAgPyAnZycgOiB2ID49IDUwID8gJ2InIDogdiA+"
    "PSAzMCA/ICdhJyA6ICdyJ30iPiR7TWF0aC5yb3VuZCh2KX08L3NwYW4+YCB9LAogICAgICB7IGtleTogInF1YXJ0aWxlIiwgbGFiZWw6ICJRdWFydGlsZSIs"
    "IHJlbmRlcjogdiA9PiAiUSIgKyB2IH0sCiAgICBdLAogIH0pOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PT09PT09CiAgIFBBUlQgNCDigJQgQ29tcGFyZSAvIEN1c3RvbWVyIC8gUmlzayAvIERhdGEtUXVhbGl0eSB0YWJzCiAgID09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KCi8qIHBlcmlvZCBtZXRyaWMgYmxvY2sgc2hhcmVk"
    "IGJ5IGNvbXBhcmUgJiBjdXN0b21lciAqLwpmdW5jdGlvbiBwZXJpb2RNZXRyaWNzKHJvd3MpIHsKICBjb25zdCBSID0gYnVpbGRSb2xsdXAocm93cyk7CiAg"
    "Y29uc3QgbGlzdCA9IFIubGlzdDsKICBjb25zdCB3b24gPSBsaXN0LmZpbHRlcihvID0+IG8ud29uKSwgbG9zdCA9IGxpc3QuZmlsdGVyKG8gPT4gby5sb3N0"
    "KTsKICBjb25zdCBmaW4gPSB3b24ubGVuZ3RoICsgbG9zdC5sZW5ndGg7CiAgY29uc3QgcG9Sb3dzID0gcm93cy5maWx0ZXIociA9PiByLmN1c3RQb05vKTsK"
    "ICBjb25zdCByZXYgPSBhZ2dyZWdhdGVEZWR1cChyb3dzLCAicmZxS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBjb2dzID0gYWdncmVnYXRlRGVk"
    "dXAocm93cywgInJmcUtleSIsICJzdXBwbGllclRvdGFsUHJpY2UiKTsKICBjb25zdCBncCA9IGFnZ3JlZ2F0ZURlZHVwKHJvd3MsICJyZnFLZXkiLCAiZ3Jv"
    "c3NQcm9maXRDYWxjIik7CiAgY29uc3QgblBPID0gdW5pcXVlQ291bnQocG9Sb3dzLCAiY3VzdFBvS2V5Iik7CiAgY29uc3QgcG9WYWwgPSBhZ2dyZWdhdGVE"
    "ZWR1cChwb1Jvd3MsICJjdXN0UG9LZXkiLCAiZXRRdW90ZWRWYWx1ZSIpOwogIGNvbnN0IHJlc3AgPSBhdmdGaWVsZChyb3dzLCAicmZxUmVzcG9uc2VEYXlz"
    "Iik7CiAgcmV0dXJuIHsKICAgIHVuaXF1ZVJGUXM6IGxpc3QubGVuZ3RoLCBxdW90ZWQ6IGxpc3QuZmlsdGVyKG8gPT4gby5xdW90ZWQpLmxlbmd0aCwgd29u"
    "OiB3b24ubGVuZ3RoLCBsb3N0OiBsb3N0Lmxlbmd0aCwKICAgIHdpblJhdGU6IGZpbiA/IHdvbi5sZW5ndGggLyBmaW4gKiAxMDAgOiAwLCBuUE8sIHJldiwg"
    "cG9WYWwsIGNvZ3MsIGdwLAogICAgbWFyZ2luOiByZXYgPyBncCAvIHJldiAqIDEwMCA6IDAsIGF2Z1JGUTogc2FmZURpdihyZXYsIGxpc3QubGVuZ3RoKSwg"
    "YXZnUE86IHNhZmVEaXYocG9WYWwsIG5QTyksCiAgICBhdmdHcFBPOiBzYWZlRGl2KGFnZ3JlZ2F0ZURlZHVwKHBvUm93cywgImN1c3RQb0tleSIsICJncm9z"
    "c1Byb2ZpdENhbGMiKSwgblBPKSwgcmVzcCwKICB9Owp9Ci8vIGZhdm91cmFibGUgZGlyZWN0aW9uOiArMSBoaWdoZXIgYmV0dGVyLCAtMSBsb3dlciBiZXR0"
    "ZXIKY29uc3QgQ09NUEFSRV9ST1dTID0gWwogIFsiVW5pcXVlIFJGUXMiLCAidW5pcXVlUkZRcyIsICJudW0iLCAxXSwgWyJRdW90ZWQgUkZRcyIsICJxdW90"
    "ZWQiLCAibnVtIiwgMV0sCiAgWyJXb24gUkZRcyIsICJ3b24iLCAibnVtIiwgMV0sIFsiTG9zdCBSRlFzIiwgImxvc3QiLCAibnVtIiwgLTFdLAogIFsiUkZR"
    "IFdpbiBSYXRlIiwgIndpblJhdGUiLCAicGN0IiwgMV0sIFsiQ3VzdG9tZXIgUE9zIiwgIm5QTyIsICJudW0iLCAxXSwKICBbIkVUIFF1b3RlZCBWYWx1ZSIs"
    "ICJyZXYiLCAiY3VyIiwgMV0sIFsiQ3VzdG9tZXIgUE8gVmFsdWUiLCAicG9WYWwiLCAiY3VyIiwgMV0sCiAgWyJTdXBwbGllciBDb3N0IiwgImNvZ3MiLCAi"
    "Y3VyIiwgMF0sIFsiR3Jvc3MgUHJvZml0IiwgImdwIiwgImN1ciIsIDFdLAogIFsiR3Jvc3MgTWFyZ2luICUiLCAibWFyZ2luIiwgInBjdCIsIDFdLCBbIkF2"
    "ZyBSRlEgVmFsdWUiLCAiYXZnUkZRIiwgImN1ciIsIDFdLAogIFsiQXZnIFBPIFZhbHVlIiwgImF2Z1BPIiwgImN1ciIsIDFdLCBbIkF2ZyBHUCAvIFBPIiwg"
    "ImF2Z0dwUE8iLCAiY3VyIiwgMV0sCiAgWyJBdmcgUkZRIFJlc3BvbnNlIChkYXlzKSIsICJyZXNwIiwgImRheXMiLCAtMV0sCl07CmZ1bmN0aW9uIGZtdEJ5"
    "VHlwZSh2LCB0KSB7IHJldHVybiB0ID09PSAiY3VyIiA/IGZtdEN1cih2KSA6IHQgPT09ICJwY3QiID8gZm10UGN0KHYpIDogdCA9PT0gImRheXMiID8gZm10"
    "RGF5cyh2KSA6IGZtdE51bShNYXRoLnJvdW5kKHYpKTsgfQpmdW5jdGlvbiB2YXJpYW5jZUNlbGwoY3VyLCBwcmV2LCBmYXZvdXIsIHR5cGUpIHsKICBpZiAo"
    "cHJldiA9PSBudWxsIHx8IGN1ciA9PSBudWxsKSByZXR1cm4gYDx0ZCBjbGFzcz0ibnVtIj7igJQ8L3RkPjx0ZCBjbGFzcz0ibnVtIj7igJQ8L3RkPmA7CiAg"
    "Y29uc3QgYWJzID0gY3VyIC0gcHJldjsKICBjb25zdCBwY3QgPSBwcmV2ID8gYWJzIC8gTWF0aC5hYnMocHJldikgKiAxMDAgOiBudWxsOwogIGxldCBjbHMg"
    "PSAibmV1dHJhbCI7CiAgaWYgKGZhdm91ciAhPT0gMCAmJiBNYXRoLmFicyhhYnMpID4gMWUtOSkgY2xzID0gKGFicyA+IDAgPyBmYXZvdXIgPiAwIDogZmF2"
    "b3VyIDwgMCkgPyAidXAtZ29vZCIgOiAiZG93bi1iYWQiOwogIGNvbnN0IGFycm93ID0gYWJzID4gMCA/ICLilrIiIDogYWJzIDwgMCA/ICLilrwiIDogIuKG"
    "kiI7CiAgY29uc3QgYWJzU3RyID0gKGFicyA+IDAgPyAiKyIgOiAiIikgKyBmbXRCeVR5cGUoYWJzLCB0eXBlKTsKICByZXR1cm4gYDx0ZCBjbGFzcz0ibnVt"
    "ICR7Y2xzfSI+JHthcnJvd30gJHthYnNTdHJ9PC90ZD48dGQgY2xhc3M9Im51bSAke2Nsc30iPiR7cGN0ID09IG51bGwgPyAi4oCUIiA6IChwY3QgPiAwID8g"
    "IisiIDogIiIpICsgcGN0LnRvRml4ZWQoMSkgKyAiJSJ9PC90ZD5gOwp9CgpmdW5jdGlvbiByZW5kZXJDb21wYXJlKHJvd3MpIHsKICBjb25zdCBlbCA9ICQo"
    "IiN0YWItY29tcGFyZSIpOwogIGNvbnN0IHllYXJzID0gQXJyYXkuZnJvbShuZXcgU2V0KFJBVy5tYXAociA9PiByLl9xeSkuZmlsdGVyKEJvb2xlYW4pKSku"
    "c29ydCgpOwogIGlmICh3aW5kb3cuX19jbXBZZWFyQ3VyID09IG51bGwpIHsgd2luZG93Ll9fY21wWWVhckN1ciA9IHllYXJzW3llYXJzLmxlbmd0aCAtIDFd"
    "OyB3aW5kb3cuX19jbXBZZWFyUHJldiA9IHllYXJzW3llYXJzLmxlbmd0aCAtIDJdIHx8IHllYXJzW3llYXJzLmxlbmd0aCAtIDFdOyB9CiAgY29uc3QgY3Vy"
    "WSA9IHdpbmRvdy5fX2NtcFllYXJDdXIsIHByZXZZID0gd2luZG93Ll9fY21wWWVhclByZXY7CiAgY29uc3QgY3VyUm93cyA9IHJvd3MuZmlsdGVyKHIgPT4g"
    "ci5fcXkgPT09ICtjdXJZKTsKICBjb25zdCBwcmV2Um93cyA9IHJvd3MuZmlsdGVyKHIgPT4gci5fcXkgPT09ICtwcmV2WSk7CiAgY29uc3QgY3VyID0gcGVy"
    "aW9kTWV0cmljcyhjdXJSb3dzKSwgcHJldiA9IHBlcmlvZE1ldHJpY3MocHJldlJvd3MpOwoKICBlbC5pbm5lckhUTUwgPSBgCiAgICA8ZGl2IGNsYXNzPSJz"
    "ZWN0aW9uLXRpdGxlIj5QZXJpb2QgQ29tcGFyaXNvbiAoYnkgRVQgUXVvdGF0aW9uIERhdGUpPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjbXAtY29udHJvbHMi"
    "PgogICAgICA8bGFiZWw+Q3VycmVudCB5ZWFyCiAgICAgICAgPHNlbGVjdCBpZD0iY21wQ3VyIj4ke3llYXJzLm1hcCh5ID0+IGA8b3B0aW9uICR7eSA9PSBj"
    "dXJZID8gInNlbGVjdGVkIiA6ICIifT4ke3l9PC9vcHRpb24+YCkuam9pbigiIil9PC9zZWxlY3Q+CiAgICAgIDwvbGFiZWw+CiAgICAgIDxsYWJlbD5Db21w"
    "YXJlIHRvCiAgICAgICAgPHNlbGVjdCBpZD0iY21wUHJldiI+JHt5ZWFycy5tYXAoeSA9PiBgPG9wdGlvbiAke3kgPT0gcHJldlkgPyAic2VsZWN0ZWQiIDog"
    "IiJ9PiR7eX08L29wdGlvbj5gKS5qb2luKCIiKX08L3NlbGVjdD4KICAgICAgPC9sYWJlbD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY2FyZCBjMTIi"
    "IHN0eWxlPSJvdmVyZmxvdzphdXRvIj4KICAgICAgPHRhYmxlIGNsYXNzPSJkYXRhIGNtcC10YWJsZSI+CiAgICAgICAgPHRoZWFkPjx0cj48dGg+TWV0cmlj"
    "PC90aD48dGggY2xhc3M9Im51bSI+JHtlc2MoY3VyWSl9PC90aD48dGggY2xhc3M9Im51bSI+JHtlc2MocHJldlkpfTwvdGg+PHRoIGNsYXNzPSJudW0iPlZh"
    "cmlhbmNlPC90aD48dGggY2xhc3M9Im51bSI+JTwvdGg+PC90cj48L3RoZWFkPgogICAgICAgIDx0Ym9keT4KICAgICAgICAgICR7Q09NUEFSRV9ST1dTLm1h"
    "cCgoW2xhYmVsLCBrZXksIHR5cGUsIGZhdl0pID0+IGAKICAgICAgICAgICAgPHRyPjx0ZD4ke2xhYmVsfTwvdGQ+CiAgICAgICAgICAgICAgPHRkIGNsYXNz"
    "PSJudW0iPiR7Zm10QnlUeXBlKGN1cltrZXldLCB0eXBlKX08L3RkPgogICAgICAgICAgICAgIDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEJ5VHlwZShwcmV2W2tl"
    "eV0sIHR5cGUpfTwvdGQ+CiAgICAgICAgICAgICAgJHt2YXJpYW5jZUNlbGwoY3VyW2tleV0sIHByZXZba2V5XSwgZmF2LCB0eXBlKX0KICAgICAgICAgICAg"
    "PC90cj5gKS5qb2luKCIiKX0KICAgICAgICA8L3Rib2R5PgogICAgICA8L3RhYmxlPgogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRs"
    "ZSI+TW9udGhseSBZb1k8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXJ0LWdyaWQiPgogICAgICAke2NhcmRTaGVsbChgTW9udGhseSBFVCBRdW90ZWQgVmFs"
    "dWUg4oCUICR7ZXNjKGN1clkpfSB2cyAke2VzYyhwcmV2WSl9YCwgImNtcFJldiIsICJjNiIpfQogICAgICAke2NhcmRTaGVsbChgTW9udGhseSBHcm9zcyBQ"
    "cm9maXQg4oCUICR7ZXNjKGN1clkpfSB2cyAke2VzYyhwcmV2WSl9YCwgImNtcEdwIiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKGBNb250aGx5IFVuaXF1"
    "ZSBSRlFzIOKAlCAke2VzYyhjdXJZKX0gdnMgJHtlc2MocHJldlkpfWAsICJjbXBSZnEiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoYE1vbnRobHkgV29u"
    "IFJGUXMg4oCUICR7ZXNjKGN1clkpfSB2cyAke2VzYyhwcmV2WSl9YCwgImNtcFdvbiIsICJjNiIpfQogICAgPC9kaXY+YDsKCiAgJCgiI2NtcEN1ciIpLm9u"
    "Y2hhbmdlID0gKGUpID0+IHsgd2luZG93Ll9fY21wWWVhckN1ciA9IGUudGFyZ2V0LnZhbHVlOyBSRU5ERVJFRC5jb21wYXJlID0gZmFsc2U7IHJlbmRlckNv"
    "bXBhcmUoYXBwbHlGaWx0ZXJzKCkpOyB9OwogICQoIiNjbXBQcmV2Iikub25jaGFuZ2UgPSAoZSkgPT4geyB3aW5kb3cuX19jbXBZZWFyUHJldiA9IGUudGFy"
    "Z2V0LnZhbHVlOyBSRU5ERVJFRC5jb21wYXJlID0gZmFsc2U7IHJlbmRlckNvbXBhcmUoYXBwbHlGaWx0ZXJzKCkpOyB9OwoKICBjb25zdCBtYyA9IG1vbnRo"
    "bHlCeVllYXIoY3VyUm93cyksIG1wID0gbW9udGhseUJ5WWVhcihwcmV2Um93cyk7CiAgY29uc3QgZHVhbCA9IChpZCwgZmllbGQsIGZtdCkgPT4gcmVuZGVy"
    "Q2hhcnQoaWQsIHsKICAgIHR5cGU6ICJiYXIiLCBkYXRhOiB7IGxhYmVsczogTU9OVEhTLCBkYXRhc2V0czogWwogICAgICB7IGxhYmVsOiBTdHJpbmcocHJl"
    "dlkpLCBkYXRhOiBNT05USFMubWFwKChfLCBpKSA9PiBtcFtmaWVsZF1baV0pLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5tdXRlZCwgYm9yZGVyUmFk"
    "aXVzOiAzIH0sCiAgICAgIHsgbGFiZWw6IFN0cmluZyhjdXJZKSwgZGF0YTogTU9OVEhTLm1hcCgoXywgaSkgPT4gbWNbZmllbGRdW2ldKSwgYmFja2dyb3Vu"
    "ZENvbG9yOiBwYWxldHRlKCkucHJpbWFyeSwgYm9yZGVyUmFkaXVzOiAzIH0sCiAgICBdIH0sIG9wdGlvbnM6IHsgc2NhbGVzOiB7IHk6IHsgYmVnaW5BdFpl"
    "cm86IHRydWUsIHRpY2tzOiB7IGNhbGxiYWNrOiB2ID0+IGZtdCA/IGZtdCh2KSA6IHYgfSB9IH0gfSwKICB9KTsKICBkdWFsKCJjbXBSZXYiLCAicmV2Iiwg"
    "Zm10Q29tcGFjdCk7IGR1YWwoImNtcEdwIiwgImdwIiwgZm10Q29tcGFjdCk7CiAgZHVhbCgiY21wUmZxIiwgInJmcSIsIG51bGwpOyBkdWFsKCJjbXBXb24i"
    "LCAid29uIiwgbnVsbCk7Cn0KZnVuY3Rpb24gbW9udGhseUJ5WWVhcihyb3dzKSB7CiAgY29uc3QgcmV2ID0gZ3JvdXBEZWR1cChyb3dzLCAicmZxS2V5Iiwg"
    "ImV0UXVvdGVkVmFsdWUiKSwgZ3AgPSBncm91cERlZHVwKHJvd3MsICJyZnFLZXkiLCAiZ3Jvc3NQcm9maXRDYWxjIik7CiAgY29uc3QgcmZxID0gTU9OVEhT"
    "Lm1hcCgoKSA9PiBuZXcgU2V0KCkpLCB3b24gPSBNT05USFMubWFwKCgpID0+IG5ldyBTZXQoKSk7CiAgY29uc3QgcnYgPSBNT05USFMubWFwKCgpID0+IDAp"
    "LCBncHYgPSBNT05USFMubWFwKCgpID0+IDApOwogIGNvbnN0IHNlZW4gPSBuZXcgU2V0KCk7CiAgZm9yIChjb25zdCByIG9mIHJvd3MpIHsKICAgIGNvbnN0"
    "IG1pID0gci5fcW0gPyByLl9xbSAtIDEgOiAoci5ldFF1b3RlRGF0ZSA/ICtyLmV0UXVvdGVEYXRlLnNsaWNlKDUsIDcpIC0gMSA6IG51bGwpOwogICAgaWYg"
    "KG1pID09IG51bGwpIGNvbnRpbnVlOwogICAgcmZxW21pXS5hZGQoci5yZnFLZXkpOwogICAgaWYgKHIuZXRRdW90ZVN0YXR1cyA9PT0gIldvbiIpIHdvbltt"
    "aV0uYWRkKHIucmZxS2V5KTsKICAgIGlmICghc2Vlbi5oYXMoci5yZnFLZXkpKSB7IHNlZW4uYWRkKHIucmZxS2V5KTsgcnZbbWldICs9IHJldi5nZXQoci5y"
    "ZnFLZXkpIHx8IDA7IGdwdlttaV0gKz0gZ3AuZ2V0KHIucmZxS2V5KSB8fCAwOyB9CiAgfQogIHJldHVybiB7IHJldjogcnYsIGdwOiBncHYsIHJmcTogcmZx"
    "Lm1hcChzID0+IHMuc2l6ZSksIHdvbjogd29uLm1hcChzID0+IHMuc2l6ZSkgfTsKfQoKLyogPT09PT09PT09PT09PT09PT0gVEFCIDcg4oCUIENVU1RPTUVS"
    "IEFOQUxZU0lTID09PT09PT09PT09PT09PT09ICovCmZ1bmN0aW9uIHJlbmRlckN1c3RvbWVyKHJvd3MpIHsKICBjb25zdCBlbCA9ICQoIiN0YWItY3VzdG9t"
    "ZXIiKTsKICBjb25zdCBjdXN0b21lcnMgPSAoTUVUQS5kaXN0aW5jdCAmJiBNRVRBLmRpc3RpbmN0LmN1c3RvbWVycykgfHwgW107CiAgaWYgKHdpbmRvdy5f"
    "X3NlbEN1c3RvbWVyID09IG51bGwpIHsKICAgIC8vIGRlZmF1bHQgdG8gdG9wIGN1c3RvbWVyIGJ5IFBPIHZhbHVlIGluIGN1cnJlbnQgc2VsZWN0aW9uCiAg"
    "ICBjb25zdCB0b3AgPSB0b3BOKGRlZHVwTWFwQnkocm93cy5maWx0ZXIociA9PiByLmN1c3RQb05vKSwgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIiwg"
    "ImN1c3RvbWVyIiksIDEpWzBdOwogICAgd2luZG93Ll9fc2VsQ3VzdG9tZXIgPSB0b3AgPyB0b3BbMF0gOiBjdXN0b21lcnNbMF07CiAgfQogIGNvbnN0IHNl"
    "bCA9IHdpbmRvdy5fX3NlbEN1c3RvbWVyOwogIGNvbnN0IGNSb3dzID0gcm93cy5maWx0ZXIociA9PiByLmN1c3RvbWVyID09PSBzZWwpOwogIGNvbnN0IHll"
    "YXJzID0gQXJyYXkuZnJvbShuZXcgU2V0KGNSb3dzLm1hcChyID0+IHIuX3F5KS5maWx0ZXIoQm9vbGVhbikpKS5zb3J0KCk7CiAgY29uc3QgbGFzdFllYXJz"
    "ID0geWVhcnMuc2xpY2UoLTQpOwoKICAvLyB0b3RhbHMgZm9yIGNvbnRyaWJ1dGlvbgogIGNvbnN0IHRvdFBPID0gYWdncmVnYXRlRGVkdXAocm93cy5maWx0"
    "ZXIociA9PiByLmN1c3RQb05vKSwgImN1c3RQb0tleSIsICJldFF1b3RlZFZhbHVlIik7CiAgY29uc3QgdG90R1AgPSBhZ2dyZWdhdGVEZWR1cChyb3dzLCAi"
    "cmZxS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpOwogIGNvbnN0IGNQTyA9IGFnZ3JlZ2F0ZURlZHVwKGNSb3dzLmZpbHRlcihyID0+IHIuY3VzdFBvTm8pLCAi"
    "Y3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBjR1AgPSBhZ2dyZWdhdGVEZWR1cChjUm93cywgInJmcUtleSIsICJncm9zc1Byb2ZpdENh"
    "bGMiKTsKCiAgY29uc3QgcGVyWWVhciA9IGxhc3RZZWFycy5tYXAoeSA9PiB7CiAgICBjb25zdCB5ciA9IGNSb3dzLmZpbHRlcihyID0+IHIuX3F5ID09PSB5"
    "KTsKICAgIHJldHVybiB7IHllYXI6IHksIC4uLnBlcmlvZE1ldHJpY3MoeXIpIH07CiAgfSk7CgogIGVsLmlubmVySFRNTCA9IGAKICAgIDxkaXYgY2xhc3M9"
    "InNlY3Rpb24tdGl0bGUiPkN1c3RvbWVyIEFuYWx5c2lzPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJjbXAtY29udHJvbHMiPgogICAgICA8bGFiZWw+Q3VzdG9t"
    "ZXIKICAgICAgICA8c2VsZWN0IGlkPSJjdXN0U2VsIj4ke2N1c3RvbWVycy5tYXAoYyA9PiBgPG9wdGlvbiAke2MgPT09IHNlbCA/ICJzZWxlY3RlZCIgOiAi"
    "In0+JHtlc2MoYyl9PC9vcHRpb24+YCkuam9pbigiIil9PC9zZWxlY3Q+CiAgICAgIDwvbGFiZWw+CiAgICAgIDxzcGFuIGNsYXNzPSJwby1zZWFyY2gtbm90"
    "ZSI+JHtsYXN0WWVhcnMubGVuZ3RoID8gIlNob3dpbmcgIiArIGxhc3RZZWFyc1swXSArICLigJMiICsgbGFzdFllYXJzW2xhc3RZZWFycy5sZW5ndGggLSAx"
    "XSA6ICJObyBkYXRlZCBhY3Rpdml0eSJ9PC9zcGFuPgogICAgPC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCI+CiAgICAgICR7a3BpKCJQTyBWYWx1"
    "ZSAoc2VsZWN0aW9uKSIsIGZtdENvbXBhY3QoY1BPKSwgZm10Q3VyKGNQTykpfQogICAgICAke2twaSgiR3Jvc3MgUHJvZml0IiwgZm10Q29tcGFjdChjR1Ap"
    "KX0KICAgICAgJHtrcGkoIkdyb3NzIE1hcmdpbiAlIiwgZm10UGN0KGNQTyA/IGNHUCAvIGNQTyAqIDEwMCA6IChhZ2dyZWdhdGVEZWR1cChjUm93cywgInJm"
    "cUtleSIsICJldFF1b3RlZFZhbHVlIikgPyBjR1AgLyBhZ2dyZWdhdGVEZWR1cChjUm93cywgInJmcUtleSIsICJldFF1b3RlZFZhbHVlIikgKiAxMDAgOiAw"
    "KSkpfQogICAgICAke2twaSgiUmV2ZW51ZSBDb250cmlidXRpb24iLCBmbXRQY3QodG90UE8gPyBjUE8gLyB0b3RQTyAqIDEwMCA6IDApLCAib2YgYWxsIFBP"
    "IHZhbHVlIil9CiAgICAgICR7a3BpKCJHUCBDb250cmlidXRpb24iLCBmbXRQY3QodG90R1AgPyBjR1AgLyB0b3RHUCAqIDEwMCA6IDApLCAib2YgYWxsIGdy"
    "b3NzIHByb2ZpdCIpfQogICAgICAke2twaSgiVW5pcXVlIFJGUXMiLCBmbXROdW0odW5pcXVlQ291bnQoY1Jvd3MsICJyZnFLZXkiKSkpfQogICAgICAke2tw"
    "aSgiQ3VzdG9tZXIgUE9zIiwgZm10TnVtKHVuaXF1ZUNvdW50KGNSb3dzLmZpbHRlcihyID0+IHIuY3VzdFBvTm8pLCAiY3VzdFBvS2V5IikpKX0KICAgICAg"
    "JHtrcGkoIlN1cHBsaWVycyBVc2VkIiwgZm10TnVtKGRpc3RpbmN0QnkoY1Jvd3MsICgpID0+IHRydWUsICJzdXBwbGllck5hbWUiKSkpfQogICAgPC9kaXY+"
    "CgogICAgJHtsYXN0WWVhcnMubGVuZ3RoID49IDIgPyBgCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5ZZWFyLW9uLVllYXIgKGxhdGVzdCB2cyBw"
    "cmV2aW91cyk8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNhcmQgYzEyIiBzdHlsZT0ib3ZlcmZsb3c6YXV0byI+CiAgICAgIDx0YWJsZSBjbGFzcz0iZGF0YSBj"
    "bXAtdGFibGUiPjx0aGVhZD48dHI+PHRoPk1ldHJpYzwvdGg+CiAgICAgICAgJHtsYXN0WWVhcnMubWFwKHkgPT4gYDx0aCBjbGFzcz0ibnVtIj4ke3l9PC90"
    "aD5gKS5qb2luKCIiKX0KICAgICAgICA8dGggY2xhc3M9Im51bSI+zpQgbGF0ZXN0PC90aD48dGggY2xhc3M9Im51bSI+JTwvdGg+PC90cj48L3RoZWFkPjx0"
    "Ym9keT4KICAgICAgICAke0NPTVBBUkVfUk9XUy5tYXAoKFtsYWJlbCwga2V5LCB0eXBlLCBmYXZdKSA9PiB7CiAgICAgICAgICBjb25zdCB2YWxzID0gcGVy"
    "WWVhci5tYXAocCA9PiBwW2tleV0pOwogICAgICAgICAgY29uc3QgY3VyID0gdmFsc1t2YWxzLmxlbmd0aCAtIDFdLCBwcmV2ID0gdmFsc1t2YWxzLmxlbmd0"
    "aCAtIDJdOwogICAgICAgICAgcmV0dXJuIGA8dHI+PHRkPiR7bGFiZWx9PC90ZD4ke3ZhbHMubWFwKHYgPT4gYDx0ZCBjbGFzcz0ibnVtIj4ke2ZtdEJ5VHlw"
    "ZSh2LCB0eXBlKX08L3RkPmApLmpvaW4oIiIpfSR7dmFyaWFuY2VDZWxsKGN1ciwgcHJldiwgZmF2LCB0eXBlKX08L3RyPmA7CiAgICAgICAgfSkuam9pbigi"
    "Iil9CiAgICAgIDwvdGJvZHk+PC90YWJsZT4KICAgIDwvZGl2PmAgOiAiIn0KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5UcmVuZHMgJiBNaXg8"
    "L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoYXJ0LWdyaWQiPgogICAgICAke2NhcmRTaGVsbCgiWWVhcmx5IFBPIFZhbHVlICYgR3Jvc3MgUHJvZml0IiwgImN1"
    "c3RZZWFybHkiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlllYXJseSBHcm9zcyBNYXJnaW4gJSIsICJjdXN0TWFyZ2luIiwgImM2Iil9CiAgICAgICR7"
    "Y2FyZFNoZWxsKCJNb250aGx5IFBPIFZhbHVlIiwgImN1c3RNb250aGx5IiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJQcm9kdWN0IENhdGVnb3J5IE1p"
    "eCIsICJjdXN0Q2F0IiwgImM2Iil9CiAgICAgICR7Y2FyZFNoZWxsKCJTdXBwbGllciBNaXgiLCAiY3VzdFN1cCIsICJjNiIpfQogICAgICAke2NhcmRTaGVs"
    "bCgiRVQgUE9DIE1peCIsICJjdXN0UG9jIiwgImM2Iil9CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5DdXN0b21lciBJbnNp"
    "Z2h0czwvZGl2PgogICAgPGRpdiBpZD0iY3VzdEluc2lnaHRzIj48L2Rpdj5gOwoKICAkKCIjY3VzdFNlbCIpLm9uY2hhbmdlID0gKGUpID0+IHsgd2luZG93"
    "Ll9fc2VsQ3VzdG9tZXIgPSBlLnRhcmdldC52YWx1ZTsgUkVOREVSRUQuY3VzdG9tZXIgPSBmYWxzZTsgcmVuZGVyQ3VzdG9tZXIoYXBwbHlGaWx0ZXJzKCkp"
    "OyB9OwoKICBpZiAoIWxhc3RZZWFycy5sZW5ndGgpIHsgJCgiI2N1c3RJbnNpZ2h0cyIpLmlubmVySFRNTCA9IGA8ZGl2IGNsYXNzPSJlbXB0eS1zdGF0ZSI+"
    "Tm8gZGF0ZWQgYWN0aXZpdHkgZm9yICR7ZXNjKHNlbCl9IGluIHRoZSBjdXJyZW50IHNlbGVjdGlvbi48L2Rpdj5gOyByZXR1cm47IH0KCiAgcmVuZGVyQ2hh"
    "cnQoImN1c3RZZWFybHkiLCB7CiAgICB0eXBlOiAiYmFyIiwgZGF0YTogeyBsYWJlbHM6IGxhc3RZZWFycywgZGF0YXNldHM6IFsKICAgICAgeyBsYWJlbDog"
    "IlBPIFZhbHVlIiwgZGF0YTogcGVyWWVhci5tYXAocCA9PiBwLnBvVmFsKSwgYmFja2dyb3VuZENvbG9yOiBwYWxldHRlKCkucHJpbWFyeSwgYm9yZGVyUmFk"
    "aXVzOiAzLCB5QXhpc0lEOiAieSIgfSwKICAgICAgeyBsYWJlbDogIkdyb3NzIFByb2ZpdCIsIHR5cGU6ICJsaW5lIiwgZGF0YTogcGVyWWVhci5tYXAocCA9"
    "PiBwLmdwKSwgYm9yZGVyQ29sb3I6IHBhbGV0dGUoKS5nb29kLCBiYWNrZ3JvdW5kQ29sb3I6IHBhbGV0dGUoKS5nb29kLCB0ZW5zaW9uOiAuMywgeUF4aXNJ"
    "RDogInkiIH0sCiAgICBdIH0sIG9wdGlvbnM6IHsgc2NhbGVzOiB7IHk6IHsgYmVnaW5BdFplcm86IHRydWUsIHRpY2tzOiB7IGNhbGxiYWNrOiB2ID0+IGZt"
    "dENvbXBhY3QodikgfSB9IH0gfSwKICB9KTsKICByZW5kZXJDaGFydCgiY3VzdE1hcmdpbiIsIHsKICAgIHR5cGU6ICJsaW5lIiwgZGF0YTogeyBsYWJlbHM6"
    "IGxhc3RZZWFycywgZGF0YXNldHM6IFt7IGxhYmVsOiAiTWFyZ2luICUiLCBkYXRhOiBwZXJZZWFyLm1hcChwID0+ICtwLm1hcmdpbi50b0ZpeGVkKDEpKSwg"
    "Ym9yZGVyQ29sb3I6IHBhbGV0dGUoKS5hY2NlbnQsIHRlbnNpb246IC4zIH1dIH0sCiAgICBvcHRpb25zOiB7IHBsdWdpbnM6IHsgbGVnZW5kOiB7IGRpc3Bs"
    "YXk6IGZhbHNlIH0gfSwgc2NhbGVzOiB7IHk6IHsgdGlja3M6IHsgY2FsbGJhY2s6IHYgPT4gdiArICIlIiB9IH0gfSB9LAogIH0pOwogIC8vIG1vbnRobHkg"
    "UE8gdmFsdWUgZm9yIHNlbGVjdGVkIGN1c3RvbWVyIChjdXN0UG9EYXRlKQogIGNvbnN0IGNtID0gcG9Nb250aGx5KGNSb3dzLmZpbHRlcihyID0+IHIuY3Vz"
    "dFBvTm8pKTsKICByZW5kZXJDaGFydCgiY3VzdE1vbnRobHkiLCB7CiAgICB0eXBlOiAiYmFyIiwgZGF0YTogeyBsYWJlbHM6IGNtLmxhYmVscywgZGF0YXNl"
    "dHM6IFt7IGRhdGE6IGNtLnZhbCwgYmFja2dyb3VuZENvbG9yOiBwYWxldHRlKCkucHJpbWFyeSwgYm9yZGVyUmFkaXVzOiAzIH1dIH0sCiAgICBvcHRpb25z"
    "OiB7IHBsdWdpbnM6IHsgbGVnZW5kOiB7IGRpc3BsYXk6IGZhbHNlIH0gfSwgc2NhbGVzOiB7IHk6IHsgYmVnaW5BdFplcm86IHRydWUsIHRpY2tzOiB7IGNh"
    "bGxiYWNrOiB2ID0+IGZtdENvbXBhY3QodikgfSB9IH0gfSwKICB9KTsKICBkb251dCgiY3VzdENhdCIsIGRlZHVwTWFwVG9Db3VudChjUm93cywgInByb2R1"
    "Y3RDYXRlZ29yeSIpLCBudWxsKTsKICBkb251dCgiY3VzdFN1cCIsIHRvcE1hcEFzTWFwKHN1bU1hcEJ5KGNSb3dzLCAic3VwcGxpZXJOYW1lIiwgInN1cHBs"
    "aWVyVG90YWxQcmljZSIpLCA4KSwgbnVsbCk7CiAgZG9udXQoImN1c3RQb2MiLCBkZWR1cE1hcFRvQ291bnQoY1Jvd3MsICJldFBPQyIpLCBudWxsKTsKCiAg"
    "JCgiI2N1c3RJbnNpZ2h0cyIpLmlubmVySFRNTCA9IGluc2lnaHRIVE1MKGN1c3RvbWVySW5zaWdodHMoc2VsLCBjUm93cywgcGVyWWVhciwgcm93cykpOwp9"
    "CmZ1bmN0aW9uIGRlZHVwTWFwVG9Db3VudChyb3dzLCBmaWVsZCkgewogIGNvbnN0IG93bmVyID0gbmV3IE1hcCgpOyByb3dzLmZvckVhY2gociA9PiB7IGlm"
    "ICghb3duZXIuaGFzKHIucmZxS2V5KSkgb3duZXIuc2V0KHIucmZxS2V5LCByW2ZpZWxkXSB8fCAi4oCUIik7IH0pOwogIGNvbnN0IG0gPSBuZXcgTWFwKCk7"
    "IGZvciAoY29uc3QgdiBvZiBvd25lci52YWx1ZXMoKSkgbS5zZXQodiwgKG0uZ2V0KHYpIHx8IDApICsgMSk7IHJldHVybiBtOwp9CmZ1bmN0aW9uIHRvcE1h"
    "cEFzTWFwKG1hcCwgbikgeyByZXR1cm4gbmV3IE1hcCh0b3BOKG1hcCwgbikpOyB9CmZ1bmN0aW9uIGN1c3RvbWVySW5zaWdodHMobmFtZSwgY1Jvd3MsIHBl"
    "clllYXIsIGFsbFJvd3MpIHsKICBjb25zdCBvdXQgPSBbXTsKICBpZiAocGVyWWVhci5sZW5ndGggPj0gMikgewogICAgY29uc3QgY3VyID0gcGVyWWVhcltw"
    "ZXJZZWFyLmxlbmd0aCAtIDFdLCBwcmV2ID0gcGVyWWVhcltwZXJZZWFyLmxlbmd0aCAtIDJdOwogICAgaWYgKGN1ci5wb1ZhbCA+IHByZXYucG9WYWwgJiYg"
    "Y3VyLm1hcmdpbiA8IHByZXYubWFyZ2luKQogICAgICBvdXQucHVzaChbIndhcm4iLCAiTWFyZ2luIiwgYFBPIHZhbHVlIHJvc2UgdG8gJHtmbXRDb21wYWN0"
    "KGN1ci5wb1ZhbCl9IGJ1dCBncm9zcyBtYXJnaW4gc2xpcHBlZCBmcm9tICR7Zm10UGN0KHByZXYubWFyZ2luKX0gdG8gJHtmbXRQY3QoY3VyLm1hcmdpbil9"
    "LmBdKTsKICAgIGlmIChjdXIublBPID4gcHJldi5uUE8gJiYgY3VyLmF2Z1BPIDwgcHJldi5hdmdQTykKICAgICAgb3V0LnB1c2goWyJ3YXJuIiwgIlBPIHNp"
    "emUiLCBgUE8gY291bnQgaW5jcmVhc2VkICgke3ByZXYublBPfeKGkiR7Y3VyLm5QT30pIHdoaWxlIGF2ZXJhZ2UgUE8gdmFsdWUgZmVsbCB0byAke2ZtdENv"
    "bXBhY3QoY3VyLmF2Z1BPKX0uYF0pOwogICAgaWYgKGN1ci5wb1ZhbCA+IHByZXYucG9WYWwgJiYgY3VyLm1hcmdpbiA+PSBwcmV2Lm1hcmdpbikKICAgICAg"
    "b3V0LnB1c2goWyJwb3MiLCAiR3Jvd3RoIiwgYFByb2ZpdGFibGUgZ3Jvd3RoIOKAlCBQTyB2YWx1ZSBhbmQgbWFyZ2luIGJvdGggaW1wcm92ZWQgeWVhci1v"
    "bi15ZWFyLmBdKTsKICAgIGlmIChjdXIudW5pcXVlUkZRcyA+IHByZXYudW5pcXVlUkZRcyAmJiBjdXIud2luUmF0ZSA8IHByZXYud2luUmF0ZSkKICAgICAg"
    "b3V0LnB1c2goWyJ3YXJuIiwgIldpbiByYXRlIiwgYFJGUSB2b2x1bWUgaXMgdXAgYnV0IHdpbiByYXRlIHdlYWtlbmVkIGZyb20gJHtmbXRQY3QocHJldi53"
    "aW5SYXRlKX0gdG8gJHtmbXRQY3QoY3VyLndpblJhdGUpfS5gXSk7CiAgfQogIC8vIGNhdGVnb3J5IGNvbmNlbnRyYXRpb24KICBjb25zdCBjYXQgPSBkZWR1"
    "cE1hcFRvQ291bnQoY1Jvd3MsICJwcm9kdWN0Q2F0ZWdvcnkiKTsKICBjb25zdCBjYXRUb3QgPSBBcnJheS5mcm9tKGNhdC52YWx1ZXMoKSkucmVkdWNlKChh"
    "LCBiKSA9PiBhICsgYiwgMCk7CiAgY29uc3QgdG9wQ2F0ID0gdG9wTihjYXQsIDEpWzBdOwogIGlmICh0b3BDYXQgJiYgY2F0VG90ICYmIHRvcENhdFsxXSAv"
    "IGNhdFRvdCA+IDAuNikgb3V0LnB1c2goWyJ3YXJuIiwgIkNvbmNlbnRyYXRpb24iLCBgJHtmbXRQY3QodG9wQ2F0WzFdIC8gY2F0VG90ICogMTAwKX0gb2Yg"
    "JHtlc2MobmFtZSl9J3MgUkZRcyBhcmUgaW4gb25lIGNhdGVnb3J5ICgke2VzYyh0b3BDYXRbMF0pfSkuYF0pOwogIC8vIHN1cHBsaWVyIHJlbGlhbmNlCiAg"
    "Y29uc3Qgc3VwID0gc3VtTWFwQnkoY1Jvd3MsICJzdXBwbGllck5hbWUiLCAic3VwcGxpZXJUb3RhbFByaWNlIik7CiAgY29uc3Qgc3VwVG90ID0gQXJyYXku"
    "ZnJvbShzdXAudmFsdWVzKCkpLnJlZHVjZSgoYSwgYikgPT4gYSArIGIsIDApOwogIGNvbnN0IHRvcFN1cCA9IHRvcE4oc3VwLCAxKVswXTsKICBpZiAodG9w"
    "U3VwICYmIHN1cFRvdCAmJiB0b3BTdXBbMV0gLyBzdXBUb3QgPiAwLjYpIG91dC5wdXNoKFsid2FybiIsICJTdXBwbGllciByZWxpYW5jZSIsIGAke2VzYyh0"
    "b3BTdXBbMF0pfSBzdXBwbGllcyAke2ZtdFBjdCh0b3BTdXBbMV0gLyBzdXBUb3QgKiAxMDApfSBvZiB0aGlzIGN1c3RvbWVyJ3MgcHJvY3VyZW1lbnQgdmFs"
    "dWUuYF0pOwogIC8vIGluYWN0aXZpdHkKICBjb25zdCBsYXN0UG8gPSBjUm93cy5maWx0ZXIociA9PiByLmN1c3RQb0RhdGUpLm1hcChyID0+IHIuY3VzdFBv"
    "RGF0ZSkuc29ydCgpLnBvcCgpOwogIGlmIChsYXN0UG8pIHsgY29uc3QgZCA9IGRheXNCZXR3ZWVuKHRvZGF5SVNPKCksIGxhc3RQbyk7IGlmIChkICE9IG51"
    "bGwgJiYgZCA+IDEyMCkgb3V0LnB1c2goWyJjcml0IiwgIkluYWN0aXZpdHkiLCBgTm8gY3VzdG9tZXIgUE8gcmVjb3JkZWQgaW4gfiR7TWF0aC5yb3VuZChk"
    "IC8gMzApfSBtb250aHMgKGxhc3QgJHtmbXREYXRlKGxhc3RQbyl9KS5gXSk7IH0KICAvLyBxdW90ZWQgbm90IGNvbnZlcnRpbmcKICBjb25zdCBSID0gYnVp"
    "bGRSb2xsdXAoY1Jvd3MpOyBjb25zdCB3b24gPSBSLmxpc3QuZmlsdGVyKG8gPT4gby53b24pOwogIGNvbnN0IHF1b3RlZFZhbCA9IGFnZ3JlZ2F0ZURlZHVw"
    "KGNSb3dzLCAicmZxS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBjb25zdCBwb1ZhbCA9IGFnZ3JlZ2F0ZURlZHVwKGNSb3dzLmZpbHRlcihyID0+IHIuY3Vz"
    "dFBvTm8pLCAiY3VzdFBvS2V5IiwgImV0UXVvdGVkVmFsdWUiKTsKICBpZiAocXVvdGVkVmFsID4gMCAmJiBwb1ZhbCAvIHF1b3RlZFZhbCA8IDAuMTUpIG91"
    "dC5wdXNoKFsib3BwIiwgIkNvbnZlcnNpb24iLCBgSGlnaCBxdW90ZWQgdmFsdWUgKCR7Zm10Q29tcGFjdChxdW90ZWRWYWwpfSkgYnV0IG9ubHkgJHtmbXRD"
    "b21wYWN0KHBvVmFsKX0gY29udmVydGVkIHRvIFBPIOKAlCBwdXJzdWUgb3BlbiBxdW90ZXMuYF0pOwogIGNvbnN0IHdvbk5vUG8gPSB3b24uZmlsdGVyKG8g"
    "PT4gIW8uaGFzUE8pOwogIGlmICh3b25Ob1BvLmxlbmd0aCkgb3V0LnB1c2goWyJjcml0IiwgIlBPIGNhcHR1cmUiLCBgJHt3b25Ob1BvLmxlbmd0aH0gd29u"
    "IFJGUShzKSBmb3IgJHtlc2MobmFtZSl9IGhhdmUgbm8gY3VzdG9tZXIgUE8gY2FwdHVyZWQuYF0pOwogIGlmICghb3V0Lmxlbmd0aCkgb3V0LnB1c2goWyJw"
    "b3MiLCAiU3RhYmxlIiwgYE5vIG1hdGVyaWFsIHJpc2sgc2lnbmFscyBkZXRlY3RlZCBmb3IgJHtlc2MobmFtZSl9IGluIHRoZSBjdXJyZW50IHNlbGVjdGlv"
    "bi5gXSk7CiAgcmV0dXJuIG91dDsKfQoKLyogPT09PT09PT09PT09PT09PT0gVEFCIDgg4oCUIFJJU0sgQU5BTFlTSVMgPT09PT09PT09PT09PT09PT0gKi8K"
    "ZnVuY3Rpb24gcmVuZGVyUmlzayhyb3dzKSB7CiAgY29uc3QgZWwgPSAkKCIjdGFiLXJpc2siKTsKICBjb25zdCByaXNrcyA9IGNvbXB1dGVSaXNrcyhyb3dz"
    "KTsKICBjb25zdCBjcml0ID0gcmlza3MuZmlsdGVyKHIgPT4gci5zZXZlcml0eSA9PT0gIkNyaXRpY2FsIiksIGhpZ2ggPSByaXNrcy5maWx0ZXIociA9PiBy"
    "LnNldmVyaXR5ID09PSAiSGlnaCIpOwogIGNvbnN0IGF0Umlza1ZhbCA9IHJpc2tzLnJlZHVjZSgoYSwgcikgPT4gYSArIChyLmV4cG9zdXJlIHx8IDApLCAw"
    "KTsKICBjb25zdCBjcml0VmFsID0gY3JpdC5yZWR1Y2UoKGEsIHIpID0+IGEgKyAoci5leHBvc3VyZSB8fCAwKSwgMCk7CiAgY29uc3QgaGlnaFZhbCA9IGhp"
    "Z2gucmVkdWNlKChhLCByKSA9PiBhICsgKHIuZXhwb3N1cmUgfHwgMCksIDApOwoKICBlbC5pbm5lckhUTUwgPSBgCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9u"
    "LXRpdGxlIj5SaXNrIE92ZXJ2aWV3PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJrcGktZ3JpZCI+CiAgICAgICR7a3BpKCJGbGFnZ2VkIFJlY29yZHMiLCBmbXRO"
    "dW0ocmlza3MubGVuZ3RoKSl9CiAgICAgICR7a3BpKCJDcml0aWNhbCIsIGZtdE51bShjcml0Lmxlbmd0aCksICIiLCBjcml0Lmxlbmd0aCA/ICJiYWQiIDog"
    "IiIpfQogICAgICAke2twaSgiSGlnaCIsIGZtdE51bShoaWdoLmxlbmd0aCksICIiLCBoaWdoLmxlbmd0aCA/ICJ3YXJuIiA6ICIiKX0KICAgICAgJHtrcGko"
    "IlRvdGFsIEV4cG9zdXJlIiwgZm10Q29tcGFjdChhdFJpc2tWYWwpLCBmbXRDdXIoYXRSaXNrVmFsKSl9CiAgICAgICR7a3BpKCJDcml0aWNhbCBFeHBvc3Vy"
    "ZSIsIGZtdENvbXBhY3QoY3JpdFZhbCksICIiLCAiYmFkIil9CiAgICAgICR7a3BpKCJIaWdoIEV4cG9zdXJlIiwgZm10Q29tcGFjdChoaWdoVmFsKSwgIiIs"
    "ICJ3YXJuIil9CiAgICA8L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5SaXNrIERpc3RyaWJ1dGlvbjwvZGl2PgogICAgPGRpdiBjbGFz"
    "cz0iY2hhcnQtZ3JpZCI+CiAgICAgICR7Y2FyZFNoZWxsKCJSaXNrIGJ5IFNldmVyaXR5IiwgInJpc2tTZXYiLCAiYzQiKX0KICAgICAgJHtjYXJkU2hlbGwo"
    "IlJpc2sgYnkgQ2F0ZWdvcnkiLCAicmlza0NhdCIsICJjNCIpfQogICAgICAke2NhcmRTaGVsbCgiRXhwb3N1cmUgYnkgQ2F0ZWdvcnkiLCAicmlza0V4cCIs"
    "ICJjNCIpfQogICAgICAke2NhcmRTaGVsbCgiUmlzayBDb3VudCBieSBDdXN0b21lciAoVG9wIDEyKSIsICJyaXNrQ3VzdCIsICJjNiIpfQogICAgICAke2Nh"
    "cmRTaGVsbCgiRXhwb3N1cmUgYnkgQ3VzdG9tZXIgKFRvcCAxMikiLCAicmlza0N1c3RFeHAiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlJpc2sgQ291"
    "bnQgYnkgRVQgUE9DIiwgInJpc2tQb2MiLCAiYzYiKX0KICAgICAgJHtjYXJkU2hlbGwoIlJpc2sgQ291bnQgYnkgU3VwcGxpZXIgKFRvcCAxMikiLCAicmlz"
    "a1N1cCIsICJjNiIpfQogICAgPC9kaXY+CgogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+UmlzayBSZWdpc3RlcjwvZGl2PgogICAgPGRpdiBjbGFz"
    "cz0iY2FyZCBjMTIiPjxkaXYgaWQ9InJpc2tUYWJsZSI+PC9kaXY+PC9kaXY+YDsKCiAgaWYgKCFyaXNrcy5sZW5ndGgpIHsgJCgiI3Jpc2tUYWJsZSIpLmlu"
    "bmVySFRNTCA9IGA8ZGl2IGNsYXNzPSJlbXB0eS1zdGF0ZSI+Tm8gcmlza3MgZmxhZ2dlZCBmb3IgdGhlIGN1cnJlbnQgc2VsZWN0aW9uLjwvZGl2PmA7IHJl"
    "dHVybjsgfQoKICBkb251dCgicmlza1NldiIsIGNvdW50QnlGaWVsZEFycihyaXNrcywgInNldmVyaXR5IiksIG51bGwsIHsgQ3JpdGljYWw6ICItLWJhZCIs"
    "IEhpZ2g6ICItLXdhcm4iLCBNb2RlcmF0ZTogIi0taW5mbyIsIExvdzogIi0tbXV0ZWQiIH0pOwogIGRvbnV0KCJyaXNrQ2F0IiwgY291bnRCeUZpZWxkQXJy"
    "KHJpc2tzLCAiY2F0ZWdvcnkiKSwgbnVsbCk7CiAgaGJhcigicmlza0V4cCIsIHRvcE4oc3VtTWFwQXJyKHJpc2tzLCAiY2F0ZWdvcnkiLCAiZXhwb3N1cmUi"
    "KSwgOCksIGZtdENvbXBhY3QpOwogIGhiYXIoInJpc2tDdXN0IiwgdG9wTihjb3VudEJ5RmllbGRBcnIocmlza3MsICJjdXN0b21lciIpLCAxMiksIG51bGwp"
    "OwogIGhiYXIoInJpc2tDdXN0RXhwIiwgdG9wTihzdW1NYXBBcnIocmlza3MsICJjdXN0b21lciIsICJleHBvc3VyZSIpLCAxMiksIGZtdENvbXBhY3QpOwog"
    "IGhiYXIoInJpc2tQb2MiLCB0b3BOKGNvdW50QnlGaWVsZEFycihyaXNrcywgImV0UE9DIiksIDEyKSwgbnVsbCk7CiAgaGJhcigicmlza1N1cCIsIHRvcE4o"
    "Y291bnRCeUZpZWxkQXJyKHJpc2tzLCAic3VwcGxpZXIiKSwgMTIpLCBudWxsKTsKCiAgbWFrZVRhYmxlKCQoIiNyaXNrVGFibGUiKSwgewogICAgcm93czog"
    "cmlza3MsIG5hbWU6ICJyaXNrX3JlZ2lzdGVyIiwgcGVyUGFnZTogMjUsCiAgICBkcmlsbDogKHIpID0+IG9wZW5Nb2RhbChyLmlkICsgIiDigJQgIiArIHIu"
    "Y2F0ZWdvcnksIHJpc2tNb2RhbEJvZHkocikpLAogICAgY29sdW1uczogWwogICAgICB7IGtleTogImlkIiwgbGFiZWw6ICJSaXNrIElEIiB9LCB7IGtleTog"
    "ImNhdGVnb3J5IiwgbGFiZWw6ICJDYXRlZ29yeSIgfSwKICAgICAgeyBrZXk6ICJzZXZlcml0eSIsIGxhYmVsOiAiU2V2ZXJpdHkiLCByZW5kZXI6IHYgPT4g"
    "YDxzcGFuIGNsYXNzPSJwaWxsICR7diA9PT0gJ0NyaXRpY2FsJyA/ICdyJyA6IHYgPT09ICdIaWdoJyA/ICdhJyA6IHYgPT09ICdNb2RlcmF0ZScgPyAnYicg"
    "OiAnJ30iPiR7ZXNjKHYpfTwvc3Bhbj5gIH0sCiAgICAgIHsga2V5OiAic2NvcmUiLCBsYWJlbDogIlNjb3JlIiwgdHlwZTogIm51bSIgfSwgeyBrZXk6ICJj"
    "dXN0b21lciIsIGxhYmVsOiAiQ3VzdG9tZXIiIH0sCiAgICAgIHsga2V5OiAicmZxIiwgbGFiZWw6ICJSRlEgTm8uIiB9LCB7IGtleTogInBvIiwgbGFiZWw6"
    "ICJDdXN0IFBPIiB9LAogICAgICB7IGtleTogImV0UE9DIiwgbGFiZWw6ICJFVCBQT0MiIH0sIHsga2V5OiAic3VwcGxpZXIiLCBsYWJlbDogIlN1cHBsaWVy"
    "IiwgdmlzOiBmYWxzZSB9LAogICAgICB7IGtleTogImRlc2NyaXB0aW9uIiwgbGFiZWw6ICJEZXNjcmlwdGlvbiIgfSwKICAgICAgeyBrZXk6ICJkYXlzT3Zl"
    "cmR1ZSIsIGxhYmVsOiAiRGF5cyBvdmVyZHVlIiwgdHlwZTogIm51bSIsIHJlbmRlcjogdiA9PiB2ID09IG51bGwgPyAi4oCUIiA6IHYgKyAiIGQiIH0sCiAg"
    "ICAgIHsga2V5OiAiZXhwb3N1cmUiLCBsYWJlbDogIkV4cG9zdXJlIiwgdHlwZTogImN1ciIgfSwgeyBrZXk6ICJhY3Rpb24iLCBsYWJlbDogIlJlY29tbWVu"
    "ZGVkIGFjdGlvbiIgfSwKICAgIF0sCiAgfSk7Cn0KZnVuY3Rpb24gY291bnRCeUZpZWxkQXJyKGFyciwgZmllbGQpIHsgY29uc3QgbSA9IG5ldyBNYXAoKTsg"
    "YXJyLmZvckVhY2gobyA9PiB7IGNvbnN0IGsgPSBvW2ZpZWxkXSB8fCAi4oCUIjsgbS5zZXQoaywgKG0uZ2V0KGspIHx8IDApICsgMSk7IH0pOyByZXR1cm4g"
    "bTsgfQpmdW5jdGlvbiBzdW1NYXBBcnIoYXJyLCBmaWVsZCwgdmFsRmllbGQpIHsgY29uc3QgbSA9IG5ldyBNYXAoKTsgYXJyLmZvckVhY2gobyA9PiB7IGNv"
    "bnN0IGsgPSBvW2ZpZWxkXSB8fCAi4oCUIjsgbS5zZXQoaywgKG0uZ2V0KGspIHx8IDApICsgKG9bdmFsRmllbGRdIHx8IDApKTsgfSk7IHJldHVybiBtOyB9"
    "CmZ1bmN0aW9uIHJpc2tNb2RhbEJvZHkocikgewogIGNvbnN0IGYgPSBbWyJSaXNrIElEIiwgci5pZF0sIFsiQ2F0ZWdvcnkiLCByLmNhdGVnb3J5XSwgWyJT"
    "ZXZlcml0eSIsIHIuc2V2ZXJpdHldLCBbIlNjb3JlIiwgci5zY29yZV0sCiAgICBbIkN1c3RvbWVyIiwgci5jdXN0b21lcl0sIFsiUkZRIE5vLiIsIHIucmZx"
    "XSwgWyJDdXN0b21lciBQTyIsIHIucG9dLCBbIkVUIFBPQyIsIHIuZXRQT0NdLCBbIlN1cHBsaWVyIiwgci5zdXBwbGllcl0sCiAgICBbIkRlc2NyaXB0aW9u"
    "Iiwgci5kZXNjcmlwdGlvbl0sIFsiRGF5cyBvdmVyZHVlIiwgci5kYXlzT3ZlcmR1ZSA9PSBudWxsID8gIuKAlCIgOiByLmRheXNPdmVyZHVlICsgIiBkIl0s"
    "CiAgICBbIkZpbmFuY2lhbCBleHBvc3VyZSIsIGZtdEN1cihyLmV4cG9zdXJlKV0sIFsiUmVjb21tZW5kZWQgYWN0aW9uIiwgci5hY3Rpb25dXTsKICByZXR1"
    "cm4gYDxkaXYgY2xhc3M9ImRldGFpbC1ncmlkIj4ke2YubWFwKChbaywgdl0pID0+IGA8ZGl2IGNsYXNzPSJkdC1yb3ciPjxzcGFuIGNsYXNzPSJkdC1rIj4k"
    "e2VzYyhrKX08L3NwYW4+PHNwYW4gY2xhc3M9ImR0LXYiPiR7ZXNjKHYgPT0gbnVsbCA/ICLigJQiIDogdil9PC9zcGFuPjwvZGl2PmApLmpvaW4oIiIpfTwv"
    "ZGl2PmA7Cn0KCi8qID09PT09PT09PT09PT09PT09IFRBQiA5IOKAlCBEQVRBIFFVQUxJVFkgJiBDT05UUk9MUyA9PT09PT09PT09PT09PT09PSAqLwpmdW5j"
    "dGlvbiByZW5kZXJEUShyb3dzKSB7CiAgY29uc3QgZWwgPSAkKCIjdGFiLWRxIik7CiAgY29uc3QgaXNzdWVzID0gZGF0YVF1YWxpdHlJc3N1ZXMocm93cyk7"
    "CiAgY29uc3QgUiA9IGJ1aWxkUm9sbHVwKHJvd3MpOwogIGNvbnN0IG1pc3NDdXN0ID0gcm93cy5maWx0ZXIociA9PiAhci5jdXN0b21lcikubGVuZ3RoOwog"
    "IGNvbnN0IG1pc3NTdXAgPSByb3dzLmZpbHRlcihyID0+ICFyLnN1cHBsaWVyTmFtZSkubGVuZ3RoOwogIGNvbnN0IG1pc3NQb2MgPSByb3dzLmZpbHRlcihy"
    "ID0+ICFyLmV0UE9DKS5sZW5ndGg7CiAgY29uc3QgbWlzc1N0YXR1cyA9IHJvd3MuZmlsdGVyKHIgPT4gIXIuZXRRdW90ZVN0YXR1cykubGVuZ3RoOwogIGNv"
    "bnN0IG5lZ0dwID0gUi5saXN0LmZpbHRlcihvID0+IG8uZ3AgPCAwKS5sZW5ndGg7CiAgY29uc3QgcG9CZWZvcmVRdW90ZSA9IHJvd3MuZmlsdGVyKHIgPT4g"
    "ci5jdXN0UG9EYXRlICYmIHIuZXRRdW90ZURhdGUgJiYgci5jdXN0UG9EYXRlIDwgci5ldFF1b3RlRGF0ZSkubGVuZ3RoOwogIGNvbnN0IGR1cFBvID0gZHVw"
    "bGljYXRlUG9BY3Jvc3NDdXN0b21lcnMocm93cyk7CiAgY29uc3QgdG90YWwgPSByb3dzLmxlbmd0aDsKICBjb25zdCBjb21wbGV0ZW5lc3MgPSBNRVRBLmNv"
    "bXBsZXRlbmVzcyB8fCB7fTsKCiAgZWwuaW5uZXJIVE1MID0gYAogICAgPGRpdiBjbGFzcz0ic2VjdGlvbi10aXRsZSI+RGF0YSBRdWFsaXR5IFN1bW1hcnk8"
    "L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImtwaS1ncmlkIj4KICAgICAgJHtrcGkoIkxpbmUgSXRlbXMiLCBmbXROdW0odG90YWwpKX0KICAgICAgJHtrcGkoIlVu"
    "aXF1ZSBSRlFzIiwgZm10TnVtKFIubGlzdC5sZW5ndGgpKX0KICAgICAgJHtrcGkoIk1pc3NpbmcgQ3VzdG9tZXIiLCBmbXROdW0obWlzc0N1c3QpLCB0b3Rh"
    "bCA/IGZtdFBjdChtaXNzQ3VzdCAvIHRvdGFsICogMTAwKSA6ICIiLCBtaXNzQ3VzdCA/ICJ3YXJuIiA6ICIiKX0KICAgICAgJHtrcGkoIk1pc3NpbmcgU3Vw"
    "cGxpZXIiLCBmbXROdW0obWlzc1N1cCksICIiLCBtaXNzU3VwID8gIndhcm4iIDogIiIpfQogICAgICAke2twaSgiTWlzc2luZyBFVCBQT0MiLCBmbXROdW0o"
    "bWlzc1BvYyksICIiLCBtaXNzUG9jID8gIndhcm4iIDogIiIpfQogICAgICAke2twaSgiTWlzc2luZyBRdW90ZSBTdGF0dXMiLCBmbXROdW0obWlzc1N0YXR1"
    "cyksICIiLCBtaXNzU3RhdHVzID8gIndhcm4iIDogIiIpfQogICAgICAke2twaSgiTmVnYXRpdmUgR1AgUkZRcyIsIGZtdE51bShuZWdHcCksICIiLCBuZWdH"
    "cCA/ICJiYWQiIDogIiIpfQogICAgICAke2twaSgiUE8gRGF0ZSA8IFF1b3RlIERhdGUiLCBmbXROdW0ocG9CZWZvcmVRdW90ZSksICIiLCBwb0JlZm9yZVF1"
    "b3RlID8gIndhcm4iIDogIiIpfQogICAgICAke2twaSgiUE8gTm8uIHJldXNlZCAobXVsdGktY3VzdG9tZXIpIiwgZm10TnVtKGR1cFBvKSwgIiIsIGR1cFBv"
    "ID8gIndhcm4iIDogIiIpfQogICAgICAke2twaSgiRmxhZ2dlZCBEUSBSZWNvcmRzIiwgZm10TnVtKGlzc3Vlcy5sZW5ndGgpKX0KICAgIDwvZGl2PgoKICAg"
    "IDxkaXYgY2xhc3M9InNlY3Rpb24tdGl0bGUiPkZpZWxkIENvbXBsZXRlbmVzczwvZGl2PgogICAgPGRpdiBjbGFzcz0iY2FyZCBjMTIiPjxkaXYgaWQ9ImRx"
    "TWF0cml4Ij48L2Rpdj48L2Rpdj4KCiAgICA8ZGl2IGNsYXNzPSJzZWN0aW9uLXRpdGxlIj5EYXRhLVF1YWxpdHkgSXNzdWVzPC9kaXY+CiAgICA8ZGl2IGNs"
    "YXNzPSJjYXJkIGMxMiI+PGRpdiBpZD0iZHFUYWJsZSI+PC9kaXY+PC9kaXY+YDsKCiAgLy8gY29tcGxldGVuZXNzIG1hdHJpeAogIGNvbnN0IGZpZWxkcyA9"
    "IE9iamVjdC5rZXlzKGNvbXBsZXRlbmVzcyk7CiAgY29uc3QgbXggPSAkKCIjZHFNYXRyaXgiKTsKICBteC5pbm5lckhUTUwgPSBgPGRpdiBjbGFzcz0iZHEt"
    "bWF0cml4Ij4ke2ZpZWxkcy5tYXAoZiA9PiB7CiAgICBjb25zdCBjID0gY29tcGxldGVuZXNzW2ZdOyBjb25zdCBwY3QgPSBjID8gYy5wY3QgOiAwOwogICAg"
    "Y29uc3QgY2xzID0gcGN0ID49IDkwID8gImciIDogcGN0ID49IDUwID8gImEiIDogInIiOwogICAgcmV0dXJuIGA8ZGl2IGNsYXNzPSJkcS1jZWxsIj4KICAg"
    "ICAgPGRpdiBjbGFzcz0iZHEtZmllbGQiPiR7ZXNjKGYpfTwvZGl2PgogICAgICA8ZGl2IGNsYXNzPSJkcS1iYXIiPjxzcGFuIGNsYXNzPSJkcS1maWxsICR7"
    "Y2xzfSIgc3R5bGU9IndpZHRoOiR7TWF0aC5tYXgoMiwgcGN0KX0lIj48L3NwYW4+PC9kaXY+CiAgICAgIDxkaXYgY2xhc3M9ImRxLXBjdCI+JHtwY3QudG9G"
    "aXhlZCgwKX0lPC9kaXY+PC9kaXY+YDsKICB9KS5qb2luKCIiKX08L2Rpdj4KICA8ZGl2IGNsYXNzPSJkcS1ub3RlIj5GaWVsZHMgYXQgMCUgKEN1c3RvbWVy"
    "IFJlcXVpcmVkIERhdGUsIEVUIFByb21pc2VkIC8gUlRTIC8gQWN0dWFsLVNoaXAgRGF0ZSwgU3VwcGxpZXIgUE8gZmllbGRzLAogIFNoaXBtZW50IEZpbmFs"
    "IFN0YXR1cykgYXJlIGVtcHR5IGluIHRoZSBjdXJyZW50IHNvdXJjZSDigJQgdGhlIGNvcnJlc3BvbmRpbmcgZGVsaXZlcnksIHNoaXBtZW50IGFuZCBzdXBw"
    "bGllci1QTwogIG1ldHJpY3MgYWNyb3NzIHRoZSBkYXNoYm9hcmQgc2hvdyBlbXB0eSBzdGF0ZXMgdW50aWwgdGhvc2UgY29sdW1ucyBhcmUgcG9wdWxhdGVk"
    "LjwvZGl2PmA7CgogIG1ha2VUYWJsZSgkKCIjZHFUYWJsZSIpLCB7CiAgICBuYW1lOiAiZGF0YV9xdWFsaXR5X2lzc3VlcyIsIHBlclBhZ2U6IDI1LAogICAg"
    "Y29sdW1uczogWwogICAgICB7IGtleTogInJvdyIsIGxhYmVsOiAiUm93IiwgdHlwZTogIm51bSIgfSwgeyBrZXk6ICJyZWNJZCIsIGxhYmVsOiAiUmVjb3Jk"
    "IiB9LAogICAgICB7IGtleTogImlzc3VlIiwgbGFiZWw6ICJJc3N1ZSIgfSwKICAgICAgeyBrZXk6ICJzZXZlcml0eSIsIGxhYmVsOiAiU2V2ZXJpdHkiLCBy"
    "ZW5kZXI6IHYgPT4gYDxzcGFuIGNsYXNzPSJwaWxsICR7diA9PT0gJ0hpZ2gnID8gJ3InIDogdiA9PT0gJ01lZGl1bScgPyAnYScgOiAnYid9Ij4ke2VzYyh2"
    "KX08L3NwYW4+YCB9LAogICAgICB7IGtleTogImZpZWxkIiwgbGFiZWw6ICJGaWVsZCIgfSwgeyBrZXk6ICJ2YWx1ZSIsIGxhYmVsOiAiQ3VycmVudCB2YWx1"
    "ZSIgfSwKICAgICAgeyBrZXk6ICJmaXgiLCBsYWJlbDogIlN1Z2dlc3RlZCBjb3JyZWN0aW9uIiB9LAogICAgXSwKICAgIHJvd3M6IGlzc3VlcywKICB9KTsK"
    "fQpmdW5jdGlvbiBkYXRhUXVhbGl0eUlzc3Vlcyhyb3dzKSB7CiAgY29uc3Qgb3V0ID0gW107CiAgY29uc3QgcHVzaCA9IChyLCBpc3N1ZSwgc2V2LCBmaWVs"
    "ZCwgdmFsdWUsIGZpeCkgPT4gb3V0LnB1c2goeyByb3c6IHIuX3JvdyB8fCByLnJmcVNObyB8fCAi4oCUIiwgcmVjSWQ6IChyLmN1c3RvbWVyIHx8ICI/Iikg"
    "KyAiIC8gIiArIChyLmN1c3RSZnFObyB8fCByLnJmcVNObyB8fCAiPyIpLCBpc3N1ZSwgc2V2ZXJpdHk6IHNldiwgZmllbGQsIHZhbHVlOiB2YWx1ZSA9PSBu"
    "dWxsID8gIuKAlCIgOiBTdHJpbmcodmFsdWUpLCBmaXggfSk7CiAgZm9yIChjb25zdCByIG9mIHJvd3MpIHsKICAgIGlmICghci5jdXN0b21lcikgcHVzaChy"
    "LCAiTWlzc2luZyBjdXN0b21lciBuYW1lIiwgIkhpZ2giLCAiY3VzdG9tZXIiLCBudWxsLCAiQWRkIGN1c3RvbWVyIG5hbWUiKTsKICAgIGlmICghci5jdXN0"
    "UmZxTm8pIHB1c2gociwgIk1pc3NpbmcgY3VzdG9tZXIgUkZRIG51bWJlciIsICJNZWRpdW0iLCAiY3VzdFJmcU5vIiwgbnVsbCwgIkNhcHR1cmUgUkZRIG51"
    "bWJlciIpOwogICAgaWYgKCFyLmV0UE9DKSBwdXNoKHIsICJNaXNzaW5nIEVUIFBPQyIsICJNZWRpdW0iLCAiZXRQT0MiLCBudWxsLCAiQXNzaWduIEVUIFBP"
    "QyIpOwogICAgaWYgKCFyLmV0UXVvdGVTdGF0dXMpIHB1c2gociwgIk1pc3NpbmcgcXVvdGUgc3RhdHVzIiwgIk1lZGl1bSIsICJldFF1b3RlU3RhdHVzIiwg"
    "bnVsbCwgIlNldCBxdW90ZSBzdGF0dXMiKTsKICAgIGlmIChyLmdyb3NzUHJvZml0Q2FsYyAhPSBudWxsICYmIHIuZ3Jvc3NQcm9maXRDYWxjIDwgMCkgcHVz"
    "aChyLCAiTmVnYXRpdmUgZ3Jvc3MgcHJvZml0IiwgIkhpZ2giLCAiZ3Jvc3NQcm9maXQiLCByLmdyb3NzUHJvZml0Q2FsYywgIlZhbGlkYXRlIHByaWNpbmcg"
    "LyBjb3N0Iik7CiAgICBpZiAoci5jdXN0UG9EYXRlICYmIHIuZXRRdW90ZURhdGUgJiYgci5jdXN0UG9EYXRlIDwgci5ldFF1b3RlRGF0ZSkgcHVzaChyLCAi"
    "Q3VzdG9tZXIgUE8gZGF0ZSBiZWZvcmUgRVQgcXVvdGUgZGF0ZSIsICJNZWRpdW0iLCAiY3VzdFBvRGF0ZSIsIHIuY3VzdFBvRGF0ZSwgIlZlcmlmeSBkYXRl"
    "IHNlcXVlbmNlIik7CiAgICBpZiAoci5ldFF1b3RlZFZhbHVlICE9IG51bGwgJiYgci5zdXBwbGllclRvdGFsUHJpY2UgIT0gbnVsbCAmJiByLmdyb3NzUHJv"
    "Zml0ICE9IG51bGwgJiYKICAgICAgTWF0aC5hYnMoKHIuZXRRdW90ZWRWYWx1ZSAtIHIuc3VwcGxpZXJUb3RhbFByaWNlKSAtIHIuZ3Jvc3NQcm9maXQpID4g"
    "TWF0aC5tYXgoMSwgci5ldFF1b3RlZFZhbHVlICogMC4wMikpCiAgICAgIHB1c2gociwgIkdQIGRvZXMgbm90IHJlY29uY2lsZSB0byB2YWx1ZSDiiJIgY29z"
    "dCIsICJNZWRpdW0iLCAiZ3Jvc3NQcm9maXQiLCByLmdyb3NzUHJvZml0LCAiUmVjb25jaWxlIEdQIik7CiAgICBpZiAoci5xdHkgIT0gbnVsbCAmJiByLnF0"
    "eSA8PSAwKSBwdXNoKHIsICJaZXJvIC8gbmVnYXRpdmUgcXVhbnRpdHkiLCAiTWVkaXVtIiwgInF0eSIsIHIucXR5LCAiQ29ycmVjdCBxdWFudGl0eSIpOwog"
    "IH0KICByZXR1cm4gb3V0LnNsaWNlKDAsIDUwMDApOwp9CmZ1bmN0aW9uIGR1cGxpY2F0ZVBvQWNyb3NzQ3VzdG9tZXJzKHJvd3MpIHsKICBjb25zdCBtID0g"
    "bmV3IE1hcCgpOwogIHJvd3MuZm9yRWFjaChyID0+IHsgaWYgKHIuY3VzdFBvTm8pIHsgY29uc3QgcyA9IG0uZ2V0KHIuY3VzdFBvTm8pIHx8IG5ldyBTZXQo"
    "KTsgcy5hZGQoci5jdXN0b21lcik7IG0uc2V0KHIuY3VzdFBvTm8sIHMpOyB9IH0pOwogIGxldCBuID0gMDsgZm9yIChjb25zdCBzIG9mIG0udmFsdWVzKCkp"
    "IGlmIChzLnNpemUgPiAxKSBuKys7IHJldHVybiBuOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09CiAgIFBBUlQgNSDigJQgZXhwb3J0cywgbWFzdGVyIHJlZnJlc2gsIHRhYiBzd2l0Y2hpbmcsIGJvb3QKICAgPT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwoKLyogLS0tLS0tLS0tLS0tLS0tLSAxMy4gRVhQT1JUUyAtLS0t"
    "LS0tLS0tLS0tLS0tICovCmZ1bmN0aW9uIGNzdkVzY2FwZSh2KSB7CiAgaWYgKHYgPT0gbnVsbCkgcmV0dXJuICIiOwogIGNvbnN0IHMgPSBTdHJpbmcodik7"
    "CiAgcmV0dXJuIC9bIixcbl0vLnRlc3QocykgPyAnIicgKyBzLnJlcGxhY2UoLyIvZywgJyIiJykgKyAnIicgOiBzOwp9CmZ1bmN0aW9uIGRvd25sb2FkRmls"
    "ZShuYW1lLCBjb250ZW50LCB0eXBlKSB7CiAgY29uc3QgYmxvYiA9IG5ldyBCbG9iKFtjb250ZW50XSwgeyB0eXBlOiB0eXBlIHx8ICJ0ZXh0L3BsYWluO2No"
    "YXJzZXQ9dXRmLTgiIH0pOwogIGNvbnN0IHVybCA9IFVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7CiAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1l"
    "bnQoImEiKTsgYS5ocmVmID0gdXJsOyBhLmRvd25sb2FkID0gbmFtZTsgYS5jbGljaygpOwogIHNldFRpbWVvdXQoKCkgPT4gVVJMLnJldm9rZU9iamVjdFVS"
    "TCh1cmwpLCAxMDAwKTsKfQpmdW5jdGlvbiBleHBvcnRSb3dzQ1NWKGNvbHVtbnMsIHJvd3MsIGZpbGVuYW1lKSB7CiAgY29uc3QgaGVhZCA9IGNvbHVtbnMu"
    "bWFwKGMgPT4gY3N2RXNjYXBlKGMubGFiZWwpKS5qb2luKCIsIik7CiAgY29uc3QgYm9keSA9IHJvd3MubWFwKHIgPT4gY29sdW1ucy5tYXAoYyA9PiB7CiAg"
    "ICBsZXQgdiA9IHJbYy5rZXldOwogICAgaWYgKGMudHlwZSA9PT0gImRhdGUiKSB2ID0gdiB8fCAiIjsKICAgIGVsc2UgaWYgKHYgIT0gbnVsbCAmJiAoYy50"
    "eXBlID09PSAiY3VyIiB8fCBjLnR5cGUgPT09ICJudW0iIHx8IGMudHlwZSA9PT0gInBjdCIgfHwgYy50eXBlID09PSAiZGF5cyIpKSB2ID0gK3Y7CiAgICBy"
    "ZXR1cm4gY3N2RXNjYXBlKHYpOwogIH0pLmpvaW4oIiwiKSkuam9pbigiXG4iKTsKICBkb3dubG9hZEZpbGUoZmlsZW5hbWUgfHwgImV4cG9ydC5jc3YiLCBo"
    "ZWFkICsgIlxuIiArIGJvZHksICJ0ZXh0L2NzdjtjaGFyc2V0PXV0Zi04Iik7Cn0KZnVuY3Rpb24gZXhwb3J0RGF0YXNldChraW5kKSB7CiAgY29uc3Qgcm93"
    "cyA9IGFwcGx5RmlsdGVycygpOwogIGlmIChraW5kID09PSAicmZxIikgewogICAgY29uc3QgUiA9IGJ1aWxkUm9sbHVwKHJvd3MpLmxpc3QubWFwKG8gPT4g"
    "KHsKICAgICAgcmZxTm86IG8ua2V5LnNwbGl0KCJ8fCIpWzFdIHx8ICIiLCBjdXN0b21lcjogby5jdXN0b21lciwgZXRQT0M6IG8uZXRQT0MsIGN1c3RSZnFE"
    "YXRlOiBvLmN1c3RSZnFEYXRlLAogICAgICBldFF1b3RlRGF0ZTogby5ldFF1b3RlRGF0ZSwgcmVzcERheXM6IG8ucmVzcERheSwgbGluZXM6IG8ubGluZXMs"
    "IHN1cHBsaWVyOiBvLnN1cHBsaWVyLAogICAgICBxdW90ZWRWYWx1ZTogby5yZXYsIHN1cHBsaWVyQ29zdDogby5jb2dzLCBncm9zc1Byb2ZpdDogby5ncCwg"
    "bWFyZ2luOiBvLm1hcmdpbiwKICAgICAgcmVzdWx0OiBvLnJlc3VsdCwgY3VzdFBvTm86IG8uY3VzdFBvTm8sCiAgICB9KSk7CiAgICBleHBvcnRSb3dzQ1NW"
    "KFsKICAgICAgeyBrZXk6ICJyZnFObyIsIGxhYmVsOiAiQ3VzdG9tZXIgUkZRIE5vIiB9LCB7IGtleTogImN1c3RvbWVyIiwgbGFiZWw6ICJDdXN0b21lciIg"
    "fSwgeyBrZXk6ICJldFBPQyIsIGxhYmVsOiAiRVQgUE9DIiB9LAogICAgICB7IGtleTogImN1c3RSZnFEYXRlIiwgbGFiZWw6ICJSRlEgRGF0ZSIsIHR5cGU6"
    "ICJkYXRlIiB9LCB7IGtleTogImV0UXVvdGVEYXRlIiwgbGFiZWw6ICJRdW90ZSBEYXRlIiwgdHlwZTogImRhdGUiIH0sCiAgICAgIHsga2V5OiAicmVzcERh"
    "eXMiLCBsYWJlbDogIlJlc3BvbnNlIERheXMiLCB0eXBlOiAibnVtIiB9LCB7IGtleTogImxpbmVzIiwgbGFiZWw6ICJMaW5lIEl0ZW1zIiwgdHlwZTogIm51"
    "bSIgfSwKICAgICAgeyBrZXk6ICJzdXBwbGllciIsIGxhYmVsOiAiU3VwcGxpZXIiIH0sIHsga2V5OiAicXVvdGVkVmFsdWUiLCBsYWJlbDogIkVUIFF1b3Rl"
    "ZCBWYWx1ZSIsIHR5cGU6ICJjdXIiIH0sCiAgICAgIHsga2V5OiAic3VwcGxpZXJDb3N0IiwgbGFiZWw6ICJTdXBwbGllciBDb3N0IiwgdHlwZTogImN1ciIg"
    "fSwgeyBrZXk6ICJncm9zc1Byb2ZpdCIsIGxhYmVsOiAiR3Jvc3MgUHJvZml0IiwgdHlwZTogImN1ciIgfSwKICAgICAgeyBrZXk6ICJtYXJnaW4iLCBsYWJl"
    "bDogIk1hcmdpbiAlIiwgdHlwZTogInBjdCIgfSwgeyBrZXk6ICJyZXN1bHQiLCBsYWJlbDogIlJlc3VsdCIgfSwgeyBrZXk6ICJjdXN0UG9ObyIsIGxhYmVs"
    "OiAiQ3VzdG9tZXIgUE8gTm8iIH0sCiAgICBdLCBSLCAicmZxX2RhdGEuY3N2Iik7CiAgfSBlbHNlIGlmIChraW5kID09PSAicG8iKSB7CiAgICBjb25zdCBw"
    "b1Jvd3MgPSByb3dzLmZpbHRlcihyID0+IHIuY3VzdFBvTm8pOwogICAgY29uc3QgdmFsID0gZ3JvdXBEZWR1cChwb1Jvd3MsICJjdXN0UG9LZXkiLCAiZXRR"
    "dW90ZWRWYWx1ZSIpLCBncCA9IGdyb3VwRGVkdXAocG9Sb3dzLCAiY3VzdFBvS2V5IiwgImdyb3NzUHJvZml0Q2FsYyIpLCBjb3N0ID0gZ3JvdXBEZWR1cChw"
    "b1Jvd3MsICJjdXN0UG9LZXkiLCAic3VwcGxpZXJUb3RhbFByaWNlIik7CiAgICBjb25zdCBzZWVuID0gbmV3IE1hcCgpOwogICAgcG9Sb3dzLmZvckVhY2go"
    "ciA9PiB7IGlmICghc2Vlbi5oYXMoci5jdXN0UG9LZXkpKSBzZWVuLnNldChyLmN1c3RQb0tleSwgcik7IH0pOwogICAgY29uc3QgZGF0YSA9IEFycmF5LmZy"
    "b20oc2Vlbi5lbnRyaWVzKCkpLm1hcCgoW2ssIHJdKSA9PiAoewogICAgICBwb05vOiByLmN1c3RQb05vLCBjdXN0b21lcjogci5jdXN0b21lciwgZXRQT0M6"
    "IHIuZXRQT0MsIHBvRGF0ZTogci5jdXN0UG9EYXRlLCBvYURhdGU6IHIuZXRPYURhdGUsCiAgICAgIHBvVG9PYTogci5wb1RvT2FEYXlzLCB2YWx1ZTogdmFs"
    "LmdldChrKSB8fCAwLCBjb3N0OiBjb3N0LmdldChrKSB8fCAwLCBncDogZ3AuZ2V0KGspIHx8IDAsCiAgICB9KSk7CiAgICBleHBvcnRSb3dzQ1NWKFsKICAg"
    "ICAgeyBrZXk6ICJwb05vIiwgbGFiZWw6ICJDdXN0b21lciBQTyBObyIgfSwgeyBrZXk6ICJjdXN0b21lciIsIGxhYmVsOiAiQ3VzdG9tZXIiIH0sIHsga2V5"
    "OiAiZXRQT0MiLCBsYWJlbDogIkVUIFBPQyIgfSwKICAgICAgeyBrZXk6ICJwb0RhdGUiLCBsYWJlbDogIlBPIERhdGUiLCB0eXBlOiAiZGF0ZSIgfSwgeyBr"
    "ZXk6ICJvYURhdGUiLCBsYWJlbDogIk9BIERhdGUiLCB0eXBlOiAiZGF0ZSIgfSwKICAgICAgeyBrZXk6ICJwb1RvT2EiLCBsYWJlbDogIlBPIHRvIE9BIERh"
    "eXMiLCB0eXBlOiAibnVtIiB9LCB7IGtleTogInZhbHVlIiwgbGFiZWw6ICJQTyBWYWx1ZSIsIHR5cGU6ICJjdXIiIH0sCiAgICAgIHsga2V5OiAiY29zdCIs"
    "IGxhYmVsOiAiU3VwcGxpZXIgQ29zdCIsIHR5cGU6ICJjdXIiIH0sIHsga2V5OiAiZ3AiLCBsYWJlbDogIkdyb3NzIFByb2ZpdCIsIHR5cGU6ICJjdXIiIH0s"
    "CiAgICBdLCBkYXRhLCAiY3VzdG9tZXJfcG9fZGF0YS5jc3YiKTsKICB9IGVsc2UgaWYgKGtpbmQgPT09ICJzdXBwbGllciIpIHsKICAgIGNvbnN0IHNjb3Jl"
    "cyA9IHN1cHBsaWVyU2NvcmVzKHJvd3MpOwogICAgZXhwb3J0Um93c0NTVihbCiAgICAgIHsga2V5OiAibmFtZSIsIGxhYmVsOiAiU3VwcGxpZXIiIH0sIHsg"
    "a2V5OiAic2VjdG9yIiwgbGFiZWw6ICJTZWN0b3IiIH0sIHsga2V5OiAic2VsQ291bnQiLCBsYWJlbDogIlNlbGVjdGVkICh3b24gUkZRcykiLCB0eXBlOiAi"
    "bnVtIiB9LAogICAgICB7IGtleTogInNwZW5kIiwgbGFiZWw6ICJQcm9jdXJlbWVudCBWYWx1ZSIsIHR5cGU6ICJjdXIiIH0sIHsga2V5OiAiZ3AiLCBsYWJl"
    "bDogIkdQIFN1cHBvcnRlZCIsIHR5cGU6ICJjdXIiIH0sCiAgICAgIHsga2V5OiAicmVzcEF2ZyIsIGxhYmVsOiAiQXZnIFF1b3RlIFJlc3AgRGF5cyIsIHR5"
    "cGU6ICJudW0iIH0sIHsga2V5OiAibWFyZ2luQXZnIiwgbGFiZWw6ICJBdmcgTWFyZ2luICUiLCB0eXBlOiAicGN0IiB9LAogICAgICB7IGtleTogInNjb3Jl"
    "IiwgbGFiZWw6ICJTY29yZSIsIHR5cGU6ICJudW0iIH0sIHsga2V5OiAiYmFuZCIsIGxhYmVsOiAiUmF0aW5nIiB9LAogICAgXSwgc2NvcmVzLCAic3VwcGxp"
    "ZXJfZGF0YS5jc3YiKTsKICB9IGVsc2UgaWYgKGtpbmQgPT09ICJyaXNrIikgewogICAgZXhwb3J0Um93c0NTVihbCiAgICAgIHsga2V5OiAiaWQiLCBsYWJl"
    "bDogIlJpc2sgSUQiIH0sIHsga2V5OiAiY2F0ZWdvcnkiLCBsYWJlbDogIkNhdGVnb3J5IiB9LCB7IGtleTogInNldmVyaXR5IiwgbGFiZWw6ICJTZXZlcml0"
    "eSIgfSwKICAgICAgeyBrZXk6ICJzY29yZSIsIGxhYmVsOiAiU2NvcmUiLCB0eXBlOiAibnVtIiB9LCB7IGtleTogImN1c3RvbWVyIiwgbGFiZWw6ICJDdXN0"
    "b21lciIgfSwgeyBrZXk6ICJyZnEiLCBsYWJlbDogIlJGUSBObyIgfSwKICAgICAgeyBrZXk6ICJwbyIsIGxhYmVsOiAiQ3VzdG9tZXIgUE8iIH0sIHsga2V5"
    "OiAiZXRQT0MiLCBsYWJlbDogIkVUIFBPQyIgfSwgeyBrZXk6ICJzdXBwbGllciIsIGxhYmVsOiAiU3VwcGxpZXIiIH0sCiAgICAgIHsga2V5OiAiZGVzY3Jp"
    "cHRpb24iLCBsYWJlbDogIkRlc2NyaXB0aW9uIiB9LCB7IGtleTogImRheXNPdmVyZHVlIiwgbGFiZWw6ICJEYXlzIE92ZXJkdWUiLCB0eXBlOiAibnVtIiB9"
    "LAogICAgICB7IGtleTogImV4cG9zdXJlIiwgbGFiZWw6ICJFeHBvc3VyZSIsIHR5cGU6ICJjdXIiIH0sIHsga2V5OiAiYWN0aW9uIiwgbGFiZWw6ICJSZWNv"
    "bW1lbmRlZCBBY3Rpb24iIH0sCiAgICBdLCBjb21wdXRlUmlza3Mocm93cyksICJyaXNrX3JlZ2lzdGVyLmNzdiIpOwogIH0gZWxzZSBpZiAoa2luZCA9PT0g"
    "ImRxIikgewogICAgZXhwb3J0Um93c0NTVihbCiAgICAgIHsga2V5OiAicm93IiwgbGFiZWw6ICJSb3ciIH0sIHsga2V5OiAicmVjSWQiLCBsYWJlbDogIlJl"
    "Y29yZCIgfSwgeyBrZXk6ICJpc3N1ZSIsIGxhYmVsOiAiSXNzdWUiIH0sCiAgICAgIHsga2V5OiAic2V2ZXJpdHkiLCBsYWJlbDogIlNldmVyaXR5IiB9LCB7"
    "IGtleTogImZpZWxkIiwgbGFiZWw6ICJGaWVsZCIgfSwgeyBrZXk6ICJ2YWx1ZSIsIGxhYmVsOiAiVmFsdWUiIH0sIHsga2V5OiAiZml4IiwgbGFiZWw6ICJT"
    "dWdnZXN0ZWQgQ29ycmVjdGlvbiIgfSwKICAgIF0sIGRhdGFRdWFsaXR5SXNzdWVzKHJvd3MpLCAiZGF0YV9xdWFsaXR5X2lzc3Vlcy5jc3YiKTsKICB9Cn0K"
    "ZnVuY3Rpb24gZXhwb3J0U3VtbWFyeSgpIHsKICBjb25zdCByb3dzID0gYXBwbHlGaWx0ZXJzKCk7CiAgY29uc3QgbSA9IGNvcmVNZXRyaWNzKHJvd3MpOwog"
    "IGNvbnN0IGluc2lnaHRzID0gZ2VuZXJhdGVJbnNpZ2h0cyhyb3dzKTsKICBjb25zdCByaXNrcyA9IGNvbXB1dGVSaXNrcyhyb3dzKTsKICBjb25zdCB0b3BS"
    "aXNrcyA9IHJpc2tzLnNsaWNlKDAsIDUpOwogIGNvbnN0IHdvbk5vUG8gPSBtLndvbi5maWx0ZXIobyA9PiAhby5oYXNQTyk7CiAgY29uc3QgaHRtbCA9IGA8"
    "IWRvY3R5cGUgaHRtbD48aHRtbD48aGVhZD48bWV0YSBjaGFyc2V0PSJ1dGYtOCI+PHRpdGxlPkVUIFJGUSAmIFBPIENvbnRyb2xsZXIg4oCUIE1hbmFnZW1l"
    "bnQgU3VtbWFyeTwvdGl0bGU+CiAgPHN0eWxlPmJvZHl7Zm9udC1mYW1pbHk6QXJpYWwsSGVsdmV0aWNhLHNhbnMtc2VyaWY7bWFyZ2luOjMycHg7Y29sb3I6"
    "IzFhMjIzMDtsaW5lLWhlaWdodDoxLjV9CiAgaDF7Zm9udC1zaXplOjIwcHg7Ym9yZGVyLWJvdHRvbToycHggc29saWQgIzJmNWJkNDtwYWRkaW5nLWJvdHRv"
    "bTo2cHh9aDJ7Zm9udC1zaXplOjE1cHg7bWFyZ2luLXRvcDoyMnB4O2NvbG9yOiMyZjViZDR9CiAgdGFibGV7Ym9yZGVyLWNvbGxhcHNlOmNvbGxhcHNlO3dp"
    "ZHRoOjEwMCU7bWFyZ2luLXRvcDo4cHh9dGQsdGh7Ym9yZGVyOjFweCBzb2xpZCAjY2NjO3BhZGRpbmc6NnB4IDlweDtmb250LXNpemU6MTNweDt0ZXh0LWFs"
    "aWduOmxlZnR9CiAgdGh7YmFja2dyb3VuZDojZjJmNWZifS5re2NvbG9yOiM1NTZ9Lm57dGV4dC1hbGlnbjpyaWdodH1saXttYXJnaW46M3B4IDB9PC9zdHls"
    "ZT48L2hlYWQ+PGJvZHk+CiAgPGgxPkVUIFJGUSAmYW1wOyBQTyBDb250cm9sbGVyIOKAlCBNYW5hZ2VtZW50IFN1bW1hcnk8L2gxPgogIDxkaXYgY2xhc3M9"
    "ImsiPkdlbmVyYXRlZCAke25ldyBEYXRlKCkudG9Mb2NhbGVTdHJpbmcoKX0gwrcgRGF0YSB0aHJvdWdoICR7Zm10RGF0ZShNRVRBLmRhdGFEYXRlTWF4KX0g"
    "wrcgJHtmbXROdW0ocm93cy5sZW5ndGgpfSBsaW5lIGl0ZW1zIGluIGN1cnJlbnQgZmlsdGVyPC9kaXY+CiAgPGgyPktleSBLUElzPC9oMj4KICA8dGFibGU+"
    "PHRib2R5PgogICAgPHRyPjx0aD5VbmlxdWUgUkZRczwvdGg+PHRkIGNsYXNzPSJuIj4ke2ZtdE51bShtLmxpc3QubGVuZ3RoKX08L3RkPjx0aD5Xb24gUkZR"
    "czwvdGg+PHRkIGNsYXNzPSJuIj4ke2ZtdE51bShtLndvbi5sZW5ndGgpfTwvdGQ+PC90cj4KICAgIDx0cj48dGg+V2luIFJhdGU8L3RoPjx0ZCBjbGFzcz0i"
    "biI+JHtmbXRQY3QobS53aW5SYXRlKX08L3RkPjx0aD5RdW90ZeKGklBPIENvbnZlcnNpb248L3RoPjx0ZCBjbGFzcz0ibiI+JHtmbXRQY3QobS5jb252KX08"
    "L3RkPjwvdHI+CiAgICA8dHI+PHRoPkVUIFF1b3RlZCBWYWx1ZTwvdGg+PHRkIGNsYXNzPSJuIj4ke2ZtdEN1cihtLnJldil9PC90ZD48dGg+Q3VzdG9tZXIg"
    "UE9zPC90aD48dGQgY2xhc3M9Im4iPiR7Zm10TnVtKG0ublBPKX08L3RkPjwvdHI+CiAgICA8dHI+PHRoPlN1cHBsaWVyIENvc3Q8L3RoPjx0ZCBjbGFzcz0i"
    "biI+JHtmbXRDdXIobS5jb2dzKX08L3RkPjx0aD5Hcm9zcyBQcm9maXQ8L3RoPjx0ZCBjbGFzcz0ibiI+JHtmbXRDdXIobS5ncCl9PC90ZD48L3RyPgogICAg"
    "PHRyPjx0aD5Hcm9zcyBNYXJnaW4gJTwvdGg+PHRkIGNsYXNzPSJuIj4ke2ZtdFBjdChtLm1hcmdpbil9PC90ZD48dGg+QXZnIFBPIFZhbHVlPC90aD48dGQg"
    "Y2xhc3M9Im4iPiR7Zm10Q3VyKG0uYXZnUE8pfTwvdGQ+PC90cj4KICA8L3Rib2R5PjwvdGFibGU+CiAgPGgyPlRvcCBJbnNpZ2h0czwvaDI+PHVsPiR7aW5z"
    "aWdodHMuc2xpY2UoMCwgNikubWFwKGkgPT4gYDxsaT48Yj4ke2lbMV19OjwvYj4gJHtpWzJdfTwvbGk+YCkuam9pbigiIil9PC91bD4KICA8aDI+VG9wIFJp"
    "c2tzPC9oMj4KICA8dGFibGU+PHRoZWFkPjx0cj48dGg+SUQ8L3RoPjx0aD5TZXZlcml0eTwvdGg+PHRoPkN1c3RvbWVyPC90aD48dGg+RGVzY3JpcHRpb248"
    "L3RoPjx0aCBjbGFzcz0ibiI+RXhwb3N1cmU8L3RoPjx0aD5BY3Rpb248L3RoPjwvdHI+PC90aGVhZD4KICA8dGJvZHk+JHt0b3BSaXNrcy5sZW5ndGggPyB0"
    "b3BSaXNrcy5tYXAociA9PiBgPHRyPjx0ZD4ke3IuaWR9PC90ZD48dGQ+JHtyLnNldmVyaXR5fTwvdGQ+PHRkPiR7ZXNjKHIuY3VzdG9tZXIpfTwvdGQ+PHRk"
    "PiR7ZXNjKHIuZGVzY3JpcHRpb24pfTwvdGQ+PHRkIGNsYXNzPSJuIj4ke2ZtdEN1cihyLmV4cG9zdXJlKX08L3RkPjx0ZD4ke2VzYyhyLmFjdGlvbil9PC90"
    "ZD48L3RyPmApLmpvaW4oIiIpIDogYDx0cj48dGQgY29sc3Bhbj0iNiI+Tm8gcmlza3MgZmxhZ2dlZC48L3RkPjwvdHI+YH08L3Rib2R5PjwvdGFibGU+CiAg"
    "PGgyPkFjdGlvbiBJdGVtczwvaDI+PHVsPgogICAgJHt3b25Ob1BvLmxlbmd0aCA/IGA8bGk+T2J0YWluIGN1c3RvbWVyIFBPcyBmb3IgJHt3b25Ob1BvLmxl"
    "bmd0aH0gd29uIFJGUShzKSB3b3J0aCAke2ZtdEN1cih3b25Ob1BvLnJlZHVjZSgoYSwgbykgPT4gYSArIG8ucmV2LCAwKSl9LjwvbGk+YCA6ICIifQogICAg"
    "PGxpPlJldmlldyAke3Jpc2tzLmZpbHRlcihyID0+IHIuc2V2ZXJpdHkgPT09ICJDcml0aWNhbCIpLmxlbmd0aH0gY3JpdGljYWwgYW5kICR7cmlza3MuZmls"
    "dGVyKHIgPT4gci5zZXZlcml0eSA9PT0gIkhpZ2giKS5sZW5ndGh9IGhpZ2gtcmlzayByZWNvcmRzIGluIHRoZSBSaXNrIHRhYi48L2xpPgogICAgPGxpPlBv"
    "cHVsYXRlIGRlbGl2ZXJ5ICZhbXA7IHN1cHBsaWVyLVBPIGNvbHVtbnMgKEFa4oCTQkopIHRvIHVubG9jayBmdWxmaWxtZW50IGFuZCBzaGlwbWVudCBhbmFs"
    "eXRpY3MuPC9saT4KICA8L3VsPgogIDxwIGNsYXNzPSJrIiBzdHlsZT0ibWFyZ2luLXRvcDoyNHB4Ij5Ob3RlOiBkZWxpdmVyeSAvIHNoaXBtZW50IC8gc3Vw"
    "cGxpZXItUE8gbWV0cmljcyBhcmUgdW5hdmFpbGFibGUgYmVjYXVzZSB0aG9zZSBzb3VyY2UgY29sdW1ucyBhcmUgZW1wdHkuIEZpbmFuY2lhbHMgdXNlIGRl"
    "ZHVwLWF3YXJlIGFnZ3JlZ2F0aW9uOyBhbGwgZGF5IG1ldHJpY3MgYXJlIGNhbGVuZGFyIGRheXMuPC9wPgogIDwvYm9keT48L2h0bWw+YDsKICBkb3dubG9h"
    "ZEZpbGUoImV0X21hbmFnZW1lbnRfc3VtbWFyeS5odG1sIiwgaHRtbCwgInRleHQvaHRtbDtjaGFyc2V0PXV0Zi04Iik7Cn0KZnVuY3Rpb24gcHJpbnRWaWV3"
    "KCkgeyB3aW5kb3cucHJpbnQoKTsgfQoKLyogLS0tLS0tLS0tLS0tLS0tLSBNQVNURVIgUkVGUkVTSCAtLS0tLS0tLS0tLS0tLS0tICovCmNvbnN0IFRBQl9S"
    "RU5ERVJFUlMgPSB7CiAgb3ZlcnZpZXc6IHJlbmRlck92ZXJ2aWV3LCByZnE6IHJlbmRlclJGUSwgc3VwcGxpZXI6IHJlbmRlclN1cHBsaWVyLCBwbzogcmVu"
    "ZGVyUE8sCiAgcG9jOiByZW5kZXJQT0MsIGNvbXBhcmU6IHJlbmRlckNvbXBhcmUsIGN1c3RvbWVyOiByZW5kZXJDdXN0b21lciwgcmlzazogcmVuZGVyUmlz"
    "aywgZHE6IHJlbmRlckRRLAp9OwpmdW5jdGlvbiByZWZyZXNoKGZvcmNlKSB7CiAgY29uc3Qgcm93cyA9IGFwcGx5RmlsdGVycygpOwogIC8vIGhlYWRlciAr"
    "IGZvb3RlcgogICQoIiNyZWNvcmRDb3VudCIpLnRleHRDb250ZW50ID0gZm10TnVtKHJvd3MubGVuZ3RoKTsKICAkKCIjZmlsdGVyQ291bnQiKS50ZXh0Q29u"
    "dGVudCA9IGFjdGl2ZUZpbHRlckNvdW50KCk7CiAgJCgiI2RhdGFSYW5nZUxhYmVsIikudGV4dENvbnRlbnQgPSBNRVRBLmRhdGFEYXRlTWluID8gYCR7Zm10"
    "RGF0ZShNRVRBLmRhdGFEYXRlTWluKX0g4oCTICR7Zm10RGF0ZShNRVRBLmRhdGFEYXRlTWF4KX1gIDogIuKAlCI7CiAgaWYgKE1FVEEuZ2VuZXJhdGVkQXQp"
    "IHsKICAgIGNvbnN0IGQgPSBuZXcgRGF0ZShNRVRBLmdlbmVyYXRlZEF0KTsKICAgICQoIiNmb290ZXJSZWZyZXNoIikudGV4dENvbnRlbnQgPSAiRGFzaGJv"
    "YXJkIHJlZnJlc2hlZCAiICsgZC50b0xvY2FsZVN0cmluZygpOwogIH0KICB1cGRhdGVDaGlwcygpOwogIGlmIChmb3JjZSkgUkVOREVSRUQgPSB7fTsKICBp"
    "ZiAoIVJFTkRFUkVEW0NVUlJFTlRfVEFCXSkgewogICAgY29uc3QgZm4gPSBUQUJfUkVOREVSRVJTW0NVUlJFTlRfVEFCXTsKICAgIGlmIChmbikgeyB0cnkg"
    "eyBmbihyb3dzKTsgUkVOREVSRURbQ1VSUkVOVF9UQUJdID0gdHJ1ZTsgfSBjYXRjaCAoZSkgeyBjb25zb2xlLmVycm9yKCJyZW5kZXIgIiArIENVUlJFTlRf"
    "VEFCLCBlKTsgfSB9CiAgfQp9CmZ1bmN0aW9uIHN3aXRjaFRhYih0YWIpIHsKICBDVVJSRU5UX1RBQiA9IHRhYjsKICAkJCgiLnRhYi1idG4iKS5mb3JFYWNo"
    "KGIgPT4gYi5jbGFzc0xpc3QudG9nZ2xlKCJhY3RpdmUiLCBiLmRhdGFzZXQudGFiID09PSB0YWIpKTsKICAkJCgiLnRhYi1wYW5lbCIpLmZvckVhY2gocCA9"
    "PiBwLmNsYXNzTGlzdC50b2dnbGUoImFjdGl2ZSIsIHAuaWQgPT09ICJ0YWItIiArIHRhYikpOwogIHJlZnJlc2goKTsKfQoKLyogLS0tLS0tLS0tLS0tLS0t"
    "LSAxNC4gQk9PVCAtLS0tLS0tLS0tLS0tLS0tICovCmZ1bmN0aW9uIHdpcmVFdmVudHMoKSB7CiAgLy8gdGFicwogICQkKCIudGFiLWJ0biIpLmZvckVhY2go"
    "YiA9PiBiLm9uY2xpY2sgPSAoKSA9PiBzd2l0Y2hUYWIoYi5kYXRhc2V0LnRhYikpOwogIC8vIGdsb2JhbCBzZWFyY2ggKGRlYm91bmNlZCkKICBjb25zdCBn"
    "cyA9ICQoIiNnbG9iYWxTZWFyY2giKTsgbGV0IGRlYjsKICBncy5vbmlucHV0ID0gKCkgPT4gewogICAgJCgiI3NlYXJjaENsZWFyIikuaGlkZGVuID0gIWdz"
    "LnZhbHVlOwogICAgY2xlYXJUaW1lb3V0KGRlYik7IGRlYiA9IHNldFRpbWVvdXQoKCkgPT4geyBGSUxURVJTLnNlYXJjaCA9IGdzLnZhbHVlOyBvbkZpbHRl"
    "ckNoYW5nZSgpOyB9LCAyMjApOwogIH07CiAgJCgiI3NlYXJjaENsZWFyIikub25jbGljayA9ICgpID0+IHsgZ3MudmFsdWUgPSAiIjsgRklMVEVSUy5zZWFy"
    "Y2ggPSAiIjsgJCgiI3NlYXJjaENsZWFyIikuaGlkZGVuID0gdHJ1ZTsgb25GaWx0ZXJDaGFuZ2UoKTsgfTsKICAvLyByZXNldAogICQoIiNyZXNldEJ0biIp"
    "Lm9uY2xpY2sgPSAoKSA9PiB7IFBPX1NFQVJDSCA9ICIiOyB3aW5kb3cuX19zZWxDdXN0b21lciA9IG51bGw7IHdpbmRvdy5fX2NtcFllYXJDdXIgPSBudWxs"
    "OyB3aW5kb3cuX19jbXBZZWFyUHJldiA9IG51bGw7IHJlc2V0RmlsdGVycygpOyB9OwogIC8vIHRoZW1lCiAgJCgiI3RoZW1lVG9nZ2xlIikub25jbGljayA9"
    "IHRvZ2dsZVRoZW1lOwogIC8vIGV4cG9ydCBtZW51CiAgY29uc3QgZW0gPSAkKCIjZXhwb3J0TWVudSIpOwogICQoIiNleHBvcnRCdG4iKS5vbmNsaWNrID0g"
    "KGUpID0+IHsgZS5zdG9wUHJvcGFnYXRpb24oKTsgZW0uaGlkZGVuID0gIWVtLmhpZGRlbjsgfTsKICBkb2N1bWVudC5hZGRFdmVudExpc3RlbmVyKCJjbGlj"
    "ayIsIChlKSA9PiB7IGlmICghZS50YXJnZXQuY2xvc2VzdCgiLmRyb3Bkb3duIikpIGVtLmhpZGRlbiA9IHRydWU7IH0pOwogIGVtLnF1ZXJ5U2VsZWN0b3JB"
    "bGwoImJ1dHRvbiIpLmZvckVhY2goYiA9PiBiLm9uY2xpY2sgPSAoKSA9PiB7CiAgICBlbS5oaWRkZW4gPSB0cnVlOwogICAgY29uc3QgayA9IGIuZGF0YXNl"
    "dC5leHBvcnQ7CiAgICBpZiAoayA9PT0gInN1bW1hcnkiKSBleHBvcnRTdW1tYXJ5KCk7CiAgICBlbHNlIGlmIChrID09PSAicHJpbnQiKSBwcmludFZpZXco"
    "KTsKICAgIGVsc2UgZXhwb3J0RGF0YXNldChrKTsKICB9KTsKICAvLyBtb2RhbCBjbG9zZQogICQoIiNtb2RhbCIpLnF1ZXJ5U2VsZWN0b3JBbGwoIltkYXRh"
    "LWNsb3NlXSIpLmZvckVhY2goeCA9PiB4Lm9uY2xpY2sgPSAoKSA9PiAkKCIjbW9kYWwiKS5oaWRkZW4gPSB0cnVlKTsKICBkb2N1bWVudC5hZGRFdmVudExp"
    "c3RlbmVyKCJrZXlkb3duIiwgKGUpID0+IHsgaWYgKGUua2V5ID09PSAiRXNjYXBlIikgJCgiI21vZGFsIikuaGlkZGVuID0gdHJ1ZTsgfSk7Cn0KCmFzeW5j"
    "IGZ1bmN0aW9uIGJvb3QoKSB7CiAgaW5pdFRoZW1lKCk7CiAgdHJ5IHsgYXdhaXQgbG9hZERhdGEoKTsgfQogIGNhdGNoIChlKSB7IHJldHVybjsgfSAgICAg"
    "ICAgICAgICAgIC8vIGxvYWRlciBzaG93cyB0aGUgZXJyb3IgbWVzc2FnZQogIGJ1aWxkU2xpY2VycygpOwogIHdpcmVFdmVudHMoKTsKICAkKCIjbG9hZGVy"
    "IikuaGlkZGVuID0gdHJ1ZTsKICByZWZyZXNoKHRydWUpOwp9CmRvY3VtZW50LmFkZEV2ZW50TGlzdGVuZXIoIkRPTUNvbnRlbnRMb2FkZWQiLCBib290KTsK"
)


def _dec(b):
    return _b64.b64decode(b.encode("ascii")).decode("utf-8")


def build_index_html(data_json_str, html_out="index.html"):
    """Assemble a single, self-contained index.html with data + CSS + JS inlined.

    The only external dependency is the Chart.js CDN (loads over https, which
    works fine from file://). No web server is needed to open the result.
    """
    shell = _dec(_SHELL_B64)
    css = _dec(_CSS_B64)
    js = _dec(_JS_B64)
    # Neutralise any "</" so embedded strings can never close the <script> tag.
    safe = data_json_str.replace("</", "<\\/")
    shell = shell.replace(
        '  <link rel="stylesheet" href="styles.css" />',
        "  <style>\n" + css + "\n  </style>",
    )
    data_block = (
        '  <script type="application/json" id="et-data">' + safe + "</script>\n"
    )
    shell = shell.replace(
        '  <script src="script.js"></script>',
        data_block + "  <script>\n" + js + "\n  </script>",
    )
    with open(html_out, "w", encoding="utf-8") as fh:
        fh.write(shell)
    return html_out


def main():
    ap = argparse.ArgumentParser(description="Convert RFQ Excel/CSV to data.json")
    ap.add_argument("source", nargs="?", help="Path to .xlsx/.xls/.csv (auto-detected if omitted)")
    ap.add_argument("--sheet", default="RFQ Tracker", help="Worksheet name (Excel only)")
    ap.add_argument("--out", default="data.json", help="Output JSON path")
    ap.add_argument("--html", default="index.html", help="Self-contained dashboard HTML output path")
    args = ap.parse_args()

    src = args.source or find_source()
    if not src or not os.path.exists(src):
        sys.exit("ERROR: No source file found. Pass a filename, e.g.  python convert.py data.xlsx")
    convert(src, args.sheet, args.out, args.html)


if __name__ == "__main__":
    main()
