# ET RFQ & PO Controller

A **self-contained management dashboard** covering the RFQ → quotation → award → customer-PO →
profitability lifecycle for ET's gas-turbine parts trading business. The dashboard runs entirely in
the browser (vanilla JavaScript + Chart.js). The finished `index.html` has the data, styling and
logic **all embedded inside it**, so it **opens by double-clicking the file** — no web server,
backend, database or API is required.

---

## 1. The four files

| File | Purpose |
|------|---------|
| **`index.html`** | The complete dashboard. Data + styles + logic are all inside this one file. **Double-click it to open.** (~13 MB — that size is normal; it contains the full dataset.) |
| **`convert.py`** | Rebuilds the dashboard from your Excel/CSV. One run regenerates **both** `data.json` and a fresh self-contained `index.html`. |
| **`data.json`** | The cleaned, normalized data on its own (records + a `meta` block with distinct values, completeness and reconciliation). Provided for reuse/auditing; the dashboard itself does **not** need it because the same data is already embedded in `index.html`. |
| **`README.md`** | This document. |

---

## 2. Open the dashboard

**Just double-click `index.html`.** It opens in your default browser with no server and no setup.

The only thing loaded from the internet is the Chart.js charting library (over `https`), so the
viewer needs an internet connection for the charts to render. Everything else — all data and
calculations — is inside the file and works offline.

---

## 3. What it does

Nine linked tabs, all driven by shared global slicers, a universal search box, and a light/dark theme:

1. **Executive Overview** — headline KPIs, monthly trends, funnel, result/status donuts, top
   customers/suppliers/POCs, RFQ ageing, and an automatic management-alerts panel.
2. **RFQ Tracker** — RFQ KPIs, response-time and deadline compliance, win-rate by
   customer/POC/category, inferred lost-reason analysis, and a full drill-down RFQ table.
3. **Supplier & Procurement** — procurement value, supplier selection frequency, quote
   responsiveness, a data-available supplier score, and a supplier table.
4. **Customer PO & Delivery** — dedicated PO-number search, PO value / OA-processing KPIs and charts,
   and a per-PO detail table.
5. **ET POC Performance** — per-POC volume, win rate, conversion, profitability, responsiveness and a
   **balanced internal indicator** (not an appraisal), with ranking and quartiles.
6. **Quarterly / Yearly** — year-vs-year variance table with favourable-direction colouring, plus YoY
   monthly charts.
7. **Customer Analysis** — searchable customer selector, multi-year trends, contribution %,
   category/supplier/POC mix, and rule-based customer insights.
8. **Risk Analysis** — risk engine that scores every RFQ (0–100, Low→Critical), risk distribution
   charts, exposure, and a drill-down risk register with recommended actions.
9. **Data Quality & Controls** — completeness matrix, DQ counters, and a downloadable issue table.

Every chart, KPI and table respects the current filters. Charts are click-to-filter; tables sort /
search / paginate / toggle columns / export CSV / open a drill-down modal; and the whole view can be
exported (CSV per dataset, an HTML management summary, and a print/PDF view).

---

## 4. Rebuild from a new/updated Excel file

You only need this when the underlying data changes.

```bash
pip install pandas openpyxl

python convert.py                                   # auto-detects the Excel/CSV in the folder
python convert.py "RFQ Tracker - 2024 (New).xlsx"   # or pass a specific file
python convert.py "data.xlsx" --sheet "RFQ Tracker" # optional: name the worksheet
```

Each run **regenerates both `data.json` and a brand-new self-contained `index.html`** and prints a
conversion summary (rows processed, unique RFQs, customer POs, suppliers, customers, invalid dates,
missing mandatory values, and the two output locations). After it finishes, just **re-open
`index.html`** (hard-refresh if it was already open) — the new data is already baked in.

Optional flags: `--out <path>` (data.json location), `--html <path>` (dashboard location).

### Source-file notes
* **Accepted formats:** `.xlsx`, `.xls`, `.csv`. The worksheet is auto-detected (falls back to the first sheet).
* **Header matching:** columns are matched by **heading text** (normalized), not by column letter, so
  slight wording differences still map correctly.
* **Dates** become ISO `YYYY-MM-DD` (unparseable → `null`). **Numbers** are cleaned of currency signs,
  separators and `%` (invalid → `null`); GP% is normalized so 26.5 = 26.5%.

---

## 5. Optional: host it online

Because `index.html` is fully self-contained, you can host it by simply putting **that one file**
anywhere that serves static files (GitHub Pages, SharePoint, any web server, or a shared drive):

1. Create a GitHub repository and upload `index.html` (the other files are optional).
2. Commit, then open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, select **main** / **/(root)**, and **Save**.
4. Wait a minute and open the Pages URL. To update later, re-run `convert.py`, replace `index.html`, and push.

---

## 6. Key assumptions

* **Calendar days.** All response, ageing and processing metrics use calendar days (no working-day
  calendar exists in the source).
* **Deduplication.** The data is line-item based. For each financial total, within one RFQ/PO group:
  if the same total repeats *identically* across every line **and** there is more than one line, it is
  counted **once**; otherwise line values are **summed**. Unique RFQs use `Customer + RFQ No.`
  (falling back to `S.No. + Customer + RFQ Date`); customer POs use `Customer + PO No.`; supplier POs
  use `Supplier + PO No.`.
* **Customer sales / PO value.** No dedicated customer-PO value column exists, so **PO value = the ET
  quoted value linked to the PO** (dedup-aware), labelled "Customer PO / Sales Value" throughout.
* **COGS / gross profit.** Supplier cost = total supplier price (dedup-aware). Missing GP is calculated
  as `ET Quoted Value − Supplier Price`; margin = `GP ÷ Quoted Value × 100` (never divided by zero).
* **Lost reasons are inferred** from status, supplier participation and deadline signals, and labelled
  **Confirmed / Inferred / Unknown** — an inferred reason is never presented as confirmed.
* **Scores** (supplier, ET POC, risk) are **internal management indicators**, reweighted for available
  data — not audited appraisals.

---

## 7. Known source-data limitation (important)

In the **current** source file the entire **delivery / shipment / supplier-PO block is empty**:

* Customer delivery timeline — Customer Required Date, ET Promised Date, ET RTS Date, ET Actual Ship
  Date (columns **AZ–BC**).
* Supplier PO & shipment — Supplier PO No., PO/required/promised/RTS/actual-ship dates and Shipment
  Final Status (columns **BD–BJ**).

Because of this, the dashboard **honestly shows empty states** for delivery status, on-time %,
lead-time, delay ageing and supplier-shipment metrics rather than fabricating them. (This is also why
the risk engine currently flags only Low/Moderate items — the high-severity triggers depend on
delivery dates that aren't present.) The RFQ, quotation, award, customer-PO, profitability, POC, risk
and data-quality analytics are fully populated. **As soon as those columns are filled in the source
and `convert.py` is re-run, the delivery and shipment sections populate automatically** — no code
change needed.

Two data-quality items in the current source: **405 rows** have no customer name and **509 rows** have
no RFQ number. Both are listed with row-level detail in the **Data Quality** tab and can be exported.

---

## 8. Troubleshooting

* **Nothing happens / blank page on open** — make sure you opened `index.html` itself. If it was
  regenerated while open, hard-refresh (Ctrl/Cmd+Shift+R).
* **Charts don't appear** — Chart.js loads from the internet; confirm the viewer is online.
* **`ModuleNotFoundError` running convert.py** — run `pip install pandas openpyxl`.
* **Invalid Excel format** — ensure a real `.xlsx/.xls/.csv`; pass the filename explicitly if
  auto-detect picks the wrong file.
* **Totals look doubled** — the dedup logic specifically prevents this; if a *new* source repeats
  totals unusually, check the Data Quality tab for aggregation-ambiguity flags.
* **`index.html` is ~13 MB** — expected; it embeds the full dataset so it can run with no server.

---

## 9. Reconciliation summary (current source)

Verified by the converter and re-checked live in the dashboard:

| Measure | Value |
|---|---|
| Source rows | 16,476 |
| Processed rows | 16,078 |
| Unique RFQs | 3,124 |
| Unique quoted RFQs | 1,970 |
| Unique won RFQs | 493 |
| Unique customer POs | 459 |
| Unique supplier POs | 0 *(supplier-PO columns empty)* |
| Total ET quoted value | $353,856,559.51 |
| Total supplier cost | $266,135,171.56 |
| Total gross profit | $96,029,276.45 |

Dashboard totals reconcile to these figures exactly under the "All" filter.

---

*All monetary values are USD. Day/delay metrics are calendar days. Financial totals use dedup-aware
aggregation. Scores are internal management indicators, not audited appraisals.*
