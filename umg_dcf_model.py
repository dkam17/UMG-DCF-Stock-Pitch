"""
Universal Music Group (UMG) - DCF valuation model

Builds UMG_DCF_Model.xlsx:
  1. Historical Financials
  2. Assumptions
  3. Financial Forecast (Bear / Base / Bull, 2027E-2031E)
  4. DCF Valuation
  5. Scenarios and Sensitivity

All calculations are written as live Excel formulas.

Usage:  python umg_dcf_model.py   (requires openpyxl)
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

OUTPUT = "UMG_DCF_Model.xlsx"

HIST = "Historical Financials"
ASM = "Assumptions"
FC = "Financial Forecast"
DCF = "DCF Valuation"
DASH = "Scenarios and Sensitivity"


def x(sheet, ref):
    return f"'{sheet}'!{ref}"


# --- formatting -------------------------------------------------------------
FONT = "Arial"
NAVY = "1F3864"
BLUE = "0000FF"
GREEN = "008000"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
SECTION_FILL = PatternFill("solid", fgColor=NAVY)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
KPI_FILL = PatternFill("solid", fgColor="EEF3FA")
THIN = Side(style="thin", color="808080")
TOP_BORDER = Border(top=THIN)
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
PCT3 = '0.000%;(0.000%);"-"'
EUR = '"€"#,##0.00;("€"#,##0.00);"-"'
BETA = "0.00"
MULT = '0.0"x"'
FACTOR = "0.0000"

YEARS = ["2027E", "2028E", "2029E", "2030E", "2031E"]
YCOLS = ["D", "E", "F", "G", "H"]
SCEN = ["Bear", "Base", "Bull"]
SCOL = {"Bear": "B", "Base": "C", "Bull": "D"}


def style(c, kind="formula", fmt=None, bold=False, size=10, italic=False, wrap=False, align=None):
    color = {"input": BLUE, "data": BLUE, "link": GREEN}.get(kind, "000000")
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if kind == "input":
        c.fill = INPUT_FILL
    if fmt:
        c.number_format = fmt
    if wrap or align:
        c.alignment = Alignment(wrap_text=wrap, horizontal=align, vertical="center" if not wrap else "top")


def put(ws, ref, value, kind="formula", fmt=None, bold=False, **kw):
    c = ws[ref]
    c.value = value
    style(c, kind, fmt, bold, **kw)
    return c


def label(ws, ref, text, bold=False, indent=0, italic=False, size=10, color="000000"):
    c = ws[ref]
    c.value = text
    c.font = Font(name=FONT, size=size, bold=bold, italic=italic, color=color)
    if indent:
        c.alignment = Alignment(indent=indent)
    return c


def note(ws, ref, text):
    if not text:
        return None
    c = ws[ref]
    c.value = text
    c.font = Font(name=FONT, size=9, italic=True, color="595959")
    c.alignment = Alignment(wrap_text=False, vertical="center")
    return c


def section(ws, row, text, first_col="A", last_col="I"):
    for col in range(ord(first_col), ord(last_col) + 1):
        ws[f"{chr(col)}{row}"].fill = SECTION_FILL
    c = ws[f"{first_col}{row}"]
    c.value = text
    c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")


def header_row(ws, row, values):
    for col, text in values:
        c = ws[f"{col}{row}"]
        c.value = text
        c.font = Font(name=FONT, size=10, bold=True)
        c.fill = SUB_FILL
        c.alignment = Alignment(horizontal="center" if col != "A" else "left")


def title(ws, text, subtitle, last_col="I"):
    for col in range(1, ord(last_col) - 64 + 1):
        ws.cell(row=1, column=col).fill = SECTION_FILL
    c = ws["A1"]
    c.value = text
    c.font = Font(name=FONT, size=14, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 24
    s = ws["A2"]
    s.value = subtitle
    s.font = Font(name=FONT, size=9, italic=True, color="404040")


def total_border(ws, row, cols):
    for col in cols:
        ws[f"{col}{row}"].border = TOP_BORDER


wb = Workbook()
ws_h = wb.active
ws_h.title = HIST
ws_a = wb.create_sheet(ASM)
ws_f = wb.create_sheet(FC)
ws_d = wb.create_sheet(DCF)
ws_s = wb.create_sheet(DASH)

for ws in wb.worksheets:
    ws.sheet_view.showGridLines = False

# ============================================================================
# 1. HISTORICAL FINANCIALS
# ============================================================================
ws = ws_h
title(ws, "Universal Music Group N.V. - Historical Financials",
      "EUR m unless stated. IFRS reported unless marked Adjusted (non-IFRS, reference only).",
      last_col="G")
for col, w in zip("ABCDEFG", [58, 12, 12, 12, 12, 9, 70]):
    ws.column_dimensions[col].width = w

header_row(ws, 4, [("A", "EUR m"), ("B", "FY2023"), ("C", "FY2024"),
                   ("D", "FY2025"), ("E", "3-yr avg"), ("F", "Unit"), ("G", "Source")])
ws.freeze_panes = "B5"

H = {}
r = 5


def hist_data(key, text, vals, unit="€m", src="", fmt=NUM, indent=0):
    global r
    H[key] = r
    label(ws, f"A{r}", text, indent=indent)
    for col, v in zip("BCD", vals):
        if v is not None:
            put(ws, f"{col}{r}", v, "data", fmt)
    label(ws, f"F{r}", unit)
    note(ws, f"G{r}", src)
    r += 1


def hist_formula(key, text, fn, unit="%", src="", fmt=PCT, avg=False, cols="BCD", bold=False, indent=0):
    global r
    H[key] = r
    label(ws, f"A{r}", text, bold=bold, indent=indent)
    for col in cols:
        put(ws, f"{col}{r}", fn(col), "formula", fmt, bold=bold)
    if avg:
        put(ws, f"E{r}", f"=AVERAGE({cols[0]}{r}:D{r})", "formula", fmt, bold=True)
    label(ws, f"F{r}", unit)
    note(ws, f"G{r}", src)
    r += 1


prev = {"C": "B", "D": "C"}

section(ws, r, "INCOME STATEMENT", "A", "G"); r += 1
hist_data("rev", "Revenue", [11108, 11834, 12507], src="S1, S2")
hist_formula("growth", "Revenue growth", lambda c: f"={c}{H['rev']}/{prev[c]}{H['rev']}-1", cols="CD", avg=True,
             src="FY2025 +8.7% in constant currency (S1)")
hist_data("ebitda", "EBITDA (reported)", [1808, 2332, 2538], src="S1, S2. After share-based compensation")
hist_formula("ebitda_m", "EBITDA margin", lambda c: f"={c}{H['ebitda']}/{c}{H['rev']}", avg=True)
hist_data("adj_ebitda", "Adjusted EBITDA (non-IFRS, reference)", [2369, 2661, 2810],
          src="S1, S2. FY2023 restated. Not used in valuation")
hist_formula("adj_m", "Adjusted EBITDA margin", lambda c: f"={c}{H['adj_ebitda']}/{c}{H['rev']}")
hist_data("ebit", "Operating profit (EBIT)", [1418, 1775, 1998], src="S1, S2")
hist_formula("ebit_m", "EBIT margin", lambda c: f"={c}{H['ebit']}/{c}{H['rev']}", avg=True)
hist_data("da", "Depreciation and amortisation", [382, 409, 446],
          src="S1, S2. Incl. catalogue amortisation and right-of-use assets")
hist_data("cat_amort", "of which: catalogue amortisation", [236, 245, 256], indent=2, src="S1, S2")
hist_formula("gap", "Other items in EBITDA (EBITDA - EBIT - D&A)",
             lambda c: f"={c}{H['ebitda']}-{c}{H['ebit']}-{c}{H['da']}", unit="€m", fmt=NUM,
             src="Mainly impairments")
hist_data("pbt", "Profit before tax", [1721, 2871, 2086], src="S1, S2. Incl. Spotify revaluation")
hist_data("tax", "Income tax expense", [458, 778, 546], src="S1, S2")
hist_formula("etr", "Effective tax rate", lambda c: f"={c}{H['tax']}/{c}{H['pbt']}", avg=True,
             src="Dutch statutory rate 25.8%")
r += 1

section(ws, r, "CASH FLOW ITEMS", "A", "G"); r += 1
hist_data("capex", "Capital expenditure", [47, 91, 70], src="S1, S2")
hist_data("cat_inv", "Catalogue investments", [178, 266, 345], src="S1, S2. FY2025 net of divestments: €280m")
hist_data("oth_int", "Other intangible investments (mainly software)", [74, 92, 125], src="S1, S2")
hist_data("roy", "Royalty advances, net of recoupments", [100, 186, 402],
          src="S1, S2. Treated as investment, not working capital")
hist_formula("alloc", "Total catalogue & intangible investment",
             lambda c: f"={c}{H['cat_inv']}+{c}{H['oth_int']}+{c}{H['roy']}", unit="€m", fmt=NUM, bold=True)
total_border(ws, H["alloc"], "BCD")
hist_data("lease_rep", "Lease principal repayments", [80, 81, 72], src="S1, S2")
hist_data("lease_int", "Lease interest paid", [14, 21, 21], src="S1, S2")
hist_formula("lease", "Total lease payments", lambda c: f"={c}{H['lease_rep']}+{c}{H['lease_int']}",
             unit="€m", fmt=NUM, bold=True, src="Leases excluded from net debt")
total_border(ws, H["lease"], "BCD")
hist_data("nwc_cf", "Other changes in net working capital (cash flow)", [164, -5, 38],
          src="S1, S2. Excl. royalty advances")
hist_formula("nwc_inc", "Increase / (decrease) in NWC", lambda c: f"=-{c}{H['nwc_cf']}", unit="€m", fmt=NUM)
hist_data("ocf", "Operating cash flow before tax paid", [2278, 2104, 2142], src="S1, S2")
hist_data("umg_fcf", "Free cash flow (UMG definition, levered)", [1082, 523, 702], src="S1, S2")
r += 1

section(ws, r, "RATIOS (% OF REVENUE)", "A", "G"); r += 1
for key, text, num in [("r_da", "D&A", "da"),
                       ("r_cat", "Catalogue investments", "cat_inv"),
                       ("r_oth", "Software and other intangibles", "oth_int"),
                       ("r_roy", "Royalty advances (net)", "roy"),
                       ("r_capex", "Capex", "capex"),
                       ("r_alloc", "Total catalogue & intangible investment", "alloc"),
                       ("r_lease", "Lease payments", "lease"),
                       ("r_nwc", "Increase in NWC", "nwc_inc")]:
    hist_formula(key, text, lambda c, n=num: f"={c}{H[n]}/{c}{H['rev']}", avg=True)
r += 1

section(ws, r, "BALANCE SHEET AND SHARES (31 DECEMBER)", "A", "G"); r += 1
hist_data("debt", "Borrowings", [2104, 2651, 2846], src="S1, S2. Excl. lease liabilities")
hist_data("cash", "Cash and cash equivalents", [413, 553, 451], src="S1, S2. Net debt table")
hist_data("netdebt", "Financial net debt (reported)", [1689, 2098, 2390], src="S1, S2")
hist_formula("nd_other", "Other net debt items (derivatives etc.)",
             lambda c: f"={c}{H['debt']}-{c}{H['cash']}-{c}{H['netdebt']}", unit="€m", fmt=NUM)
hist_formula("lev", "Net debt / EBITDA", lambda c: f"={c}{H['netdebt']}/{c}{H['ebitda']}", unit="x", fmt=MULT,
             src="Reported 0.9x at FY2024 and FY2025 (S1)")
hist_data("shares_w", "Weighted average basic shares", [1819, 1827, 1833], unit="m", fmt=NUM, src="S1, S2")
r += 1

section(ws, r, "2026 INTERIM RESULTS (UNAUDITED)", "A", "G"); r += 1
header_row(ws, r, [("A", "Item"), ("B", "Value"), ("F", "Unit"), ("G", "Source")]); r += 1


def single(key, text, v, unit="€m", src="", fmt=NUM, kind="data"):
    global r
    H[key] = r
    label(ws, f"A{r}", text)
    put(ws, f"B{r}", v, kind, fmt)
    label(ws, f"F{r}", unit)
    note(ws, f"G{r}", src)
    r += 1


single("h1_rev", "H1 2026 revenue", 6194, src="S3. Downtown consolidated from 20 Feb 2026")
single("h1_rev25", "H1 2025 revenue", 5881, src="S3")
single("h1_ebitda", "H1 2026 EBITDA", 1181, src="S3")
single("h1_adj", "H1 2026 Adjusted EBITDA", 1310, src="S3")
single("h1_ebit", "H1 2026 EBIT", 901, src="S3")
single("h1_ebit_m", "H1 2026 EBIT margin", f"=B{H['h1_ebit']}/B{H['h1_rev']}", unit="%", fmt=PCT, kind="formula")
single("h1_roy", "H1 2026 royalty advances (net)", 292, src="S3")
single("h1_roy25", "H1 2025 royalty advances (net)", 377, src="S3")
single("h1_fcf", "H1 2026 free cash flow (UMG definition)", 24, src="S3")
single("h1_nd", "Financial net debt, 30 Jun 2026", 4131, src="S3")
single("m9_rev", "9M 2026 revenue", 9588, src="S4")
single("m9_rev25", "9M 2025 revenue", 8902, src="S4")
single("m9_g", "9M 2026 revenue growth", f"=B{H['m9_rev']}/B{H['m9_rev25']}-1", unit="%", fmt=PCT,
       kind="formula", src="+11.1% constant currency (S4)")
single("m9_org", "9M 2026 growth ex-Downtown, constant currency", 0.055, unit="%", fmt=PCT, src="S4")
single("q3_rev", "Q3 2026 revenue", 3394, src="S4")
single("q3_g", "Q3 2026 revenue growth", 0.123, unit="%", fmt=PCT, src="S4. +5.3% ex-Downtown (S16)")
single("q4_rev25", "Q4 2025 revenue", 3605, src="S1. Incl. one-off legal resolution benefit")
single("dt_rev", "Downtown revenue, 9M 2026", 484, src="S5")
single("dt_ebitda", "Downtown Adjusted EBITDA, 9M 2026", 26, src="S5")
single("dt_m", "Downtown Adjusted EBITDA margin", f"=B{H['dt_ebitda']}/B{H['dt_rev']}", unit="%", fmt=PCT,
       kind="formula")
single("eps_b", "FY2025 EPS, basic", 0.84, unit="€", fmt=EUR, src="S1")
single("eps_d", "FY2025 EPS, diluted", 0.83, unit="€", fmt=EUR, src="S1")
single("dil", "Implied dilution", f"=B{H['eps_b']}/B{H['eps_d']}-1", unit="%", fmt=PCT, kind="formula")
r += 1

section(ws, r, "2026 SHARE BUYBACKS AND DIVIDENDS", "A", "G"); r += 1
single("bb1_sh", "Programme 1 (Mar-Jul 2026): shares repurchased", 26.707529, unit="m", fmt='#,##0.000', src="S6")
single("bb1_cost", "Programme 1: total cost", 499.2, fmt=NUM1, src="S6")
single("bb1_h1", "Programme 1: spent by 30 Jun 2026", 485, src="S6")
single("bb2_sh", "Pershing Square block purchase (4 Jun 2026): shares", 14.156285, unit="m", fmt='#,##0.000',
       src="S6. €250m at €17.66")
single("bb3_sh", "Programme 2 (Aug-Sep 2026): shares repurchased", 16.552027, unit="m", fmt='#,##0.000', src="S6")
single("bb3_cost", "Programme 2: total cost", 250.0, fmt=NUM1, src="S6")
single("bb_tot", "Total shares repurchased in 2026",
       f"=B{H['bb1_sh']}+B{H['bb2_sh']}+B{H['bb3_sh']}", unit="m", fmt='#,##0.000', kind="formula")
single("div_int", "H1 2026 interim dividend (paid Oct 2026)", 432, src="S3. €0.24/share, ex-div 5 Oct 2026")
r += 1

section(ws, r, "CONSENSUS, COMPANY TARGETS AND INDUSTRY DATA", "A", "G"); r += 1
header_row(ws, r, [("A", "Item"), ("B", "Value"), ("F", "Unit"), ("G", "Source")]); r += 1
single("cons26", "Consensus revenue, 2026E", 13593, src="S14")
single("cons27", "Consensus revenue, 2027E", 14623, src="S14")
single("cons_g", "Consensus revenue growth, 2027E", f"=B{H['cons27']}/B{H['cons26']}-1", unit="%", fmt=PCT,
       kind="formula")
single("cmd_rev", "UMG target: revenue CAGR to FY2028", 0.07, unit="%", fmt=PCT, src="S13. 7%+")
single("cmd_ebitda", "UMG target: Adjusted EBITDA CAGR to FY2028", 0.10, unit="%", fmt=PCT, src="S13. 10%+")
single("gs_rm", "Goldman Sachs: recorded music revenue CAGR 2025-35", 0.059, unit="%", fmt=PCT, src="S15")
single("gs_sub", "Goldman Sachs: subscription streaming CAGR 2025-35", 0.071, unit="%", fmt=PCT, src="S15")
single("q3_sub", "Q3 2026 subscription growth ex-Downtown, constant currency", 0.095, unit="%", fmt=PCT, src="S16")
single("q3_rm", "Q3 2026 Recorded Music Adj. EBITDA margin ex-Downtown", 0.262, unit="%", fmt=PCT,
       src="S16. +1.2pp y/y")
single("cost27", "Incremental cost savings expected in 2027 (mid-point)", 40, src="S16. €35-45m")
single("cat_share", "Catalogue share of recorded music revenue, 2025", 0.66, unit="%", fmt=PCT, src="S17")
r += 1

section(ws, r, "SOURCES (ACCESSED 13-14 NOV 2026)", "A", "G"); r += 1
SOURCES = [
    ("S1", "UMG: Q4 and FY2025 results (5 Mar 2026), unaudited",
     "https://tools.prnewswire.com/en-us/live/20823/release/20260305EN02412"),
    ("S2", "UMG: Q4 and FY2024 results incl. FY2023 comparatives (6 Mar 2025)",
     "https://finanzen.net/nachricht/aktien/universal-music-group-n-v-reports-financial-results-for-the-fourth-quarter-and-full-year-ended-december-31-2024-14297870"),
    ("S3", "UMG: Q2 and H1 2026 results (30 Jul 2026)",
     "https://www.trivano.com/persbericht/749323-749323.html"),
    ("S4", "UMG Q3 2026 results (29 Oct 2026) - Securities.io, Music Week",
     "https://www.securities.io/universal-music-group-q3-revenue-rises-12-3-as-downtown-adds-195-million/"),
    ("S5", "Downtown 9M / Q3 2026 contribution - Ideal Investisseur, MBW (Nov 2026)",
     "https://www.ideal-investisseur.fr/en/stock-news/universal-music-group-in-the-third-quarter-downtown-weighs-on-margins-despite-growth-in-subscriptions/28426.html"),
    ("S6", "2026 share buybacks - MBW (Sep 2026), UMG (7 Sep 2026), TheWrap (Jun 2026)",
     "https://www.musicbusinessworldwide.com/umg-has-now-bought-back-e1-billion-of-its-shares-this-year/"),
    ("S7", "Share price: Euronext Amsterdam close, 10 Nov 2026 (MarketScreener)",
     "https://www.marketscreener.com/quote/stock/UNIVERSAL-MUSIC-GROUP-N-V-127080831/"),
    ("S8", "German 10-year Bund yield (Trading Economics, 4 Nov 2026)",
     "https://tradingeconomics.com/germany/government-bond-yield/news/598117"),
    ("S9", "Damodaran, Betas by Sector (Global), Jan 2026 - Entertainment",
     "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/BetasGlobal.html"),
    ("S10", "Damodaran, implied equity risk premium, 1 Nov 2026",
     "https://pages.stern.nyu.edu/~adamodar/New_Home_Page/home.htm"),
    ("S11", "UMG €1.0bn Eurobond issue: 3.375% 2030 / 4.125% 2036 (9 Jun 2026)",
     "https://tools.prnewswire.com/en-us/live/20813/release/20260609EN79622"),
    ("S12", "UMG residual Spotify stake (MBW, Aug 2026)",
     "https://www.musicbusinessworldwide.com/universal-has-sold-a-chunk-of-spotify-stock-for-around-467m-artists-will-likely-be-paid-over-100m-of-it/"),
    ("S13", "UMG Capital Markets Day targets through FY2028 (17 Sep 2024)",
     "https://prnewswire.com/news-releases/universal-music-group-nv-provides-new-financial-targets-through-fy-2028-ahead-of-its-2024-capital-markets-day-302249686.html"),
    ("S14", "Analyst consensus revenue (Stockopedia, 14 Nov 2026)",
     "https://www.stockopedia.com/share-prices/universal-music-nv-AMS:UMG/"),
    ("S15", "Goldman Sachs 'Music in the Air' 2026 (via Hypebot)",
     "https://www.hypebot.com/goldman-sachs-2026-music-report-the-monetization-era/"),
    ("S16", "UMG Q3 2026 earnings call (MBW; GuruFocus)",
     "https://www.musicbusinessworldwide.com/umgs-share-price-closes-up-7-9-adding-around-2-5b-in-value-after-q3-subscription-growth-beats-forecasts/"),
    ("S17", "Catalogue share of UMG recorded music revenue (MBW, May 2026)",
     "https://www.musicbusinessworldwide.com/universal-spotify-virgin-and-the-leakage-causing-analyst-discomfort/"),
]
for code, desc, url in SOURCES:
    label(ws, f"A{r}", f"{code}: {desc}")
    c = ws[f"G{r}"]
    c.value = url
    c.hyperlink = url
    c.font = Font(name=FONT, size=9, color="0563C1", underline="single")
    r += 1


def h(key, col="D"):
    return x(HIST, f"${col}${H[key]}")


# ============================================================================
# 2. ASSUMPTIONS
# ============================================================================
ws = ws_a
title(ws, "Assumptions", "EUR m unless stated.", last_col="I")
for col, w in zip("ABCDEFGHI", [52, 11, 11, 11, 11, 11, 11, 11, 80]):
    ws.column_dimensions[col].width = w

put(ws, "A3", "Input", "input")
put(ws, "B3", "Formula", "formula")
put(ws, "C3", "Link", "link")
ws.freeze_panes = "A5"

A = {}
r = 5


def asm(key, text, value, kind="input", fmt=NUM, unit="", ref=None, ref_fmt=None, why=""):
    global r
    A[key] = r
    label(ws, f"A{r}", text)
    put(ws, f"B{r}", value, kind, fmt)
    label(ws, f"C{r}", unit)
    if ref is not None:
        put(ws, f"D{r}", ref, "link", ref_fmt or fmt)
    note(ws, f"I{r}", why)
    r += 1


def a(key, col="B"):
    return x(ASM, f"${col}${A[key]}")


HDR = [("A", "Item"), ("B", "Value"), ("C", "Unit"), ("D", "Reference"), ("I", "Comment")]

section(ws, r, "1. MARKET DATA AND VALUATION DATE", "A", "I"); r += 1
header_row(ws, r, HDR); r += 1
asm("val_date", "Valuation date", "31-Dec-2026", fmt="@", unit="date")
asm("price", "Share price", 17.05, fmt=EUR, unit="€", why="S7. Euronext Amsterdam close")
asm("price_date", "Share price date", "10-Nov-2026", fmt="@", unit="date", why="S7")
asm("midyear", "Mid-year convention (1 = on, 0 = off)", 1, fmt="0", unit="switch",
    why="Terminal value discounted 5 full years in both cases")
r += 1

section(ws, r, "2. 2026E REVENUE", "A", "I"); r += 1
header_row(ws, r, HDR); r += 1
asm("m9", "9M 2026 revenue (actual)", f"={h('m9_rev', 'B')}", kind="link", unit="€m", why="S4")
asm("q4_25", "Q4 2025 revenue (actual)", f"={h('q4_rev25', 'B')}", kind="link", unit="€m", why="S1")
asm("q4_g", "Q4 2026 revenue growth", 0.111, fmt=PCT, unit="%",
    ref=f"={h('q3_g', 'B')}", ref_fmt=PCT,
    why="In line with 9M constant-currency growth (S4); full quarter of Downtown. Ref = Q3 reported growth")
asm("rev26", "2026E revenue", f"=B{A['m9']}+B{A['q4_25']}*(1+B{A['q4_g']})",
    kind="formula", unit="€m", ref=f"={h('cons26', 'B')}", ref_fmt=NUM,
    why="Ref = consensus (S14). Includes Downtown")
ws[f"B{A['rev26']}"].font = Font(name=FONT, size=10, bold=True)
asm("g26", "2026E growth", f"=B{A['rev26']}/{h('rev')}-1", kind="formula", fmt=PCT, unit="%")
r += 1

section(ws, r, "3. SCENARIO ASSUMPTIONS", "A", "I"); r += 1
header_row(ws, r, [("A", "Assumption"), ("B", "Bear"), ("C", "Base"), ("D", "Bull"), ("I", "Comment")]); r += 1
SC = {}


def scen(key, text, vals, fmt, why):
    global r
    SC[key] = r
    label(ws, f"A{r}", text)
    for s_, v_ in zip(SCEN, vals):
        put(ws, f"{SCOL[s_]}{r}", v_, "input", fmt)
    note(ws, f"I{r}", why)
    r += 1


scen("wacc_adj", "WACC adjustment", [0.005, 0.0, -0.005], PCT2, "Applied to calculated WACC")
scen("tg", "Terminal growth rate", [0.025, 0.030, 0.0325], PCT2,
     "Below 3.3% risk-free rate. Industry growth 5.9% p.a. to 2035 (S15); mainly US$ revenue")
r += 1

section(ws, r, "4. ANNUAL ASSUMPTIONS", "A", "I"); r += 1
header_row(ws, r, [("A", "Assumption"), ("B", "2025A"), ("C", "3-yr avg")] +
           [(c_, y_) for c_, y_ in zip(YCOLS, YEARS)] + [("I", "Comment")]); r += 1
YA = {}


def yrow(key, text, vals, why, fmt=PCT, b=None, c=None, kind="input", bold=False, formula=None):
    global r
    YA[key] = r
    label(ws, f"A{r}", text, bold=bold)
    if b is not None:
        put(ws, f"B{r}", b, "link", fmt)
    if c is not None:
        put(ws, f"C{r}", c, "link", fmt)
    for i_, col_ in enumerate(YCOLS):
        val = formula(col_) if formula else vals[i_]
        put(ws, f"{col_}{r}", val, kind, fmt, bold=bold)
    note(ws, f"I{r}", why)
    r += 1


label(ws, f"A{r}", "Revenue growth", bold=True); r += 1
yrow("g_Bear", "  Bear", [0.050, 0.045, 0.040, 0.035, 0.030],
     "Share loss, AI dilution of royalty pools, weaker ad-supported streaming",
     b=f"={h('growth')}")
yrow("g_Base", "  Base", [0.075, 0.070, 0.065, 0.060, 0.055],
     "2027 in line with consensus (S14); 2028 at UMG 7%+ target (S13); fades toward 5.9% industry rate (S15)",
     b=f"={h('growth')}")
yrow("g_Bull", "  Bull", [0.085, 0.080, 0.075, 0.070, 0.065],
     "Streaming price increases, superfan tiers, AI licensing (S16)",
     b=f"={h('growth')}")
label(ws, f"A{r}", "EBIT margin", bold=True); r += 1
yrow("m_Bear", "  Bear", [0.150, 0.150, 0.1525, 0.1525, 0.155],
     "Close to H1 2026 level; continued Downtown dilution",
     b=f"={h('ebit_m')}", c=f"={h('ebit_m', 'E')}")
yrow("m_Base", "  Base", [0.1625, 0.166, 0.169, 0.172, 0.175],
     "€250m savings programme complete, €35-45m extra in 2027 (S16); EBITDA target growth above revenue (S13)",
     b=f"={h('ebit_m')}", c=f"={h('ebit_m', 'E')}")
yrow("m_Bull", "  Bull", [0.1675, 0.173, 0.179, 0.184, 0.19],
     "Operating leverage from subscription pricing; Downtown synergies (S16)",
     b=f"={h('ebit_m')}", c=f"={h('ebit_m', 'E')}")
label(ws, f"A{r}", "Catalogue & intangible investment (% of revenue)", bold=True); r += 1
yrow("cat", "  Catalogue acquisitions", [0.020, 0.019, 0.018, 0.017, 0.016],
     "Capital shifting toward buybacks (S6)",
     b=f"={h('r_cat')}", c=f"={h('r_cat', 'E')}")
yrow("sw", "  Software and other intangibles", [0.0085, 0.0080, 0.0080, 0.0075, 0.0075],
     "Normalises after 2025-26 platform upgrade",
     b=f"={h('r_oth')}", c=f"={h('r_oth', 'E')}")
yrow("roy", "  Royalty advances (net)", [0.022, 0.021, 0.020, 0.019, 0.018],
     "H1 2026 €292m vs €377m in H1 2025; higher recoupment (S3)",
     b=f"={h('r_roy')}", c=f"={h('r_roy', 'E')}")
yrow("inv", "  Total", None, "",
     b=f"={h('r_alloc')}", c=f"={h('r_alloc', 'E')}", kind="formula", bold=True,
     formula=lambda col_: f"={col_}{YA['cat']}+{col_}{YA['sw']}+{col_}{YA['roy']}")
total_border(ws, YA["inv"], YCOLS)
r += 1

section(ws, r, "5. OTHER OPERATING ASSUMPTIONS", "A", "I"); r += 1
header_row(ws, r, [("A", "Assumption"), ("B", "Value"), ("C", "Unit"), ("D", "3-yr avg"), ("I", "Comment")]); r += 1
asm("da", "D&A", 0.035, fmt=PCT, unit="% rev", ref=f"={h('r_da', 'E')}", ref_fmt=PCT)
asm("capex", "Capex", 0.006, fmt=PCT, unit="% rev", ref=f"={h('r_capex', 'E')}", ref_fmt=PCT)
asm("lease", "Lease payments", 0.008, fmt=PCT, unit="% rev", ref=f"={h('r_lease', 'E')}", ref_fmt=PCT,
    why="Leases excluded from net debt")
asm("nwc", "Increase in net working capital", 0.0, fmt=PCT2, unit="% rev",
    ref=f"={h('r_nwc', 'E')}", ref_fmt=PCT2, why="Historical average was a cash release")
asm("tax", "Tax rate", 0.265, fmt=PCT, unit="%", ref=f"={h('etr', 'E')}", ref_fmt=PCT)
r += 1

section(ws, r, "6. COST OF CAPITAL", "A", "I"); r += 1
header_row(ws, r, [("A", "Input"), ("B", "Value"), ("C", "Unit"), ("D", "Reference"), ("I", "Source")]); r += 1
asm("rf", "Risk-free rate (10y Bund)", 0.033, fmt=PCT2, unit="%", why="S8")
asm("beta_u", "Unlevered beta (Entertainment, Global)", 0.98, fmt=BETA, unit="x",
    ref=1.01, ref_fmt=BETA, why="S9. Ref = industry levered beta")
ws[f"D{A['beta_u']}"].font = Font(name=FONT, size=10, color=BLUE)
asm("erp", "Equity risk premium", 0.0441, fmt=PCT2, unit="%", why="S10")
asm("kd", "Pre-tax cost of debt", 0.04125, fmt=PCT3, unit="%", why="S11. 4.125% notes due 2036; Baa1/BBB+")
asm("tax_k", "Tax rate", f"=B{A['tax']}", kind="formula", fmt=PCT, unit="%")
r += 1

section(ws, r, "7. NET DEBT AND SHARE COUNT", "A", "I"); r += 1
header_row(ws, r, HDR); r += 1
asm("debt", "Borrowings, 31 Dec 2025", f"={h('debt')}", kind="link", unit="€m", why="S1")
asm("cash", "Cash, 31 Dec 2025", f"={h('cash')}", kind="link", unit="€m", why="S1")
asm("nd_oth", "Other net debt items, 31 Dec 2025", f"={h('nd_other')}", kind="link", unit="€m", why="S1")
asm("nd25", "Net debt, 31 Dec 2025", f"=B{A['debt']}-B{A['cash']}-B{A['nd_oth']}", kind="formula", unit="€m")
asm("nd_h1", "Increase in net debt, H1 2026", f"={h('h1_nd', 'B')}-B{A['nd25']}", kind="formula", unit="€m",
    why="To €4,131m at 30 Jun 2026 (S3), incl. Downtown")
asm("bb_post", "Buybacks after 30 Jun 2026",
    f"=({h('bb1_cost', 'B')}-{h('bb1_h1', 'B')})+{h('bb3_cost', 'B')}", kind="formula", unit="€m", why="S6")
asm("div_post", "Interim dividend paid Oct 2026", f"={h('div_int', 'B')}", kind="link", unit="€m", why="S3")
asm("adj_nd", "Adjusted net debt",
    f"=B{A['nd25']}+B{A['nd_h1']}+B{A['bb_post']}+B{A['div_post']}", kind="formula", unit="€m",
    why="Excludes Q3 2026 operating cash flow")
ws[f"B{A['adj_nd']}"].font = Font(name=FONT, size=10, bold=True)
asm("spot_usd", "Spotify stake - market value", 2600, fmt=NUM, unit="US$m", why="S12. As at 4 Aug 2026")
asm("fx", "USD / EUR", 1.15, fmt="0.00", unit="$/€")
asm("spot_net", "Share attributable to UMG shareholders", 0.55, fmt=PCT, unit="%",
    why="Estimate after artist share and tax (S12)")
asm("spot_on", "Include Spotify stake (1 = yes, 0 = no)", 1, fmt="0", unit="switch")
asm("spot", "Spotify stake (net)",
    f"=B{A['spot_usd']}/B{A['fx']}*B{A['spot_net']}*B{A['spot_on']}", kind="formula", unit="€m")
asm("sh_basic25", "Weighted average basic shares, FY2025", f"={h('shares_w')}", kind="link", fmt=NUM1, unit="m",
    why="S1")
asm("sh_bb", "Less: 2026 buybacks", f"={h('bb_tot', 'B')}", kind="link", fmt=NUM1, unit="m", why="S6")
asm("sh_basic", "Basic shares", f"=B{A['sh_basic25']}-B{A['sh_bb']}", kind="formula", fmt=NUM1, unit="m")
asm("dil", "Dilution allowance", 0.01, fmt=PCT, unit="%", ref=f"={h('dil', 'B')}", ref_fmt=PCT,
    why="Ref = FY2025 basic vs diluted EPS")
asm("sh_dil", "Diluted shares", f"=B{A['sh_basic']}*(1+B{A['dil']})", kind="formula", fmt=NUM1, unit="m")
ws[f"B{A['sh_dil']}"].font = Font(name=FONT, size=10, bold=True)

dv = DataValidation(type="whole", operator="between", formula1="0", formula2="1", allow_blank=False)
ws.add_data_validation(dv)
dv.add(f"B{A['spot_on']}")
dv.add(f"B{A['midyear']}")


def sc(key, s):
    return x(ASM, f"${SCOL[s]}${SC[key]}")


def ya(key, col):
    return x(ASM, f"{col}${YA[key]}")


def ya_rng(key):
    return x(ASM, f"$D${YA[key]}:$H${YA[key]}")


# ============================================================================
# 3. FINANCIAL FORECAST
# ============================================================================
ws = ws_f
title(ws, "Financial Forecast 2027E-2031E", "EUR m.", last_col="H")
for col, w in zip("ABCDEFGH", [46, 11, 11, 11, 11, 11, 11, 11]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B4"

F = {s: {} for s in SCEN}
r = 4
for s in SCEN:
    section(ws, r, f"{s.upper()} CASE", "A", "H")
    r += 1
    header_row(ws, r, [("A", "EUR m"), ("B", "2025A"), ("C", "2026E")] +
               [(c, y) for c, y in zip(YCOLS, YEARS)])
    r += 1
    R = F[s]

    def frow(key, text, fn, fmt=NUM, kind="formula", bold=False, b=None, c=None, indent=0):
        global r
        R[key] = r
        label(ws, f"A{r}", text, bold=bold, indent=indent)
        if b is not None:
            put(ws, f"B{r}", b[0], b[1], fmt, bold=bold)
        if c is not None:
            put(ws, f"C{r}", c[0], c[1], fmt, bold=bold)
        for col in YCOLS:
            put(ws, f"{col}{r}", fn(col), kind, fmt, bold=bold)
        r += 1

    prevc = {"D": "C", "E": "D", "F": "E", "G": "F", "H": "G"}
    frow("g", "Revenue growth", lambda col: f"={ya('g_' + s, col)}", fmt=PCT, kind="link",
         b=(f"={h('growth')}", "link"), c=(f"={a('g26')}", "link"))
    frow("rev", "Revenue", lambda col: f"={prevc[col]}{r}*(1+{col}{R['g']})", bold=True,
         b=(f"={h('rev')}", "link"), c=(f"={a('rev26')}", "link"))
    frow("m", "EBIT margin", lambda col: f"={ya('m_' + s, col)}", fmt=PCT, kind="link",
         b=(f"={h('ebit_m')}", "link"))
    frow("ebit", "EBIT", lambda col: f"={col}{R['rev']}*{col}{R['m']}", bold=True,
         b=(f"={h('ebit')}", "link"))
    frow("da", "D&A", lambda col: f"={col}{R['rev']}*{a('da')}", b=(f"={h('da')}", "link"))
    frow("ebitda", "EBITDA", lambda col: f"={col}{R['ebit']}+{col}{R['da']}", bold=True)
    frow("ebitda_m", "EBITDA margin", lambda col: f"={col}{R['ebitda']}/{col}{R['rev']}", fmt=PCT)
    r += 1
    label(ws, f"A{r}", "Free cash flow", bold=True); r += 1
    frow("ebit2", "EBIT", lambda col: f"={col}{R['ebit']}")
    frow("taxe", "Less: tax", lambda col: f"=-{col}{R['ebit2']}*{a('tax')}")
    frow("nopat", "NOPAT", lambda col: f"={col}{R['ebit2']}+{col}{R['taxe']}", bold=True)
    total_border(ws, R["nopat"], YCOLS)
    frow("da2", "Add: D&A", lambda col: f"={col}{R['da']}")
    frow("capex", "Less: capex", lambda col: f"=-{col}{R['rev']}*{a('capex')}")
    frow("cat", "Less: catalogue acquisitions", lambda col: f"=-{col}{R['rev']}*{ya('cat', col)}")
    frow("sw", "Less: software and other intangibles", lambda col: f"=-{col}{R['rev']}*{ya('sw', col)}")
    frow("roy", "Less: royalty advances (net)", lambda col: f"=-{col}{R['rev']}*{ya('roy', col)}")
    frow("lease", "Less: lease payments", lambda col: f"=-{col}{R['rev']}*{a('lease')}")
    frow("nwc", "Less: increase in NWC", lambda col: f"=-{col}{R['rev']}*{a('nwc')}")
    frow("fcff", "FCFF",
         lambda col: (f"={col}{R['nopat']}+{col}{R['da2']}+{col}{R['capex']}+{col}{R['cat']}+{col}{R['sw']}"
                      f"+{col}{R['roy']}+{col}{R['lease']}+{col}{R['nwc']}"),
         bold=True)
    total_border(ws, R["fcff"], YCOLS)
    frow("inv_pct", "Catalogue & intangible investment % revenue",
         lambda col: f"=-({col}{R['cat']}+{col}{R['sw']}+{col}{R['roy']})/{col}{R['rev']}", fmt=PCT)
    frow("fcff_m", "FCFF margin", lambda col: f"={col}{R['fcff']}/{col}{R['rev']}", fmt=PCT)
    frow("conv", "FCFF / EBITDA", lambda col: f"={col}{R['fcff']}/{col}{R['ebitda']}", fmt=PCT)
    r += 1


def f(s, key, col):
    return x(FC, f"${col}${F[s][key]}")


# ============================================================================
# 4. DCF VALUATION
# ============================================================================
ws = ws_d
title(ws, "DCF Valuation", "EUR m unless stated. Valuation date 31 Dec 2026.", last_col="H")
for col, w in zip("ABCDEFGH", [44, 13, 13, 13, 12, 12, 12, 12]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "B4"

D = {}
r = 4
section(ws, r, "A. COST OF CAPITAL", "A", "H"); r += 1
header_row(ws, r, [("A", "EUR m"), ("B", "Value")]); r += 1


def drow(key, text, formula, kind="formula", fmt=PCT2, bold=False):
    global r
    D[key] = r
    label(ws, f"A{r}", text, bold=bold)
    put(ws, f"B{r}", formula, kind, fmt, bold=bold)
    r += 1


drow("rf", "Risk-free rate", f"={a('rf')}", "link")
drow("bu", "Unlevered beta", f"={a('beta_u')}", "link", BETA)
drow("price", "Share price (€)", f"={a('price')}", "link", EUR)
drow("shares", "Diluted shares (m)", f"={a('sh_dil')}", "link", NUM1)
drow("E", "Market capitalisation (E)", f"=B{D['price']}*B{D['shares']}", fmt=NUM)
drow("Dv", "Net debt (D)", f"={a('adj_nd')}", "link", NUM)
drow("de", "D / E", f"=B{D['Dv']}/B{D['E']}", fmt=PCT)
drow("t", "Tax rate", f"={a('tax_k')}", "link", PCT)
drow("bl", "Levered beta", f"=B{D['bu']}*(1+(1-B{D['t']})*B{D['de']})", fmt=BETA)
drow("erp", "Equity risk premium", f"={a('erp')}", "link")
drow("ke", "Cost of equity", f"=B{D['rf']}+B{D['bl']}*B{D['erp']}", bold=True)
drow("kd", "Pre-tax cost of debt", f"={a('kd')}", "link", PCT3)
drow("kd_at", "After-tax cost of debt", f"=B{D['kd']}*(1-B{D['t']})")
drow("we", "E / (D + E)", f"=B{D['E']}/(B{D['E']}+B{D['Dv']})", fmt=PCT)
drow("wd", "D / (D + E)", f"=B{D['Dv']}/(B{D['E']}+B{D['Dv']})", fmt=PCT)
drow("wacc", "WACC", f"=B{D['we']}*B{D['ke']}+B{D['wd']}*B{D['kd_at']}", bold=True)
ws[f"B{D['wacc']}"].fill = KPI_FILL
r += 1
header_row(ws, r, [("A", "By scenario"), ("B", "Bear"), ("C", "Base"), ("D", "Bull")]); r += 1
D["sw"] = r
label(ws, f"A{r}", "WACC", bold=True)
for s, col in zip(SCEN, "BCD"):
    put(ws, f"{col}{r}", f"=$B${D['wacc']}+{sc('wacc_adj', s)}", "formula", PCT2, bold=True)
r += 1
D["sg"] = r
label(ws, f"A{r}", "Terminal growth rate")
for s, col in zip(SCEN, "BCD"):
    put(ws, f"{col}{r}", f"={sc('tg', s)}", "link", PCT2)
r += 1
D["spread"] = r
label(ws, f"A{r}", "WACC - g")
for col in "BCD":
    put(ws, f"{col}{r}", f"={col}{D['sw']}-{col}{D['sg']}", "formula", PCT2)
r += 2

WCOL = {"Bear": "B", "Base": "C", "Bull": "D"}
V = {s: {} for s in SCEN}
for idx, s in enumerate(SCEN):
    R = V[s]
    wacc_ref = f"${WCOL[s]}${D['sw']}"
    g_ref = f"${WCOL[s]}${D['sg']}"
    section(ws, r, f"{'BCD'[idx]}. {s.upper()} CASE", "A", "H"); r += 1
    header_row(ws, r, [("A", "EUR m"), ("B", "Value")] + [(c, y) for c, y in zip(YCOLS, YEARS)]); r += 1

    R["t"] = r
    label(ws, f"A{r}", "Discount period (years)")
    for i, col in enumerate(YCOLS):
        put(ws, f"{col}{r}", f"=1-0.5*{a('midyear')}" if i == 0 else f"={YCOLS[i-1]}{r}+1", "formula", "0.0")
    r += 1
    R["fcff"] = r
    label(ws, f"A{r}", "FCFF")
    for col in YCOLS:
        put(ws, f"{col}{r}", f"={f(s, 'fcff', col)}", "link", NUM)
    r += 1
    R["df"] = r
    label(ws, f"A{r}", "Discount factor")
    for col in YCOLS:
        put(ws, f"{col}{r}", f"=1/(1+{wacc_ref})^{col}{R['t']}", "formula", FACTOR)
    r += 1
    R["pv"] = r
    label(ws, f"A{r}", "PV of FCFF", bold=True)
    for col in YCOLS:
        put(ws, f"{col}{r}", f"={col}{R['fcff']}*{col}{R['df']}", "formula", NUM, bold=True)
    r += 2

    def vrow(key, text, formula, fmt=NUM, kind="formula", bold=False):
        global r
        R[key] = r
        label(ws, f"A{r}", text, bold=bold)
        put(ws, f"B{r}", formula, kind, fmt, bold=bold)
        r += 1

    vrow("sum_pv", "Sum of PV of FCFF", f"=SUM(D{R['pv']}:H{R['pv']})", bold=True)
    vrow("fcff5", "FCFF 2031E", f"=H{R['fcff']}")
    vrow("g", "Terminal growth rate", f"={g_ref}", PCT2)
    vrow("w", "WACC", f"={wacc_ref}", PCT2)
    vrow("tv", "Terminal value (Gordon growth)", f"=B{R['fcff5']}*(1+B{R['g']})/(B{R['w']}-B{R['g']})")
    vrow("df5", "Terminal value discount factor", f"=1/(1+B{R['w']})^5", FACTOR)
    vrow("pv_tv", "PV of terminal value", f"=B{R['tv']}*B{R['df5']}", bold=True)
    vrow("ev", "Enterprise value", f"=B{R['sum_pv']}+B{R['pv_tv']}", bold=True)
    ws[f"B{R['ev']}"].fill = KPI_FILL
    total_border(ws, R["ev"], "B")
    vrow("tv_pct", "Terminal value % of EV", f"=B{R['pv_tv']}/B{R['ev']}", PCT)
    vrow("ev_ebitda", "EV / 2027E EBITDA", f"=B{R['ev']}/{f(s, 'ebitda', 'D')}", MULT)
    r += 1
    label(ws, f"A{r}", "Equity bridge", bold=True); r += 1
    vrow("b_ev", "Enterprise value", f"=B{R['ev']}")
    vrow("b_debt", "Less: borrowings (31 Dec 2025)", f"=-{a('debt')}", kind="link")
    vrow("b_cash", "Add: cash (31 Dec 2025)", f"={a('cash')}", kind="link")
    vrow("b_oth", "Add: other net debt items (31 Dec 2025)", f"={a('nd_oth')}", kind="link")
    vrow("b_h1", "Less: increase in net debt, H1 2026", f"=-{a('nd_h1')}", kind="link")
    vrow("b_post", "Less: buybacks and dividend after 30 Jun 2026", f"=-({a('bb_post')}+{a('div_post')})",
         kind="link")
    vrow("b_spot", "Add: Spotify stake (net)", f"={a('spot')}", kind="link")
    vrow("eq", "Equity value", f"=SUM(B{R['b_ev']}:B{R['b_spot']})", bold=True)
    ws[f"B{R['eq']}"].fill = KPI_FILL
    total_border(ws, R["eq"], "B")
    vrow("sh", "Diluted shares (m)", f"={a('sh_dil')}", NUM1, kind="link")
    vrow("px", "Implied share price (€)", f"=B{R['eq']}/B{R['sh']}", EUR, bold=True)
    ws[f"B{R['px']}"].fill = KPI_FILL
    vrow("cur", "Current share price (€)", f"={a('price')}", EUR, kind="link")
    vrow("up", "Upside / (downside)", f"=B{R['px']}/B{R['cur']}-1", PCT, bold=True)
    r += 1


def v(s, key):
    return x(DCF, f"$B${V[s][key]}")


# --- checks -----------------------------------------------------------------
section(ws, r, "E. CHECKS", "A", "H"); r += 1
header_row(ws, r, [("A", "Check"), ("B", "Value"), ("C", "Status")]); r += 1
TOL = 0.01
CHECK_ROWS = []


def check(text, value_formula, status_formula, fmt=NUM1):
    global r
    label(ws, f"A{r}", text)
    put(ws, f"B{r}", value_formula, "formula", fmt)
    put(ws, f"C{r}", status_formula, "formula", None, bold=True)
    ws[f"C{r}"].alignment = Alignment(horizontal="center")
    CHECK_ROWS.append(r)
    r += 1


rev26 = a("rev26")
diffs = "+".join(f"ABS({f(s, 'rev', 'H')}-{rev26}*" + "*".join(f"(1+{ya('g_' + s, c)})" for c in YCOLS) + ")"
                 for s in SCEN)
check("Revenue build", f"={diffs}", f'=IF(B{r}<{TOL},"OK","ERROR")')

parts = []
for s in SCEN:
    rr = F[s]
    parts.append(f"SUMPRODUCT(ABS('{FC}'!$D${rr['ebitda']}:$H${rr['ebitda']}-'{FC}'!$D${rr['ebit']}:$H${rr['ebit']}"
                 f"-'{FC}'!$D${rr['da']}:$H${rr['da']}))")
check("EBITDA = EBIT + D&A", "=" + "+".join(parts), f'=IF(B{r}<{TOL},"OK","ERROR")')

parts = []
for s in SCEN:
    rr = F[s]
    rg = lambda k: f"'{FC}'!$D${rr[k]}:$H${rr[k]}"
    parts.append(f"SUMPRODUCT(ABS({rg('ebit')}*(1-{a('tax')})+{rg('da')}-{rg('rev')}*({a('capex')}+{ya_rng('cat')}"
                 f"+{ya_rng('sw')}+{ya_rng('roy')}+{a('lease')}+{a('nwc')})-{rg('fcff')}))")
check("FCFF recalculation", "=" + "+".join(parts), f'=IF(B{r}<{TOL},"OK","ERROR")')

mins = ",".join(f"'{FC}'!$D${F[s]['nwc']}:$H${F[s]['nwc']}" for s in SCEN)
check("NWC sign", f"={a('nwc')}", f'=IF(IF({a("nwc")}>=0,MAX({mins})<=0,MIN({mins})>=0),"OK","ERROR")', PCT2)

indep = (f"({a('price')}*{a('sh_dil')})/({a('price')}*{a('sh_dil')}+{a('adj_nd')})*({a('rf')}+{a('beta_u')}*"
         f"(1+(1-{a('tax')})*{a('adj_nd')}/({a('price')}*{a('sh_dil')}))*{a('erp')})+{a('adj_nd')}/"
         f"({a('price')}*{a('sh_dil')}+{a('adj_nd')})*{a('kd')}*(1-{a('tax')})")
check("WACC recalculation and weights",
      f"=ABS($B${D['wacc']}-({indep}))+ABS($B${D['we']}+$B${D['wd']}-1)", f'=IF(B{r}<0.000001,"OK","ERROR")',
      '0.000000')

check("g < WACC (all scenarios)", f"=MIN($B${D['spread']}:$D${D['spread']})", f'=IF(B{r}>0,"OK","ERROR")', PCT2)

parts = "+".join(f"ABS($B${V[s]['pv_tv']}-$B${V[s]['tv']}/(1+$B${V[s]['w']})^5)" for s in SCEN)
check("Terminal value discounting", "=" + parts, f'=IF(B{r}<{TOL},"OK","ERROR")')

parts = "+".join(
    f"ABS($B${V[s]['ev']}-SUMPRODUCT($D${V[s]['fcff']}:$H${V[s]['fcff']}/(1+$B${V[s]['w']})^$D${V[s]['t']}:$H${V[s]['t']})"
    f"-$B${V[s]['pv_tv']})" for s in SCEN)
check("EV recalculation", "=" + parts, f'=IF(B{r}<{TOL},"OK","ERROR")')

parts = "+".join(f"ABS($B${V[s]['eq']}-($B${V[s]['ev']}-{a('adj_nd')}+{a('spot')}))" for s in SCEN)
check("Equity bridge", "=" + parts, f'=IF(B{r}<{TOL},"OK","ERROR")')
check("Net debt ties to 30 Jun 2026 reported",
      f"=ABS(-SUM($B${V['Base']['b_debt']}:$B${V['Base']['b_h1']})-{h('h1_nd', 'B')})", f'=IF(B{r}<{TOL},"OK","ERROR")')

parts = "+".join(f"ABS($B${V[s]['px']}*{a('sh_dil')}-$B${V[s]['eq']})" for s in SCEN)
check("Diluted share count", "=" + parts, f'=IF(AND(B{r}<{TOL},{a("sh_dil")}>={a("sh_basic")}),"OK","ERROR")')

CHK_LINK_ROW = r
check("Dashboard ties to DCF", "=0", f'=IF(B{r}<{TOL},"OK","ERROR")')
r += 1
D["overall"] = r
label(ws, f"A{r}", "Overall", bold=True)
first_chk, last_chk = CHECK_ROWS[0], CHECK_ROWS[-1]
put(ws, f"B{r}", f'=IF(COUNTIF(C{first_chk}:C{last_chk},"ERROR")=0,"OK",'
                 f'COUNTIF(C{first_chk}:C{last_chk},"ERROR")&" ERROR(S)")', "formula", None, bold=True)
CHK_RANGE = f"C{first_chk}:C{last_chk}"
ws.conditional_formatting.add(CHK_RANGE, CellIsRule(operator="equal", formula=['"OK"'],
                              fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
ws.conditional_formatting.add(CHK_RANGE, CellIsRule(operator="equal", formula=['"ERROR"'],
                              fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'B{r}="OK"'],
                              fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add(f"B{r}", FormulaRule(formula=[f'B{r}<>"OK"'],
                              fill=PatternFill("solid", fgColor="FFC7CE")))

# ============================================================================
# 5. SCENARIOS AND SENSITIVITY
# ============================================================================
ws = ws_s
title(ws, "UMG Valuation Summary", "Share prices in €; EV and equity value in €bn; other values in €m.", last_col="P")
for col, w in zip("ABCDEFGH", [40, 20, 15, 15, 15, 15, 15, 4]):
    ws.column_dimensions[col].width = w

section(ws, 4, "BASE CASE", "A", "B")
KPI = {}
kpis = [
    ("cur", "Current share price (€)", f"={a('price')}", EUR, "link"),
    ("base", "Implied share price - base (€)", f"={v('Base', 'px')}", EUR, "link"),
    ("bear", "Implied share price - bear (€)", f"={v('Bear', 'px')}", EUR, "link"),
    ("bull", "Implied share price - bull (€)", f"={v('Bull', 'px')}", EUR, "link"),
    ("up", "Upside / (downside) - base", None, PCT, "formula"),
    ("ev", "Enterprise value (€bn)", f"={v('Base', 'ev')}/1000", '"€"#,##0.0"bn"', "link"),
    ("eq", "Equity value (€bn)", f"={v('Base', 'eq')}/1000", '"€"#,##0.0"bn"', "link"),
    ("wacc", "WACC", f"={x(DCF, '$C$' + str(D['sw']))}", PCT2, "link"),
    ("g", "Terminal growth", f"={x(DCF, '$C$' + str(D['sg']))}", PCT2, "link"),
    ("verdict", "Valuation (±10%)", None, None, "formula"),
    ("disc", "Discounting", f'=IF({a("midyear")}=1,"Mid-year","Year-end")', None, "link"),
    ("status", "Checks", f"={x(DCF, '$B$' + str(D['overall']))}", None, "link"),
]
rr = 5
for key, text, formula, fmt, kind in kpis:
    KPI[key] = rr
    label(ws, f"A{rr}", text, bold=True)
    ws[f"A{rr}"].fill = KPI_FILL
    if formula:
        put(ws, f"B{rr}", formula, kind, fmt, bold=True, size=11)
    ws[f"B{rr}"].fill = KPI_FILL
    ws[f"B{rr}"].alignment = Alignment(horizontal="right")
    rr += 1
put(ws, f"B{KPI['up']}", f"=B{KPI['base']}/B{KPI['cur']}-1", "formula", PCT, bold=True, size=11)
ws[f"B{KPI['up']}"].fill = KPI_FILL
put(ws, f"B{KPI['verdict']}",
    f'=IF(B{KPI["up"]}>0.1,"UNDERVALUED",IF(B{KPI["up"]}<-0.1,"OVERVALUED","FAIRLY VALUED"))',
    "formula", None, bold=True, size=11)
ws[f"B{KPI['verdict']}"].fill = KPI_FILL
ws[f"B{KPI['verdict']}"].alignment = Alignment(horizontal="right")
ws.conditional_formatting.add(f"B{KPI['up']}", CellIsRule(operator="lessThan", formula=["0"],
                              font=Font(color="C00000", bold=True)))
ws.conditional_formatting.add(f"B{KPI['up']}", CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                              font=Font(color="008000", bold=True)))
for rr_ in range(5, rr):
    ws[f"A{rr_}"].border = BOX
    ws[f"B{rr_}"].border = BOX

# scenario comparison
SC0 = 18
section(ws, SC0, "SCENARIO COMPARISON", "A", "G")
header_row(ws, SC0 + 1, [("A", "Metric"), ("B", "Bear"), ("C", "Base"), ("D", "Bull")])
comp = [
    ("Revenue CAGR 2026E-2031E", lambda s: f"=({f(s, 'rev', 'H')}/{f(s, 'rev', 'C')})^(1/5)-1", PCT),
    ("EBIT margin 2027E", lambda s: f"={f(s, 'm', 'D')}", PCT),
    ("EBIT margin 2031E", lambda s: f"={f(s, 'm', 'H')}", PCT),
    ("WACC", lambda s: f"={x(DCF, '$' + WCOL[s] + '$' + str(D['sw']))}", PCT2),
    ("Terminal growth", lambda s: f"={x(DCF, '$' + WCOL[s] + '$' + str(D['sg']))}", PCT2),
    ("Revenue 2031E (€m)", lambda s: f"={f(s, 'rev', 'H')}", NUM),
    ("EBIT 2031E (€m)", lambda s: f"={f(s, 'ebit', 'H')}", NUM),
    ("EBITDA 2031E (€m)", lambda s: f"={f(s, 'ebitda', 'H')}", NUM),
    ("FCFF 2031E (€m)", lambda s: f"={f(s, 'fcff', 'H')}", NUM),
    ("Terminal value (€m)", lambda s: f"={v(s, 'tv')}", NUM),
    ("Enterprise value (€m)", lambda s: f"={v(s, 'ev')}", NUM),
    ("Equity value (€m)", lambda s: f"={v(s, 'eq')}", NUM),
    ("Implied share price (€)", lambda s: f"={v(s, 'px')}", EUR),
    ("Upside / (downside)", lambda s: f"={v(s, 'up')}", PCT),
]
CMP = {}
rr = SC0 + 2
for text, fn, fmt in comp:
    CMP[text] = rr
    label(ws, f"A{rr}", text, bold=text.startswith("Implied"))
    for s, col in zip(SCEN, "BCD"):
        put(ws, f"{col}{rr}", fn(s), "link", fmt, bold=text.startswith("Implied"))
    rr += 1
ws.conditional_formatting.add(f"B{CMP['Upside / (downside)']}:D{CMP['Upside / (downside)']}",
                              CellIsRule(operator="lessThan", formula=["0"], font=Font(color="C00000")))

# sensitivity
SR = SC0 + 2 + len(comp) + 1
section(ws, SR, "SENSITIVITY - BASE CASE IMPLIED SHARE PRICE (€)", "A", "G")
label(ws, f"A{SR + 1}", "WACC (rows) vs terminal growth (columns)", italic=True, size=9)
hdr = SR + 2
put(ws, f"A{hdr}", "WACC  \\  g", "formula", None, bold=True)
ws[f"A{hdr}"].fill = SUB_FILL
ws[f"A{hdr}"].alignment = Alignment(horizontal="right")
g_steps = [-0.01, -0.005, 0.0, 0.005, 0.01]
w_steps = [-0.01, -0.005, 0.0, 0.005, 0.01]
gcols = ["B", "C", "D", "E", "F"]
base_w = x(DCF, f"$C${D['sw']}")
base_g = x(DCF, f"$C${D['sg']}")
for col, st in zip(gcols, g_steps):
    put(ws, f"{col}{hdr}", f"={base_g}+({st})", "formula", PCT2, bold=True)
    ws[f"{col}{hdr}"].fill = SUB_FILL
    ws[f"{col}{hdr}"].alignment = Alignment(horizontal="center")
fc_rng = f"'{DCF}'!$D${V['Base']['fcff']}:$H${V['Base']['fcff']}"
t_rng = f"'{DCF}'!$D${V['Base']['t']}:$H${V['Base']['t']}"
fc5 = f"'{DCF}'!$H${V['Base']['fcff']}"
for i, st in enumerate(w_steps):
    row = hdr + 1 + i
    put(ws, f"A{row}", f"={base_w}+({st})", "formula", PCT2, bold=True)
    ws[f"A{row}"].fill = SUB_FILL
    ws[f"A{row}"].alignment = Alignment(horizontal="right")
    for col in gcols:
        w_ = f"$A{row}"
        g_ = f"{col}${hdr}"
        formula = (f'=IF({w_}<={g_},"n/a",(SUMPRODUCT({fc_rng}/(1+{w_})^{t_rng})'
                   f"+{fc5}*(1+{g_})/({w_}-{g_})/(1+{w_})^5"
                   f"-{a('adj_nd')}+{a('spot')})/{a('sh_dil')})")
        put(ws, f"{col}{row}", formula, "formula", EUR)
        ws[f"{col}{row}"].border = BOX
        ws[f"{col}{row}"].alignment = Alignment(horizontal="center")
sens_rng = f"B{hdr + 1}:F{hdr + 5}"
centre = f"D{hdr + 3}"
ws[centre].font = Font(name=FONT, size=10, bold=True)
cur_ref = f"$B${KPI['cur']}"
ws.conditional_formatting.add(sens_rng, FormulaRule(formula=[f"AND(ISNUMBER(B{hdr + 1}),B{hdr + 1}>={cur_ref})"],
                              fill=PatternFill("solid", fgColor="C6EFCE")))
ws.conditional_formatting.add(sens_rng, FormulaRule(formula=[f"AND(ISNUMBER(B{hdr + 1}),B{hdr + 1}<{cur_ref})"],
                              fill=PatternFill("solid", fgColor="FFC7CE")))

# chart data
CD = hdr + 8
section(ws, CD, "CHART DATA", "A", "G")
rv_h = CD + 1
label(ws, f"A{rv_h}", "Revenue (€m)", bold=True)
rev_years = ["2026E"] + YEARS
rev_cols_src = ["C"] + YCOLS
for i, y in enumerate(rev_years):
    c = ws.cell(row=rv_h, column=2 + i, value=y)
    c.font = Font(name=FONT, bold=True)
    c.alignment = Alignment(horizontal="center")
for j, s in enumerate(SCEN):
    row = rv_h + 1 + j
    label(ws, f"A{row}", s)
    for i, src_col in enumerate(rev_cols_src):
        put(ws, f"{get_column_letter(2 + i)}{row}", f"={f(s, 'rev', src_col)}", "link", NUM)

fc_h = rv_h + 5
label(ws, f"A{fc_h}", "FCFF (€m)", bold=True)
for i, y in enumerate(YEARS):
    c = ws.cell(row=fc_h, column=2 + i, value=y)
    c.font = Font(name=FONT, bold=True)
    c.alignment = Alignment(horizontal="center")
for j, s in enumerate(SCEN):
    row = fc_h + 1 + j
    label(ws, f"A{row}", s)
    for i, src_col in enumerate(YCOLS):
        put(ws, f"{get_column_letter(2 + i)}{row}", f"={f(s, 'fcff', src_col)}", "link", NUM)

px_h = fc_h + 5
label(ws, f"A{px_h}", "Share price (€)", bold=True)
label(ws, f"B{px_h}", "€ per share", bold=True)
px_rows = [("Bear", f"={v('Bear', 'px')}"), ("Base", f"={v('Base', 'px')}"),
           ("Bull", f"={v('Bull', 'px')}"), ("Current", f"={a('price')}")]
for j, (lab, fml) in enumerate(px_rows):
    label(ws, f"A{px_h + 1 + j}", lab)
    put(ws, f"B{px_h + 1 + j}", fml, "link", EUR)


def tidy(chart, ytitle):
    chart.roundedCorners = False
    if chart.legend is not None:
        chart.legend.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.title = ytitle
    chart.height = 7.5
    chart.width = 17


lc = LineChart()
lc.title = "Revenue by scenario (€m)"
lc.add_data(Reference(ws, min_col=1, min_row=rv_h + 1, max_col=7, max_row=rv_h + 3), from_rows=True,
            titles_from_data=True)
lc.set_categories(Reference(ws, min_col=2, min_row=rv_h, max_col=7, max_row=rv_h))
for ser, color in zip(lc.series, ["C00000", NAVY, "00A050"]):
    ser.graphicalProperties.line.solidFill = color
    ser.graphicalProperties.line.width = 28000
    ser.smooth = False
lc.y_axis.numFmt = "#,##0"
lc.y_axis.majorGridlines = None
lc.y_axis.scaling.min = 10000
tidy(lc, "€m")
lc.legend.position = "b"
ws.add_chart(lc, "I4")

bc = BarChart()
bc.type = "col"
bc.title = "FCFF by scenario (€m)"
bc.add_data(Reference(ws, min_col=1, min_row=fc_h + 1, max_col=6, max_row=fc_h + 3), from_rows=True,
            titles_from_data=True)
bc.set_categories(Reference(ws, min_col=2, min_row=fc_h, max_col=6, max_row=fc_h))
for ser, color in zip(bc.series, ["C00000", NAVY, "00A050"]):
    ser.graphicalProperties.solidFill = color
    ser.graphicalProperties.line.solidFill = color
bc.y_axis.numFmt = "#,##0"
bc.y_axis.majorGridlines = None
tidy(bc, "€m")
bc.legend.position = "b"
ws.add_chart(bc, "I21")

pc = BarChart()
pc.type = "col"
pc.title = "Implied vs current share price (€)"
pc.add_data(Reference(ws, min_col=2, min_row=px_h, max_row=px_h + 4), titles_from_data=True)
pc.set_categories(Reference(ws, min_col=1, min_row=px_h + 1, max_row=px_h + 4))
ser = pc.series[0]
for idx, color in enumerate(["C00000", NAVY, "00A050", "7F7F7F"]):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = color
    pt.graphicalProperties.line.solidFill = color
    ser.dPt.append(pt)
pc.y_axis.numFmt = '"€"0'
pc.y_axis.majorGridlines = None
pc.legend = None
pc.roundedCorners = False
pc.x_axis.delete = False
pc.y_axis.delete = False
pc.y_axis.title = "€ per share"
pc.height = 7.5
pc.width = 17
ws.add_chart(pc, "I38")

# dashboard tie-out check
ws_d[f"B{CHK_LINK_ROW}"].value = (
    f"=ABS('{DASH}'!$B${KPI['base']}-$B${V['Base']['px']})"
    f"+ABS('{DASH}'!{centre}-$B${V['Base']['px']})"
    f"+ABS('{DASH}'!$B${CMP['Implied share price (€)']}-$B${V['Bear']['px']})"
    f"+ABS('{DASH}'!$D${CMP['Implied share price (€)']}-$B${V['Bull']['px']})")

wb.active = wb.sheetnames.index(DASH)
for wsx in wb.worksheets:
    wsx.sheet_view.tabSelected = (wsx.title == DASH)
ws_s.sheet_properties.tabColor = NAVY
ws_a.sheet_properties.tabColor = "FFC000"

wb.properties.creator = ""
wb.save(OUTPUT)
print(f"Saved {OUTPUT}")
