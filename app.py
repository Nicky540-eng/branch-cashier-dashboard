"""
Playbet — Branch & Cashier Report Builder
Upload the Cash Operations Summary and Slip Summary CSVs, view the dashboard,
download the full Excel workbook.
"""
import io
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Playbet — Branch & Cashier Report", layout="wide")

# ----------------------------------------------------------------- styling
st.markdown("""
<style>
  .block-container {padding-top: 2rem;}
  div[data-testid="stMetricValue"] {font-size: 1.5rem;}
  .small-note {color:#6b7280; font-size:0.82rem;}
</style>
""", unsafe_allow_html=True)

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F3864")
LBL_FONT = Font(name=FONT, bold=True, size=10)
BODY = Font(name=FONT, size=10)
NOTE_F = Font(name=FONT, size=9, italic=True, color="595959")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INT_FMT = '#,##0;(#,##0);-'
MON_FMT = 'R #,##0.00;(R #,##0.00);-'
PCT_FMT = '0.00"%";(0.00"%");-'

CASH_REQUIRED = ["Cashier", "Shop", "Game", "Paid In Count",
                 "Paid Out - Revoked Count", "Revoked Sum"]
SLIP_REQUIRED = ["Game", "Shop", "User", "Bet Slips", "First Slip Issued",
                 "Last Slip Issued", "Paid In", "Net Win"]


def num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce").fillna(0)


def load_csv(f):
    f.seek(0)
    try:
        return pd.read_csv(f, encoding="utf-8-sig")
    except UnicodeDecodeError:
        f.seek(0)
        return pd.read_csv(f, encoding="latin-1")


def missing(df, cols):
    return [c for c in cols if c not in df.columns]


def prep_cash(df, drop_managers):
    d = df.copy()
    d["Bets"] = num(d["Paid In Count"])
    d["Revokes"] = num(d["Paid Out - Revoked Count"])
    d["RevokedSum"] = num(d["Revoked Sum"])
    d["IsManager"] = d["Cashier"].astype(str).str.contains("manager", case=False, na=False)
    if drop_managers:
        d = d[~d["IsManager"]]
    return d


def prep_slip(df):
    d = df.copy()
    d["BetSlips"] = num(d["Bet Slips"])
    d["PaidIn"] = num(d["Paid In"])
    d["NetWin"] = num(d["Net Win"])
    for src, dst in (("First Slip Issued", "FirstDT"), ("Last Slip Issued", "LastDT")):
        parsed = pd.to_datetime(d[src], format="%d/%m/%y %H:%M:%S", errors="coerce")
        if parsed.isna().all():
            parsed = pd.to_datetime(d[src], errors="coerce", dayfirst=True)
        d[dst] = parsed
    return d


def money(v):
    return f"R {v:,.2f}"


# ================================================================= EXCEL
def build_workbook(cash, slip):
    """Values-only workbook. Every figure is computed in Python and written as a
    literal number, so all figures display immediately with no need to enable
    editing or recalculate."""
    BRANCHES = sorted(cash["Shop"].dropna().unique())

    cs = cash.groupby(["Shop", "Cashier"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"),
        RevSum=("RevokedSum", "sum"), IsMgr=("IsManager", "max"))
    cg = cash.groupby(["Shop", "Game"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"))

    wb = Workbook()

    def hrow(s, row, headers, widths=None):
        for i, h in enumerate(headers, start=1):
            c = s.cell(row=row, column=i, value=h)
            c.fill = HDR_FILL; c.font = HDR_FONT; c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            for i, w in enumerate(widths, start=1):
                s.column_dimensions[get_column_letter(i)].width = w
        return row + 1

    def put(s, r, c, val, fmt=None, font=BODY, fill=None, border=True):
        cell = s.cell(row=r, column=c, value=val)
        cell.font = font
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        if border:
            cell.border = BOX
        return cell

    # raw data (audit)
    ws = wb.active
    ws.title = "CashData"
    ccols = ["Cashier", "Shop", "Game", "Bets", "Revokes", "RevokedSum"]
    ws.append(ccols)
    for c in range(1, len(ccols) + 1):
        cell = ws.cell(row=1, column=c); cell.fill = HDR_FILL; cell.font = HDR_FONT
    for _, rw in cash[ccols].iterrows():
        ws.append([rw[k] for k in ccols])
    for row in ws.iter_rows(min_row=2, max_row=len(cash) + 1, max_col=len(ccols)):
        for cell in row:
            cell.font = BODY
            if cell.column in (4, 5): cell.number_format = INT_FMT
            elif cell.column == 6: cell.number_format = MON_FMT
    ws.freeze_panes = "A2"
    for i, w in enumerate([26, 15, 18, 12, 12, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("SlipData")
    scols = ["User", "Shop", "Game", "BetSlips", "PaidIn", "NetWin"]
    ws2.append(scols)
    for c in range(1, len(scols) + 1):
        cell = ws2.cell(row=1, column=c); cell.fill = HDR_FILL; cell.font = HDR_FONT
    for _, rw in slip[scols].iterrows():
        ws2.append([rw[k] for k in scols])
    for row in ws2.iter_rows(min_row=2, max_row=len(slip) + 1, max_col=len(scols)):
        for cell in row:
            cell.font = BODY
            if cell.column == 4: cell.number_format = INT_FMT
            elif cell.column >= 5: cell.number_format = MON_FMT
    ws2.freeze_panes = "A2"
    for i, w in enumerate([26, 15, 18, 12, 14, 14], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    for br in BRANCHES:
        s = wb.create_sheet(str(br)[:31])
        s.cell(row=1, column=1, value=f"{br} — Cashier & Game Report").font = TITLE_FONT
        r = 3
        s.cell(row=r, column=1, value="Cashier performance").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Cashier", "Total Bets", "Total Revokes", "Revoked Amount"], [30, 14, 15, 18])
        sub = cs[cs["Shop"] == br].sort_values("Bets", ascending=False)
        for _, row_ in sub.iterrows():
            put(s, r, 1, row_["Cashier"])
            put(s, r, 2, int(row_["Bets"]), INT_FMT)
            put(s, r, 3, int(row_["Revokes"]), INT_FMT)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT)
            r += 1
        put(s, r, 1, "BRANCH TOTAL", font=LBL_FONT, fill=SUB_FILL)
        put(s, r, 2, int(sub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 3, int(sub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 4, float(sub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
        r += 1
        put(s, r, 1, "BRANCH AVERAGE (per cashier)", font=LBL_FONT, fill=SUB_FILL)
        put(s, r, 2, round(float(sub["Bets"].mean()), 0) if len(sub) else 0, INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 3, round(float(sub["Revokes"].mean()), 1) if len(sub) else 0, '#,##0.0;(#,##0.0);-', LBL_FONT, SUB_FILL)
        put(s, r, 4, round(float(sub["RevSum"].mean()), 2) if len(sub) else 0, MON_FMT, LBL_FONT, SUB_FILL)
        r += 1
        put(s, r, 1, "Cashiers counted", border=False)
        put(s, r, 2, int(len(sub)), INT_FMT, border=False)
        r += 3

        subn = sub[~sub["IsMgr"]] if "IsMgr" in sub.columns else sub
        subn = subn if len(subn) else sub
        mb = sub.loc[sub["Bets"].idxmax()]
        lb = subn.loc[subn["Bets"].idxmin()]
        mr = subn.loc[subn["Revokes"].idxmax()]
        s.cell(row=r, column=1, value="Branch highlights").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Measure", "Cashier", "Value"], [40, 30, 18])
        for label, who, val, fmt in [
            ("Most bets (cashier)", mb["Cashier"], int(mb["Bets"]), INT_FMT),
            ("Least bets (cashier, managers excluded)", lb["Cashier"], int(lb["Bets"]), INT_FMT),
            ("Most revokes (cashier, managers excluded)", mr["Cashier"], int(mr["Revokes"]), INT_FMT),
            ("Revoked amount of that cashier", None, float(mr["RevSum"]), MON_FMT),
        ]:
            put(s, r, 1, label)
            if who is not None: put(s, r, 2, who)
            else: s.cell(row=r, column=2).border = BOX
            put(s, r, 3, val, fmt)
            r += 1
        r += 2

        s.cell(row=r, column=1, value="Bets & revokes per game").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Game", "Bets", "Revokes", "Revoked Amount"], [30, 14, 15, 18])
        gsub = cg[cg["Shop"] == br].sort_values("Bets", ascending=False)
        for _, row_ in gsub.iterrows():
            put(s, r, 1, row_["Game"])
            put(s, r, 2, int(row_["Bets"]), INT_FMT)
            put(s, r, 3, int(row_["Revokes"]), INT_FMT)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT)
            r += 1
        put(s, r, 1, "TOTAL", font=LBL_FONT, fill=SUB_FILL)
        put(s, r, 2, int(gsub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 3, int(gsub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 4, float(gsub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
        r += 2
        if len(gsub):
            gb = gsub.loc[gsub["Bets"].idxmax()]; gr = gsub.loc[gsub["Revokes"].idxmax()]
            put(s, r, 1, "Game with most bets", font=LBL_FONT, border=False)
            put(s, r, 2, gb["Game"], border=False)
            put(s, r, 3, int(gb["Bets"]), INT_FMT, border=False)
            r += 1
            put(s, r, 1, "Game with most revokes", font=LBL_FONT, border=False)
            put(s, r, 2, gr["Game"], border=False)
            put(s, r, 3, int(gr["Revokes"]), INT_FMT, border=False)
        s.freeze_panes = "A4"

    ac = wb.create_sheet("All Cashiers")
    ac.cell(row=1, column=1, value="All Cashiers — All Branches").font = TITLE_FONT
    r = hrow(ac, 3, ["Cashier", "Branch", "Total Bets", "Total Revokes", "Revoked Amount"], [30, 16, 14, 15, 18])
    allc = cs.sort_values("Bets", ascending=False)
    for _, row_ in allc.iterrows():
        put(ac, r, 1, row_["Cashier"]); put(ac, r, 2, row_["Shop"])
        put(ac, r, 3, int(row_["Bets"]), INT_FMT)
        put(ac, r, 4, int(row_["Revokes"]), INT_FMT)
        put(ac, r, 5, float(row_["RevSum"]), MON_FMT)
        r += 1
    put(ac, r, 1, "GRAND TOTAL", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 2, "", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 3, int(allc["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 4, int(allc["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 5, float(allc["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
    ac.freeze_panes = "A4"

    bgs = wb.create_sheet("Bets per Game")
    bgs.cell(row=1, column=1, value="Bets per Game — by Branch").font = TITLE_FONT
    r = hrow(bgs, 3, ["Game"] + list(BRANCHES) + ["All Branches"], [26] + [16] * len(BRANCHES) + [16])
    gpiv = cash.pivot_table(index="Game", columns="Shop", values="Bets", aggfunc="sum", fill_value=0)
    gpiv["_all"] = gpiv.sum(axis=1)
    gpiv = gpiv.sort_values("_all", ascending=False)
    for game, row_ in gpiv.iterrows():
        put(bgs, r, 1, game)
        for i, br in enumerate(BRANCHES, start=2):
            put(bgs, r, i, int(row_.get(br, 0)), INT_FMT)
        put(bgs, r, len(BRANCHES) + 2, int(row_["_all"]), INT_FMT)
        r += 1
    put(bgs, r, 1, "TOTAL", font=LBL_FONT, fill=SUB_FILL)
    for i, br in enumerate(BRANCHES, start=2):
        put(bgs, r, i, int(gpiv[br].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(bgs, r, len(BRANCHES) + 2, int(gpiv["_all"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    bgs.freeze_panes = "B4"

    sl = wb.create_sheet("Slip Summary")
    sl.cell(row=1, column=1, value="Slip Report Summary — by Branch").font = TITLE_FONT
    r = hrow(sl, 3, ["Branch", "Total Betslips", "Paid In", "Net Win", "GW Margin %",
                     "Net Win Margin %", "First Slip Issued", "Last Slip Issued"],
             [16, 16, 16, 16, 14, 16, 20, 20])
    for br in BRANCHES:
        sub = slip[slip["Shop"] == br]
        betslips = int(sub["BetSlips"].sum()); paidin = float(sub["PaidIn"].sum()); netwin = float(sub["NetWin"].sum())
        gw = round(netwin / paidin * 100, 2) if paidin else 0
        put(sl, r, 1, br); put(sl, r, 2, betslips, INT_FMT)
        put(sl, r, 3, paidin, MON_FMT); put(sl, r, 4, netwin, MON_FMT)
        put(sl, r, 5, gw, PCT_FMT); put(sl, r, 6, gw, PCT_FMT)
        fd, ld = sub["FirstDT"].min(), sub["LastDT"].max()
        c7 = put(sl, r, 7, fd.to_pydatetime() if pd.notna(fd) else None)
        c8 = put(sl, r, 8, ld.to_pydatetime() if pd.notna(ld) else None)
        c7.number_format = "dd/mm/yyyy hh:mm"; c8.number_format = "dd/mm/yyyy hh:mm"
        r += 1
    tb = int(slip["BetSlips"].sum()); tin = float(slip["PaidIn"].sum()); tnw = float(slip["NetWin"].sum())
    tgw = round(tnw / tin * 100, 2) if tin else 0
    put(sl, r, 1, "ALL BRANCHES", font=LBL_FONT, fill=SUB_FILL)
    put(sl, r, 2, tb, INT_FMT, LBL_FONT, SUB_FILL)
    put(sl, r, 3, tin, MON_FMT, LBL_FONT, SUB_FILL)
    put(sl, r, 4, tnw, MON_FMT, LBL_FONT, SUB_FILL)
    put(sl, r, 5, tgw, PCT_FMT, LBL_FONT, SUB_FILL)
    put(sl, r, 6, tgw, PCT_FMT, LBL_FONT, SUB_FILL)
    fd, ld = slip["FirstDT"].min(), slip["LastDT"].max()
    c7 = put(sl, r, 7, fd.to_pydatetime() if pd.notna(fd) else None, font=LBL_FONT, fill=SUB_FILL)
    c8 = put(sl, r, 8, ld.to_pydatetime() if pd.notna(ld) else None, font=LBL_FONT, fill=SUB_FILL)
    c7.number_format = "dd/mm/yyyy hh:mm"; c8.number_format = "dd/mm/yyyy hh:mm"
    sl.freeze_panes = "A4"

    sm = wb.create_sheet("Summary", 0)
    sm.cell(row=1, column=1, value="Branch & Cashier Performance").font = TITLE_FONT
    sm.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d %b %Y %H:%M} from the Cash Operations and Slip Summary reports.").font = NOTE_F
    r = hrow(sm, 4, ["Branch", "Cashiers", "Total Bets", "Total Revokes", "Revoked Amount",
                     "Avg Bets / Cashier", "Avg Revokes / Cashier", "Total Betslips", "GW Margin %"],
             [16, 11, 14, 14, 16, 17, 19, 14, 13])
    for br in BRANCHES:
        sub = cs[cs["Shop"] == br]; subs = slip[slip["Shop"] == br]
        ncash = int(sub["Cashier"].nunique()); bets = int(sub["Bets"].sum()); revs = int(sub["Revokes"].sum())
        rsum = float(sub["RevSum"].sum()); paidin = float(subs["PaidIn"].sum()); netwin = float(subs["NetWin"].sum())
        gw = round(netwin / paidin * 100, 2) if paidin else 0
        put(sm, r, 1, br); put(sm, r, 2, ncash, INT_FMT)
        put(sm, r, 3, bets, INT_FMT); put(sm, r, 4, revs, INT_FMT); put(sm, r, 5, rsum, MON_FMT)
        put(sm, r, 6, round(bets / ncash, 0) if ncash else 0, INT_FMT)
        put(sm, r, 7, round(revs / ncash, 1) if ncash else 0, '#,##0.0;(#,##0.0);-')
        put(sm, r, 8, int(subs["BetSlips"].sum()), INT_FMT); put(sm, r, 9, gw, PCT_FMT)
        r += 1
    tcash = int(cs["Cashier"].nunique()); tbets = int(cs["Bets"].sum()); trev = int(cs["Revokes"].sum())
    trsum = float(cs["RevSum"].sum()); tslip = int(slip["BetSlips"].sum())
    tin = float(slip["PaidIn"].sum()); tnw = float(slip["NetWin"].sum()); tgw = round(tnw / tin * 100, 2) if tin else 0
    put(sm, r, 1, "ALL BRANCHES", font=LBL_FONT, fill=SUB_FILL)
    put(sm, r, 2, tcash, INT_FMT, LBL_FONT, SUB_FILL); put(sm, r, 3, tbets, INT_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 4, trev, INT_FMT, LBL_FONT, SUB_FILL); put(sm, r, 5, trsum, MON_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 6, round(tbets / tcash, 0) if tcash else 0, INT_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 7, round(trev / tcash, 1) if tcash else 0, '#,##0.0;(#,##0.0);-', LBL_FONT, SUB_FILL)
    put(sm, r, 8, tslip, INT_FMT, LBL_FONT, SUB_FILL); put(sm, r, 9, tgw, PCT_FMT, LBL_FONT, SUB_FILL)
    r += 3

    csn = cs[~cs["IsMgr"]] if "IsMgr" in cs.columns else cs
    csn = csn if len(csn) else cs
    mb = cs.loc[cs["Bets"].idxmax()]; lb = csn.loc[csn["Bets"].idxmin()]
    mr = csn.loc[csn["Revokes"].idxmax()]; ma = csn.loc[csn["RevSum"].idxmax()]
    gt = cash.groupby("Game", as_index=False)["Bets"].sum(); gtop = gt.loc[gt["Bets"].idxmax()]
    sm.cell(row=r, column=1, value="Overall highlights (all branches)").font = LBL_FONT
    r += 1
    r = hrow(sm, r, ["Measure", "Cashier / Game", "Branch", "Value"], [40, 28, 16, 18])
    for label, who, brc, val, fmt in [
        ("Most bets — cashier", mb["Cashier"], mb["Shop"], int(mb["Bets"]), INT_FMT),
        ("Least bets — cashier (managers excluded)", lb["Cashier"], lb["Shop"], int(lb["Bets"]), INT_FMT),
        ("Most revokes — cashier (managers excluded)", mr["Cashier"], mr["Shop"], int(mr["Revokes"]), INT_FMT),
        ("Revoked amount of that cashier", None, None, float(mr["RevSum"]), MON_FMT),
        ("Highest revoked amount (managers excluded)", ma["Cashier"], ma["Shop"], float(ma["RevSum"]), MON_FMT),
        ("Game with most bets", gtop["Game"], "All branches", int(gtop["Bets"]), INT_FMT),
    ]:
        put(sm, r, 1, label)
        if who is not None: put(sm, r, 2, who)
        else: sm.cell(row=r, column=2).border = BOX
        if brc is not None: put(sm, r, 3, brc)
        else: sm.cell(row=r, column=3).border = BOX
        put(sm, r, 4, val, fmt)
        r += 1
    r += 1
    sm.cell(row=r, column=1,
            value='Definitions — "Bets" = Paid In Count. "Revokes" = Paid Out Revoked Count. '
                  '"Revoked Amount" = Revoked Sum. Margins = Net Win / Paid In from the Slip report.').font = NOTE_F
    r += 1
    sm.cell(row=r, column=1,
            value='Manager accounts are excluded from "most revokes", "least bets" and '
                  '"highest revoked amount" (revokes are manager-authorised). They remain in the cashier lists.').font = NOTE_F
    sm.freeze_panes = "A5"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ================================================================= UI
st.title("Branch & Cashier Report Builder")
st.markdown('<p class="small-note">Upload the two monthly reports to view the analysis and '
            'download the full Excel workbook.</p>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Reports")
    cash_file = st.file_uploader("Cash Operations Summary (CSV)", type=["csv"])
    slip_file = st.file_uploader("Slip Summary (CSV)", type=["csv"])
    st.divider()
    st.header("Options")
    drop_mgr = st.checkbox("Exclude Manager accounts from cashier rankings", value=False,
                           help="Manager logins often carry most revokes because revokes are "
                                "manager-authorised. Tick this to rank tellers only.")
    period = st.text_input("Report period (label only)", value="")

if not cash_file or not slip_file:
    st.info("Upload both reports in the sidebar to begin.")
    st.markdown("""
**Cash Operations Summary** needs: `Cashier`, `Shop`, `Game`, `Paid In Count`,
`Paid Out - Revoked Count`, `Revoked Sum`.

**Slip Summary** needs: `Game`, `Shop`, `User`, `Bet Slips`, `First Slip Issued`,
`Last Slip Issued`, `Paid In`, `Net Win`.
""")
    st.stop()

try:
    cash_raw = load_csv(cash_file)
    slip_raw = load_csv(slip_file)
except Exception as e:
    st.error(f"Could not read a file: {e}")
    st.stop()

mc, ms = missing(cash_raw, CASH_REQUIRED), missing(slip_raw, SLIP_REQUIRED)
if mc or ms:
    if mc:
        st.error(f"Cash Operations report is missing: {', '.join(mc)}")
    if ms:
        st.error(f"Slip report is missing: {', '.join(ms)}")
    st.caption("Check you haven't swapped the two files in the uploader.")
    st.stop()

cash = prep_cash(cash_raw, drop_mgr)
slip = prep_slip(slip_raw)
if cash.empty:
    st.warning("No cashier rows left after filtering.")
    st.stop()

BRANCHES = sorted(cash["Shop"].dropna().unique())
cs = cash.groupby(["Shop", "Cashier"], as_index=False).agg(
    Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"),
    IsMgr=("IsManager", "max"))
cg = cash.groupby(["Shop", "Game"], as_index=False).agg(
    Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"))

hdr = f"Period: {period}" if period.strip() else ""
if hdr:
    st.caption(hdr)
if drop_mgr:
    n_mgr = prep_cash(cash_raw, False)["IsManager"].sum()
    st.caption(f"Manager accounts excluded ({int(n_mgr)} source rows).")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total bets", f"{int(cash['Bets'].sum()):,}")
k2.metric("Total revokes", f"{int(cash['Revokes'].sum()):,}")
k3.metric("Revoked amount", money(cash["RevokedSum"].sum()))
k4.metric("Total betslips", f"{int(slip['BetSlips'].sum()):,}")
gw = slip["NetWin"].sum() / slip["PaidIn"].sum() * 100 if slip["PaidIn"].sum() else 0
k5.metric("GW margin", f"{gw:.2f}%")

st.divider()

tab_over, tab_branch, tab_games, tab_slip, tab_cash = st.tabs(
    ["Overview", "By branch", "Games", "Slip report", "All cashiers"])

with tab_over:
    st.subheader("Branch overview")
    bo = cash.groupby("Shop", as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"))
    bo["Cashiers"] = bo["Shop"].map(cs.groupby("Shop")["Cashier"].nunique())
    bo["Avg bets/cashier"] = (bo["Bets"] / bo["Cashiers"]).round(0)
    bo["Avg revokes/cashier"] = (bo["Revokes"] / bo["Cashiers"]).round(1)
    sb = slip.groupby("Shop", as_index=False).agg(
        Betslips=("BetSlips", "sum"), PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"))
    sb["GW margin %"] = (sb["NetWin"] / sb["PaidIn"] * 100).round(2)
    bo = bo.merge(sb[["Shop", "Betslips", "GW margin %"]], on="Shop", how="left")
    bo = bo.rename(columns={"Shop": "Branch", "RevSum": "Revoked amount"})
    st.dataframe(
        bo[["Branch", "Cashiers", "Bets", "Revokes", "Revoked amount",
            "Avg bets/cashier", "Avg revokes/cashier", "Betslips", "GW margin %"]],
        use_container_width=True, hide_index=True,
        column_config={
            "Bets": st.column_config.NumberColumn(format="%d"),
            "Revokes": st.column_config.NumberColumn(format="%d"),
            "Revoked amount": st.column_config.NumberColumn(format="R %.2f"),
            "Betslips": st.column_config.NumberColumn(format="%d"),
            "GW margin %": st.column_config.NumberColumn(format="%.2f%%"),
        })

    c1, c2 = st.columns(2)
    with c1:
        st.caption("Bets by branch")
        _b = bo[["Branch", "Bets"]].sort_values("Bets", ascending=False)
        st.dataframe(_b, use_container_width=True, hide_index=True,
                     column_config={"Bets": st.column_config.ProgressColumn(
                         "Bets", format="%d", min_value=0, max_value=int(_b["Bets"].max()))})
    with c2:
        st.caption("Revokes by branch")
        _r = bo[["Branch", "Revokes"]].sort_values("Revokes", ascending=False)
        st.dataframe(_r, use_container_width=True, hide_index=True,
                     column_config={"Revokes": st.column_config.ProgressColumn(
                         "Revokes", format="%d", min_value=0, max_value=int(_r["Revokes"].max()))})

    st.subheader("Overall highlights")
    csn = cs[~cs["IsMgr"]] if "IsMgr" in cs.columns else cs
    csn = csn if len(csn) else cs
    top = cs.loc[cs["Bets"].idxmax()]
    low = csn.loc[csn["Bets"].idxmin()]
    mrev = csn.loc[csn["Revokes"].idxmax()]
    mamt = csn.loc[csn["RevSum"].idxmax()]
    gtot = cash.groupby("Game", as_index=False)["Bets"].sum()
    gtop = gtot.loc[gtot["Bets"].idxmax()]
    hi = pd.DataFrame([
        ["Most bets — cashier", top["Cashier"], top["Shop"], f"{int(top['Bets']):,}"],
        ["Least bets — cashier (managers excluded)", low["Cashier"], low["Shop"], f"{int(low['Bets']):,}"],
        ["Most revokes — cashier (managers excluded)", mrev["Cashier"], mrev["Shop"],
         f"{int(mrev['Revokes']):,}  ({money(mrev['RevSum'])})"],
        ["Highest revoked amount (managers excluded)", mamt["Cashier"], mamt["Shop"], money(mamt["RevSum"])],
        ["Game with most bets", gtop["Game"], "All branches", f"{int(gtop['Bets']):,}"],
    ], columns=["Measure", "Cashier / Game", "Branch", "Value"])
    st.dataframe(hi, use_container_width=True, hide_index=True)
    st.caption("Manager accounts are excluded from 'most revokes', 'least bets' and "
               "'highest revoked amount' (revokes are manager-authorised). They remain in the cashier lists.")

with tab_branch:
    br = st.selectbox("Branch", BRANCHES)
    sub = cs[cs["Shop"] == br].sort_values("Bets", ascending=False)
    gsub = cg[cg["Shop"] == br].sort_values("Bets", ascending=False)
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cashiers", f"{sub['Cashier'].nunique():,}")
    m2.metric("Total bets", f"{int(sub['Bets'].sum()):,}")
    m3.metric("Avg bets/cashier", f"{sub['Bets'].mean():,.0f}")
    m4.metric("Total revokes", f"{int(sub['Revokes'].sum()):,}")

    st.subheader("Cashiers")
    disp = sub.rename(columns={"Cashier": "Cashier", "Bets": "Bets",
                               "Revokes": "Revokes", "RevSum": "Revoked amount"})
    st.dataframe(disp[["Cashier", "Bets", "Revokes", "Revoked amount"]],
                 use_container_width=True, hide_index=True,
                 column_config={
                     "Bets": st.column_config.NumberColumn(format="%d"),
                     "Revokes": st.column_config.NumberColumn(format="%d"),
                     "Revoked amount": st.column_config.NumberColumn(format="R %.2f")})

    c1, c2 = st.columns(2)
    with c1:
        st.caption("Top 15 cashiers by bets")
        _t = sub.head(15)[["Cashier", "Bets"]]
        st.dataframe(_t, use_container_width=True, hide_index=True,
                     column_config={"Bets": st.column_config.ProgressColumn(
                         "Bets", format="%d", min_value=0, max_value=int(_t["Bets"].max()) if len(_t) else 1)})
    with c2:
        st.caption("Bets per game")
        _g = gsub[["Game", "Bets"]].sort_values("Bets", ascending=False)
        st.dataframe(_g, use_container_width=True, hide_index=True,
                     column_config={"Bets": st.column_config.ProgressColumn(
                         "Bets", format="%d", min_value=0, max_value=int(_g["Bets"].max()) if len(_g) else 1)})

    st.subheader("Branch highlights")
    subn = sub[~sub["IsMgr"]] if "IsMgr" in sub.columns else sub
    subn = subn if len(subn) else sub
    _mr = subn.loc[subn['Revokes'].idxmax()]
    hb = pd.DataFrame([
        ["Most bets", sub.iloc[0]["Cashier"], f"{int(sub.iloc[0]['Bets']):,}"],
        ["Least bets (managers excluded)", subn.loc[subn['Bets'].idxmin(), "Cashier"],
         f"{int(subn['Bets'].min()):,}"],
        ["Most revokes (managers excluded)", _mr["Cashier"],
         f"{int(_mr['Revokes']):,}  ({money(_mr['RevSum'])})"],
        ["Game with most bets", gsub.iloc[0]["Game"], f"{int(gsub.iloc[0]['Bets']):,}"],
        ["Game with most revokes", gsub.loc[gsub['Revokes'].idxmax(), "Game"],
         f"{int(gsub['Revokes'].max()):,}"],
    ], columns=["Measure", "Cashier / Game", "Value"])
    st.dataframe(hb, use_container_width=True, hide_index=True)

with tab_games:
    st.subheader("Bets per game, by branch")
    piv = cash.pivot_table(index="Game", columns="Shop", values="Bets",
                           aggfunc="sum", fill_value=0)
    piv["All branches"] = piv.sum(axis=1)
    piv = piv.sort_values("All branches", ascending=False)
    st.dataframe(piv, use_container_width=True)
    st.caption("Revokes per game, by branch")
    pivr = cash.pivot_table(index="Game", columns="Shop", values="Revokes",
                            aggfunc="sum", fill_value=0)
    pivr["All branches"] = pivr.sum(axis=1)
    st.dataframe(pivr.sort_values("All branches", ascending=False), use_container_width=True)

with tab_slip:
    st.subheader("Slip report by branch")
    ss = slip.groupby("Shop", as_index=False).agg(
        Betslips=("BetSlips", "sum"), PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"),
        First=("FirstDT", "min"), Last=("LastDT", "max"))
    ss["GW margin %"] = (ss["NetWin"] / ss["PaidIn"] * 100).round(2)
    ss["Net win margin %"] = ss["GW margin %"]
    ss = ss.rename(columns={"Shop": "Branch", "PaidIn": "Paid in", "NetWin": "Net win",
                            "First": "First slip issued", "Last": "Last slip issued"})
    st.dataframe(ss[["Branch", "Betslips", "Paid in", "Net win", "GW margin %",
                     "Net win margin %", "First slip issued", "Last slip issued"]],
                 use_container_width=True, hide_index=True,
                 column_config={
                     "Betslips": st.column_config.NumberColumn(format="%d"),
                     "Paid in": st.column_config.NumberColumn(format="R %.2f"),
                     "Net win": st.column_config.NumberColumn(format="R %.2f"),
                     "GW margin %": st.column_config.NumberColumn(format="%.2f%%"),
                     "Net win margin %": st.column_config.NumberColumn(format="%.2f%%")})
    st.caption("Margins are computed as Net Win / Paid In. If your source reports GW Margin % "
               "and Net Win Margin identically, both columns will match.")

with tab_cash:
    st.subheader("All cashiers, all branches")
    q = st.text_input("Search cashier")
    allc = cs.sort_values("Bets", ascending=False).rename(
        columns={"Shop": "Branch", "RevSum": "Revoked amount"})
    if q.strip():
        allc = allc[allc["Cashier"].str.contains(q.strip(), case=False, na=False)]
    st.dataframe(allc[["Cashier", "Branch", "Bets", "Revokes", "Revoked amount"]],
                 use_container_width=True, hide_index=True,
                 column_config={
                     "Bets": st.column_config.NumberColumn(format="%d"),
                     "Revokes": st.column_config.NumberColumn(format="%d"),
                     "Revoked amount": st.column_config.NumberColumn(format="R %.2f")})

st.divider()
st.subheader("Download")
label = period.strip().replace(" ", "_") or datetime.now().strftime("%Y_%m")
if st.button("Build Excel workbook", type="primary"):
    with st.spinner("Building workbook…"):
        buf = build_workbook(cash, slip)
    st.download_button("Download Excel", data=buf,
                       file_name=f"Branch_Cashier_Report_{label}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("The workbook is formula-driven: edit the CashData / SlipData sheets and "
               "every figure recalculates.")
