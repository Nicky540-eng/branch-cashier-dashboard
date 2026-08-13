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
    BRANCHES = sorted(cash["Shop"].dropna().unique())
    wb = Workbook()

    ws = wb.active
    ws.title = "CashData"
    ccols = ["Cashier", "Shop", "Game", "Bets", "Revokes", "RevokedSum"]
    ws.append(ccols)
    for c in range(1, len(ccols) + 1):
        cell = ws.cell(row=1, column=c); cell.fill = HDR_FILL; cell.font = HDR_FONT
    for _, rw in cash[ccols].iterrows():
        ws.append(list(rw.values))
    CN = len(cash) + 1
    for row in ws.iter_rows(min_row=2, max_row=CN, max_col=len(ccols)):
        for cell in row:
            cell.font = BODY
            if cell.column in (4, 5):
                cell.number_format = INT_FMT
            elif cell.column == 6:
                cell.number_format = MON_FMT
    ws.freeze_panes = "A2"
    for i, w in enumerate([26, 15, 18, 12, 12, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws2 = wb.create_sheet("SlipData")
    scols = ["User", "Shop", "Game", "BetSlips", "PaidIn", "NetWin"]
    ws2.append(scols)
    for c in range(1, len(scols) + 1):
        cell = ws2.cell(row=1, column=c); cell.fill = HDR_FILL; cell.font = HDR_FONT
    for _, rw in slip[scols].iterrows():
        ws2.append(list(rw.values))
    SN = len(slip) + 1
    for row in ws2.iter_rows(min_row=2, max_row=SN, max_col=len(scols)):
        for cell in row:
            cell.font = BODY
            if cell.column == 4:
                cell.number_format = INT_FMT
            elif cell.column >= 5:
                cell.number_format = MON_FMT
    ws2.freeze_panes = "A2"
    for i, w in enumerate([26, 15, 18, 12, 14, 14], start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    CD = f"CashData!$B$2:$B${CN}"
    CD_CASH = f"CashData!$A$2:$A${CN}"
    CD_GAME = f"CashData!$C$2:$C${CN}"
    CD_BETS = f"CashData!$D$2:$D${CN}"
    CD_REV = f"CashData!$E$2:$E${CN}"
    CD_RSUM = f"CashData!$F$2:$F${CN}"
    SD_SHOP = f"SlipData!$B$2:$B${SN}"
    SD_SLIPS = f"SlipData!$D$2:$D${SN}"
    SD_IN = f"SlipData!$E$2:$E${SN}"
    SD_NET = f"SlipData!$F$2:$F${SN}"

    def hrow(s, row, headers, widths=None):
        for i, h in enumerate(headers, start=1):
            c = s.cell(row=row, column=i, value=h)
            c.fill = HDR_FILL; c.font = HDR_FONT; c.border = BOX
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            for i, w in enumerate(widths, start=1):
                s.column_dimensions[get_column_letter(i)].width = w
        return row + 1

    cs = cash.groupby(["Shop", "Cashier"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"),
        RevSum=("RevokedSum", "sum"), IsMgr=("IsManager", "max"))
    cg = cash.groupby(["Shop", "Game"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"))

    for br in BRANCHES:
        s = wb.create_sheet(str(br)[:31])
        s.cell(row=1, column=1, value=f"{br} — Cashier & Game Report").font = TITLE_FONT
        r = 3
        s.cell(row=r, column=1, value="Cashier performance").font = LBL_FONT
        r += 1
        keyrow = r
        s.cell(row=keyrow, column=7, value="Branch key:").font = NOTE_F
        s.cell(row=keyrow, column=8, value=br).font = LBL_FONT
        r = hrow(s, r, ["Cashier", "Total Bets", "Total Revokes", "Revoked Amount"], [30, 14, 15, 18])
        first = r
        for n in cs[cs["Shop"] == br].sort_values("Bets", ascending=False)["Cashier"]:
            s.cell(row=r, column=1, value=n).font = BODY
            s.cell(row=r, column=2,
                   value=f'=SUMIFS({CD_BETS},{CD},$H${keyrow},{CD_CASH},$A{r})').number_format = INT_FMT
            s.cell(row=r, column=3,
                   value=f'=SUMIFS({CD_REV},{CD},$H${keyrow},{CD_CASH},$A{r})').number_format = INT_FMT
            s.cell(row=r, column=4,
                   value=f'=SUMIFS({CD_RSUM},{CD},$H${keyrow},{CD_CASH},$A{r})').number_format = MON_FMT
            for c in range(1, 5):
                s.cell(row=r, column=c).border = BOX; s.cell(row=r, column=c).font = BODY
            r += 1
        last = r - 1
        for label, f2, f3, f4 in [
            ("BRANCH TOTAL", f'=SUM(B{first}:B{last})', f'=SUM(C{first}:C{last})', f'=SUM(D{first}:D{last})'),
            ("BRANCH AVERAGE (per cashier)", f'=IFERROR(AVERAGE(B{first}:B{last}),0)',
             f'=IFERROR(AVERAGE(C{first}:C{last}),0)', f'=IFERROR(AVERAGE(D{first}:D{last}),0)'),
        ]:
            s.cell(row=r, column=1, value=label).font = LBL_FONT
            for c, f in ((2, f2), (3, f3), (4, f4)):
                cell = s.cell(row=r, column=c, value=f)
                cell.font = LBL_FONT
                cell.number_format = MON_FMT if c == 4 else INT_FMT
            for c in range(1, 5):
                s.cell(row=r, column=c).fill = SUB_FILL; s.cell(row=r, column=c).border = BOX
            r += 1
        s.cell(row=r, column=1, value="Cashiers counted").font = BODY
        s.cell(row=r, column=2, value=f'=COUNTA(A{first}:A{last})').number_format = INT_FMT
        r += 3

        s.cell(row=r, column=1, value="Branch highlights").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Measure", "Cashier", "Value"], [40, 30, 18])
        _all = cs[cs["Shop"] == br]
        _nom = _all[~_all["IsMgr"]] if "IsMgr" in _all.columns else _all
        _nom = _nom if len(_nom) else _all
        _mb = _all.loc[_all["Bets"].idxmax()]
        _lb = _nom.loc[_nom["Bets"].idxmin()]
        _mr = _nom.loc[_nom["Revokes"].idxmax()]
        for label, who, val, fmt in [
            ("Most bets (cashier)", _mb["Cashier"], float(_mb["Bets"]), INT_FMT),
            ("Least bets (cashier, managers excluded)", _lb["Cashier"], float(_lb["Bets"]), INT_FMT),
            ("Most revokes (cashier, managers excluded)", _mr["Cashier"], float(_mr["Revokes"]), INT_FMT),
            ("Revoked amount of that cashier", None, float(_mr["RevSum"]), MON_FMT),
        ]:
            s.cell(row=r, column=1, value=label).font = BODY
            if who is not None:
                s.cell(row=r, column=2, value=who).font = BODY
            c = s.cell(row=r, column=3, value=val); c.font = BODY; c.number_format = fmt
            for cc in range(1, 4):
                s.cell(row=r, column=cc).border = BOX
            r += 1
        r += 2

        s.cell(row=r, column=1, value="Bets & revokes per game").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Game", "Bets", "Revokes", "Revoked Amount"], [30, 14, 15, 18])
        gfirst = r
        for g in cg[cg["Shop"] == br].sort_values("Bets", ascending=False)["Game"]:
            s.cell(row=r, column=1, value=g).font = BODY
            s.cell(row=r, column=2,
                   value=f'=SUMIFS({CD_BETS},{CD},$H${keyrow},{CD_GAME},$A{r})').number_format = INT_FMT
            s.cell(row=r, column=3,
                   value=f'=SUMIFS({CD_REV},{CD},$H${keyrow},{CD_GAME},$A{r})').number_format = INT_FMT
            s.cell(row=r, column=4,
                   value=f'=SUMIFS({CD_RSUM},{CD},$H${keyrow},{CD_GAME},$A{r})').number_format = MON_FMT
            for c in range(1, 5):
                s.cell(row=r, column=c).border = BOX; s.cell(row=r, column=c).font = BODY
            r += 1
        glast = r - 1
        s.cell(row=r, column=1, value="TOTAL").font = LBL_FONT
        for c, f in ((2, f'=SUM(B{gfirst}:B{glast})'), (3, f'=SUM(C{gfirst}:C{glast})'),
                     (4, f'=SUM(D{gfirst}:D{glast})')):
            cell = s.cell(row=r, column=c, value=f); cell.font = LBL_FONT
            cell.number_format = MON_FMT if c == 4 else INT_FMT
        for c in range(1, 5):
            s.cell(row=r, column=c).fill = SUB_FILL; s.cell(row=r, column=c).border = BOX
        r += 2
        for label, col in (("Game with most bets", "B"), ("Game with most revokes", "C")):
            s.cell(row=r, column=1, value=label).font = LBL_FONT
            s.cell(row=r, column=2,
                   value=f'=INDEX(A{gfirst}:A{glast},MATCH(MAX({col}{gfirst}:{col}{glast}),{col}{gfirst}:{col}{glast},0))').font = BODY
            c = s.cell(row=r, column=3, value=f'=MAX({col}{gfirst}:{col}{glast})')
            c.number_format = INT_FMT; c.font = BODY
            r += 1
        s.freeze_panes = "A4"

    ac = wb.create_sheet("All Cashiers")
    ac.cell(row=1, column=1, value="All Cashiers — All Branches").font = TITLE_FONT
    r = hrow(ac, 3, ["Cashier", "Branch", "Total Bets", "Total Revokes", "Revoked Amount"],
             [30, 16, 14, 15, 18])
    afirst = r
    for _, rw in cs.sort_values("Bets", ascending=False).iterrows():
        ac.cell(row=r, column=1, value=rw["Cashier"]).font = BODY
        ac.cell(row=r, column=2, value=rw["Shop"]).font = BODY
        ac.cell(row=r, column=3,
                value=f'=SUMIFS({CD_BETS},{CD_CASH},$A{r},{CD},$B{r})').number_format = INT_FMT
        ac.cell(row=r, column=4,
                value=f'=SUMIFS({CD_REV},{CD_CASH},$A{r},{CD},$B{r})').number_format = INT_FMT
        ac.cell(row=r, column=5,
                value=f'=SUMIFS({CD_RSUM},{CD_CASH},$A{r},{CD},$B{r})').number_format = MON_FMT
        for c in range(1, 6):
            ac.cell(row=r, column=c).border = BOX; ac.cell(row=r, column=c).font = BODY
        r += 1
    alast = r - 1
    ac.cell(row=r, column=1, value="GRAND TOTAL").font = LBL_FONT
    for c in (3, 4, 5):
        L = get_column_letter(c)
        cell = ac.cell(row=r, column=c, value=f'=SUM({L}{afirst}:{L}{alast})')
        cell.font = LBL_FONT; cell.number_format = MON_FMT if c == 5 else INT_FMT
    for c in range(1, 6):
        ac.cell(row=r, column=c).fill = SUB_FILL; ac.cell(row=r, column=c).border = BOX
    ac.freeze_panes = "A4"

    bgs = wb.create_sheet("Bets per Game")
    bgs.cell(row=1, column=1, value="Bets per Game — by Branch").font = TITLE_FONT
    r = hrow(bgs, 3, ["Game"] + list(BRANCHES) + ["All Branches"],
             [26] + [16] * len(BRANCHES) + [16])
    gf = r
    for g in cash.groupby("Game", as_index=False)["Bets"].sum().sort_values(
            "Bets", ascending=False)["Game"]:
        bgs.cell(row=r, column=1, value=g).font = BODY
        for i in range(2, len(BRANCHES) + 2):
            col = get_column_letter(i)
            bgs.cell(row=r, column=i,
                     value=f'=SUMIFS({CD_BETS},{CD_GAME},$A{r},{CD},{col}$3)').number_format = INT_FMT
        endc = get_column_letter(len(BRANCHES) + 1)
        bgs.cell(row=r, column=len(BRANCHES) + 2,
                 value=f'=SUM(B{r}:{endc}{r})').number_format = INT_FMT
        for c in range(1, len(BRANCHES) + 3):
            bgs.cell(row=r, column=c).border = BOX; bgs.cell(row=r, column=c).font = BODY
        r += 1
    gl = r - 1
    bgs.cell(row=r, column=1, value="TOTAL").font = LBL_FONT
    for c in range(2, len(BRANCHES) + 3):
        L = get_column_letter(c)
        cell = bgs.cell(row=r, column=c, value=f'=SUM({L}{gf}:{L}{gl})')
        cell.font = LBL_FONT; cell.number_format = INT_FMT
    for c in range(1, len(BRANCHES) + 3):
        bgs.cell(row=r, column=c).fill = SUB_FILL; bgs.cell(row=r, column=c).border = BOX
    bgs.freeze_panes = "B4"

    sl = wb.create_sheet("Slip Summary")
    sl.cell(row=1, column=1, value="Slip Report Summary — by Branch").font = TITLE_FONT
    r = hrow(sl, 3, ["Branch", "Total Betslips", "Paid In", "Net Win", "GW Margin %",
                     "Net Win Margin %", "First Slip Issued", "Last Slip Issued"],
             [16, 16, 16, 16, 14, 16, 20, 20])
    sf = r
    for br in BRANCHES:
        sub = slip[slip["Shop"] == br]
        sl.cell(row=r, column=1, value=br).font = BODY
        sl.cell(row=r, column=2, value=f'=SUMIFS({SD_SLIPS},{SD_SHOP},$A{r})').number_format = INT_FMT
        sl.cell(row=r, column=3, value=f'=SUMIFS({SD_IN},{SD_SHOP},$A{r})').number_format = MON_FMT
        sl.cell(row=r, column=4, value=f'=SUMIFS({SD_NET},{SD_SHOP},$A{r})').number_format = MON_FMT
        sl.cell(row=r, column=5, value=f'=IFERROR(D{r}/C{r}*100,0)').number_format = PCT_FMT
        sl.cell(row=r, column=6, value=f'=IFERROR(D{r}/C{r}*100,0)').number_format = PCT_FMT
        fd, ld = sub["FirstDT"].min(), sub["LastDT"].max()
        for c, v in ((7, fd), (8, ld)):
            cell = sl.cell(row=r, column=c,
                           value=v.to_pydatetime() if pd.notna(v) else None)
            cell.number_format = "dd/mm/yyyy hh:mm"; cell.font = BODY
        for c in range(1, 9):
            sl.cell(row=r, column=c).border = BOX; sl.cell(row=r, column=c).font = BODY
        r += 1
    sll = r - 1
    sl.cell(row=r, column=1, value="ALL BRANCHES").font = LBL_FONT
    for c in (2, 3, 4):
        L = get_column_letter(c)
        cell = sl.cell(row=r, column=c, value=f'=SUM({L}{sf}:{L}{sll})')
        cell.font = LBL_FONT; cell.number_format = INT_FMT if c == 2 else MON_FMT
    for c in (5, 6):
        cell = sl.cell(row=r, column=c, value=f'=IFERROR(D{r}/C{r}*100,0)')
        cell.font = LBL_FONT; cell.number_format = PCT_FMT
    for c, v in ((7, slip["FirstDT"].min()), (8, slip["LastDT"].max())):
        cell = sl.cell(row=r, column=c, value=v.to_pydatetime() if pd.notna(v) else None)
        cell.number_format = "dd/mm/yyyy hh:mm"; cell.font = LBL_FONT
    for c in range(1, 9):
        sl.cell(row=r, column=c).fill = SUB_FILL; sl.cell(row=r, column=c).border = BOX

    sm = wb.create_sheet("Summary", 0)
    sm.cell(row=1, column=1, value="Branch & Cashier Performance").font = TITLE_FONT
    sm.cell(row=2, column=1,
            value=f"Generated {datetime.now():%d %b %Y %H:%M} from the Cash Operations and Slip Summary reports.").font = NOTE_F
    r = hrow(sm, 4, ["Branch", "Cashiers", "Total Bets", "Total Revokes", "Revoked Amount",
                     "Avg Bets / Cashier", "Avg Revokes / Cashier", "Total Betslips", "GW Margin %"],
             [16, 11, 14, 14, 16, 17, 19, 14, 13])
    bf = r
    counts = cs.groupby("Shop")["Cashier"].nunique().to_dict()
    for br in BRANCHES:
        sm.cell(row=r, column=1, value=br).font = BODY
        sm.cell(row=r, column=2, value=int(counts.get(br, 0))).number_format = INT_FMT
        sm.cell(row=r, column=3, value=f'=SUMIFS({CD_BETS},{CD},$A{r})').number_format = INT_FMT
        sm.cell(row=r, column=4, value=f'=SUMIFS({CD_REV},{CD},$A{r})').number_format = INT_FMT
        sm.cell(row=r, column=5, value=f'=SUMIFS({CD_RSUM},{CD},$A{r})').number_format = MON_FMT
        sm.cell(row=r, column=6, value=f'=IFERROR(C{r}/B{r},0)').number_format = INT_FMT
        sm.cell(row=r, column=7, value=f'=IFERROR(D{r}/B{r},0)').number_format = '#,##0.0;(#,##0.0);-'
        sm.cell(row=r, column=8, value=f'=SUMIFS({SD_SLIPS},{SD_SHOP},$A{r})').number_format = INT_FMT
        sm.cell(row=r, column=9,
                value=f'=IFERROR(SUMIFS({SD_NET},{SD_SHOP},$A{r})/SUMIFS({SD_IN},{SD_SHOP},$A{r})*100,0)').number_format = PCT_FMT
        for c in range(1, 10):
            sm.cell(row=r, column=c).border = BOX; sm.cell(row=r, column=c).font = BODY
        r += 1
    bl = r - 1
    sm.cell(row=r, column=1, value="ALL BRANCHES").font = LBL_FONT
    for c in (2, 3, 4, 5, 8):
        L = get_column_letter(c)
        cell = sm.cell(row=r, column=c, value=f'=SUM({L}{bf}:{L}{bl})')
        cell.number_format = MON_FMT if c == 5 else INT_FMT
    sm.cell(row=r, column=6, value=f'=IFERROR(C{r}/B{r},0)').number_format = INT_FMT
    sm.cell(row=r, column=7, value=f'=IFERROR(D{r}/B{r},0)').number_format = '#,##0.0;(#,##0.0);-'
    sm.cell(row=r, column=9,
            value=f'=IFERROR(SUM({SD_NET})/SUM({SD_IN})*100,0)').number_format = PCT_FMT
    for c in range(1, 10):
        sm.cell(row=r, column=c).fill = SUB_FILL; sm.cell(row=r, column=c).border = BOX
        sm.cell(row=r, column=c).font = LBL_FONT
    r += 3
    sm.cell(row=r, column=1, value="Overall highlights (all branches)").font = LBL_FONT
    r += 1
    r = hrow(sm, r, ["Measure", "Cashier / Game", "Branch", "Value"], [40, 28, 16, 18])
    _csn = cs[~cs["IsMgr"]] if "IsMgr" in cs.columns else cs
    _csn = _csn if len(_csn) else cs
    _mb = cs.loc[cs["Bets"].idxmax()]
    _lb = _csn.loc[_csn["Bets"].idxmin()]
    _mr = _csn.loc[_csn["Revokes"].idxmax()]
    _ma = _csn.loc[_csn["RevSum"].idxmax()]
    _gt = cash.groupby("Game", as_index=False)["Bets"].sum()
    _gtop = _gt.loc[_gt["Bets"].idxmax()]
    for label, who, brc, val, fmt in [
        ("Most bets — cashier", _mb["Cashier"], _mb["Shop"], float(_mb["Bets"]), INT_FMT),
        ("Least bets — cashier (managers excluded)", _lb["Cashier"], _lb["Shop"], float(_lb["Bets"]), INT_FMT),
        ("Most revokes — cashier (managers excluded)", _mr["Cashier"], _mr["Shop"], float(_mr["Revokes"]), INT_FMT),
        ("Revoked amount of that cashier", None, None, float(_mr["RevSum"]), MON_FMT),
        ("Highest revoked amount (managers excluded)", _ma["Cashier"], _ma["Shop"], float(_ma["RevSum"]), MON_FMT),
        ("Game with most bets", _gtop["Game"], "All branches", float(_gtop["Bets"]), INT_FMT),
    ]:
        sm.cell(row=r, column=1, value=label).font = BODY
        if who is not None:
            sm.cell(row=r, column=2, value=who).font = BODY
        if brc is not None:
            sm.cell(row=r, column=3, value=brc).font = BODY
        c = sm.cell(row=r, column=4, value=val); c.font = BODY; c.number_format = fmt
        for cc in range(1, 5):
            sm.cell(row=r, column=cc).border = BOX
        r += 1
    r += 1
    sm.cell(row=r, column=1,
            value='Definitions — "Bets" = Paid In Count. "Revokes" = Paid Out Revoked Count. '
                  '"Revoked Amount" = Revoked Sum. Margins = Net Win / Paid In from the Slip report.').font = NOTE_F
    r += 1
    sm.cell(row=r, column=1,
            value='Manager accounts are excluded from the "most revokes", "least bets" and '
                  '"highest revoked amount" figures because revokes are manager-authorised. '
                  'They remain in the full cashier lists.').font = NOTE_F
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
        st.bar_chart(bo.set_index("Branch")["Bets"])
    with c2:
        st.caption("Revokes by branch")
        st.bar_chart(bo.set_index("Branch")["Revokes"])

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
        st.bar_chart(sub.head(15).set_index("Cashier")["Bets"])
    with c2:
        st.caption("Bets per game")
        st.bar_chart(gsub.set_index("Game")["Bets"])

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
