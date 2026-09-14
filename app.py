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
HDR_FILL = PatternFill("solid", fgColor="000000")            # black header
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="D9E1F2")
TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F3864")
LBL_FONT = Font(name=FONT, bold=True, size=10)
BODY = Font(name=FONT, size=10)
NOTE_F = Font(name=FONT, size=9, italic=True, color="595959")
THIN = Side(style="thin", color="000000")                    # black table lines
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INT_FMT = '#,##0;[Red](#,##0);-'
MON_FMT = 'R #,##0.00;[Red](R #,##0.00);-'
PCT_FMT = '0.00"%";-0.00"%";-'

CASH_REQUIRED = ["Cashier", "Shop", "Game",
                 "Paid Out - Revoked Count", "Revoked Sum"]
# "Paid Out Sum" is optional; pulled in when present for the paid-out highlight.
SLIP_REQUIRED = ["Game", "Shop", "User", "Bet Slips", "First Slip Issued",
                 "Last Slip Issued", "Paid In", "Net Win"]


def num(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.strip(),
        errors="coerce").fillna(0)


def load_csv(f):
    f.seek(0)
    try:
        df = pd.read_csv(f, encoding="utf-8-sig")
    except UnicodeDecodeError:
        f.seek(0)
        df = pd.read_csv(f, encoding="latin-1")
    return canon_columns(df)


def canon_columns(df):
    """Map common header variations to the exact names the report expects, so exports
    from different months / slightly different configs still work. Matching is done on a
    squashed key (lowercased, spaces/underscores/dashes removed) against known aliases."""
    aliases = {
        # target -> list of accepted variants (do NOT cross-map Cashier<->User:
        # the cash file uses "Cashier", the slip file uses "User"; keep them distinct)
        "Cashier": ["cashiername", "teller", "operator"],
        "User": ["username"],
        "Shop": ["branch", "shopname", "branchname", "store", "location"],
        "Game": ["gamename", "product"],
        "Bet Slips": ["betslip", "betslips", "slips", "totalbetslips", "numberofbetslips"],
        "Paid In": ["paidinsum", "totalpaidin", "stakes", "turnover"],
        "Net Win": ["netwinsum", "netwinnings"],
        "GW Margin %": ["gwmargin", "grosswinmargin", "grosswinmarginpct", "gwmarginpct"],
        "Net Win Margin": ["netwinmarginpct", "nwmargin"],
        "Winnings - Unpaid": ["winningsunpaid", "unpaidwinnings", "unpaid", "winningsminusunpaid"],
        "Paid Out - Revoked Count": ["paidoutrevokedcount", "revokedcount", "revokecount",
                                     "paidoutrevoked"],
        "Revoked Sum": ["revokedamount", "revokeamount", "revokedtotal"],
        "Paid Out": ["payout", "payouts"],
        "Paid Out Sum": ["paidoutsum"],
        "First Slip Issued": ["firstslip", "firstissued"],
        "Last Slip Issued": ["lastslip", "lastissued"],
    }

    def squash(x):
        return "".join(ch for ch in str(x).lower() if ch.isalnum())

    present = {squash(c): c for c in df.columns}
    rename = {}
    for target, variants in aliases.items():
        if target in df.columns:
            continue  # already exactly right
        for v in variants:
            if squash(v) in present:
                rename[present[squash(v)]] = target
                break
    if rename:
        df = df.rename(columns=rename)
    return df


def missing(df, cols):
    return [c for c in cols if c not in df.columns]


def prep_cash(cash_df, slip_df, drop_managers):
    """Combined per Shop+Cashier+Game frame.
    Bets come from the Slip report's Bet Slips (keyed on User) — matches Aardvark.
    Revokes and revoked amounts come from the Cash Operations report (per Cashier).
    Paid In / Net Win / a turnover-weighted GW numerator also ride along so GW Margin %
    and Net Win Margin % can be aggregated per cashier / branch / game."""
    c = cash_df.copy()
    c["Revokes"] = num(c["Paid Out - Revoked Count"])
    c["RevokedSum"] = num(c["Revoked Sum"])
    c["PaidOut"] = num(c["Paid Out Sum"]) if "Paid Out Sum" in c.columns else 0.0
    c = c.groupby(["Shop", "Cashier", "Game"], as_index=False).agg(
        Revokes=("Revokes", "sum"), RevokedSum=("RevokedSum", "sum"),
        PaidOut=("PaidOut", "sum"))

    sp = slip_df.copy()
    sp["Bets"] = num(sp["Bet Slips"])
    sp["PaidIn"] = num(sp["Paid In"])
    sp["NetWin"] = num(sp["Net Win"])
    sp["Unpaid"] = num(sp["Winnings - Unpaid"]) if "Winnings - Unpaid" in sp.columns else 0.0
    sp = sp.groupby(["Shop", "User", "Game"], as_index=False).agg(
        Bets=("Bets", "sum"), PaidIn=("PaidIn", "sum"),
        NetWin=("NetWin", "sum"), Unpaid=("Unpaid", "sum"))
    sp = sp.rename(columns={"User": "Cashier"})

    d = pd.merge(sp, c, on=["Shop", "Cashier", "Game"], how="outer")
    for col in ("Bets", "Revokes", "RevokedSum", "PaidIn", "NetWin", "Unpaid", "PaidOut"):
        if col in d.columns:
            d[col] = d[col].fillna(0)
    d["IsManager"] = d["Cashier"].astype(str).str.contains("manager", case=False, na=False)
    if drop_managers:
        d = d[~d["IsManager"]]
    return d


def gw_pct(frame):
    """GW Margin % = (Net Win - Winnings Unpaid) / Paid In * 100, aggregated over the
    subset. This is Aardvark's exact formula (verified to match the source file's own
    GW Margin % in 100% of rows). Money is summed first, then the ratio is taken."""
    pin = frame["PaidIn"].sum()
    if not pin:
        return 0.0
    gross = frame["NetWin"].sum() - frame["Unpaid"].sum()
    return round(gross / pin * 100, 2)


def nwm_pct(frame):
    """Net Win Margin % = Net Win / Paid In * 100, aggregated over the subset."""
    pin = frame["PaidIn"].sum()
    return round(frame["NetWin"].sum() / pin * 100, 2) if pin else 0.0


def prep_slip(df):
    d = df.copy()
    d["BetSlips"] = num(d["Bet Slips"])
    d["PaidIn"] = num(d["Paid In"])
    d["NetWin"] = num(d["Net Win"])
    d["GWpct"] = num(d["GW Margin %"]) if "GW Margin %" in d.columns else 0.0
    d["NWMraw"] = num(d["Net Win Margin"]) if "Net Win Margin" in d.columns else 0.0
    d["Unpaid"] = num(d["Winnings - Unpaid"]) if "Winnings - Unpaid" in d.columns else 0.0
    for src, dst in (("First Slip Issued", "FirstDT"), ("Last Slip Issued", "LastDT")):
        parsed = pd.to_datetime(d[src], format="%d/%m/%y %H:%M:%S", errors="coerce")
        if parsed.isna().all():
            parsed = pd.to_datetime(d[src], errors="coerce", dayfirst=True)
        d[dst] = parsed
    return d


def money(v):
    return f"R {v:,.2f}"


def gw_margin(sub):
    """GW Margin %: weighted average of the source GW Margin % column, weighted by
    Paid In. Uses Aardvark's own per-row figure rather than recomputing (the raw
    columns don't reconstruct GW% reliably). Net Win Margin is Net Win / Paid In."""
    w = sub["PaidIn"].sum()
    if w == 0:
        return 0.0
    return round((sub["GWpct"] * sub["PaidIn"]).sum() / w, 2)


def nwm_margin(sub):
    pin = sub["PaidIn"].sum()
    if pin == 0:
        return 0.0
    return round(sub["NetWin"].sum() / pin * 100, 2)


# ================================================================= EXCEL
def _margin_cf(ws, cell_range):
    """Dark, standout font colors on margin cells: forest green for positive,
    bold dark red for negative (via conditional formatting so exact hex is used)."""
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Font as _F
    green = _F(color="1B7A2F", bold=True)   # strong dark green
    red = _F(color="B00020", bold=True)     # bold dark red
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="greaterThan", formula=["0"], font=green))
    ws.conditional_formatting.add(cell_range, CellIsRule(operator="lessThan", formula=["0"], font=red))


def _add_lookup_card(wb, title, data_key, rows, value_specs):
    """rows: list of [Cashier, Shop, Game, v1, v2, ...]; value_specs: list of
    (label, fmt) for the value columns after Game. Builds a hidden data sheet + a
    hidden per-cashier game list + a neat lookup card sheet."""
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.formula import ArrayFormula
    FONT = "Calibri"
    BLACK = PatternFill("solid", fgColor="000000")
    TITLEF = Font(name=FONT, bold=True, size=18, color="FFFFFF")
    LBL = Font(name=FONT, bold=True, size=11, color="1F3864")
    BIG = Font(name=FONT, bold=True, size=20, color="1F3864")
    CARD = PatternFill("solid", fgColor="F2F5FB")
    PICKF = PatternFill("solid", fgColor="FFF2CC")
    med = Side(style="medium", color="000000"); thin = Side(style="thin", color="000000")
    BOXM = Border(left=med, right=med, top=med, bottom=med)
    BOXT = Border(left=thin, right=thin, top=thin, bottom=thin)
    INTf = '#,##0;(#,##0);-'; MONf = 'R #,##0.00;(R #,##0.00);-'; PCTf = '0.00"%";-0.00"%";-'

    # hidden data sheet
    dname = data_key + "_data"
    dd = wb.create_sheet(dname)
    ncols = 3 + len(value_specs)
    dd.append(["Cashier", "Shop", "Game"] + [lbl for lbl, _ in value_specs])
    for row in rows:
        dd.append(row)
    Ndd = len(rows) + 1
    dd.sheet_state = "hidden"

    shops = sorted({r[1] for r in rows if r[1]})
    lname = data_key + "_lists"
    ll = wb.create_sheet(lname)
    # col A: all shops (shop dropdown source)
    for i, sh in enumerate(shops, 1):
        ll.cell(row=i, column=1, value=sh)

    # Master cashier+shop table on the lists sheet (cols E, F) so formulas can filter it.
    # Distinct cashiers per shop, in shop order.
    from collections import defaultdict
    shop_cashiers = {}
    for sh in shops:
        shop_cashiers[sh] = sorted({r[0] for r in rows if r[1] == sh and r[0]})
    master = []  # (cashier, shop)
    for sh in shops:
        for name in shop_cashiers[sh]:
            master.append((name, sh))
    Mrow = len(master)
    for i, (name, sh) in enumerate(master, 1):
        ll.cell(row=i, column=5, value=name)   # E: cashier
        ll.cell(row=i, column=6, value=sh)      # F: shop

    # Col B: cashiers for the SELECTED shop, packed to the top (no blanks between).
    # Uses IFERROR+SMALL+IF over the master table — reads the shop from the card's B6.
    maxc = max(len(v) for v in shop_cashiers.values()) if shop_cashiers else 1
    for k in range(maxc):
        f = (f"IFERROR(INDEX({lname}!$E$1:$E${Mrow},"
             f"SMALL(IF({lname}!$F$1:$F${Mrow}='{title}'!$B$6,ROW({lname}!$F$1:$F${Mrow})),{k+1})),\"\")")
        ll[f'B{1+k}'] = ArrayFormula(f'B{1+k}', f'={f}')

    # Master cashier+game table (cols H, I) for the game dropdown.
    cg_pairs = []
    seen = set()
    for r in rows:
        if r[0] and r[2] and (r[0], r[2]) not in seen:
            seen.add((r[0], r[2])); cg_pairs.append((r[2], r[0]))  # (game, cashier)
    Grow = len(cg_pairs)
    for i, (g, name) in enumerate(cg_pairs, 1):
        ll.cell(row=i, column=8, value=g)       # H: game
        ll.cell(row=i, column=9, value=name)    # I: cashier
    # Col C: games for the SELECTED cashier (card B9), packed to top.
    from collections import Counter as _Counter
    _gc = _Counter(name for (g, name) in cg_pairs)
    maxg = max(max(_gc.values()) if _gc else 1, 1)
    for k in range(maxg):
        f = (f"IFERROR(INDEX({lname}!$H$1:$H${Grow},"
             f"SMALL(IF({lname}!$I$1:$I${Grow}='{title}'!$B$9,ROW({lname}!$I$1:$I${Grow})),{k+1})),\"\")")
        ll[f'C{1+k}'] = ArrayFormula(f'C{1+k}', f'={f}')
    ll.sheet_state = "hidden"

    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    for col_, w in zip('ABCDEFGH', [3, 22, 22, 18, 18, 4, 3, 3]):
        ws.column_dimensions[col_].width = w
    ws.merge_cells('B2:E3')
    t = ws['B2']; t.value = title; t.font = TITLEF; t.fill = BLACK
    t.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    for cc in ['B2','C2','D2','E2','B3','C3','D3','E3']:
        ws[cc].fill = BLACK
    # Row 5/6: SHOP selector
    ws['B5'] = "SHOP"; ws['B5'].font = LBL
    b6 = ws['B6']; b6.value = shops[0] if shops else ""
    b6.fill = PICKF; b6.font = Font(name=FONT, bold=True, size=12); b6.border = BOXM
    b6.alignment = Alignment(indent=1, vertical="center")
    ws.merge_cells('B6:C6')
    dvS = DataValidation(type="list", formula1=f"={lname}!$A$1:$A${len(shops)}", allow_blank=False)
    ws.add_data_validation(dvS); dvS.add(b6)
    ws.row_dimensions[6].height = 24
    # Row 8/9: CASHIER selector (fixed helper range B), GAME selector (fixed helper range C).
    ws['B8'] = "CASHIER"; ws['B8'].font = LBL
    ws['D8'] = "GAME"; ws['D8'].font = LBL
    b9 = ws['B9']; b9.fill = PICKF; b9.font = Font(name=FONT, bold=True, size=12); b9.border = BOXM
    b9.alignment = Alignment(indent=1, vertical="center")
    ws.merge_cells('B9:C9')
    dvA = DataValidation(type="list", formula1=f"={lname}!$B$1:$B${maxc}", allow_blank=True)
    ws.add_data_validation(dvA); dvA.add(b9)
    d9 = ws['D9']; d9.fill = PICKF; d9.font = Font(name=FONT, bold=True, size=12); d9.border = BOXM
    d9.alignment = Alignment(indent=1, vertical="center")
    ws.merge_cells('D9:E9')
    dvC = DataValidation(type="list", formula1=f"={lname}!$C$1:$C${maxg}", allow_blank=True)
    ws.add_data_validation(dvC); dvC.add(d9)
    ws.row_dimensions[9].height = 24

    # value cards — cashier is B9, game is D9. Stat boxes start at row 12.
    key = '$B$9&"|"&$D$9'
    positions = ['B', 'C', 'D', 'E', 'B', 'C', 'D']
    rowsets = [(12, 13), (12, 13), (12, 13), (12, 13), (15, 16), (15, 16), (15, 16)]
    for idx, (lbl, fmt) in enumerate(value_specs):
        col = positions[idx]; lr, vr = rowsets[idx]
        datacol = get_column_letter(4 + idx)  # D onward in data sheet
        l = ws[f'{col}{lr}']; l.value = lbl; l.font = Font(name=FONT, bold=True, size=9, color="FFFFFF")
        l.fill = BLACK; l.alignment = Alignment(horizontal="center")
        v = ws[f'{col}{vr}']
        if fmt == PCTf:
            v.value = (f'=IF($D$9="","",IFERROR(INDEX({dname}!${datacol}$2:${datacol}${Ndd},'
                       f'MATCH({key},INDEX({dname}!$A$2:$A${Ndd}&"|"&{dname}!$C$2:$C${Ndd},0),0)),0))')
        else:
            v.value = (f'=IF($D$9="","",SUMIFS({dname}!${datacol}$2:${datacol}${Ndd},'
                       f'{dname}!$A$2:$A${Ndd},$B$9,{dname}!$C$2:$C${Ndd},$D$9))')
        v.number_format = fmt; v.font = BIG; v.fill = CARD
        v.alignment = Alignment(horizontal="center", vertical="center"); v.border = BOXT
        ws.row_dimensions[vr].height = 30
        if fmt == PCTf:
            _margin_cf(ws, f'{col}{vr}')
    ws['B18'] = "Pick a shop, then a cashier (only that shop's cashiers show), then a game. Green = positive, red = negative."
    ws['B18'].font = Font(name=FONT, italic=True, size=9, color="808080")
    return ws


def build_workbook(cash, slip):
    """Values-only workbook. Every figure is computed in Python and written as a
    literal number, so all figures display immediately with no need to enable
    editing or recalculate."""
    BRANCHES = sorted(cash["Shop"].dropna().unique())

    cs = cash.groupby(["Shop", "Cashier"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"),
        PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"), Unpaid=("Unpaid", "sum"),
        PaidOut=("PaidOut", "sum") if "PaidOut" in cash.columns else ("Bets", "sum"),
        IsMgr=("IsManager", "max"))
    if "PaidOut" not in cash.columns:
        cs["PaidOut"] = 0.0
    cs["GWpct"] = cs.apply(lambda r: round((r["NetWin"] - r["Unpaid"]) / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
    cs["NWM"] = cs.apply(lambda r: round(r["NetWin"] / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
    cg = cash.groupby(["Shop", "Game"], as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"),
        PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"), Unpaid=("Unpaid", "sum"))
    cg["GWpct"] = cg.apply(lambda r: round((r["NetWin"] - r["Unpaid"]) / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
    cg["NWM"] = cg.apply(lambda r: round(r["NetWin"] / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)

    wb = Workbook()

    # Branch-sheet headers use black fill; other tabs keep navy.
    BR_HDR_FILL = PatternFill("solid", fgColor="000000")
    BR_HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
    BR_THIN = Side(style="thin", color="000000")
    BR_BOX = Border(left=BR_THIN, right=BR_THIN, top=BR_THIN, bottom=BR_THIN)

    def hrow(s, row, headers, widths=None, black=False):
        fill = BR_HDR_FILL if black else HDR_FILL
        font = BR_HDR_FONT if black else HDR_FONT
        brd = BR_BOX if black else BOX
        for i, h in enumerate(headers, start=1):
            c = s.cell(row=row, column=i, value=h)
            c.fill = fill; c.font = font; c.border = brd
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if widths:
            for i, w in enumerate(widths, start=1):
                s.column_dimensions[get_column_letter(i)].width = w
        return row + 1

    def put(s, r, c, val, fmt=None, font=BODY, fill=None, border=True, black_border=False):
        cell = s.cell(row=r, column=c, value=val)
        cell.font = font
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        if border:
            cell.border = BR_BOX if black_border else BOX
        return cell

    # A blank starter sheet exists (wb.active); we'll remove it at the end.
    _starter = wb.active

    PCTf = '0.00"%";-0.00"%";-'
    INTf = '#,##0;[Red](#,##0);-'
    MONf = 'R #,##0.00;[Red](R #,##0.00);-'

    # ---- Cash Lookup card (bets, revokes, revoked sum, margins) ----
    cashd = cash.copy()
    cashd["GWpct"] = cashd.apply(lambda r: round((r["NetWin"] - r["Unpaid"]) / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
    cashd["NWM"] = cashd.apply(lambda r: round(r["NetWin"] / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
    cashd = cashd[(cashd["Bets"] > 0) | (cashd["Revokes"] > 0) | (cashd["RevokedSum"] > 0)]
    cash_rows = [[r["Cashier"], r["Shop"], r["Game"], int(r["Bets"]), int(r["Revokes"]),
                  float(r["RevokedSum"]), float(r["GWpct"]), float(r["NWM"])]
                 for _, r in cashd.iterrows()]
    _add_lookup_card(wb, "Cashier Stats", "cash", cash_rows,
                     [("BETS", INTf), ("REVOKES", INTf), ("REVOKED SUM", MONf),
                      ("GW MARGIN %", PCTf), ("NET WIN MARGIN %", PCTf)])

    for br in BRANCHES:
        s = wb.create_sheet(str(br)[:31])
        s.sheet_view.showGridLines = True                       # whole sheet has clean lines
        s.cell(row=1, column=1, value=f"{br} — Cashier & Game Report").font = TITLE_FONT
        r = 3
        s.cell(row=r, column=1, value="Cashier performance").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Cashier", "Total Bets", "Total Revokes", "Revoked Amount",
                        "Paid Out Amount", "Trents comment", "Branch manager feedback"],
                 [30, 14, 15, 18, 16, 30, 30], black=True)
        sub = cs[cs["Shop"] == br].sort_values("Bets", ascending=False)
        # Branch per-cashier averages drive the red/orange highlight thresholds.
        avg_bets = float(sub["Bets"].mean()) if len(sub) else 0.0
        avg_revokes = float(sub["Revokes"].mean()) if len(sub) else 0.0
        # Highlight fills (black text inside every highlighted block).
        BLUE_H = PatternFill("solid", fgColor="5B9BD5")    # paid out (darker blue)
        RED_H = PatternFill("solid", fgColor="D0342C")     # below threshold (true red)
        ORANGE_H = PatternFill("solid", fgColor="ED9C28")  # above-3 revokes (darker orange)
        BLACKB = Font(name=FONT, size=10, color="000000")
        # Cashier-row threshold highlight: 5000+ -> green on Name + Total Bets,
        # below 5000 -> red on Name + Total Bets (all black bold text now).
        GREEN_NAME = PatternFill("solid", fgColor="92D050")   # the shade you picked
        RED_NAME = PatternFill("solid", fgColor="D0342C")     # true red
        NAME_BLACK = Font(name=FONT, bold=True, size=10, color="000000")
        for _, row_ in sub.iterrows():
            paid_out = float(row_.get("PaidOut", 0) or 0)
            bets = int(row_["Bets"]); revokes = int(row_["Revokes"])
            # Revokes cell still gets the orange flag when > 3.
            rev_fill = ORANGE_H if revokes > 3 else None
            # Paid Out cell still gets the blue flag when > 10,000.
            po_fill = BLUE_H if paid_out > 10000 else None
            # Name + Total Bets highlight by threshold (black bold text in both cases).
            name_fill = GREEN_NAME if bets >= 5000 else RED_NAME
            put(s, r, 1, row_["Cashier"], None, NAME_BLACK, name_fill, black_border=True)
            put(s, r, 2, bets, INT_FMT, NAME_BLACK, name_fill, black_border=True)
            put(s, r, 3, revokes, INT_FMT, BLACKB if rev_fill else BODY, rev_fill, black_border=True)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT, border=True, black_border=True)
            put(s, r, 5, paid_out if paid_out > 0 else None, MON_FMT,
                BLACKB if po_fill else BODY, po_fill, black_border=True)
            put(s, r, 6, None, border=True, black_border=True)
            put(s, r, 7, None, border=True, black_border=True)
            r += 1
        brf = cash[cash["Shop"] == br]
        put(s, r, 1, "BRANCH TOTAL", font=LBL_FONT, fill=SUB_FILL, black_border=True)
        put(s, r, 2, int(sub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 3, int(sub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 4, float(sub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 5, float(sub["PaidOut"].sum()) if "PaidOut" in sub.columns else None, MON_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 6, None, border=True, black_border=True)
        put(s, r, 7, None, border=True, black_border=True)
        r += 1
        # ---- Branch average per cashier: colour the Avg Bets cell only ----
        # above 5000 -> green ; 4301 to 5000 -> orange ; 4300 and below -> red
        GREEN_AVG = PatternFill("solid", fgColor="00B050")
        ORANGE_AVG = PatternFill("solid", fgColor="FFA500")
        RED_AVG = PatternFill("solid", fgColor="FF0000")
        avg_bets_rounded = round(avg_bets, 0)
        if avg_bets_rounded > 5000:
            avg_fill = GREEN_AVG
        elif avg_bets_rounded > 4300:
            avg_fill = ORANGE_AVG
        else:
            avg_fill = RED_AVG
        # Black bold text on the average value + counted rows.
        avg_font = Font(name=FONT, bold=True, size=10, color="000000")
        put(s, r, 1, "BRANCH AVERAGE (per cashier)", font=avg_font, black_border=True)
        put(s, r, 2, avg_bets_rounded, INT_FMT, avg_font, avg_fill, black_border=True)
        put(s, r, 3, round(avg_revokes, 1), '#,##0.0;(#,##0.0);-', avg_font, black_border=True)
        put(s, r, 4, round(float(sub["RevSum"].mean()), 2) if len(sub) else 0, MON_FMT, avg_font, black_border=True)
        put(s, r, 5, None, black_border=True)
        put(s, r, 6, None, black_border=True)
        put(s, r, 7, None, black_border=True)
        r += 1
        ORANGE = PatternFill("solid", fgColor="E8730C")
        put(s, r, 1, "Cashiers counted", font=avg_font, fill=ORANGE, black_border=True)
        put(s, r, 2, int(len(sub)), INT_FMT, avg_font, ORANGE, black_border=True)
        put(s, r, 3, None, black_border=True)
        put(s, r, 4, None, black_border=True)
        put(s, r, 5, None, black_border=True)
        put(s, r, 6, None, black_border=True)
        put(s, r, 7, None, black_border=True)
        r += 3

        subn = sub[~sub["IsMgr"]] if "IsMgr" in sub.columns else sub
        subn = subn if len(subn) else sub
        mb = subn.loc[subn["Bets"].idxmax()]
        lb = subn.loc[subn["Bets"].idxmin()]
        mr = subn.loc[subn["Revokes"].idxmax()]
        s.cell(row=r, column=1, value="Branch highlights").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Measure", "Cashier", "Value"], [40, 30, 18], black=True)
        for label, who, val, fmt in [
            ("Most bets (cashier)", mb["Cashier"], int(mb["Bets"]), INT_FMT),
            ("Least bets (cashier)", lb["Cashier"], int(lb["Bets"]), INT_FMT),
            ("Most revokes (cashier)", mr["Cashier"], int(mr["Revokes"]), INT_FMT),
            ("Revoked amount of that cashier", None, float(mr["RevSum"]), MON_FMT),
        ]:
            put(s, r, 1, label, black_border=True)
            if who is not None:
                put(s, r, 2, who, black_border=True)
            else:
                s.cell(row=r, column=2).border = BR_BOX
            put(s, r, 3, val, fmt, black_border=True)
            r += 1
        r += 2

        s.cell(row=r, column=1, value="Bets & revokes per game").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Game", "Bets", "Revokes", "Revoked Amount",
                        "GW Margin %"], [30, 14, 15, 18, 14], black=True)
        gsub = cg[cg["Shop"] == br].sort_values("Bets", ascending=False)
        gfirst = r
        for _, row_ in gsub.iterrows():
            put(s, r, 1, row_["Game"], black_border=True)
            put(s, r, 2, int(row_["Bets"]), INT_FMT, black_border=True)
            put(s, r, 3, int(row_["Revokes"]), INT_FMT, black_border=True)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT, black_border=True)
            put(s, r, 5, float(row_["GWpct"]), PCT_FMT, black_border=True)
            r += 1
        if r > gfirst:
            _margin_cf(s, f"E{gfirst}:E{r-1}")
        brf = cash[cash["Shop"] == br]
        put(s, r, 1, "TOTAL", font=LBL_FONT, fill=SUB_FILL, black_border=True)
        put(s, r, 2, int(gsub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 3, int(gsub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 4, float(gsub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL, black_border=True)
        put(s, r, 5, gw_pct(brf), PCT_FMT, LBL_FONT, SUB_FILL, black_border=True)
        r += 2
        if len(gsub):
            gb = gsub.loc[gsub["Bets"].idxmax()]; gr = gsub.loc[gsub["Revokes"].idxmax()]
            BRIGHT = PatternFill("solid", fgColor="FFEB00")  # bright yellow
            BLKB = Font(name="Calibri", bold=True, color="000000")
            put(s, r, 1, "Game with most bets", font=BLKB, fill=BRIGHT, black_border=True)
            put(s, r, 2, gb["Game"], font=BLKB, fill=BRIGHT, black_border=True)
            put(s, r, 3, int(gb["Bets"]), INT_FMT, BLKB, BRIGHT, black_border=True)
            r += 1
            put(s, r, 1, "Game with most revokes", font=BLKB, fill=BRIGHT, black_border=True)
            put(s, r, 2, gr["Game"], font=BLKB, fill=BRIGHT, black_border=True)
            put(s, r, 3, int(gr["Revokes"]), INT_FMT, BLKB, BRIGHT, black_border=True)
        s.freeze_panes = "A4"

    ac = wb.create_sheet("All Cashiers")
    ac.cell(row=1, column=1, value="All Cashiers — All Branches").font = TITLE_FONT
    r = 3
    ac.cell(row=r, column=1, value="Cashier performance").font = LBL_FONT
    r += 1
    r = hrow(ac, r, ["Cashier", "Branch", "Total Bets", "Total Revokes", "Revoked Amount",
                     "GW Margin %", "Net Win Margin %"], [30, 16, 14, 15, 18, 14, 16])
    allc = cs.sort_values("Bets", ascending=False)
    ac_first = r
    for _, row_ in allc.iterrows():
        put(ac, r, 1, row_["Cashier"]); put(ac, r, 2, row_["Shop"])
        put(ac, r, 3, int(row_["Bets"]), INT_FMT)
        put(ac, r, 4, int(row_["Revokes"]), INT_FMT)
        put(ac, r, 5, float(row_["RevSum"]), MON_FMT)
        put(ac, r, 6, float(row_["GWpct"]), PCT_FMT)
        put(ac, r, 7, float(row_["NWM"]), PCT_FMT)
        r += 1
    if r > ac_first:
        _margin_cf(ac, f"F{ac_first}:G{r-1}")
    put(ac, r, 1, "GRAND TOTAL", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 2, "", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 3, int(allc["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 4, int(allc["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 5, float(allc["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 6, gw_pct(cash), PCT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 7, nwm_pct(cash), PCT_FMT, LBL_FONT, SUB_FILL)
    r += 1
    put(ac, r, 1, "AVERAGE (per cashier)", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 2, "", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 3, round(float(allc["Bets"].mean()), 0), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 4, round(float(allc["Revokes"].mean()), 1), '#,##0.0;(#,##0.0);-', LBL_FONT, SUB_FILL)
    put(ac, r, 5, round(float(allc["RevSum"].mean()), 2), MON_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 6, "", fill=SUB_FILL); put(ac, r, 7, "", fill=SUB_FILL)
    r += 3

    # All-branches highlights (managers excluded where noted)
    csn = allc[~allc["IsMgr"]] if "IsMgr" in allc.columns else allc
    csn = csn if len(csn) else allc
    mb = allc.loc[allc["Bets"].idxmax()]
    lb = csn.loc[csn["Bets"].idxmin()]
    mr = csn.loc[csn["Revokes"].idxmax()]
    ma = csn.loc[csn["RevSum"].idxmax()]
    ac.cell(row=r, column=1, value="Overall highlights").font = LBL_FONT
    r += 1
    r = hrow(ac, r, ["Measure", "Cashier", "Branch", "Value"], [40, 28, 16, 18])
    for label, who, brc, val, fmt in [
        ("Most bets (cashier)", mb["Cashier"], mb["Shop"], int(mb["Bets"]), INT_FMT),
        ("Least bets (cashier, managers excluded)", lb["Cashier"], lb["Shop"], int(lb["Bets"]), INT_FMT),
        ("Most revokes (cashier, managers excluded)", mr["Cashier"], mr["Shop"], int(mr["Revokes"]), INT_FMT),
        ("Revoked amount of that cashier", None, None, float(mr["RevSum"]), MON_FMT),
        ("Highest revoked amount (managers excluded)", ma["Cashier"], ma["Shop"], float(ma["RevSum"]), MON_FMT),
    ]:
        put(ac, r, 1, label)
        if who is not None: put(ac, r, 2, who)
        else: ac.cell(row=r, column=2).border = BOX
        if brc is not None: put(ac, r, 3, brc)
        else: ac.cell(row=r, column=3).border = BOX
        put(ac, r, 4, val, fmt)
        r += 1
    r += 2

    # Per-game breakdown across all branches
    ac.cell(row=r, column=1, value="Bets & revokes per game (all branches)").font = LBL_FONT
    r += 1
    r = hrow(ac, r, ["Game", "Bets", "Revokes", "Revoked Amount", "GW Margin %", "Net Win Margin %"],
             [30, 14, 15, 18, 14, 16])
    allg = cash.groupby("Game", as_index=False).agg(
        Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"),
        PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"), Unpaid=("Unpaid", "sum"))
    allg["GWpct"] = allg.apply(lambda x: round((x["NetWin"] - x["Unpaid"]) / x["PaidIn"] * 100, 2) if x["PaidIn"] else 0.0, axis=1)
    allg["NWM"] = allg.apply(lambda x: round(x["NetWin"] / x["PaidIn"] * 100, 2) if x["PaidIn"] else 0.0, axis=1)
    allg = allg.sort_values("Bets", ascending=False)
    allg_first = r
    for _, row_ in allg.iterrows():
        put(ac, r, 1, row_["Game"])
        put(ac, r, 2, int(row_["Bets"]), INT_FMT)
        put(ac, r, 3, int(row_["Revokes"]), INT_FMT)
        put(ac, r, 4, float(row_["RevSum"]), MON_FMT)
        put(ac, r, 5, float(row_["GWpct"]), PCT_FMT)
        put(ac, r, 6, float(row_["NWM"]), PCT_FMT)
        r += 1
    if r > allg_first:
        _margin_cf(ac, f"E{allg_first}:F{r-1}")
    put(ac, r, 1, "TOTAL", font=LBL_FONT, fill=SUB_FILL)
    put(ac, r, 2, int(allg["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 3, int(allg["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 4, float(allg["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 5, gw_pct(cash), PCT_FMT, LBL_FONT, SUB_FILL)
    put(ac, r, 6, nwm_pct(cash), PCT_FMT, LBL_FONT, SUB_FILL)

    sm = wb.create_sheet("Branch Performance", 0)
    sm.cell(row=1, column=1, value="Branch Performance").font = TITLE_FONT
    r = hrow(sm, 4, ["Branch", "Cashiers", "Total Bets", "Total Revokes", "Revoked Amount",
                     "Avg Bets / Cashier", "Avg Revokes / Cashier", "Total Betslips",
                     "GW Margin %"],
             [16, 11, 14, 14, 16, 17, 19, 14, 13])
    sm_first = r
    for br in BRANCHES:
        sub = cs[cs["Shop"] == br]; subs = slip[slip["Shop"] == br]
        ncash = int(sub["Cashier"].nunique()); bets = int(sub["Bets"].sum()); revs = int(sub["Revokes"].sum())
        rsum = float(sub["RevSum"].sum())
        put(sm, r, 1, br); put(sm, r, 2, ncash, INT_FMT)
        put(sm, r, 3, bets, INT_FMT); put(sm, r, 4, revs, INT_FMT); put(sm, r, 5, rsum, MON_FMT)
        put(sm, r, 6, round(bets / ncash, 0) if ncash else 0, INT_FMT)
        put(sm, r, 7, round(revs / ncash, 1) if ncash else 0, '#,##0.0;(#,##0.0);-')
        put(sm, r, 8, int(subs["BetSlips"].sum()), INT_FMT)
        put(sm, r, 9, gw_margin(subs), PCT_FMT)
        r += 1
    tcash = int(cs["Cashier"].nunique()); tbets = int(cs["Bets"].sum()); trev = int(cs["Revokes"].sum())
    trsum = float(cs["RevSum"].sum()); tslip = int(slip["BetSlips"].sum())
    put(sm, r, 1, "ALL BRANCHES", font=LBL_FONT, fill=SUB_FILL)
    put(sm, r, 2, tcash, INT_FMT, LBL_FONT, SUB_FILL); put(sm, r, 3, tbets, INT_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 4, trev, INT_FMT, LBL_FONT, SUB_FILL); put(sm, r, 5, trsum, MON_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 6, round(tbets / tcash, 0) if tcash else 0, INT_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 7, round(trev / tcash, 1) if tcash else 0, '#,##0.0;(#,##0.0);-', LBL_FONT, SUB_FILL)
    put(sm, r, 8, tslip, INT_FMT, LBL_FONT, SUB_FILL)
    put(sm, r, 9, gw_margin(slip), PCT_FMT, LBL_FONT, SUB_FILL)
    _margin_cf(sm, f"I{sm_first}:I{r}")
    r += 3

    # ---- Games & bets per game, per branch ----
    sm.cell(row=r, column=1, value="Bets & revokes per game — by branch").font = LBL_FONT
    r += 1
    for br in BRANCHES:
        sm.cell(row=r, column=1, value=str(br)).font = LBL_FONT
        r += 1
        r = hrow(sm, r, ["Game", "Bets", "Revokes", "Revoked Amount", "GW Margin %"],
                 [30, 14, 15, 18, 14])
        gsub = cg[cg["Shop"] == br].sort_values("Bets", ascending=False)
        gfirst = r
        for _, row_ in gsub.iterrows():
            put(sm, r, 1, row_["Game"])
            put(sm, r, 2, int(row_["Bets"]), INT_FMT)
            put(sm, r, 3, int(row_["Revokes"]), INT_FMT)
            put(sm, r, 4, float(row_["RevSum"]), MON_FMT)
            put(sm, r, 5, float(row_["GWpct"]), PCT_FMT)
            r += 1
        if r > gfirst:
            _margin_cf(sm, f"E{gfirst}:E{r-1}")
        brf = cash[cash["Shop"] == br]
        put(sm, r, 1, "TOTAL", font=LBL_FONT, fill=SUB_FILL)
        put(sm, r, 2, int(gsub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(sm, r, 3, int(gsub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(sm, r, 4, float(gsub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
        put(sm, r, 5, gw_pct(brf), PCT_FMT, LBL_FONT, SUB_FILL)
        r += 2

    csn = cs[~cs["IsMgr"]] if "IsMgr" in cs.columns else cs
    csn = csn if len(csn) else cs
    mb = csn.loc[csn["Bets"].idxmax()]; lb = csn.loc[csn["Bets"].idxmin()]
    mr = csn.loc[csn["Revokes"].idxmax()]; ma = csn.loc[csn["RevSum"].idxmax()]
    gt = cash.groupby("Game", as_index=False)["Bets"].sum(); gtop = gt.loc[gt["Bets"].idxmax()]
    sm.cell(row=r, column=1, value="Overall highlights (all branches)").font = LBL_FONT
    r += 1
    r = hrow(sm, r, ["Measure", "Cashier / Game", "Branch", "Value"], [40, 28, 16, 18])
    for label, who, brc, val, fmt in [
        ("Most bets — cashier", mb["Cashier"], mb["Shop"], int(mb["Bets"]), INT_FMT),
        ("Least bets — cashier", lb["Cashier"], lb["Shop"], int(lb["Bets"]), INT_FMT),
        ("Most revokes — cashier", mr["Cashier"], mr["Shop"], int(mr["Revokes"]), INT_FMT),
        ("Revoked amount of that cashier", None, None, float(mr["RevSum"]), MON_FMT),
        ("Highest revoked amount", ma["Cashier"], ma["Shop"], float(ma["RevSum"]), MON_FMT),
        ("Game with most bets", gtop["Game"], "All branches", int(gtop["Bets"]), INT_FMT),
    ]:
        put(sm, r, 1, label)
        if who is not None: put(sm, r, 2, who)
        else: sm.cell(row=r, column=2).border = BOX
        if brc is not None: put(sm, r, 3, brc)
        else: sm.cell(row=r, column=3).border = BOX
        put(sm, r, 4, val, fmt)
        r += 1

    # remove the blank starter sheet openpyxl created
    if _starter in wb.worksheets:
        wb.remove(_starter)

    # Order tabs: Branch Performance, All Cashiers, then the branch sheets in the
    # order the boss specified (Potchefstroom last), then the Cashier Stats card.
    branch_order = ["Malvern", "Randburg", "Pretoria", "White River", "Potchefstroom"]
    ordered_branches = [b for b in branch_order if b in [str(x)[:31] for x in BRANCHES]]
    ordered_branches += [str(b)[:31] for b in BRANCHES if str(b)[:31] not in ordered_branches]
    desired = ["Branch Performance", "All Cashiers"] + ordered_branches + ["Cashier Stats"]
    order = {name: i for i, name in enumerate(desired)}
    wb._sheets.sort(key=lambda ws: order.get(ws.title, len(order)))

    # Force Excel to fully recalculate when the file opens, so the dependent-dropdown
    # array formulas (shop -> cashier -> game) populate immediately instead of showing
    # an empty list until a manual recalc.
    try:
        wb.calculation.calcMode = "auto"
        wb.calculation.fullCalcOnLoad = True
    except Exception:
        from openpyxl.workbook.properties import CalcProperties
        wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True)

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
**Cash Operations Summary** needs: `Cashier`, `Shop`, `Game`, `Paid Out - Revoked Count`, `Revoked Sum`.

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

cash = prep_cash(cash_raw, slip_raw, drop_mgr)
slip = prep_slip(slip_raw)
if cash.empty:
    st.warning("No cashier rows left after filtering.")
    st.stop()

BRANCHES = sorted(cash["Shop"].dropna().unique())
cs = cash.groupby(["Shop", "Cashier"], as_index=False).agg(
    Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"),
    PaidIn=("PaidIn", "sum"), NetWin=("NetWin", "sum"), Unpaid=("Unpaid", "sum"),
    IsMgr=("IsManager", "max"))
cs["GW %"] = cs.apply(lambda r: round((r["NetWin"] - r["Unpaid"]) / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
cs["Net win %"] = cs.apply(lambda r: round(r["NetWin"] / r["PaidIn"] * 100, 2) if r["PaidIn"] else 0.0, axis=1)
cg = cash.groupby(["Shop", "Game"], as_index=False).agg(
    Bets=("Bets", "sum"), Revokes=("Revokes", "sum"), RevSum=("RevokedSum", "sum"))

hdr = f"Period: {period}" if period.strip() else ""
if hdr:
    st.caption(hdr)
if drop_mgr:
    n_mgr = prep_cash(cash_raw, slip_raw, False)["IsManager"].sum()
    st.caption(f"Manager accounts excluded ({int(n_mgr)} source rows).")

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total bets", f"{int(cash['Bets'].sum()):,}")
k2.metric("Total revokes", f"{int(cash['Revokes'].sum()):,}")
k3.metric("Revoked amount", money(cash["RevokedSum"].sum()))
k4.metric("Total betslips", f"{int(slip['BetSlips'].sum()):,}")
gw = gw_margin(slip)
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
    sb["GW margin %"] = sb["Shop"].map(lambda b: gw_margin(slip[slip["Shop"] == b]))
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
    disp = sub.rename(columns={"RevSum": "Revoked amount"})
    st.dataframe(disp[["Cashier", "Bets", "Revokes", "Revoked amount", "GW %", "Net win %"]],
                 use_container_width=True, hide_index=True,
                 column_config={
                     "Bets": st.column_config.NumberColumn(format="%d"),
                     "Revokes": st.column_config.NumberColumn(format="%d"),
                     "Revoked amount": st.column_config.NumberColumn(format="R %.2f"),
                     "GW %": st.column_config.NumberColumn(format="%.2f%%"),
                     "Net win %": st.column_config.NumberColumn(format="%.2f%%")})

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
    ss["GW margin %"] = ss["Shop"].map(lambda b: gw_margin(slip[slip["Shop"] == b]))
    ss["Net win margin %"] = ss["Shop"].map(lambda b: nwm_margin(slip[slip["Shop"] == b]))
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
    st.dataframe(allc[["Cashier", "Branch", "Bets", "Revokes", "Revoked amount", "GW %", "Net win %"]],
                 use_container_width=True, hide_index=True,
                 column_config={
                     "Bets": st.column_config.NumberColumn(format="%d"),
                     "Revokes": st.column_config.NumberColumn(format="%d"),
                     "Revoked amount": st.column_config.NumberColumn(format="R %.2f"),
                     "GW %": st.column_config.NumberColumn(format="%.2f%%"),
                     "Net win %": st.column_config.NumberColumn(format="%.2f%%")})

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
