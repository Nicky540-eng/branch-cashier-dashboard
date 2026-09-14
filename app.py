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
    NAVY = PatternFill("solid", fgColor="1F3864")
    TITLEF = Font(name=FONT, bold=True, size=18, color="FFFFFF")
    LBL = Font(name=FONT, bold=True, size=11, color="1F3864")
    BIG = Font(name=FONT, bold=True, size=20, color="1F3864")
    CARD = PatternFill("solid", fgColor="F2F5FB")
    PICKF = PatternFill("solid", fgColor="FFF2CC")
    med = Side(style="medium", color="1F3864"); thin = Side(style="thin", color="C9D3E8")
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
    t = ws['B2']; t.value = title; t.font = TITLEF; t.fill = NAVY
    t.alignment = Alignment(vertical="center", horizontal="left", indent=1)
    for cc in ['B2','C2','D2','E2','B3','C3','D3','E3']:
        ws[cc].fill = NAVY
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
        l.fill = NAVY; l.alignment = Alignment(horizontal="center")
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
        s.cell(row=1, column=1, value=f"{br} — Cashier & Game Report").font = TITLE_FONT
        r = 3
        s.cell(row=r, column=1, value="Cashier performance").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Cashier", "Total Bets", "Total Revokes", "Revoked Amount",
                        "Paid Out Amount", "Trents comment", "Branch manager feedback"],
                 [30, 14, 15, 18, 16, 30, 30])
        sub = cs[cs["Shop"] == br].sort_values("Bets", ascending=False)
        # Branch per-cashier averages drive the red/orange highlight thresholds.
        avg_bets = float(sub["Bets"].mean()) if len(sub) else 0.0
        avg_revokes = float(sub["Revokes"].mean()) if len(sub) else 0.0
        # Highlight fills (black text inside every highlighted block).
        BLUE_H = PatternFill("solid", fgColor="5B9BD5")    # paid out (darker blue)
        RED_H = PatternFill("solid", fgColor="D0342C")     # below-avg bets (true red)
        ORANGE_H = PatternFill("solid", fgColor="ED9C28")  # above-3 revokes (darker orange)
        BLACKB = Font(name=FONT, size=10, color="000000")
        for _, row_ in sub.iterrows():
            paid_out = float(row_.get("PaidOut", 0) or 0)
            bets = int(row_["Bets"]); revokes = int(row_["Revokes"])
            # Per-cell highlights so every flag on a cashier shows at once:
            #   paid out  -> blue on the Paid Out cell
            #   below-avg bets -> red on the Bets cell
            #   more than 3 revokes -> orange on the Revokes cell
            bets_fill = RED_H if bets < avg_bets else None
            rev_fill = ORANGE_H if revokes > 3 else None
            po_fill = BLUE_H if paid_out > 10000 else None
            put(s, r, 1, row_["Cashier"])
            put(s, r, 2, bets, INT_FMT, BLACKB if bets_fill else BODY, bets_fill)
            put(s, r, 3, revokes, INT_FMT, BLACKB if rev_fill else BODY, rev_fill)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT)
            put(s, r, 5, paid_out if paid_out > 0 else None, MON_FMT,
                BLACKB if po_fill else BODY, po_fill)
            put(s, r, 6, None)
            put(s, r, 7, None)
            r += 1
        brf = cash[cash["Shop"] == br]
        put(s, r, 1, "BRANCH TOTAL", font=LBL_FONT, fill=SUB_FILL)
        put(s, r, 2, int(sub["Bets"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 3, int(sub["Revokes"].sum()), INT_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 4, float(sub["RevSum"].sum()), MON_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 5, float(sub["PaidOut"].sum()) if "PaidOut" in sub.columns else None, MON_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 6, None, None, LBL_FONT, SUB_FILL)
        put(s, r, 7, None, None, LBL_FONT, SUB_FILL)
        r += 1
        # ---- Branch average per cashier: colour the Avg Bets cell ----
        GREEN_AVG = PatternFill("solid", fgColor="00B050")   # green: above 5000
        ORANGE_AVG = PatternFill("solid", fgColor="FFA500")  # orange: 4301 - 5000
        RED_AVG = PatternFill("solid", fgColor="FF0000")     # red: 4300 and below
        avg_bets_rounded = round(avg_bets, 0)
        if avg_bets_rounded > 5000:
            avg_fill = GREEN_AVG
        elif avg_bets_rounded >= 4301:
            avg_fill = ORANGE_AVG
        else:
            avg_fill = RED_AVG
        avg_font = Font(name=FONT, bold=True, size=10, color="FFFFFF")
        put(s, r, 1, "BRANCH AVERAGE (per cashier)", font=LBL_FONT, fill=SUB_FILL)
        put(s, r, 2, avg_bets_rounded, INT_FMT, avg_font, avg_fill)
        put(s, r, 3, round(avg_revokes, 1), '#,##0.0;(#,##0.0);-', LBL_FONT, SUB_FILL)
        put(s, r, 4, round(float(sub["RevSum"].mean()), 2) if len(sub) else 0, MON_FMT, LBL_FONT, SUB_FILL)
        put(s, r, 5, None, None, LBL_FONT, SUB_FILL)
        put(s, r, 6, None, None, LBL_FONT, SUB_FILL)
        put(s, r, 7, None, None, LBL_FONT, SUB_FILL)
        r += 1
        ORANGE = PatternFill("solid", fgColor="E8730C")
        WHITEB2 = Font(name="Calibri", bold=True, color="FFFFFF")
        put(s, r, 1, "Cashiers counted", font=WHITEB2, fill=ORANGE)
        put(s, r, 2, int(len(sub)), INT_FMT, WHITEB2, ORANGE)
        r += 3

        subn = sub[~sub["IsMgr"]] if "IsMgr" in sub.columns else sub
        subn = subn if len(subn) else sub
        mb = subn.loc[subn["Bets"].idxmax()]
        lb = subn.loc[subn["Bets"].idxmin()]
        mr = subn.loc[subn["Revokes"].idxmax()]
        s.cell(row=r, column=1, value="Branch highlights").font = LBL_FONT
        r += 1
        r = hrow(s, r, ["Measure", "Cashier", "Value"], [40, 30, 18])
        for label, who, val, fmt in [
            ("Most bets (cashier)", mb["Cashier"], int(mb["Bets"]), INT_FMT),
            ("Least bets (cashier)", lb["Cashier"], int(lb["Bets"]), INT_FMT),
            ("Most revokes (cashier)", mr["Cashier"], int(mr["Revokes"]), INT_FMT),
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
        r = hrow(s, r, ["Game", "Bets", "Revokes", "Revoked Amount",
                        "GW Margin %"], [30, 14, 15, 18, 14])
        gsub = cg[cg["Shop"] == br].sort_values("Bets", ascending=False)
        gfirst = r
        for _, row_ in gsub.iterrows():
            put(s, r, 1, row_["Game"])
            put(s, r, 2, int(row_["Bets"]), INT_FMT)
            put(s, r, 3, int(row_["Revokes"]), INT_FMT)
            put(s, r, 4, float(row_["RevSum"]), MON_FMT)
            put(s, r, 5, float(row_["GWpct"]), PCT_FMT)
            r += 1
        if r > gfirst:
            _margin_cf(s, f"E{gfirst}:E{r-1}")
        brf = cash[cash["Shop"] == br]
        put(s,
