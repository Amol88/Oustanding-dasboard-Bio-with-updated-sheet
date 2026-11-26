import io
import re
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(
    page_title="Outstanding Tracker Dashboard",
    page_icon="📊",
    layout="wide",
)

# ---------- HELPER FUNCTIONS ----------

def find_header_row(file_bytes: bytes, required_cols=("Type", "Serial")) -> int:
    """Excel sheet mein automatically header row ka index dhundhta hai."""
    preview = pd.read_excel(io.BytesIO(file_bytes), header=None, nrows=20, dtype=str)
    for idx, row in preview.iterrows():
        vals = {str(v).strip().lower() for v in row.dropna().values}
        if all(c.lower() in vals for c in required_cols):
            return int(idx)
    return 6


def read_register(file_bytes: bytes) -> pd.DataFrame:
    header_idx = find_header_row(file_bytes)
    df = pd.read_excel(io.BytesIO(file_bytes), header=header_idx, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def normalize_bill_key(doc_type: str, number: str) -> str:
    t = str(doc_type).strip().upper() if pd.notna(doc_type) else ""
    n = str(number).strip().upper() if pd.notna(number) else ""
    if not t or not n or t == "NAN" or n == "NAN":
        return ""
    n_clean = n.lstrip("0") or "0"
    return f"{t}-{n_clean}"


def extract_paid_bills_from_ledger(ml_df: pd.DataFrame) -> set:
    """Multi Ledger ke ICICI rows ki Narration se saare paid Bill No(s) nikalta hai."""
    paid_keys = set()
    if ml_df is None or ml_df.empty:
        return paid_keys
    if "Particulars" not in ml_df.columns or "Narration" not in ml_df.columns:
        return paid_keys

    icici_mask = ml_df["Particulars"].astype(str).str.contains("ICICI", case=False, na=False)
    for narr in ml_df.loc[icici_mask, "Narration"].dropna():
        clean = str(narr).replace("_x000D_", " ").replace("\r", " ").replace("\n", " ")
        match = re.search(r"Bill\s*No\(s\)\s*:\s*(.*)", clean, re.IGNORECASE)
        if not match:
            continue
        raw_part = match.group(1)
        tokens = [tok.strip().upper() for tok in re.split(r"[,;\s]+", raw_part) if tok.strip()]
        for tok in tokens:
            paid_keys.add(tok)
            if "-" in tok:
                p1, p2 = tok.split("-", 1)
                paid_keys.add(f"{p1}-{p2.lstrip('0') or '0'}")
    return paid_keys


def merge_registers(old_df: pd.DataFrame | None, new_df: pd.DataFrame | None) -> pd.DataFrame:
    """Purane Master register aur naye register ko invoice-wise deduplicate karke merge karta hai."""
    if old_df is None or old_df.empty:
        return new_df.copy() if new_df is not None else pd.DataFrame()
    if new_df is None or new_df.empty:
        return old_df.copy()

    def tag_invoice_keys(df: pd.DataFrame) -> pd.DataFrame:
        temp = df.copy()
        has_party = temp["Party Name"].notna() & (temp["Party Name"].astype(str).str.strip() != "")
        temp["_grp"] = has_party.cumsum()
        hdrs = temp[has_party].copy()
        prefix_col = hdrs["Prefix"].fillna("").str.strip().str.upper() if "Prefix" in hdrs.columns else ""
        hdrs["_inv_uid"] = (
            hdrs["Type"].fillna("").str.strip().str.upper()
            + "|"
            + prefix_col
            + "|"
            + hdrs["Serial"].fillna("").str.strip().str.upper()
        )
        grp_to_uid = dict(zip(hdrs["_grp"], hdrs["_inv_uid"]))
        temp["_inv_uid"] = temp["_grp"].map(grp_to_uid).fillna("")
        return temp

    old_t = tag_invoice_keys(old_df)
    new_t = tag_invoice_keys(new_df)

    existing_uids = set(old_t["_inv_uid"].dropna().unique()) - {""}
    new_only = new_t[~new_t["_inv_uid"].isin(existing_uids)].copy()

    merged = pd.concat(
        [old_t.drop(columns=["_grp", "_inv_uid"]), new_only.drop(columns=["_grp", "_inv_uid"])],
        ignore_index=True,
    )
    return merged


def read_master_backup(backup_bytes: bytes) -> tuple[pd.DataFrame, pd.DataFrame, set]:
    """Purani Master Backup (.xlsx) se Sales_Raw, Purchase_Raw, aur Paid_Bills load karta hai."""
    xls = pd.ExcelFile(io.BytesIO(backup_bytes))
    sales_raw = pd.read_excel(xls, sheet_name="Sales_Raw", dtype=str) if "Sales_Raw" in xls.sheet_names else pd.DataFrame()
    purchase_raw = pd.read_excel(xls, sheet_name="Purchase_Raw", dtype=str) if "Purchase_Raw" in xls.sheet_names else pd.DataFrame()
    paid_keys = set()
    if "Paid_Bills" in xls.sheet_names:
        pb_df = pd.read_excel(xls, sheet_name="Paid_Bills", dtype=str)
        if "Paid_Key" in pb_df.columns:
            paid_keys = set(pb_df["Paid_Key"].dropna().astype(str).str.strip().str.upper())
    return sales_raw, purchase_raw, paid_keys


def build_master_backup_bytes(sales_raw: pd.DataFrame, purchase_raw: pd.DataFrame, paid_keys: set) -> bytes:
    """Agli baar bina re-upload ke incremental update ke liye Master Backup (.xlsx) banata hai."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        sales_raw.to_excel(writer, sheet_name="Sales_Raw", index=False)
        purchase_raw.to_excel(writer, sheet_name="Purchase_Raw", index=False)
        pd.DataFrame({"Paid_Key": sorted(paid_keys)}).to_excel(writer, sheet_name="Paid_Bills", index=False)
    buf.seek(0)
    return buf.getvalue()


def extract_parent_party(party_name: str) -> str:
    """Branch/unit naam hata kar Main Parent Company (Group) naam nikalta hai (e.g. CIPLA)."""
    if not isinstance(party_name, str) or not party_name.strip():
        return "UNKNOWN"
    s = party_name.strip().upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.split(r"[-/,]", s)[0]
    s = re.sub(
        r"\b(PVT\.?|PRIVATE|LTD\.?|LIMITED|LLP|INC\.?|CORP\.?|CORPORATION|"
        r"UNIT[\s\-]*[IVX0-9]+|PHARMA(CEUTICALS?)?|LABORATORIES|LABS?|"
        r"INDUSTRIES|INDIA|SOLUTIONS|HEALTHCARE|LIFESCIENCES|BIOSCIENCES|"
        r"FORMULATION|R&D|CWH)\b",
        " ",
        s,
    )
    s = re.sub(r"[^A-Z0-9\s&]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    words = s.split()
    if not words:
        return party_name.strip().upper()
    return " ".join(words[:2]) if len(words) > 1 and len(words[0]) <= 3 else words[0]


def get_financial_year(dt) -> str:
    """Indian Financial Year (1 Apr - 31 Mar) return karta hai, e.g. 2025-26."""
    if pd.isna(dt):
        return "Unknown"
    yr = dt.year
    if dt.month >= 4:
        return f"{yr}-{str(yr + 1)[-2:]}"
    return f"{yr - 1}-{str(yr)[-2:]}"


def get_last_financial_year(as_on: date) -> str:
    curr_start = as_on.year if as_on.month >= 4 else as_on.year - 1
    prev_start = curr_start - 1
    return f"{prev_start}-{str(prev_start + 1)[-2:]}"


def get_overdue_bucket(days: int) -> str:
    if pd.isna(days):
        return "0-30 Days"
    if days <= 30:
        return "0-30 Days"
    if days <= 45:
        return "31-45 Days"
    if days <= 90:
        return "46-90 Days"
    return "90+ Days"


def process_register_outstanding(
    reg_df: pd.DataFrame,
    paid_keys: set,
    as_on_date: date,
    is_purchase: bool = False,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if reg_df is None or reg_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    main_df = reg_df.copy()

    party_col = "Party Name"
    has_party = main_df[party_col].notna() & (main_df[party_col].astype(str).str.strip() != "")
    main_df["_Invoice_Group"] = has_party.cumsum()

    valid_df = main_df[main_df["_Invoice_Group"] > 0].copy()
    inv_headers = valid_df[has_party[valid_df.index]].copy()

    inv_headers["Bill No"] = (
        inv_headers["Type"].fillna("").str.strip()
        + "-"
        + inv_headers["Serial"].fillna("").str.strip()
    )

    key_serial = [
        normalize_bill_key(t, s)
        for t, s in zip(inv_headers["Type"].fillna(""), inv_headers["Serial"].fillna(""))
    ]

    if is_purchase and "Bill.No" in inv_headers.columns:
        key_billno = [
            normalize_bill_key(t, b)
            for t, b in zip(inv_headers["Type"].fillna(""), inv_headers["Bill.No"].fillna(""))
        ]
        is_paid = [(ks in paid_keys) or (kb in paid_keys) for ks, kb in zip(key_serial, key_billno)]
    else:
        is_paid = [(ks in paid_keys) for ks in key_serial]

    inv_headers["Payment Status"] = ["Paid" if p else "Outstanding" for p in is_paid]
    out_df = inv_headers[inv_headers["Payment Status"] == "Outstanding"].copy()

    date_col = "Bill.Date" if (is_purchase and "Bill.Date" in out_df.columns) else "Doc.Date"
    out_df["Invoice Date"] = pd.to_datetime(out_df[date_col], errors="coerce")
    out_df["Financial Year"] = out_df["Invoice Date"].apply(get_financial_year)
    out_df["Invoice Value"] = pd.to_numeric(out_df.get("Invoice Value", 0), errors="coerce").fillna(0.0)

    if (out_df["Invoice Value"] == 0).any() and "Item Value" in valid_df.columns:
        valid_df["_item_val_num"] = pd.to_numeric(valid_df["Item Value"], errors="coerce").fillna(0.0)
        grp_sums = valid_df.groupby("_Invoice_Group")["_item_val_num"].sum()
        zero_mask = out_df["Invoice Value"] == 0
        out_df.loc[zero_mask, "Invoice Value"] = out_df.loc[zero_mask, "_Invoice_Group"].map(grp_sums).fillna(0.0)

    as_on_ts = pd.Timestamp(as_on_date)
    out_df["Overdue Days"] = (as_on_ts - out_df["Invoice Date"]).dt.days.fillna(0).astype(int)
    out_df["Overdue Bucket"] = out_df["Overdue Days"].apply(get_overdue_bucket)

    out_df["Party Name (Branch)"] = out_df["Party Name"].astype(str).str.strip()
    out_df["Party Group"] = out_df["Party Name (Branch)"].apply(extract_parent_party)

    keep_cols = [
        "Party Group",
        "Party Name (Branch)",
        "Financial Year",
        "Type",
        "Serial",
        "Bill No",
    ]
    if is_purchase and "Bill.No" in out_df.columns:
        keep_cols.append("Bill.No")
    keep_cols += [
        "Invoice Date",
        "City",
        "State",
        "Invoice Value",
        "Overdue Days",
        "Overdue Bucket",
    ]
    keep_cols = [c for c in keep_cols if c in out_df.columns]
    out_df = out_df[keep_cols].sort_values(
        by=["Party Group", "Party Name (Branch)", "Overdue Days"],
        ascending=[True, True, False],
    ).reset_index(drop=True)

    main_clean = main_df.drop(columns=["_Invoice_Group"])
    return main_clean, out_df


# ---------- CLEAN EXCEL CHART HELPER ----------

def make_clean_bar_chart(
    chart_type: str,
    title: str,
    data_ref: Reference,
    cats_ref: Reference,
    bar_hex: str,
    width: float,
    height: float,
) -> BarChart:
    ch = BarChart()
    ch.type = chart_type
    ch.grouping = "clustered"
    ch.title = title
    ch.title.overlay = False
    ch.legend = None
    ch.plot_vis_only = False
    ch.varyColors = False

    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)

    if ch.series:
        ch.series[0].graphicalProperties.solidFill = bar_hex
        ch.series[0].graphicalProperties.line.solidFill = bar_hex

    dl = DataLabelList()
    dl.showVal = True
    dl.showSerName = False
    dl.showCatName = False
    dl.showPercent = False
    dl.showLegendKey = False
    ch.dataLabels = dl

    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.x_axis.tickLblPos = "low"
    ch.y_axis.majorGridlines = None
    ch.x_axis.majorGridlines = None

    ch.width = width
    ch.height = height
    return ch


# ---------- LEFT-PANEL MULTI-SELECT + SEARCH EXCEL DASHBOARD GENERATOR ----------

def write_raw_sheet(wb: Workbook, title: str, df: pd.DataFrame):
    ws = wb.create_sheet(title=title)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)

    for c_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=2):
        for c_idx, val in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=None if pd.isna(val) else val)

    ws.freeze_panes = "A2"


def write_outstanding_dashboard_sheet(
    wb: Workbook,
    title: str,
    df: pd.DataFrame,
    data_sheet_name: str,
    last_fy: str,
):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = False

    ds = wb.create_sheet(title=data_sheet_name)
    ds.sheet_state = "hidden"

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 27
    ws.column_dimensions["C"].width = 11
    ws.column_dimensions["D"].width = 3

    # Styles
    banner_fill = PatternFill("solid", fgColor="1F4E78")
    banner_font = Font(bold=True, size=15, color="FFFFFF")
    panel_fill = PatternFill("solid", fgColor="F2F5F9")
    filter_hdr_fill = PatternFill("solid", fgColor="1F4E78")
    filter_hdr_font = Font(bold=True, size=11, color="FFFFFF")
    filter_lbl_font = Font(bold=True, size=9, color="1F4E78")
    search_fill = PatternFill("solid", fgColor="FFFBEB")
    dropdown_fill = PatternFill("solid", fgColor="FFFFFF")
    dropdown_font = Font(bold=True, size=10, color="000000")
    mode_fill = PatternFill("solid", fgColor="E2E8F0")
    mode_font = Font(bold=True, size=9, color="1F4E78")
    kpi_label_font = Font(bold=True, size=9, color="595959")
    kpi_val_font = Font(bold=True, size=13, color="1F4E78")
    kpi_amber_font = Font(bold=True, size=13, color="B45309")
    kpi_red_font = Font(bold=True, size=13, color="C00000")
    kpi_fill = PatternFill("solid", fgColor="EDF2F8")
    kpi_amber_fill = PatternFill("solid", fgColor="FEF3C7")
    kpi_red_fill = PatternFill("solid", fgColor="FCE4D6")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF", size=10)
    amber_fill = PatternFill("solid", fgColor="FEF3C7")
    amber_font = Font(bold=True, color="B45309")
    red_fill = PatternFill("solid", fgColor="FCE4D6")
    red_font = Font(bold=True, color="C00000")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    box_border = Border(
        left=Side(style="medium", color="1F4E78"),
        right=Side(style="medium", color="1F4E78"),
        top=Side(style="medium", color="1F4E78"),
        bottom=Side(style="medium", color="1F4E78"),
    )
    search_border = Border(
        left=Side(style="medium", color="D97706"),
        right=Side(style="medium", color="D97706"),
        top=Side(style="medium", color="D97706"),
        bottom=Side(style="medium", color="D97706"),
    )
    kpi_sep_border = Border(
        left=Side(style="thick", color="FFFFFF"),
        right=Side(style="thick", color="FFFFFF"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    col_start = 5  # Column E
    last_tbl_idx = col_start + max(len(df.columns) - 1, 11)
    last_tbl_ltr = get_column_letter(last_tbl_idx)

    # Top Title Banner
    ws.merge_cells(f"E1:{last_tbl_ltr}2")
    ws["E1"] = f"  {title} Dashboard"
    ws["E1"].fill = banner_fill
    ws["E1"].font = banner_font
    ws["E1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 18

    # Left Multi-Select Filter Panel Background (A1:C38)
    for r in range(1, 39):
        for c in range(1, 4):
            ws.cell(row=r, column=c).fill = panel_fill

    ws.merge_cells("B2:C2")
    ws["B2"] = "SEARCH & MULTI-FILTERS"
    ws["B2"].fill = filter_hdr_fill
    ws["B2"].font = filter_hdr_font
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")

    # Quick Search Box (B4 label, B5:C5 input)
    ws.merge_cells("B4:C4")
    ws["B4"] = "🔍 Quick Search (Party / Branch)"
    ws["B4"].font = Font(bold=True, size=9, color="B45309")
    ws.merge_cells("B5:C5")
    ws["B5"] = None
    ws["B5"].fill = search_fill
    ws["B5"].font = dropdown_font
    ws["B5"].border = search_border
    ws["B5"].alignment = Alignment(horizontal="left", vertical="center")

    def setup_filter_block(header_row: int, label_text: str, slot_rows: list):
        ws[f"B{header_row}"] = label_text
        ws[f"B{header_row}"].font = filter_lbl_font
        ws[f"C{header_row}"] = "Include"
        ws[f"C{header_row}"].fill = mode_fill
        ws[f"C{header_row}"].font = mode_font
        ws[f"C{header_row}"].border = thin_border
        ws[f"C{header_row}"].alignment = Alignment(horizontal="center", vertical="center")

        for idx, r_num in enumerate(slot_rows):
            cell = ws[f"B{r_num}"]
            cell.value = "All" if idx == 0 else None
            cell.fill = dropdown_fill
            cell.font = dropdown_font
            cell.border = box_border
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # 4 Multi-Select Filter Blocks
    setup_filter_block(7, "1. Party / Company Group", [8, 9, 10, 11, 12])
    setup_filter_block(14, "2. Branch / Sub-Party", [15, 16, 17, 18, 19])
    setup_filter_block(21, "3. Overdue Days Bucket", [22, 23, 24, 25, 26])
    setup_filter_block(28, "4. Financial Year", [29, 30, 31])

    ws["B33"] = "How to use:"
    ws["B33"].font = Font(bold=True, size=8, color="1F4E78")
    ws["B34"] = "• Search box mein naam type karein."
    ws["B34"].font = Font(italic=True, size=8, color="595959")
    ws["B35"] = "• Multi-select ke liye alag boxes"
    ws["B35"].font = Font(italic=True, size=8, color="595959")
    ws["B36"] = "  mein options chunein."
    ws["B36"].font = Font(italic=True, size=8, color="595959")
    ws["B37"] = "• Deselect ke liye 'Exclude' chunein."
    ws["B37"].font = Font(italic=True, size=8, color="595959")

    if len(df) == 0:
        ws["E4"] = "Koi pending invoice nahi hai."
        return

    # ---- POPULATE BACKEND DATA SHEET (ds) ----
    ds["A1"] = "Match_Flag"
    ds["B1"] = 0
    for idx, col_name in enumerate(df.columns, start=3):
        ds.cell(row=1, column=idx, value=col_name)

    grp_ds_ltr = get_column_letter(3 + df.columns.get_loc("Party Group"))
    br_ds_ltr = get_column_letter(3 + df.columns.get_loc("Party Name (Branch)"))
    fy_ds_ltr = get_column_letter(3 + df.columns.get_loc("Financial Year"))
    amt_ds_ltr = get_column_letter(3 + df.columns.get_loc("Invoice Value"))
    days_ds_ltr = get_column_letter(3 + df.columns.get_loc("Overdue Days"))
    bkt_ds_ltr = get_column_letter(3 + df.columns.get_loc("Overdue Bucket"))

    sh_ref = f"'{title}'"
    ds_ref = f"'{data_sheet_name}'"

    n_rows = len(df)
    r_start = 2
    r_end = n_rows + 1

    for i, row in enumerate(df.itertuples(index=False, name=None), start=2):
        cond_search = (
            f'IF(TRIM({sh_ref}!$B$5)="",TRUE,'
            f'OR(ISNUMBER(SEARCH({sh_ref}!$B$5,{grp_ds_ltr}{i})),'
            f'ISNUMBER(SEARCH({sh_ref}!$B$5,{br_ds_ltr}{i}))))'
        )
        cond_grp = (
            f'IF(COUNTIFS({sh_ref}!$B$8:$B$12,"<>",{sh_ref}!$B$8:$B$12,"<>All")=0,TRUE,'
            f'IF({sh_ref}!$C$7="Exclude",'
            f'COUNTIF({sh_ref}!$B$8:$B$12,{grp_ds_ltr}{i})=0,'
            f'COUNTIF({sh_ref}!$B$8:$B$12,{grp_ds_ltr}{i})>0))'
        )
        cond_br = (
            f'IF(COUNTIFS({sh_ref}!$B$15:$B$19,"<>",{sh_ref}!$B$15:$B$19,"<>All")=0,TRUE,'
            f'IF({sh_ref}!$C$14="Exclude",'
            f'COUNTIF({sh_ref}!$B$15:$B$19,{br_ds_ltr}{i})=0,'
            f'COUNTIF({sh_ref}!$B$15:$B$19,{br_ds_ltr}{i})>0))'
        )
        cond_bkt = (
            f'IF(COUNTIFS({sh_ref}!$B$22:$B$26,"<>",{sh_ref}!$B$22:$B$26,"<>All")=0,TRUE,'
            f'IF({sh_ref}!$C$21="Exclude",'
            f'COUNTIF({sh_ref}!$B$22:$B$26,{bkt_ds_ltr}{i})=0,'
            f'COUNTIF({sh_ref}!$B$22:$B$26,{bkt_ds_ltr}{i})>0))'
        )
        cond_fy = (
            f'IF(COUNTIFS({sh_ref}!$B$29:$B$31,"<>",{sh_ref}!$B$29:$B$31,"<>All")=0,TRUE,'
            f'IF({sh_ref}!$C$28="Exclude",'
            f'COUNTIF({sh_ref}!$B$29:$B$31,{fy_ds_ltr}{i})=0,'
            f'COUNTIF({sh_ref}!$B$29:$B$31,{fy_ds_ltr}{i})>0))'
        )
        ds.cell(
            row=i,
            column=1,
            value=f"=IF(AND({cond_search},{cond_grp},{cond_br},{cond_bkt},{cond_fy}),1,0)",
        )
        ds.cell(row=i, column=2, value=f"=B{i-1}+A{i}")

        for offset, val in enumerate(row):
            c = ds.cell(row=i, column=3 + offset)
            if isinstance(val, pd.Timestamp):
                c.value = val.to_pydatetime() if pd.notna(val) else None
            else:
                c.value = None if pd.isna(val) else val

    flag_rng = f"{ds_ref}!$A${r_start}:$A${r_end}"
    cnt_rng = f"{ds_ref}!$B${r_start}:$B${r_end}"
    grp_rng = f"{ds_ref}!${grp_ds_ltr}${r_start}:${grp_ds_ltr}${r_end}"
    br_rng = f"{ds_ref}!${br_ds_ltr}${r_start}:${br_ds_ltr}${r_end}"
    fy_rng = f"{ds_ref}!${fy_ds_ltr}${r_start}:${fy_ds_ltr}${r_end}"
    amt_rng = f"{ds_ref}!${amt_ds_ltr}${r_start}:${amt_ds_ltr}${r_end}"
    day_rng = f"{ds_ref}!${days_ds_ltr}${r_start}:${days_ds_ltr}${r_end}"
    bkt_rng = f"{ds_ref}!${bkt_ds_ltr}${r_start}:${bkt_ds_ltr}${r_end}"

    # ---- 5 EVENLY ALIGNED KPI CARDS ----
    kpi_specs = [
        ("E", "E", "Total Outstanding (₹)",
         f"=SUMIF({flag_rng},1,{amt_rng})",
         "₹#,##0.00", kpi_fill, kpi_val_font),
        ("F", "F", "46-90 Days Overdue (₹)",
         f'=SUMIFS({amt_rng},{flag_rng},1,{bkt_rng},"46-90 Days")',
         "₹#,##0.00", kpi_amber_fill, kpi_amber_font),
        ("G", "I", "90+ Days Overdue (₹)",
         f'=SUMIFS({amt_rng},{flag_rng},1,{bkt_rng},"90+ Days")',
         "₹#,##0.00", kpi_red_fill, kpi_red_font),
        ("J", "L", "Avg Overdue Days",
         f"=IFERROR(AVERAGEIF({flag_rng},1,{day_rng}),0)",
         '0" Days"', kpi_fill, kpi_val_font),
        ("M", last_tbl_ltr, "Pending Invoices",
         f"=COUNTIF({flag_rng},1)",
         "#,##0", kpi_fill, kpi_val_font),
    ]
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 26
    for c1, c2, lbl, formula, num_fmt, fill_st, val_font in kpi_specs:
        if c1 != c2:
            ws.merge_cells(f"{c1}4:{c2}4")
            ws.merge_cells(f"{c1}5:{c2}5")
        ws[f"{c1}4"] = lbl
        ws[f"{c1}4"].font = kpi_label_font
        ws[f"{c1}4"].fill = fill_st
        ws[f"{c1}4"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c1}4"].border = kpi_sep_border

        ws[f"{c1}5"] = formula
        ws[f"{c1}5"].font = val_font
        ws[f"{c1}5"].fill = fill_st
        ws[f"{c1}5"].number_format = num_fmt
        ws[f"{c1}5"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c1}5"].border = kpi_sep_border

    # ---- DROPDOWN LISTS ON ds (errorStyle="information" allows typing/searching freely) ----
    groups_list = ["All"] + sorted(df["Party Group"].dropna().unique().tolist())
    for idx, val in enumerate(groups_list, start=1):
        ds[f"T{idx}"] = val
    dv_grp = DataValidation(
        type="list", formula1=f"={ds_ref}!$T$1:$T${len(groups_list)}", allow_blank=True, errorStyle="information"
    )
    ws.add_data_validation(dv_grp)
    dv_grp.add("B8:B12")

    branches_list = ["All"] + sorted(df["Party Name (Branch)"].dropna().unique().tolist())
    for idx, val in enumerate(branches_list, start=1):
        ds[f"U{idx}"] = val
    dv_br = DataValidation(
        type="list", formula1=f"={ds_ref}!$U$1:$U${len(branches_list)}", allow_blank=True, errorStyle="information"
    )
    ws.add_data_validation(dv_br)
    dv_br.add("B15:B19")

    buckets_dd = ["All", "0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
    for idx, val in enumerate(buckets_dd, start=1):
        ds[f"V{idx}"] = val
    dv_bkt = DataValidation(
        type="list", formula1=f"={ds_ref}!$V$1:$V${len(buckets_dd)}", allow_blank=True, errorStyle="information"
    )
    ws.add_data_validation(dv_bkt)
    dv_bkt.add("B22:B26")

    fy_list = ["All"] + sorted(df["Financial Year"].dropna().unique().tolist())
    for idx, val in enumerate(fy_list, start=1):
        ds[f"W{idx}"] = val
    dv_fy = DataValidation(
        type="list", formula1=f"={ds_ref}!$W$1:$W${len(fy_list)}", allow_blank=True, errorStyle="information"
    )
    ws.add_data_validation(dv_fy)
    dv_fy.add("B29:B31")

    ds["X1"] = "Include"
    ds["X2"] = "Exclude"
    dv_mode = DataValidation(type="list", formula1=f"={ds_ref}!$X$1:$X$2", allow_blank=False)
    ws.add_data_validation(dv_mode)
    for cell_ref in ["C7", "C14", "C21", "C28"]:
        dv_mode.add(cell_ref)

    # ---- LIVE CHART SUMMARIES ON ds ----
    ds["Y1"] = "Overdue Bucket"
    ds["Z1"] = "Outstanding"
    for i, b_name in enumerate(["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"], start=2):
        ds[f"Y{i}"] = b_name
        ds[f"Z{i}"] = f"=SUMIFS({amt_rng},{flag_rng},1,{bkt_rng},Y{i})"
        ds[f"Z{i}"].number_format = "#,##0"

    ds["AA1"] = "Party Group"
    ds["AB1"] = "Outstanding"
    top10_groups = (
        df.groupby("Party Group")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, grp_name in enumerate(top10_groups, start=2):
        ds[f"AA{i}"] = grp_name
        ds[f"AB{i}"] = f"=SUMIFS({amt_rng},{flag_rng},1,{grp_rng},AA{i})"
        ds[f"AB{i}"].number_format = "#,##0"

    ds["AC1"] = "Branch"
    ds["AD1"] = "Outstanding"
    top10_branches = (
        df.groupby("Party Name (Branch)")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, br_name in enumerate(top10_branches, start=2):
        ds[f"AC{i}"] = br_name[:28] + ("..." if len(br_name) > 28 else "")
        ds[f"AE{i}"] = br_name
        ds[f"AD{i}"] = f"=SUMIFS({amt_rng},{flag_rng},1,{br_rng},AE{i})"
        ds[f"AD{i}"].number_format = "#,##0"

    ds["AF1"] = "Party Group"
    ds["AG1"] = f"FY {last_fy} Pending"
    last_fy_df = df[df["Financial Year"] == last_fy]
    top10_last_fy = (
        last_fy_df.groupby("Party Group")["Invoice Value"]
        .sum()
        .sort_values(ascending=False)
        .head(10)
        .index.tolist()
    )
    for i, grp_name in enumerate(top10_last_fy, start=2):
        ds[f"AF{i}"] = grp_name
        ds[f"AG{i}"] = f'=SUMIFS({amt_rng},{flag_rng},1,{fy_rng},"{last_fy}",{grp_rng},AF{i})'
        ds[f"AG{i}"].number_format = "#,##0"

    # ---- 4 CLEAN, ALIGNED CHARTS IN 2x2 GRID ----
    ch1 = make_clean_bar_chart(
        chart_type="col",
        title="Overdue Bucket-wise Outstanding (₹)",
        data_ref=Reference(ds, min_col=26, min_row=1, max_row=5),
        cats_ref=Reference(ds, min_col=25, min_row=2, max_row=5),
        bar_hex="1F4E78",
        width=14.0,
        height=8.8,
    )
    ws.add_chart(ch1, "E7")

    if top10_groups:
        max_r2 = 1 + len(top10_groups)
        ch2 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Parties by Outstanding (₹)",
            data_ref=Reference(ds, min_col=28, min_row=1, max_row=max_r2),
            cats_ref=Reference(ds, min_col=27, min_row=2, max_row=max_r2),
            bar_hex="0F766E",
            width=15.5,
            height=8.8,
        )
        ws.add_chart(ch2, "I7")

    if top10_branches:
        max_r3 = 1 + len(top10_branches)
        ch3 = make_clean_bar_chart(
            chart_type="bar",
            title="Top 10 Pending Branches (₹)",
            data_ref=Reference(ds, min_col=30, min_row=1, max_row=max_r3),
            cats_ref=Reference(ds, min_col=29, min_row=2, max_row=max_r3),
            bar_hex="4338CA",
            width=14.0,
            height=8.8,
        )
        ws.add_chart(ch3, "E25")

    if top10_last_fy:
        max_r4 = 1 + len(top10_last_fy)
        ch4 = make_clean_bar_chart(
            chart_type="bar",
            title=f"Last FY ({last_fy}) Pending by Party (₹)",
            data_ref=Reference(ds, min_col=33, min_row=1, max_row=max_r4),
            cats_ref=Reference(ds, min_col=32, min_row=2, max_row=max_r4),
            bar_hex="B91C1C",
            width=15.5,
            height=8.8,
        )
        ws.add_chart(ch4, "I25")

    # ---- DYNAMIC FILTERED INVOICES TABLE (Row 43 onwards) ----
    start_row = 43

    for offset, col_name in enumerate(df.columns):
        c_idx = col_start + offset
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    amt_col_offset = df.columns.get_loc("Invoice Value")
    date_col_offset = df.columns.get_loc("Invoice Date")
    bkt_tbl_ltr = get_column_letter(col_start + df.columns.get_loc("Overdue Bucket"))
    first_tbl_ltr = get_column_letter(col_start)

    for r_idx in range(1, n_rows + 1):
        excel_row = start_row + r_idx
        for offset, col_name in enumerate(df.columns):
            ds_col_ltr = get_column_letter(3 + offset)
            c_idx = col_start + offset
            cell = ws.cell(
                row=excel_row,
                column=c_idx,
                value=f'=IFERROR(INDEX({ds_ref}!${ds_col_ltr}${r_start}:${ds_col_ltr}${r_end},MATCH({r_idx},{cnt_rng},0)),"")',
            )
            if offset == amt_col_offset:
                cell.number_format = "#,##0.00"
            elif offset == date_col_offset:
                cell.number_format = "dd-mmm-yyyy"
            cell.border = thin_border

    tbl_data_rng = f"{first_tbl_ltr}{start_row + 1}:{last_tbl_ltr}{start_row + n_rows}"
    ws.conditional_formatting.add(
        tbl_data_rng,
        FormulaRule(formula=[f'${bkt_tbl_ltr}{start_row + 1}="90+ Days"'], fill=red_fill, font=red_font),
    )
    ws.conditional_formatting.add(
        tbl_data_rng,
        FormulaRule(formula=[f'${bkt_tbl_ltr}{start_row + 1}="46-90 Days"'], fill=amber_fill, font=amber_font),
    )

    for offset, col_name in enumerate(df.columns):
        c_idx = col_start + offset
        min_w = 24 if offset == 0 else 12
        max_len = max(len(str(col_name)), min_w)
        sample_max = df[col_name].astype(str).str.len().max()
        max_len = min(max(max_len, int(sample_max) + 3), 36)
        ws.column_dimensions[get_column_letter(c_idx)].width = max_len


def build_excel_workbook(
    sales_main: pd.DataFrame,
    purchase_main: pd.DataFrame,
    sales_out: pd.DataFrame,
    purchase_out: pd.DataFrame,
    last_fy: str,
) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)

    write_raw_sheet(wb, "Sales Main Sheet", sales_main)
    write_raw_sheet(wb, "Purchase Main Sheet", purchase_main)
    write_outstanding_dashboard_sheet(wb, "Sales Outstanding", sales_out, "_Sales_Data", last_fy)
    write_outstanding_dashboard_sheet(wb, "Purchase Outstanding", purchase_out, "_Purchase_Data", last_fy)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf.getvalue()


# ---------- INTERACTIVE STREAMLIT DASHBOARD RENDERER (SEARCH + MULTI-SELECT) ----------

def render_outstanding_tab(df: pd.DataFrame, key_prefix: str, label: str, last_fy: str):
    if df.empty:
        st.success(f"Koi {label} pending nahi hai!")
        return

    slicer_col, main_col = st.columns([1.2, 3.8], gap="large")

    with slicer_col:
        st.markdown("### 🎛️ Search & Multi-Select")

        search_txt = st.text_input(
            "🔍 Quick Search (Party / Branch / Bill No)",
            value="",
            key=f"{key_prefix}_search",
            placeholder="e.g. Cipla, Goa, ISS-0145...",
        ).strip()

        all_groups = sorted(df["Party Group"].dropna().unique().tolist())
        selected_groups = st.multiselect(
            "1. Party / Company Group (Khali = All)",
            options=all_groups,
            default=[],
            key=f"{key_prefix}_groups",
            help="Type karke search karein ya dropdown se multiple select karein.",
        )

        group_df = df if not selected_groups else df[df["Party Group"].isin(selected_groups)]

        available_branches = sorted(group_df["Party Name (Branch)"].dropna().unique().tolist())
        default_branches = available_branches if selected_groups else []
        selected_branches = st.multiselect(
            "2. Branches / Sub-Parties (Khali = All)",
            options=available_branches,
            default=default_branches,
            key=f"{key_prefix}_branches_{'_'.join(selected_groups)}",
            help="Company select karte hi uski branches aa jayengi; kisi ko bhi deselect karein.",
        )

        bucket_options = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
        selected_buckets = st.multiselect(
            "3. Overdue Days Bucket (Khali = All)",
            options=bucket_options,
            default=[],
            key=f"{key_prefix}_buckets",
        )

        all_fys = sorted(df["Financial Year"].dropna().unique().tolist())
        selected_fys = st.multiselect(
            "4. Financial Year (Khali = All)",
            options=all_fys,
            default=[],
            key=f"{key_prefix}_fys",
        )

    filtered = group_df.copy()
    if search_txt:
        mask = (
            filtered["Party Group"].astype(str).str.contains(search_txt, case=False, na=False)
            | filtered["Party Name (Branch)"].astype(str).str.contains(search_txt, case=False, na=False)
            | filtered["Bill No"].astype(str).str.contains(search_txt, case=False, na=False)
        )
        filtered = filtered[mask]
    if selected_branches:
        filtered = filtered[filtered["Party Name (Branch)"].isin(selected_branches)]
    if selected_buckets:
        filtered = filtered[filtered["Overdue Bucket"].isin(selected_buckets)]
    if selected_fys:
        filtered = filtered[filtered["Financial Year"].isin(selected_fys)]

    with main_col:
        total_amt = filtered["Invoice Value"].sum()
        total_bills = len(filtered)
        amt_46_90 = filtered.loc[filtered["Overdue Bucket"] == "46-90 Days", "Invoice Value"].sum()
        amt_90_plus = filtered.loc[filtered["Overdue Bucket"] == "90+ Days", "Invoice Value"].sum()
        avg_days = int(filtered["Overdue Days"].mean()) if total_bills > 0 else 0

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total Outstanding", f"₹{total_amt:,.2f}")
        k2.metric("46-90 Days Overdue", f"₹{amt_46_90:,.2f}")
        k3.metric("90+ Days Overdue", f"₹{amt_90_plus:,.2f}")
        k4.metric("Avg Overdue Days", f"{avg_days} Days")
        k5.metric("Pending Invoices", f"{total_bills:,}")

        c1, c2 = st.columns(2)
        with c1:
            bucket_order = ["0-30 Days", "31-45 Days", "46-90 Days", "90+ Days"]
            b_sum = filtered.groupby("Overdue Bucket", as_index=False)["Invoice Value"].sum()
            b_sum["Overdue Bucket"] = pd.Categorical(
                b_sum["Overdue Bucket"], categories=bucket_order, ordered=True
            )
            b_sum = b_sum.sort_values("Overdue Bucket")
            fig_b = px.bar(
                b_sum,
                x="Overdue Bucket",
                y="Invoice Value",
                title="Overdue Bucket-wise Outstanding (₹)",
                text_auto=".2s",
                color="Overdue Bucket",
                color_discrete_map={
                    "0-30 Days": "#2E7D32",
                    "31-45 Days": "#F9A825",
                    "46-90 Days": "#EF6C00",
                    "90+ Days": "#C62828",
                },
            )
            fig_b.update_layout(showlegend=False, height=320)
            st.plotly_chart(fig_b, use_container_width=True)

        with c2:
            top_grp = (
                filtered.groupby("Party Group", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            fig_g = px.bar(
                top_grp,
                x="Invoice Value",
                y="Party Group",
                orientation="h",
                title="Top 10 Parties by Outstanding (₹)",
                text_auto=".2s",
                color_discrete_sequence=["#0F766E"],
            )
            fig_g.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
            st.plotly_chart(fig_g, use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            top_br = (
                filtered.groupby("Party Name (Branch)", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            fig_p = px.bar(
                top_br,
                x="Invoice Value",
                y="Party Name (Branch)",
                orientation="h",
                title="Top 10 Pending Branches (₹)",
                text_auto=".2s",
                color_discrete_sequence=["#4338CA"],
            )
            fig_p.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
            st.plotly_chart(fig_p, use_container_width=True)

        with c4:
            last_fy_filtered = filtered[filtered["Financial Year"] == last_fy]
            top_last_fy = (
                last_fy_filtered.groupby("Party Group", as_index=False)["Invoice Value"]
                .sum()
                .sort_values("Invoice Value", ascending=False)
                .head(10)
            )
            if not top_last_fy.empty:
                fig_lfy = px.bar(
                    top_last_fy,
                    x="Invoice Value",
                    y="Party Group",
                    orientation="h",
                    title=f"Last FY ({last_fy}) Pending by Party (₹)",
                    text_auto=".2s",
                    color_discrete_sequence=["#B91C1C"],
                )
                fig_lfy.update_layout(yaxis={"categoryorder": "total ascending"}, height=320)
                st.plotly_chart(fig_lfy, use_container_width=True)
            else:
                st.info(f"Selected filters mein Last FY ({last_fy}) ka koi pending invoice nahi hai.")

        def highlight_overdue(row):
            if row["Overdue Bucket"] == "90+ Days":
                return ["background-color: #FCE4D6; color: #C00000; font-weight: bold"] * len(row)
            if row["Overdue Bucket"] == "46-90 Days":
                return ["background-color: #FEF3C7; color: #B45309; font-weight: bold"] * len(row)
            return [""] * len(row)

        display_df = filtered.copy()
        display_df["Invoice Date"] = display_df["Invoice Date"].dt.strftime("%d-%b-%Y")

        st.dataframe(
            display_df.style.apply(highlight_overdue, axis=1).format({"Invoice Value": "₹{:,.2f}"}),
            use_container_width=True,
            height=420,
        )


# ---------- MAIN APP UI (WITH MASTER BACKUP + INCREMENTAL UPDATE) ----------

st.title("📊 Sales & Purchase Outstanding Dashboard")

with st.sidebar:
    st.header("1. Previous Master Backup (Optional)")
    backup_file = st.file_uploader(
        "Master Data Backup (.xlsx)",
        type=["xlsx"],
        help="Agar aapke paas pichli baar ka Master Backup hai toh yahan upload karein, fir neeche sirf naya data daalein.",
    )

    st.divider()
    st.header("2. Upload New / Full Excel Files")
    ml_file = st.file_uploader("Multi Ledger Sheet (.xlsx)", type=["xlsx", "xls"])
    pur_file = st.file_uploader("Purchase Record Sheet (.xlsx)", type=["xlsx", "xls"])
    sal_file = st.file_uploader("Sales Record Sheet (.xlsx)", type=["xlsx", "xls"])

    st.divider()
    as_on_date = st.date_input("As-On Date (Overdue calculate karne ke liye)", value=date.today())

# Validation: Either Master Backup OR all 3 initial files required
if not backup_file and not (ml_file and pur_file and sal_file):
    st.info(
        "Pehli baar teeno files (**Multi Ledger, Purchase, Sales**) upload karein. "
        "Agli baar sirf **Master Data Backup (.xlsx)** aur naye mahine ki files upload karke data update kar sakte hain."
    )
    st.stop()

last_fy = get_last_financial_year(as_on_date)

with st.spinner("Data merge ho raha hai aur Excel Dashboard + Master Backup generate ho rahe hain..."):
    old_sales, old_purchase, old_paid = (pd.DataFrame(), pd.DataFrame(), set())
    if backup_file:
        old_sales, old_purchase, old_paid = read_master_backup(backup_file.getvalue())

    new_paid = extract_paid_bills_from_ledger(read_register(ml_file.getvalue())) if ml_file else set()
    new_pur = read_register(pur_file.getvalue()) if pur_file else pd.DataFrame()
    new_sal = read_register(sal_file.getvalue()) if sal_file else pd.DataFrame()

    combined_paid = old_paid | new_paid
    combined_sales = merge_registers(old_sales, new_sal)
    combined_purchase = merge_registers(old_purchase, new_pur)

    sales_main, sales_out = process_register_outstanding(combined_sales, combined_paid, as_on_date, is_purchase=False)
    purchase_main, purchase_out = process_register_outstanding(combined_purchase, combined_paid, as_on_date, is_purchase=True)

    excel_bytes = build_excel_workbook(sales_main, purchase_main, sales_out, purchase_out, last_fy)
    backup_bytes = build_master_backup_bytes(sales_main, purchase_main, combined_paid)

# Two Download Buttons Side-by-Side
d1, d2 = st.columns(2)
with d1:
    st.download_button(
        label="⬇️ Download 4-Sheet Outstanding Dashboard (.xlsx)",
        data=excel_bytes,
        file_name=f"Outstanding_Dashboard_{as_on_date.strftime('%d-%b-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
with d2:
    st.download_button(
        label="💾 Download Master Data Backup (.xlsx) — Agli Baar Update Ke Liye",
        data=backup_bytes,
        file_name=f"Master_Data_Backup_{as_on_date.strftime('%d-%b-%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

tab1, tab2 = st.tabs([
    f"📈 Sales Outstanding ({len(sales_out)} Bills)",
    f"📉 Purchase Outstanding ({len(purchase_out)} Bills)",
])

with tab1:
    render_outstanding_tab(sales_out, "sales", "Sales Outstanding", last_fy)

with tab2:
    render_outstanding_tab(purchase_out, "purchase", "Purchase Outstanding", last_fy)
