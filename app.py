import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Meesho Reconciliation Tool", page_icon="🧾", layout="wide")

# ─────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-title { font-size: 2rem; font-weight: 700; color: #6C3FC5; margin-bottom: 0; }
    .sub-title  { font-size: 1rem; color: #888; margin-bottom: 1.5rem; }
    .section-header { font-size: 1.1rem; font-weight: 600; color: #444; border-left: 4px solid #6C3FC5;
                      padding-left: 10px; margin: 1.2rem 0 0.6rem 0; }
    .stat-box { background: #f8f5ff; border-radius: 10px; padding: 14px 18px;
                border: 1px solid #e0d4f7; text-align: center; }
    .stat-num  { font-size: 1.6rem; font-weight: 700; color: #6C3FC5; }
    .stat-label{ font-size: 0.78rem; color: #777; margin-top: 2px; }
    .profit    { color: #1a8c4e; font-weight: 600; }
    .loss      { color: #d93025; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">🧾 Meesho Reconciliation Tool</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated price reconciliation for YG · PE · AG accounts</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPER: detect account from filename
# ─────────────────────────────────────────────
ACCOUNT_MAP = {
    "_YG": ("Yash Gallery", "YG"),
    "_PE": ("Pushpa",       "PE"),
    "_AG": ("Aashirwad",    "AG"),
}

def detect_account(filename):
    fn = filename.upper()
    for key, val in ACCOUNT_MAP.items():
        if key in fn:
            return val
    return ("Unknown", "UNK")

# ─────────────────────────────────────────────
# HELPER: load reference data
# ─────────────────────────────────────────────
@st.cache_data
def load_replace_sku(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    dfs = {}
    sheet_account = {
        "Meesho YG":     ("YG",  "SELLER SKU",  "OMS SKU"),
        "Meesho Pushpa": ("PE",  "MESSHO SKU",  "OMSSKU"),
        "Messho Ag":     ("AG",  "MESSHO SKU",  "OMSSKU"),
    }
    for sheet, (acct, col_meesho, col_oms) in sheet_account.items():
        if sheet in xl.sheet_names:
            df = xl.parse(sheet)
            df = df[[col_meesho, col_oms]].dropna(subset=[col_meesho, col_oms])
            df.columns = ["MEESHO_SKU", "OMS_SKU"]
            df["MEESHO_SKU"] = df["MEESHO_SKU"].astype(str).str.strip()
            df["OMS_SKU"]    = df["OMS_SKU"].astype(str).str.strip()
            dfs[acct] = df.set_index("MEESHO_SKU")["OMS_SKU"].to_dict()
    return dfs

@st.cache_data
def load_pwn(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), header=2)   # row 3 is real header
    df.columns = ["Parent_SKU", "OMS_Child_SKU", "PWN_10"]
    df = df.dropna(subset=["OMS_Child_SKU", "PWN_10"])
    df["OMS_Child_SKU"] = df["OMS_Child_SKU"].astype(str).str.strip()
    df["PWN_10"] = pd.to_numeric(df["PWN_10"], errors="coerce")
    return df.set_index("OMS_Child_SKU")["PWN_10"].to_dict()

@st.cache_data
def load_closed_sku(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    closed = {}

    # Sheet1: SKU | PWN
    if "Sheet1" in xl.sheet_names:
        df1 = xl.parse("Sheet1", usecols=[0, 1])
        df1.columns = ["SKU", "Price"]
        df1 = df1.dropna(subset=["SKU", "Price"])
        df1["SKU"]   = df1["SKU"].astype(str).str.strip()
        df1["Price"] = pd.to_numeric(df1["Price"], errors="coerce")
        for _, row in df1.iterrows():
            sku = row["SKU"]
            if sku not in closed:
                closed[sku] = []
            closed[sku].append(row["Price"])

    # Sheet2: SKU | Price
    if "Sheet2" in xl.sheet_names:
        df2 = xl.parse("Sheet2", usecols=[0, 1])
        df2.columns = ["SKU", "Price"]
        df2 = df2.dropna(subset=["SKU", "Price"])
        df2["SKU"]   = df2["SKU"].astype(str).str.strip()
        df2["Price"] = pd.to_numeric(df2["Price"], errors="coerce")
        for _, row in df2.iterrows():
            sku = row["SKU"]
            if sku not in closed:
                closed[sku] = []
            closed[sku].append(row["Price"])

    # Keep the minimum price found across both sheets
    return {sku: min(prices) for sku, prices in closed.items() if prices}

# ─────────────────────────────────────────────
# HELPER: process one CSV file
# ─────────────────────────────────────────────
def process_csv(df_raw, account_code, company_name, sku_map, pwn_map, closed_map):
    records = []

    for _, row in df_raw.iterrows():
        reason       = row.get("Reason for Credit Entry", "")
        sub_order    = row.get("Sub Order No", "")
        order_date   = row.get("Order Date", "")
        state        = row.get("Customer State", "")
        product_name = row.get("Product Name", "")
        meesho_sku   = str(row.get("SKU", "")).strip()
        size         = row.get("Size", "")
        qty          = pd.to_numeric(row.get("Quantity", 1), errors="coerce") or 1
        listed_price = pd.to_numeric(row.get("Supplier Listed Price (Incl. GST + Commission)", 0), errors="coerce") or 0
        disc_price   = pd.to_numeric(row.get("Supplier Discounted Price (Incl GST and Commision)", 0), errors="coerce") or 0
        packet_id    = row.get("Packet Id", "")

        # Step 1: Build OMS SKU = SKU-Size format for lookup
        meesho_sku_size = f"{meesho_sku}-{size}".strip("-")

        # Step 2: Replace SKU using account-specific map
        account_map = sku_map.get(account_code, {})
        oms_sku = account_map.get(meesho_sku_size, account_map.get(meesho_sku, meesho_sku_size))

        # Step 3: Check if OMS SKU is in Closed SKU list
        sku_status   = "On Going"
        pwn_10_price = None

        if oms_sku in closed_map:
            sku_status   = "Closed"
            pwn_10_price = closed_map[oms_sku]
        else:
            # Step 4: Lookup PWN+10% from PWN file
            pwn_val = pwn_map.get(oms_sku)
            if pwn_val is not None:
                pwn_10_price = pwn_val
            else:
                pwn_10_price = "SKU Not Found"

        # Step 5: Calculate amounts
        disc_price_qty = disc_price * qty

        if isinstance(pwn_10_price, (int, float)):
            pwn_10_qty    = pwn_10_price * qty
            difference    = disc_price_qty - pwn_10_qty
        else:
            pwn_10_qty    = "SKU Not Found"
            difference    = "SKU Not Found"

        records.append({
            "Company":                          company_name,
            "Reason for Credit Entry":          reason,
            "Sub Order No":                     sub_order,
            "Order Date":                       order_date,
            "Customer State":                   state,
            "Product Name":                     product_name,
            "SKU":                              meesho_sku,
            "OMS SKU":                          oms_sku,
            "Size":                             size,
            "Quantity":                         qty,
            "Supplier Listed Price":            listed_price,
            "Supplier Discounted Price":        disc_price,
            "Supplier Discounted Price * Qty":  disc_price_qty,
            "Update PWN+10%":                   pwn_10_price,
            "Update PWN+10% * Qty":             pwn_10_qty,
            "Difference Amount":                difference,
            "SKU Status":                       sku_status,
            "Packet Id":                        packet_id,
        })

    return pd.DataFrame(records)

# ─────────────────────────────────────────────
# HELPER: export to styled Excel
# ─────────────────────────────────────────────
def export_excel(sheets_dict):
    wb = Workbook()
    wb.remove(wb.active)

    # Colours
    HEADER_FILL   = PatternFill("solid", fgColor="6C3FC5")
    PROFIT_FILL   = PatternFill("solid", fgColor="C6EFCE")
    LOSS_FILL     = PatternFill("solid", fgColor="FFC7CE")
    NOTFOUND_FILL = PatternFill("solid", fgColor="FFEB9C")
    CLOSED_FILL   = PatternFill("solid", fgColor="BDD7EE")
    ALT_FILL      = PatternFill("solid", fgColor="F3EFFF")
    WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")

    HEADER_FONT   = Font(bold=True, color="FFFFFF", size=10)
    BOLD_FONT     = Font(bold=True, size=10)
    NORMAL_FONT   = Font(size=10)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name[:31])

        cols = list(df.columns)
        # Write header
        for c_idx, col in enumerate(cols, 1):
            cell = ws.cell(1, c_idx, col)
            cell.fill   = HEADER_FILL
            cell.font   = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[1].height = 35

        diff_col = cols.index("Difference Amount") + 1 if "Difference Amount" in cols else None
        status_col = cols.index("SKU Status") + 1 if "SKU Status" in cols else None

        for r_idx, (_, row) in enumerate(df.iterrows(), 2):
            base_fill = ALT_FILL if r_idx % 2 == 0 else WHITE_FILL
            diff_val  = row.get("Difference Amount", None)
            status_val = row.get("SKU Status", "")

            for c_idx, col in enumerate(cols, 1):
                val  = row[col]
                cell = ws.cell(r_idx, c_idx, val if val != "SKU Not Found" or c_idx == diff_col else val)

                # Row fill logic
                if val == "SKU Not Found":
                    cell.fill = NOTFOUND_FILL
                elif status_val == "Closed":
                    cell.fill = CLOSED_FILL
                elif c_idx == diff_col and isinstance(diff_val, (int, float)):
                    cell.fill = PROFIT_FILL if diff_val >= 0 else LOSS_FILL
                else:
                    cell.fill = base_fill

                cell.font      = NORMAL_FONT
                cell.border    = border
                cell.alignment = Alignment(vertical="center")

                # Format numbers
                if col in ("Supplier Discounted Price * Qty", "Update PWN+10% * Qty",
                           "Difference Amount", "Supplier Listed Price",
                           "Supplier Discounted Price", "Update PWN+10%") and isinstance(val, (int, float)):
                    cell.number_format = "#,##0.00"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        # Auto column widths
        for c_idx, col in enumerate(cols, 1):
            max_len = len(str(col))
            for r_idx in range(2, ws.max_row + 1):
                v = ws.cell(r_idx, c_idx).value
                if v:
                    max_len = max(max_len, len(str(v)))
            ws.column_dimensions[get_column_letter(c_idx)].width = min(max_len + 3, 40)

        ws.freeze_panes = "A2"

    # Summary sheet
    ws_sum = wb.create_sheet(title="Summary", index=0)
    sum_headers = ["Account", "Total Orders", "Profit Orders", "Loss Orders",
                   "SKU Not Found", "Closed SKUs", "Total Difference (₹)"]
    for c, h in enumerate(sum_headers, 1):
        cell = ws_sum.cell(1, c, h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
        cell.border = border

    for r, (sheet_name, df) in enumerate(sheets_dict.items(), 2):
        diff_series = pd.to_numeric(df["Difference Amount"], errors="coerce")
        not_found   = (df["Difference Amount"] == "SKU Not Found").sum()
        closed_cnt  = (df["SKU Status"] == "Closed").sum()
        profit_cnt  = (diff_series >= 0).sum()
        loss_cnt    = (diff_series < 0).sum()
        total_diff  = diff_series.sum()

        row_data = [sheet_name, len(df), profit_cnt, loss_cnt, not_found, closed_cnt, round(total_diff, 2)]
        for c, val in enumerate(row_data, 1):
            cell = ws_sum.cell(r, c, val)
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if c == 7 and isinstance(val, float):
                cell.number_format = "#,##0.00"
                cell.font = Font(bold=True, color="1a8c4e" if val >= 0 else "d93025")

    for c in range(1, len(sum_headers) + 1):
        ws_sum.column_dimensions[get_column_letter(c)].width = 22
    ws_sum.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# SIDEBAR — Upload Reference Files
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 Reference Files")
    st.caption("Upload once — reused for every reconciliation")

    ref_replace = st.file_uploader("Replace SKU (.xlsx)",     type="xlsx", key="ref1")
    ref_pwn     = st.file_uploader("PWN + 10% (.xlsx)",       type="xlsx", key="ref2")
    ref_closed  = st.file_uploader("Meesho Closed SKU (.xlsx)", type="xlsx", key="ref3")

    st.markdown("---")
    st.markdown("### ℹ️ Account Detection")
    st.markdown("""
    Filename must contain:
    - `_YG` → Yash Gallery  
    - `_PE` → Pushpa  
    - `_AG` → Aashirwad  
    """)

# ─────────────────────────────────────────────
# MAIN — Upload Order CSVs
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">📂 Upload Meesho Order CSV Files</div>', unsafe_allow_html=True)
order_files = st.file_uploader(
    "Upload one or more order CSV files (YG / PE / AG)",
    type="csv", accept_multiple_files=True, key="orders"
)

# ─────────────────────────────────────────────
# PROCESS
# ─────────────────────────────────────────────
if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
    errors = []
    if not ref_replace: errors.append("Replace SKU file")
    if not ref_pwn:     errors.append("PWN+10% file")
    if not ref_closed:  errors.append("Meesho Closed SKU file")
    if not order_files: errors.append("At least one order CSV")

    if errors:
        st.error(f"Please upload: {', '.join(errors)}")
    else:
        with st.spinner("Loading reference files..."):
            sku_map    = load_replace_sku(ref_replace.read())
            pwn_map    = load_pwn(ref_pwn.read())
            closed_map = load_closed_sku(ref_closed.read())

        st.success(f"✅ Reference data loaded — {len(pwn_map):,} PWN prices | {len(closed_map):,} closed SKUs")

        all_sheets = {}
        all_stats  = []

        progress = st.progress(0)
        for i, f in enumerate(order_files):
            account_name, account_code = detect_account(f.name)
            st.markdown(f'<div class="section-header">Processing: {f.name} → {account_name}</div>',
                        unsafe_allow_html=True)

            try:
                df_raw = pd.read_csv(f)
                df_out = process_csv(df_raw, account_code, account_name,
                                     sku_map, pwn_map, closed_map)

                sheet_key = f"{account_code}_{f.name[:20]}"
                all_sheets[sheet_key] = df_out

                # Stats
                diff_series = pd.to_numeric(df_out["Difference Amount"], errors="coerce")
                total_diff  = diff_series.sum()
                profit_cnt  = (diff_series >= 0).sum()
                loss_cnt    = (diff_series < 0).sum()
                not_found   = (df_out["Difference Amount"] == "SKU Not Found").sum()
                closed_cnt  = (df_out["SKU Status"] == "Closed").sum()

                all_stats.append({
                    "file": f.name, "account": account_name,
                    "orders": len(df_out), "profit": profit_cnt,
                    "loss": loss_cnt, "not_found": not_found,
                    "closed": closed_cnt, "total_diff": total_diff
                })

                # Show stats cards
                c1, c2, c3, c4, c5 = st.columns(5)
                with c1:
                    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(df_out):,}</div><div class="stat-label">Total Orders</div></div>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<div class="stat-box"><div class="stat-num profit">{profit_cnt:,}</div><div class="stat-label">Profit Orders</div></div>', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'<div class="stat-box"><div class="stat-num loss">{loss_cnt:,}</div><div class="stat-label">Loss Orders</div></div>', unsafe_allow_html=True)
                with c4:
                    st.markdown(f'<div class="stat-box"><div class="stat-num">{closed_cnt:,}</div><div class="stat-label">Closed SKUs</div></div>', unsafe_allow_html=True)
                with c5:
                    color_class = "profit" if total_diff >= 0 else "loss"
                    st.markdown(f'<div class="stat-box"><div class="stat-num {color_class}">₹{total_diff:,.0f}</div><div class="stat-label">Net Difference</div></div>', unsafe_allow_html=True)

                # Preview table
                with st.expander(f"👁️ Preview — {account_name} (first 20 rows)"):
                    st.dataframe(df_out.head(20), use_container_width=True)

            except Exception as e:
                st.error(f"Error processing {f.name}: {e}")

            progress.progress((i + 1) / len(order_files))

        # ── Overall Summary ──
        if all_stats:
            st.markdown('<div class="section-header">📊 Overall Summary</div>', unsafe_allow_html=True)
            total_orders = sum(s["orders"] for s in all_stats)
            total_profit = sum(s["profit"] for s in all_stats)
            total_loss   = sum(s["loss"] for s in all_stats)
            grand_diff   = sum(s["total_diff"] for s in all_stats)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown(f'<div class="stat-box"><div class="stat-num">{total_orders:,}</div><div class="stat-label">Total Orders (All Accounts)</div></div>', unsafe_allow_html=True)
            with c2:
                st.markdown(f'<div class="stat-box"><div class="stat-num profit">{total_profit:,}</div><div class="stat-label">Total Profit Orders</div></div>', unsafe_allow_html=True)
            with c3:
                st.markdown(f'<div class="stat-box"><div class="stat-num loss">{total_loss:,}</div><div class="stat-label">Total Loss Orders</div></div>', unsafe_allow_html=True)
            with c4:
                color_class = "profit" if grand_diff >= 0 else "loss"
                st.markdown(f'<div class="stat-box"><div class="stat-num {color_class}">₹{grand_diff:,.0f}</div><div class="stat-label">Grand Net Difference</div></div>', unsafe_allow_html=True)

        # ── Download ──
        if all_sheets:
            st.markdown("---")
            st.markdown('<div class="section-header">⬇️ Download Reconciliation Report</div>', unsafe_allow_html=True)
            excel_buf = export_excel(all_sheets)
            st.download_button(
                label="📥 Download Excel Report (All Accounts)",
                data=excel_buf,
                file_name="Meesho_Reconciliation_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<center style='color:#aaa; font-size:0.8rem'>Meesho Reconciliation Tool · Built with Streamlit</center>",
    unsafe_allow_html=True
)
