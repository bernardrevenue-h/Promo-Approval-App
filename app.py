import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Promo Approval Tool", page_icon="📊", layout="wide")
st.title("📊 Promo Approval Tool")
st.markdown("Upload satu file Excel dengan sheet: Promo Input, ME Per Store, Sales Mix.")
st.divider()

# ── Helpers ──────────────────────────────────────────────────────
def read_sheet(file, sheet_name):
    df = pd.read_excel(file, sheet_name=sheet_name, header=1)
    df.columns = df.columns.str.strip()
    return df

def round_to_900(price):
    base = round(price / 1000) * 1000
    candidate = base - 100
    return candidate if abs(price - candidate) <= abs(price - (candidate + 1000)) else candidate + 1000

def to_excel_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Output")
    return buffer.getvalue()

def calc_predicted_me(new_final_prices, book_prices, qtys, store_sales, me_current, pc_store_current):
    pc_new       = (book_prices - new_final_prices) / book_prices
    price_chg    = (new_final_prices - book_prices * (1 - pc_store_current)) / (book_prices * (1 - pc_store_current) + 1e-9)
    qty_new      = np.maximum(qtys * (1 - price_chg * 2.5), 0)
    sm_new       = (qty_new * book_prices) / store_sales
    sm_new       = sm_new / sm_new.sum() if sm_new.sum() > 0 else sm_new
    pc_store_new = (pc_new * sm_new).sum()
    predicted_me = me_current - (pc_store_new - pc_store_current) * 0.70
    return predicted_me, pc_new, sm_new, qty_new

def solve_new_prices(book_prices, final_prices, qtys, store_sales, me_current, pc_store_current, me_target):
    pc_current_arr = (book_prices - final_prices) / book_prices
    best_result, best_diff = None, 999
    for delta in np.arange(-0.20, 0.20, 0.0005):
        new_pc    = np.clip(pc_current_arr + delta, 0, 0.6)
        new_final = np.array([round_to_900(p) for p in book_prices * (1 - new_pc)])
        pred_me, pc_new, sm_new, qty_new = calc_predicted_me(
            new_final, book_prices, qtys, store_sales, me_current, pc_store_current
        )
        diff = abs(pred_me - me_target)
        if diff < best_diff:
            best_diff   = diff
            best_result = (new_final, pred_me, pc_new, sm_new, qty_new)
    return best_result, best_diff

# ── Upload File ───────────────────────────────────────────────────
st.subheader("📁 Upload File")
uploaded_file = st.file_uploader(
    "Upload file Excel (.xlsx) — sheet: Promo Input, ME Per Store, Sales Mix",
    type=["xlsx"]
)

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        required_sheets = ["Promo Input", "ME Per Store", "Sales Mix"]
        missing = [s for s in required_sheets if s not in xl.sheet_names]
        if missing:
            st.error(f"❌ Sheet tidak ditemukan: {', '.join(missing)}")
            st.stop()
        else:
            st.success(f"✅ File berhasil dibaca — sheet: {', '.join(required_sheets)}")
    except Exception as e:
        st.error(f"❌ Gagal membaca file: {e}")
        st.stop()

st.divider()

# ── ME Target ────────────────────────────────────────────────────
st.subheader("🎯 ME Target")
me_target_input = st.number_input(
    "Masukkan ME Target (%)", min_value=0.0, max_value=100.0,
    value=28.0, step=0.1, format="%.1f"
)
me_target = me_target_input / 100

st.divider()

# ── Generate ─────────────────────────────────────────────────────
if st.button("🚀 Generate Output", type="primary", use_container_width=True, disabled=not uploaded_file):
    with st.spinner("Memproses data..."):
        try:
            df_promo    = read_sheet(uploaded_file, "Promo Input")
            df_me_store = read_sheet(uploaded_file, "ME Per Store")
            df_sales    = read_sheet(uploaded_file, "Sales Mix")

            # ── ME Per Store: ambil latest per store + platform ──
            df_me_store['monday_of_week'] = pd.to_datetime(df_me_store['monday_of_week'])
            df_me_store = df_me_store[df_me_store['visit_purpose_name'] == 'Grab']
            df_me_store_latest = (
                df_me_store
                .sort_values('monday_of_week')
                .groupby(['visit_purpose_name', 'store_brand_owner'])
                .last()
                .reset_index()
                .rename(columns={
                    'visit_purpose_name': 'Platform',
                    'store_brand_owner':  'Store Brand',
                    'store_me_percent':   'ME_Store_Pct',
                    'gross_sales':        'Store_Sales'
                })
            )

            # ── Sales Mix: qty per menu_code + platform ──────────
            df_sales['monday_of_week'] = pd.to_datetime(df_sales['monday_of_week'])
            df_qty = (
                df_sales[df_sales['visit_purpose_name'] == 'Grab']
                .groupby(['menu_code', 'visit_purpose_name'])['qty_total']
                .sum()
                .reset_index()
                .rename(columns={
                    'menu_code':           'Menu Code Child',
                    'visit_purpose_name':  'Platform',
                    'qty_total':           'Qty'
                })
            )

            # ── Build base (Grab only) ────────────────────────────
            output = df_promo[df_promo['Platform'] == 'Grab'].copy()
            output = output.merge(
                df_me_store_latest[['Platform', 'Store Brand', 'ME_Store_Pct', 'Store_Sales']],
                on=['Platform', 'Store Brand'], how='left'
            )
            output = output.merge(df_qty, on=['Menu Code Child', 'Platform'], how='left')

            # ── Current calculations ──────────────────────────────
            output['ME_SKU'] = pd.to_numeric(output['M/E (%)'], errors='coerce')
            output['Price_Cut_Current'] = (output['Book Price'] - output['Final Price']) / output['Book Price']
            output['Sales_Mix_Current'] = (output['Qty'] * output['Book Price']) / output['Store_Sales']
            total_sm = output['Sales_Mix_Current'].sum()
            output['Sales_Mix_Current'] = output['Sales_Mix_Current'] / total_sm
            pc_store_current = (output['Price_Cut_Current'] * output['Sales_Mix_Current']).sum()

            book_prices  = output['Book Price'].values.astype(float)
            final_prices = output['Final Price'].values.astype(float)
            qtys         = output['Qty'].values.astype(float)
            store_sales  = output['Store_Sales'].iloc[0]
            me_current   = output['ME_Store_Pct'].iloc[0]
            me_sku       = output['ME_SKU'].values.astype(float)

            # ── Solve ─────────────────────────────────────────────
            best_result, best_diff = solve_new_prices(
                book_prices, final_prices, qtys, store_sales,
                me_current, pc_store_current, me_target
            )
            new_finals, pred_me, pc_new, sm_new, qty_new = best_result

            output['New_Final_Price']    = new_finals
            output['New_Price_Cut']      = pc_new
            output['New_Sales_Mix']      = sm_new
            output['Predicted_ME_Store'] = pred_me

            # Max diff check (3%)
            me_range = me_sku.max() - me_sku.min()
            me_flag  = me_range > 0.03

            # ── Display table ─────────────────────────────────────
            display = pd.DataFrame({
                'Platform':        output['Platform'],
                'Store Brand':     output['Store Brand'],
                'Menu Name':       output['Menu Name'],
                'Platform Price':  output['Platform Price'].apply(lambda x: f"Rp {x:,.0f}"),
                'Book Price':      output['Book Price'].apply(lambda x: f"Rp {x:,.0f}"),
                'Final Price':     output['Final Price'].apply(lambda x: f"Rp {x:,.0f}"),
                'Price Cut (cur)': output['Price_Cut_Current'].apply(lambda x: f"{x*100:.2f}%"),
                'ME SKU (cur)':    output['ME_SKU'].apply(lambda x: f"{x*100:.2f}%"),
                'Sales Mix (cur)': output['Sales_Mix_Current'].apply(lambda x: f"{x*100:.2f}%"),
                'New Final Price':  output['New_Final_Price'].apply(lambda x: f"Rp {x:,.0f}"),
                'New Price Cut':   output['New_Price_Cut'].apply(lambda x: f"{x*100:.2f}%"),
                'New Sales Mix':   output['New_Sales_Mix'].apply(lambda x: f"{x*100:.2f}%"),
                'Predicted ME':    output['Predicted_ME_Store'].apply(lambda x: f"{x*100:.2f}%"),
            })

            # ── Results ───────────────────────────────────────────
            st.divider()
            st.subheader("📤 Output")

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("ME Current",     f"{me_current*100:.2f}%")
            m2.metric("ME Target",      f"{me_target*100:.2f}%")
            m3.metric("Predicted ME",   f"{pred_me*100:.2f}%")
            m4.metric("Diff vs Target", f"{abs(pred_me - me_target)*100:.3f}%")

            if me_flag:
                st.warning(f"⚠️ Gap ME antar SKU = {me_range*100:.2f}% — melebihi batas 3%!")
            else:
                st.success(f"✅ Gap ME antar SKU = {me_range*100:.2f}% — dalam batas 3%")

            if abs(pred_me - me_target) <= 0.002:
                st.success(f"✅ Predicted ME {pred_me*100:.2f}% — dalam toleransi ±0.2% dari target")
            else:
                st.warning(f"⚠️ Predicted ME {pred_me*100:.2f}% — di luar toleransi ±0.2% dari target")

            st.dataframe(display, use_container_width=True)

            st.download_button(
                label="⬇️ Download Output (Excel)",
                data=to_excel_download(display),
                file_name="promo_approval_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.exception(e)

if not uploaded_file:
    st.caption("⬆️ Upload file Excel dulu untuk mengaktifkan tombol.")
