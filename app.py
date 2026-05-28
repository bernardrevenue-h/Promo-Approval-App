import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Promo Approval Tool", page_icon="📊", layout="wide")

st.title("📊 Promo Approval Tool")
st.markdown("Upload semua file yang dibutuhkan, lalu klik **Proses** untuk generate output.")
st.divider()

# ── Helper ───────────────────────────────────────────────────────
def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    else:
        return pd.read_excel(file)

def to_excel_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Output")
    return buffer.getvalue()

# ── Upload Files ─────────────────────────────────────────────────
st.subheader("📁 Upload Files")

col1, col2, col3, col4 = st.columns(4)

with col1:
    promo_input = st.file_uploader("1. Promo Input", type=["xlsx", "csv"])
with col2:
    promo_me = st.file_uploader("2. M/E Component (Promo M/E)", type=["xlsx", "csv"])
with col3:
    me_per_store = st.file_uploader("3. M/E Per Store (Raw Data)", type=["xlsx", "csv"])
with col4:
    sales_mix = st.file_uploader("4. Sales Mix (Raw Data)", type=["xlsx", "csv"])

st.divider()

# ── Status ───────────────────────────────────────────────────────
st.subheader("📋 Status Upload")
status_cols = st.columns(4)
files = {
    "Promo Input": promo_input,
    "M/E Component": promo_me,
    "M/E Per Store": me_per_store,
    "Sales Mix": sales_mix,
}
for i, (name, file) in enumerate(files.items()):
    with status_cols[i]:
        if file:
            st.success(f"✅ {name}")
        else:
            st.warning(f"⏳ {name}")

st.divider()

# ── Process ──────────────────────────────────────────────────────
all_uploaded = all(files.values())

if st.button("🚀 Generate Output", type="primary", use_container_width=True, disabled=not all_uploaded):
    with st.spinner("Memproses data..."):
        try:
            df_promo    = read_file(promo_input)
            df_me       = read_file(promo_me)
            df_me_store = read_file(me_per_store)
            df_sales    = read_file(sales_mix)

            # Get latest M/E per store
            df_me_store['Month'] = pd.to_datetime(df_me_store['Month'])
            df_me_store_latest = (
                df_me_store
                .sort_values('Month')
                .groupby(['Platform', 'Store Brand'])
                .last()
                .reset_index()
            )

            # Aggregate qty total from sales mix
            df_qty = (
                df_sales
                .groupby(['menu_code', 'visit_purpose_name'])['qty_total']
                .sum()
                .reset_index()
            )
            df_qty.columns = ['Menu Code Child', 'Platform', 'Qty Total']

            # Build output
            output = df_promo[[
                'Platform', 'Store Brand', 'Menu Name',
                'Platform Price', 'Book Price', 'Final Price',
                'Menu Code Child', 'Net Price'
            ]].copy()

            # Join M/E %
            output = output.merge(
                df_me[['Menu Code Child', 'Platform', 'M/E Parent']],
                on=['Menu Code Child', 'Platform'],
                how='left'
            )

            # Join M/E per store
            output = output.merge(
                df_me_store_latest[['Platform', 'Store Brand', 'M/E Store']],
                on=['Platform', 'Store Brand'],
                how='left'
            )

            # Join Qty Total
            output = output.merge(
                df_qty,
                on=['Menu Code Child', 'Platform'],
                how='left'
            )

            # Rename columns
            output = output.rename(columns={'M/E Parent': 'M/E (%)'})

            st.success(f"✅ Output berhasil dibuat! {len(output)} baris data.")
            st.divider()

            # ── Preview ──────────────────────────────────────────
            st.subheader("👀 Preview Output")
            st.dataframe(output, use_container_width=True)

            # ── Download ─────────────────────────────────────────
            st.download_button(
                label="⬇️ Download Output (Excel)",
                data=to_excel_download(output),
                file_name="promo_approval_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")

if not all_uploaded:
    st.caption("⬆️ Upload semua 4 file dulu untuk mengaktifkan tombol.")
