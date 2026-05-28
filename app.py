import streamlit as st
import pandas as pd

st.set_page_config(page_title="Promo Approval Tool", page_icon="📊", layout="wide")

st.title("📊 Promo Approval Tool")
st.markdown("Upload semua file yang dibutuhkan, lalu klik **Proses** untuk generate output.")
st.divider()

# ── Upload Files ────────────────────────────────────────────────
st.subheader("📁 Upload Files")

col1, col2, col3, col4 = st.columns(4)

with col1:
    promo_list = st.file_uploader("1. Promo List", type=["xlsx", "csv"])

with col2:
    store_component = st.file_uploader("2. Store Component", type=["xlsx", "csv"])

with col3:
    me_component = st.file_uploader("3. M/E Component", type=["xlsx", "csv"])

with col4:
    sales_mix = st.file_uploader("4. Sales Mix Component", type=["xlsx", "csv"])

st.divider()

# ── M/E Direction ───────────────────────────────────────────────
st.subheader("🎯 M/E Direction")

me_direction = st.selectbox(
    "Pilih M/E Direction:",
    options=["-- Pilih --", "Increase", "Decrease", "Maintain"],
)

st.divider()

# ── Output ──────────────────────────────────────────────────────
st.subheader("📤 Output")

all_filled = all([promo_list, store_component, me_component, sales_mix]) and me_direction != "-- Pilih --"

if st.button("🚀 Generate Output", type="primary", use_container_width=True, disabled=not all_filled):
    st.info("⚙️ Logic processing akan ditambahkan setelah template output tersedia.")

if not all_filled:
    st.caption("⬆️ Upload semua file & pilih M/E Direction dulu untuk mengaktifkan tombol.")
