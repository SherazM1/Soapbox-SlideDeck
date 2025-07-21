# streamlit_app.py

import os
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import date


# ─────────────────────────────────────────────────────────────────────────────
# Page config & branding
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SOAPBOX Recap Deck App ",
    page_icon="✎",
    layout="wide",
)

if os.path.exists("logo.png"):
    st.image("logo.png", width=180)

st.title("Recap Deck Editor")
st.markdown(
    "Upload your data, edit your recaps, and export the full slideshow all in here!"
)

# ─────────────────────────────────────────────────────────────────────────────
# Group Management
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# Inputs: File, Client, Date
# ─────────────────────────────────────────────────────────────────────────────
st.header("Inputs")
uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
if not uploaded:
    st.info("Please upload a data file to see the dashboard below.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Load & Preview Data
# ─────────────────────────────────────────────────────────────────────────────
df = load_dataframe(uploaded)
st.subheader("Data Preview")
st.dataframe(df.head(10), height=200)

# ─────────────────────────────────────────────────────────────────────────────
# Compute & Display Metrics
# ─────────────────────────────────────────────────────────────────────────────
m = compute_metrics(df)

col1, col2 = st.columns([3, 1])
with col1:
    st.subheader(f"{m['above']}/{m['total']} ({m['pct_above']:.1f}%) products ≥ {int(m['threshold'])}%")
    pie_buf = make_pie_bytes(m)
    st.image(pie_buf, caption="Score Distribution", use_column_width=False)

with col2:
    st.subheader("Key Metrics")
    st.write(f"- **Average CQS:** {m['avg_cqs']:.1f}%")
    st.write(f"- **SKUs ≥ {int(m['threshold'])}%:** {m['above']}")
    st.write(f"- **SKUs < {int(m['threshold'])}%:** {m['below']}")
    st.write(f"- **Buybox Ownership:** {m['buybox']:.1f}%")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# Top 5 & Below Tables
# ─────────────────────────────────────────────────────────────────────────────
st.subheader("Top 5 SKUs by Content Quality Score")
top5 = get_top_skus(df)
st.dataframe(top5, height=200)

st.subheader(f"SKUs Below {int(m['threshold'])}%")
skus_below = get_skus_below(df)
st.dataframe(skus_below, height=300)

# Export SKUs Below CSV
csv_data = skus_below.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Download SKUs Below CSV",
    data=csv_data,
    file_name="skus_below.csv",
    mime="text/csv"
)

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# Export Full Dashboard PDF
# ─────────────────────────────────────────────────────────────────────────────
st.header("Export Recap Deck Powerpoint")
if st.button("Generate Powerpoint"):
    pdf_bytes = generate_full_report(uploaded, client_name, rpt_date, client_notes)
    st.success("✅ Powerpoint ready!")
    st.download_button(
        "⬇️ Download Powerpoint",
        data=pdf_bytes,
        file_name="powepoint.pptx",
        mime="application/pdf"
    )
