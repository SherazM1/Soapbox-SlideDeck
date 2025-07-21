import os
import streamlit as st
import pandas as pd
from app import load_dataframe, populate_pptx_from_excel, extract_proposed_metrics_anywhere

st.set_page_config(
    page_title="SOAPBOX Recap Deck App",
    page_icon="✎",
    layout="wide",
)

if os.path.exists("logo.png"):
    st.image("logo.png", width=180)

st.title("Recap Deck Editor")
st.markdown(
    "Upload your Excel, see a live preview of what will go into your slide, and download your PowerPoint recap deck."
)

st.header("Step 1: Upload Data File")
uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
if not uploaded:
    st.info("Please upload your Excel/CSV to generate your recap deck.")
    st.stop()

df = load_dataframe(uploaded)
st.subheader("Preview: First 10 Rows of Data")
st.dataframe(df.head(10), height=250)

# ─────────────────────────────────────────────────────────────────────────────
# Text Preview of Slide 1 (always works!)
# ─────────────────────────────────────────────────────────────────────────────
try:
    metrics = extract_proposed_metrics_anywhere(df)
except Exception as e:
    metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
    st.warning("Could not extract Proposed Metrics from Excel.")

st.markdown("---")
st.header("Slide 1: Proposed Program Details (Preview)")

with st.container():
    st.markdown("#### What will appear on the slide:")
    st.markdown(f"""
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}
""")
    st.caption("This text will be auto-populated into your PowerPoint slide.")

# ─────────────────────────────────────────────────────────────────────────────
# Export PowerPoint Deck
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Step 2: Download Your Recap Deck PowerPoint")

pptx_template_path = "template.pptx"
output_path = "recap_deck_output.pptx"
pptx_file = populate_pptx_from_excel(df, pptx_template_path, output_path, mapping_config={})

with open(pptx_file, "rb") as f:
    st.download_button(
        "⬇️ Download PowerPoint",
        data=f,
        file_name="recap_deck.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
