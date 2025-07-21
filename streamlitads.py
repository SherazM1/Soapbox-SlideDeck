import os
import streamlit as st
import pandas as pd
from datetime import datetime
from app import load_dataframe, populate_pptx_from_excel, extract_proposed_metrics_anywhere, mapping_config

# ─────────────────────────────────────────────────────────────────────────────
# Page Setup
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SOAPBOX Recap Deck App",
    page_icon="✎",
    layout="wide",
)

if os.path.exists("logo.png"):
    st.image("logo.png", width=180)

st.title("Recap Deck Editor")
st.markdown("Upload your Excel, see a live preview of your slide, and download your PowerPoint recap deck.")

# ─────────────────────────────────────────────────────────────────────────────
# Step 1: Upload Excel
# ─────────────────────────────────────────────────────────────────────────────
st.header("Step 1: Upload Data File")
uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
if not uploaded:
    st.info("Please upload your Excel/CSV to generate your recap deck.")
    st.stop()

df = load_dataframe(uploaded)
st.subheader("Preview: First 10 Rows of Data")
st.dataframe(df.head(10), height=250)

# ─────────────────────────────────────────────────────────────────────────────
# Slide 1 Preview
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Slide 1: Proposed Program Details (Preview)")

try:
    metrics = extract_proposed_metrics_anywhere(df)
except Exception as e:
    metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
    st.warning("Could not extract 'Proposed Metrics' from Excel. Please check formatting.")

with st.container():
    st.markdown("#### What will appear on the slide:")
    st.markdown(f"""
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}
""")
    st.caption("These values will be automatically inserted into Slide 1 of your recap deck.")

# ─────────────────────────────────────────────────────────────────────────────
# Step 2: Generate + Download PowerPoint
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Step 2: Download Recap Deck")

pptx_template_path = "template.pptx"

if st.button("Generate PowerPoint Recap Deck"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"recap_deck_output_{timestamp}.pptx"

    pptx_file = populate_pptx_from_excel(
        excel_df=df,
        pptx_template_path=pptx_template_path,
        output_path=output_path,
        mapping_config= mapping_config
        
          # Expand later as more slides are added
    )

    with open(pptx_file, "rb") as f:
        st.success("✅ Your recap deck is ready!")
        st.download_button(
            "⬇️ Download PowerPoint",
            data=f,
            file_name=f"recap_deck_{timestamp}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
