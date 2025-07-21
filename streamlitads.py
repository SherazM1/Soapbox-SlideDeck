import os
import streamlit as st
import pandas as pd
from datetime import datetime
from app import load_dataframe, populate_pptx_from_excel, extract_proposed_metrics_anywhere

# ─────────────────────────────────────────────────────────────────────────────
# Page Setup
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="SOAPBOX Recap Deck App", page_icon="✎", layout="wide")
if os.path.exists("logo.png"):
    st.image("logo.png", width=180)

st.title("Recap Deck Editor")
st.markdown("Upload your Excel, see a live preview of your slide, and download your PowerPoint recap deck.")

st.header("Step 1: Upload Data File")
uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
if not uploaded:
    st.info("Please upload your Excel/CSV to generate your recap deck.")
    st.stop()

df = load_dataframe(uploaded)
st.subheader("Preview: First 10 Rows of Data")
st.dataframe(df.head(10), height=250)

st.markdown("---")
st.header("Slide 4: Program Overview (Preview)")

try:
    metrics = extract_proposed_metrics_anywhere(df)
except Exception:
    metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
    st.warning("Could not extract 'Proposed Metrics' from Excel. Please check formatting.")

# Social Posts & Stories
social_posts_value = ""
for _, row in df.iterrows():
    if str(row["Organic & Total"]).strip() == "Total Number of Posts With Stories":
        social_posts_value = row["Unnamed: 11"]
        break

# Percent Increases
engagements_increase = ""
impressions_increase = ""

try:
    perc_row_idx = None
    for idx, row in df.iterrows():
        if str(row["Unnamed: 15"]).strip().lower() == "percentage increase":
            perc_row_idx = idx
            break
    if perc_row_idx is not None:
        for offset in range(1, 10):
            row_idx = perc_row_idx + offset
            if row_idx >= len(df):
                break
            label = str(df.at[row_idx, "Unnamed: 14"]).strip().lower()
            raw = df.at[row_idx, "Unnamed: 15"]
            if pd.notna(raw):
                try:
                    val = float(raw) * 100
                    pct_str = f"{val:.1f}%"
                    if label == "engagements":
                        engagements_increase = pct_str
                    elif label == "impressions":
                        impressions_increase = pct_str
                except Exception:
                    continue
except Exception as e:
    st.warning("Error extracting percentage increases.")

with st.container():
    st.markdown("#### What will appear on the slide:")
    st.markdown(f'''
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}
- **Social Posts & Stories:** {social_posts_value}
- **Engagements:** ({engagements_increase} increase)
- **Impressions:** ({impressions_increase} increase)
''')
    st.caption("These values will be automatically inserted into Slide 4 of your recap deck.")

st.markdown("---")
st.header("Step 2: Download Recap Deck")
pptx_template_path = "template.pptx"

if st.button("Generate PowerPoint Recap Deck"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"recap_deck_output_{timestamp}.pptx"

    populate_pptx_from_excel(df, pptx_template_path, output_path)
    with open(output_path, "rb") as f:
        st.success("✅ Your recap deck is ready!")
        st.download_button("⬇️ Download PowerPoint", data=f, file_name=f"recap_deck_{timestamp}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")