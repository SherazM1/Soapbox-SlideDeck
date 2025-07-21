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
# Slide 4 Preview
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Slide 4: Program Overview (Preview)")

try:
    metrics = extract_proposed_metrics_anywhere(df)
except Exception as e:
    metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
    st.warning("Could not extract 'Proposed Metrics' from Excel. Please check formatting.")

# ----- Extract new values for TextBox 15 -----
# Social Posts & Stories
social_posts_value = ""
if "Organic & Total" in df.columns:
    for idx, row in df.iterrows():
        if str(row["Organic & Total"]).strip().lower() == "total number of posts with stories":
            social_posts_value = row.iloc[1]
            break

# Engagement Rate
engagement_rate_value = ""
for idx, row in df.iterrows():
    if str(row.iloc[0]).strip().lower() == "program er":
        raw_val = row.iloc[1]
        if isinstance(raw_val, str) and raw_val.startswith("#"):
            engagement_rate_value = ""
        else:
            engagement_rate_value = raw_val
        break

# Engagements value
engagements_value = ""
if "Organic & Total" in df.columns:
    for idx, row in df.iterrows():
        if str(row["Organic & Total"]).strip().lower() == "total engagements":
            engagements_value = row.iloc[1]
            break

# Engagements % increase
engagements_increase = ""
if "Proposed Metrics" in df.columns and "Percentage Increase" in df.columns:
    for idx, row in df.iterrows():
        if str(row["Proposed Metrics"]).strip().lower() == "engagements":
            engagements_increase = row["Percentage Increase"]
            break

# Impressions value (w/ fallback)
impressions_value = ""
for idx, row in df.iterrows():
    first_col_val = str(row.iloc[0]).strip().lower()
    if first_col_val == "total impressions":
        impressions_value = row.iloc[1]
        break
    elif first_col_val == "total":
        impressions_value = row.iloc[1]  # fallback

# Impressions % increase
impressions_increase = ""
if "Proposed Metrics" in df.columns and "Percentage Increase" in df.columns:
    for idx, row in df.iterrows():
        if str(row["Proposed Metrics"]).strip().lower() == "impressions":
            impressions_increase = row["Percentage Increase"]
            break

with st.container():
    st.markdown("#### What will appear on the slide:")
    st.markdown(f"""
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}

- **Social Posts & Stories:** {social_posts_value}
- **Engagement Rate:** {engagement_rate_value}%
- **Engagements:** {engagements_value} ({engagements_increase}% increase)
- **Impressions:** {impressions_value} ({impressions_increase}% increase)
""")
    st.caption("These values will be automatically inserted into Slide 4 of your recap deck.")


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
