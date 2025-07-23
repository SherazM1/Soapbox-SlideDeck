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
st.subheader("Preview: First 50 Rows of Data")
st.dataframe(df.head(50), height=250)

st.markdown("---")
st.header("Slides: Data Preview")

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

engagements_value = ""
for _, row in df.iterrows():
    if str(row["Organic & Total"]).strip() == "Total Engagements":
        engagements_value = row["Unnamed: 11"]
        break

engagement_rate_value = ""
for _, row in df.iterrows():
    if str(row["Organic & Total"]).strip() == "Program ER":
        engagement_rate_value = float(row["Unnamed: 11"]) * 100
        # Round to two decimal places, then remove leading zero
        engagement_rate_value = f"{engagement_rate_value:.2f}"
        if engagement_rate_value.startswith("0"):
            engagement_rate_value = engagement_rate_value[1:]
        break


impressions_value = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            cell_value = str(row["Organic & Total"]).strip()
            if cell_value in ("Total", "Total Impressions"):
                impressions_value = row["Unnamed: 11"]
                break


# Percent Increases
engagements_increase = ""
impressions_increase = ""

try:
    engagement_val = df.at[5, "Unnamed: 15"]
    impression_val = df.at[4, "Unnamed: 15"]

    if pd.notna(engagement_val):
        engagements_increase = f"{float(engagement_val) * 100:.1f}%"
    if pd.notna(impression_val):
        impressions_increase = f"{float(impression_val) * 100:.1f}%"
except Exception as e:
    st.warning("⚠️ Could not extract fixed-position % increases.")

            
organic_likes = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Likes":
                organic_likes = row["Unnamed: 11"]
                break


organic_comments = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Comments":
                organic_comments = row["Unnamed: 11"]
                break

organic_shares = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Shares":
                organic_shares = row["Unnamed: 11"]
                break

organic_saves = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Saves":
                organic_saves = row["Unnamed: 11"]
                break


organic_views_impressions = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Organic (Views)":
                organic_views_impressions = row["Unnamed: 11"]
                break



organic_reach_impressions = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
    for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Organic (Reach)":
                organic_reach_impressions = row["Unnamed: 11"]
                break

                    
impressions_paid = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
    for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Paid":
                impressions_paid = row["Unnamed: 11"]
                break



col1, col2, col3 = st.columns(3)

with col1:
    with st.container():
        st.markdown("#### What will appear on **Slide 4:**")
        st.markdown(f'''
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}
- **Social Posts & Stories:** {social_posts_value}
- **Engagement Rate:** {engagement_rate_value}
- **Engagements:** {engagements_value} ({engagements_increase} increase)
- **Impressions:** {impressions_value} ({impressions_increase} increase)
''')
        st.caption("These values will be automatically inserted into Slide 4 of your recap deck.")

with col2:
    with st.container():
        st.markdown("#### What will appear on **Slide 9:**")
        st.markdown(f'''
- **Organic Likes:** {organic_likes}
- **Organic Comments:** {organic_comments}
- **Organic Shares:** {organic_shares}
- **Organic Saves:** {organic_saves}
''')
        st.caption("These values will be automatically inserted into Slide 9 of your recap deck.")

with col3:
    with st.container():
        st.markdown("#### What will appear on **Slides 10 and 11:**")
        st.markdown(f'''
- **Influencer Reach:** {organic_reach_impressions}
- **Ad Impressions:** {impressions_paid}
- **Total Views:** {organic_views_impressions}
- **Total Impressions:** {impressions_value}
         
    ''')
        



slide_6_img = st.file_uploader("Upload image for Slide 6", type=["png", "jpg", "jpeg"])

col_left, col_right = st.columns(2)
with col_left:
    slide_7_left_img = st.file_uploader("Slide 7 — Upload image for LEFT box - Organic", type=["png", "jpg", "jpeg"], key="slide7left")
with col_right:
    slide_7_right_img = st.file_uploader("Slide 7 — Upload image for RIGHT box - Paid", type=["png", "jpg", "jpeg"], key="slide7right")


col1, col2, col3, col4 = st.columns(4)
with col1:
    slide_8_first_img = st.file_uploader("Slide 8 — 1st image (farthest left)", type=["png", "jpg", "jpeg"], key="slide8first")
with col2:
    slide_8_second_img = st.file_uploader("Slide 8 — 2nd image", type=["png", "jpg", "jpeg"], key="slide8second")
with col3:
    slide_8_third_img = st.file_uploader("Slide 8 — 3rd image", type=["png", "jpg", "jpeg"], key="slide8third")
with col4:
    slide_8_fourth_img = st.file_uploader("Slide 8 — 4th image (farthest right)", type=["png", "jpg", "jpeg"], key="slide8fourth")



images = {
    "slide_6": slide_6_img,
    "slide_7_left": slide_7_left_img,
    "slide_7_right": slide_7_right_img,
    "slide_8_first": slide_8_first_img,
    "slide_8_second": slide_8_second_img,
    "slide_8_third": slide_8_third_img,
    "slide_8_fourth": slide_8_fourth_img,

}

st.markdown("---")
st.header("Step 2: Download Recap Deck")
pptx_template_path = "template.pptx"

# 2. Create the images dictionary before calling the function


# 3. Generate PowerPoint with images passed in
if st.button("Generate PowerPoint Recap Deck"):
    images = {
    "slide_6": slide_6_img,
    "slide_7_left": slide_7_left_img,
    "slide_7_right": slide_7_right_img,
    "slide_8_first": slide_8_first_img,
    "slide_8_second": slide_8_second_img,
    "slide_8_third": slide_8_third_img,
    "slide_8_fourth": slide_8_fourth_img
}
    from datetime import datetime  # Make sure this is imported!
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"recap_deck_output_{timestamp}.pptx"

    populate_pptx_from_excel(df, pptx_template_path, output_path, images=images)

    with open(output_path, "rb") as f:
        st.success("✅ Your recap deck is ready!")
        st.download_button("⬇️ Download PowerPoint", data=f, file_name=f"recap_deck_{timestamp}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")