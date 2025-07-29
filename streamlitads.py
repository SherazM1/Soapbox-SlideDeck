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

paid_likes = ""
if "Dates" in df.columns and "Unnamed: 14" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Reactions":
                paid_likes = row["Dates"]
                break

paid_comments = ""
if "Unnamed: 14" in df.columns and "Dates" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Comments":
                paid_comments = row["Dates"]
                break
    
paid_shares = ""
if "Unnamed: 14" in df.columns and "Dates" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Shares":
                paid_shares = row["Dates"]
                break

paid_saves = ""
if "Unnamed: 14" in df.columns and "Dates" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Saves":
                paid_saves = row["Dates"]
                break 
    
paid_threesec = ""
if "Unnamed: 14" in df.columns and "Dates" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "3 sec vid views":
                paid_threesec = row["Dates"]
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

##total_post_engagements = (
    ##int(organic_likes) + int(organic_comments) + int(organic_shares) + int(organic_saves)
   ##+ int(paid_likes) + int(paid_comments) + int(paid_shares) + int(paid_saves) + int(paid_threesec)
##)

story_engagements = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
    for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Story Engagements":
                story_engagements = row["Unnamed: 11"]
                break

paid_engagements = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
    for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Paid Engagements":
                paid_engagements = row["Unnamed: 11"]
                break

total_engagements = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Engagements":
                total_engagements = row["Unnamed: 11"]

cpe = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "CPE":
                cpe = row["Unnamed: 17"]
                break
    
cpc = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "CPC":
                cpc = row["Unnamed: 17"]
                break

ctr = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "CTR":
                ctr = row["Unnamed: 17"]
                break

cpm = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "CPM":
                cpm = row["Unnamed: 17"]

thruplays = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "ThruPlays":
                thruplays = row["Unnamed: 17"]
                break

p25 = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "0.25":
                p25 = row["Unnamed: 17"]
                break

p50 = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "0.5":
             p50 = row["Unnamed: 17"]
             break

p75 = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
            if str(row["Unnamed: 18"]).strip() == "0.75":
             p75 = row["Unnamed: 17"]
             break

p100 = ""
if "Unnamed: 18" in df.columns and "Unnamed: 17" in df.columns:
        for _, row in df.iterrows():
         if str(row["Unnamed: 18"]).strip() == "1":
            p100 = row["Unnamed: 17"]

diversity_value = ""
if "Diversity" in df.columns:
    col = df["Diversity"]
    for idx, val in enumerate(col):
        if str(val).strip() == "Diversity":
            diversity_value = col.iloc[idx + 1]
            break


col1, col2, col3, col4, col5 = st.columns(5)

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
- **Diversity Rate:** {diversity_value}

''')
        st.caption("These values will be automatically inserted into Slide 4 of your recap deck.")

with col2:
     with st.container():
          st.markdown("#### What will appear on **Slide 7:**")
          st.markdown(f'''
- **Paid Impressions:** {impressions_paid}
- **Paid Engagements:** {paid_engagements}

''')

                      
with col3:
    with st.container():
        st.markdown("#### What will appear on **Slide 9:**")
        st.markdown("##### **MAKE SURE TO MANUALLY ADD CART TRANSFERS**  ")
        st.markdown(f'''
- **Organic Likes:** {organic_likes}
- **Organic Comments:** {organic_comments}
- **Organic Shares:** {organic_shares}
- **Organic Saves:** {organic_saves}
- **Paid Likes:** {paid_likes}
- **Paid Comments:** {paid_comments}
- **Paid Shares:** {paid_shares}
- **Paid Saves:** {paid_saves}
- **3 Second Video Views:** {paid_threesec}
- **Total Post Engagements** 
- **Total Story Engagements** {story_engagements}
- **Total Engagements** {total_engagements}
''')
        st.caption("These values will be automatically inserted into Slide 9 of your recap deck.")

with col4:
    with st.container():
        st.markdown("#### What will appear on **Slides 10 and 11:**")
        st.markdown(f'''
- **Influencer Reach:** {organic_reach_impressions}
- **Ad Impressions:** {impressions_paid}
- **Total Views:** {organic_views_impressions}
- **Total Impressions:** {impressions_value}
         
    ''')
        
with col5:
     with st.container():
          st.markdown("#### What will appear on **Slide 12:**")
          st.markdown(f'''
                      
- **CPE:** {cpe}
- **CPC:** {cpc}
- **CTR:** {ctr}
- **CPM:** {cpm}
- **ThruPlays:** {thruplays}
- **Plays at 25%:** {p25}
- **Plays at 50%:** {p50}
- **Plays at 75%:** {p75}
- **Plays at 100%:** {p100}

''')
          st.caption("These values will be automatically inserted into Slide 12 of your recap deck.")


influencer_slide_6 = st.text_input("Enter the Influencer Handle for Slide 6", value="@influencerhandle")
influencer_slide_7_left = st.text_input("Enter the Influencer Handle for the Organic Part of Slide 7", value="@influencerhandle")
influencer_slide_7_right = st.text_input("Enter the Influencer Handle for the Paid Part of Slide 7", value="@influencerhandle")
slide_7_likestext = st.text_input("Enter the Likes for the Organic Part of Slide 7", value="1")
slide_7_commentstext = st.text_input("Enter the Comments for the Organic Part of Slide 7", value="2")
slide_7_viewstext = st.text_input("Enter the Views for the Organic Part of Slide 7", value="3")
slide_7_reachtext = st.text_input("Enter the Social Reach for the Organic Part of Slide 7", value="4")
slide_7_engage = st.text_input("Enter the Engagements for the Paid Part of Slide 7", value="5")
slide_7_impress = st.text_input("Enter the Impressions for the Paid Part of Slide 7", value="6")
slide_9_text = st.text_input("Enter the Line for the Top Text of Slide 9", value="hey this is a placeholder")

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



col1, col2, col3, col4 = st.columns(4)
with col1:
     slide_11_first_img = st.file_uploader("Slide 11 - 1st image (farthest left)", type=["png", "jpg", "jpeg"], key="slide11first")
with col2:
     slide_11_second_img = st.file_uploader("Slide 11 — 2nd image", type=["png", "jpg", "jpeg"], key="slide11second")
with col3:
     slide_11_third_img = st.file_uploader("Slide 11 — 3nd image", type=["png", "jpg", "jpeg"], key="slide11third")
with col4:
     slide_11_fourth_img = st.file_uploader("Slide 11 — 4th image (farthest right)", type=["png", "jpg", "jpeg"], key="slide11fourth")
                                           


influencer_boxes = [
    {"label": "Influencer Box 1", "key": "box_1", "textbox": "TextBox 62"},
    {"label": "Influencer Box 2", "key": "box_2", "textbox": "TextBox 13"},
    {"label": "Influencer Box 3", "key": "box_3", "textbox": "TextBox 9"},
    {"label": "Influencer Box 4", "key": "box_4", "textbox": "TextBox 15"},
    {"label": "Influencer Box 5", "key": "box_5", "textbox": "TextBox 11"},
    {"label": "Influencer Box 6", "key": "box_6", "textbox": "TextBox 17"},
]

st.markdown("### Enter Influencer Data for Each Box (for Slide 5)")
influencer_inputs = {}

for box in influencer_boxes:
    with st.container():
        st.markdown(f"**{box['label']}**")
        handle = st.text_input(f"{box['label']} - Handle", key=f"{box['key']}_handle")
        reach = st.text_input(f"{box['label']} - Social Reach", key=f"{box['key']}_reach")
        city = st.text_input(f"{box['label']} - City", key=f"{box['key']}_city")
        state = st.text_input(f"{box['label']} - State", key=f"{box['key']}_state")
        verbatim = st.text_input(f"{box['label']} - Verbatim", key=f"{box['key']}_verbatim")
        influencer_inputs[box["textbox"]] = {
            "@influencerhandle": handle,
            "##": reach,
            "City": city,
            "State": state,
            '"Verbatim"': verbatim,
        }



images = {
    "slide_6": slide_6_img,
    "slide_7_left": slide_7_left_img,
    "slide_7_right": slide_7_right_img,
    "slide_8_first": slide_8_first_img,
    "slide_8_second": slide_8_second_img,
    "slide_8_third": slide_8_third_img,
    "slide_8_fourth": slide_8_fourth_img,
    "slide_11_first": slide_11_first_img,
    "slide_11_second": slide_11_second_img,
    "slide_11_third": slide_11_third_img,
    "slide_11_fourth": slide_11_fourth_img,


}

text_inputs = {
         
    "slide_6": influencer_slide_6,
    "slide_7_left": influencer_slide_7_left,
    "slide_7_right": influencer_slide_7_right,
    "slide_7_like": slide_7_likestext,
    "slide_7_comment": slide_7_commentstext,
    "slide_7_view": slide_7_viewstext,
    "slide_7_reaches": slide_7_reachtext,
    "slide_7_eng": slide_7_engage,
    "slide_7_impr": slide_7_impress,
    "slide_9": slide_9_text,
    "influencer_boxes": influencer_inputs


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
    "slide_8_fourth": slide_8_fourth_img,
    "slide_11_first": slide_11_first_img,
    "slide_11_second": slide_11_second_img,
    "slide_11_third": slide_11_third_img,
    "slide_11_fourth": slide_11_fourth_img,
}
    text_inputs = {
         
    "slide_6": influencer_slide_6,
    "slide_7_left": influencer_slide_7_left,
    "slide_7_right": influencer_slide_7_right,
    "slide_7_like": slide_7_likestext,
    "slide_7_comment": slide_7_commentstext,
    "slide_7_view": slide_7_viewstext,
    "slide_7_reaches": slide_7_reachtext,
    "slide_7_eng": slide_7_engage,
    "slide_7_impr": slide_7_impress,
    "slide_9": slide_9_text,
    "influencer_boxes": influencer_inputs

    
    }     



    from datetime import datetime  # Make sure this is imported!
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"recap_deck_output_{timestamp}.pptx"

    populate_pptx_from_excel(df, pptx_template_path, output_path, images=images, text_inputs=text_inputs)

    with open(output_path, "rb") as f:
        st.success("✅ Your recap deck is ready!")
        st.download_button("⬇️ Download PowerPoint", data=f, file_name=f"recap_deck_{timestamp}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")