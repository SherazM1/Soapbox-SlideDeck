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

diversity_value = ""
for col in df.columns:
    for i, value in enumerate(df[col]):
        # Check if cell matches "Diversity" (case and whitespace insensitive)
        if str(value).strip().lower() == "diversity":
            # Try to read the cell directly below (i+1)
            if i + 1 < len(df[col]):
                diversity_value = df[col].iloc[i + 1]
            break  # Stop after first match
    if diversity_value:
        break  # Stop searching after found


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


influencer_count = ""
if "Dates" in df.columns and "Unnamed: 14" in df.columns:
        for _, row in df.iterrows():
            if str(row["Dates"]).strip() == "Influencers":
                influencer_count = row["Unnamed: 14"]
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

c2c_transfer = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "C2C Transfers":
                c2c_transfer = row["Unnamed: 11"]
                break
    
c2c_value = ""
if "Organic & Total" in df.columns and "Unnamed: 11" in df.columns:
        for _, row in df.iterrows():
            if str(row["Organic & Total"]).strip() == "C2C Value": 
                c2c_value = row["Unnamed: 11"]


col1, col2, col3, col4, col5, col6 = st.columns(6)

with col1:
    with st.container():
        st.markdown("#### What will appear on **Slide 4:**")
        st.markdown(f'''
- **Proposed Influencers:** {metrics.get('Influencers','')}
- **Proposed Engagements:** {metrics.get('Engagements','')}
- **Proposed Impressions:** {metrics.get('Impressions','')}
- **Influencer Count:** {influencer_count}
- **Diversity Rate:** {diversity_value}
- **Social Posts & Stories:** {social_posts_value}
- **Engagement Rate:** {engagement_rate_value}
- **Engagements:** {engagements_value} ({engagements_increase} increase)
- **Impressions:** {impressions_value} ({impressions_increase} increase)


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


with col6:
     with st.container():
          st.markdown("#### What will appear on **Slide 13:**")
          st.markdown(f'''
- **C2C Transfers:** {c2c_transfer}
- **C2C Value:** {c2c_value}
                      
 ''' )
          st.caption("These values will be automatically inserted into Slide 13 of your recap deck.")
        


st.header("Enter Title Slide Info")

cols = st.columns(3)  # 3 columns for 3 slides

with cols[0]:
    st.subheader("Slide 1")
    slide_1_date = st.text_input("Date", value="January 1, 2025", key="slide1date")
    slide_1_hashtag = st.text_input("Hashtag", value="#CampaignHashtag", key="slide1hashtag")

with cols[1]:
    st.subheader("Slide 2")
    slide_2_date = st.text_input("Date", value="January 1, 2025", key="slide2date")
    slide_2_hashtag = st.text_input("Hashtag", value="#CampaignHashtag", key="slide2hashtag")

with cols[2]:
    st.subheader("Slide 3")
    slide_3_date = st.text_input("Date", value="January 1, 2025", key="slide3date")
    slide_3_hashtag = st.text_input("Hashtag", value="#CampaignHashtag", key="slide3hashtag")

influencer_slide_6 = st.text_input("Enter the Influencer Handle for Slide 6", value="@influencerhandle")
slide_9_text = st.text_input("Enter the Line for the Top Text of Slide 9", value="hey this is a placeholder")
slide_13_text = st.text_input("Enter the date for Slide 13", value="12/1/26")
slide_15_text = st.text_input("Enter the question for Slide 15", value="On a scale of 1 to 10...")
slide_16_text = st.text_input("Enter the question for Slide 16", value="What were your favorite parts...")



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



influencer_boxestwo = [
    
    {"label": "Influencer Box 1 (leftmost)", "key": "box2_1"},
    {"label": "Influencer Box 2", "key": "box2_2"},
    {"label": "Influencer Box 3", "key": "box2_3"},
    {"label": "Influencer Box 4 (rightmost)", "key": "box2_4"},


]

st.markdown("### Enter Influencer Data for Each Box (for Slide 5)")

# Show headers
headers = ["Influencer", "Handle", "Reach", "City", "State", "Verbatim"]
header_cols = st.columns(len(headers))
for i, header in enumerate(headers):
    header_cols[i].markdown(f"**{header}**")

influencer_inputs = {}

for box in influencer_boxes:
    cols = st.columns(len(headers))
    cols[0].markdown(box['label'])
    handle = cols[1].text_input("", key=f"{box['key']}_handle")
    reach = cols[2].text_input("", key=f"{box['key']}_reach")
    city = cols[3].text_input("", key=f"{box['key']}_city")
    state = cols[4].text_input("", key=f"{box['key']}_state")
    verbatim = cols[5].text_input("", key=f"{box['key']}_verbatim")
    influencer_inputs[box["textbox"]] = {
        "influencerhandle": handle,
        "##": reach,
        "City": city,
        "State": state,
        "Verbatim": verbatim,
    }


st.markdown("### Enter Influencer Data for Slide 7")

headers7 = [
    "Type",
    "Influencer Handle",
    "Likes",
    "Comments",
    "Views",
    "Social Reach",
    "Engagements",
    "Impressions",
]
header_cols7 = st.columns(len(headers7))
for i, header in enumerate(headers7):
    header_cols7[i].markdown(f"**{header}**")

# Organic row
row1 = st.columns(len(headers7))
row1[0].markdown("Organic")
influencer_slide_7_left = row1[1].text_input("", key="slide_7_left_handle", value="@influencerhandle")
slide_7_likestext = row1[2].text_input("", key="slide_7_likes", value="1")
slide_7_commentstext = row1[3].text_input("", key="slide_7_comments", value="2")
slide_7_viewstext = row1[4].text_input("", key="slide_7_views", value="3")
slide_7_reachtext = row1[5].text_input("", key="slide_7_reach", value="4")
row1[6].markdown("-")  # Engagements N/A for Organic
row1[7].markdown("-")  # Impressions N/A for Organic

# Paid row
row2 = st.columns(len(headers7))
row2[0].markdown("Paid")
influencer_slide_7_right = row2[1].text_input("", key="slide_7_right_handle", value="@influencerhandle")
row2[2].markdown("-")  # Likes N/A for Paid
row2[3].markdown("-")  # Comments N/A for Paid
row2[4].markdown("-")  # Views N/A for Paid
row2[5].markdown("-")  # Social Reach N/A for Paid
slide_7_engage = row2[6].text_input("", key="slide_7_engage", value="5")
slide_7_impress = row2[7].text_input("", key="slide_7_impress", value="6")



st.markdown("### Enter Influencer Data for Each Box (for Slide 8)")

headers8 = ["Influencer", "Handle", "# Likes", "# Comments", "# Views", "# Social Reach"]
header_cols8 = st.columns(len(headers8))
for i, header in enumerate(headers8):
    header_cols8[i].markdown(f"**{header}**")

influencer_boxestwo_inputs = []

for box in influencer_boxestwo:
    cols = st.columns(len(headers8))
    cols[0].markdown(box['label'])
    handle = cols[1].text_input("", key=f"{box['key']}_handle")
    likes = cols[2].text_input("", key=f"{box['key']}_likes")
    comments = cols[3].text_input("", key=f"{box['key']}_comments")
    views = cols[4].text_input("", key=f"{box['key']}_views")
    reach = cols[5].text_input("", key=f"{box['key']}_reach")
    influencer_boxestwo_inputs.append({
        "influencerhandle": handle,
        "# Likes": likes,
        "# Comments": comments,
        "# Views": views,
        "# Social Reach": reach,
    })

# Save in text_inputs for backend use




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
    "slide_13": slide_13_text,
    "slide_15": slide_15_text,
    "slide_16": slide_16_text,
    "influencer_boxes": influencer_inputs,
    "influencer_boxestwo": influencer_boxestwo_inputs,
    "slide_1_d": slide_1_date,
    "slide_1_htg": slide_1_hashtag,
    "slide_2_d": slide_2_date,
    "slide_2_htg": slide_2_hashtag,
    "slide_3_d": slide_3_date,
    "slide_3_htg": slide_3_hashtag




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
    "slide_13": slide_13_text,
    "slide_15": slide_15_text,
    "slide_16": slide_16_text,
    "influencer_boxes": influencer_inputs,
    "influencer_boxestwo": influencer_boxestwo_inputs,
    "slide_1_d": slide_1_date,
    "slide_1_htg": slide_1_hashtag,
    "slide_2_d": slide_2_date,
    "slide_2_htg": slide_2_hashtag,
    "slide_3_d": slide_3_date,
    "slide_3_htg": slide_3_hashtag


    }     



    from datetime import datetime  # Make sure this is imported!
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"recap_deck_output_{timestamp}.pptx"

    populate_pptx_from_excel(df, pptx_template_path, output_path, images=images, text_inputs=text_inputs)

    with open(output_path, "rb") as f:
        st.success("✅ Your recap deck is ready!")
        st.download_button("⬇️ Download PowerPoint", data=f, file_name=f"recap_deck_{timestamp}.pptx",
                           mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")