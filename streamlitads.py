import os
import streamlit as st
import pandas as pd
from app import load_dataframe, populate_pptx_from_excel
import aspose.slides as slides

# ─────────────────────────────────────────────────────────────────────────────
# Helper for slide preview (requires aspose.slides)
# ─────────────────────────────────────────────────────────────────────────────


def save_slide1_as_png(pptx_path, out_path):
    with slides.Presentation(pptx_path) as presentation:
        slide = presentation.slides[0]
        img = slide.get_thumbnail(1280, 720)  # 2x scale for better quality
        img.save(out_path, "PNG")

# ─────────────────────────────────────────────────────────────────────────────
# Page config & branding
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SOAPBOX Recap Deck App",
    page_icon="✎",
    layout="wide",
)

if os.path.exists("logo.png"):
    st.image("logo.png", width=180)

st.title("Recap Deck Editor")
st.markdown(
    "Upload your Excel, preview, edit fields (coming soon), and export your influencer recap deck."
)

# ─────────────────────────────────────────────────────────────────────────────
# Inputs: File, Client, Date
# ─────────────────────────────────────────────────────────────────────────────
st.header("Step 1: Upload Data File")
uploaded = st.file_uploader("Upload Excel or CSV", type=["xlsx", "csv"])
if not uploaded:
    st.info("Please upload your Excel/CSV to see and generate your recap deck.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# Load & Preview Data
# ─────────────────────────────────────────────────────────────────────────────
df = load_dataframe(uploaded)
st.subheader("Preview: First 10 Rows of Data")
st.dataframe(df.head(10), height=250)

# ─────────────────────────────────────────────────────────────────────────────
# Populate slide 1 and show preview
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.header("Live Preview: Slide 1 (Proposed Program Details)")

pptx_template_path = "template.pptx"
temp_pptx = "temp_slide1_populated.pptx"
slide_img = "slide1_preview.png"

# Populate only slide 1 with "Proposed Program Details"
# (mapping_config is not needed for slide 1 now)
_ = populate_pptx_from_excel(df, pptx_template_path, temp_pptx, mapping_config={})

# Save slide 1 as PNG
save_slide1_as_png(temp_pptx, slide_img)

if os.path.exists(slide_img):
    st.image(slide_img, use_column_width=True, caption="Slide 1 Preview")
else:
    st.warning("Could not generate slide preview image.")

st.markdown("---")

# ─────────────────────────────────────────────────────────────────────────────
# Export PowerPoint Deck
# ─────────────────────────────────────────────────────────────────────────────
st.header("Step 2: Export Recap Deck PowerPoint")

if st.button("Generate PowerPoint Recap Deck"):
    output_path = "recap_deck_output.pptx"
    pptx_file = populate_pptx_from_excel(df, pptx_template_path, output_path, mapping_config={})
    with open(pptx_file, "rb") as f:
        st.success("✅ PowerPoint deck is ready!")
        st.download_button(
            "⬇️ Download PowerPoint",
            data=f,
            file_name="recap_deck.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
