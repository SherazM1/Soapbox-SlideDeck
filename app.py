# app.py

import os
import sys
import json
import pandas as pd
import io
from io import BytesIO
from pptx import Presentation


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def resource_path(rel_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, rel_path)

# ─────────────────────────────────────────────────────────────────────────────
# Constants & Persistence
# ─────────────────────────────────────────────────────────────────────────────
BATCHES_PATH = "dashboards/batches.json"

def load_batches(path: str = BATCHES_PATH) -> list:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []

def save_batches(batches: list, path: str = BATCHES_PATH) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(batches, f, indent=2, default=str)

# ─────────────────────────────────────────────────────────────────────────────
# Data Loading
# ─────────────────────────────────────────────────────────────────────────────
def load_dataframe(src) -> pd.DataFrame:
    # Handles file upload object or file path
    if hasattr(src, "read") and hasattr(src, "name"):
        data = src.getvalue()
        ext  = os.path.splitext(src.name)[1].lower()
        if ext == ".csv":
            return pd.read_csv(io.BytesIO(data), encoding="utf-8", engine="python", on_bad_lines="skip")
        elif ext in (".xls", ".xlsx"):
            return pd.read_excel(io.BytesIO(data))
        else:
            raise ValueError(f"Unsupported file type: {ext}")
    ext = os.path.splitext(src)[1].lower()
    if ext == ".csv":
        return pd.read_csv(src)
    elif ext in (".xls", ".xlsx"):
        return pd.read_excel(src)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

# ─────────────────────────────────────────────────────────────────────────────
# Proposed Metrics Extraction
# ─────────────────────────────────────────────────────────────────────────────
def extract_proposed_metrics_anywhere(df):
    """
    Find 'Proposed Metrics' anywhere in the sheet, then extract the next 3 rows
    for 'Impressions', 'Engagements', 'Influencers' in the same column.
    Returns a dict: {'Impressions': ..., 'Engagements': ..., 'Influencers': ...}
    """
    found = False
    for col in df.columns:
        col_series = df[col].astype(str).str.strip().str.lower()
        idxs = col_series[col_series == "proposed metrics"].index.tolist()
        if idxs:
            found = True
            idx = idxs[0]
            col_num = df.columns.get_loc(col)
            names = []
            values = []
            for offset in range(1, 4):
                name = str(df.iloc[idx+offset, col_num]).strip()
                value = df.iloc[idx+offset, col_num+1]
                names.append(name)
                values.append(value)
            break
    if not found:
        raise ValueError("'Proposed Metrics' not found in any column.")
    return dict(zip(names, values))

# ─────────────────────────────────────────────────────────────────────────────
def format_compact_number(n):
    try:
        n = float(n)
        if n >= 1_000_000:
            return f"{n / 1_000_000:.1f}MM"
        elif n >= 1_000:
            return f"{n / 1_000:.1f}K"
        else:
            return str(int(n))
    except:
        return str(n)

# PowerPoint Deck Generation
# ─────────────────────────────────────────────────────────────────────────────
def populate_pptx_from_excel(excel_df, pptx_template_path, output_path, images=None, text_inputs=None):
    prs = Presentation(pptx_template_path)
    handle_slide_6 = text_inputs.get("slide_6", "@default")


    # ---------- Extract Proposed Metrics Block (TextBox 2) ----------
    try:
        metrics = extract_proposed_metrics_anywhere(excel_df)
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    print("COLUMNS:", list(excel_df.columns))

    # Slide 6
    if images and "slide_6" in images and images["slide_6"] is not None:
        slide = prs.slides[5]
        img_bytes = images["slide_6"].read()
        temp_img_path = "temp_slide_6_img.jpg"
        with open(temp_img_path, "wb") as f:
            f.write(img_bytes)
        for shape in slide.shapes:
            if shape.name == "Picture 2":
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(temp_img_path, left, top, width=width, height=height)
                break

    # Slide 7
    slide = prs.slides[6]
    if images and "slide_7_left" in images and images["slide_7_left"] is not None:
        img_bytes = images["slide_7_left"].read()
        temp_img_path = "temp_slide_7_left.jpg"
        with open(temp_img_path, "wb") as f:
            f.write(img_bytes)
        for shape in slide.shapes:
            if shape.name == "Picture 3":
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(temp_img_path, left, top, width=width, height=height)
                break

    if images and "slide_7_right" in images and images["slide_7_right"] is not None:
        img_bytes = images["slide_7_right"].read()
        temp_img_path = "temp_slide_7_right.jpg"
        with open(temp_img_path, "wb") as f:
            f.write(img_bytes)
        for shape in slide.shapes:
            if shape.name == "Picture 2":
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                slide.shapes._spTree.remove(shape._element)
                slide.shapes.add_picture(temp_img_path, left, top, width=width, height=height)
                break

    # Slide 8 (FOUR images)
    slide = prs.slides[7]
    slide8_img_configs = [
        ("slide_8_first",  "Picture 12", "temp_slide_8_first.jpg"),
        ("slide_8_second", "Picture 13", "temp_slide_8_second.jpg"),
        ("slide_8_third",  "Picture 16", "temp_slide_8_third.jpg"),
        ("slide_8_fourth", "Picture 17", "temp_slide_8_fourth.jpg"),
    ]
    for img_key, shape_name, temp_path in slide8_img_configs:
        if images and img_key in images and images[img_key] is not None:
            img_bytes = images[img_key].read()
            with open(temp_path, "wb") as f:
                f.write(img_bytes)
            for shape in slide.shapes:
                if shape.name == shape_name:
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    slide.shapes._spTree.remove(shape._element)
                    slide.shapes.add_picture(temp_path, left, top, width=width, height=height)
                    break


    #slide 11 (Four images)
    slide = prs.slides[10]
    slide11_img_configs = [
        ("slide_11_first", "Picture 26", "temp_slide_11_first.jpg"),
        ("slide_11_second", "Picture 15", "temp_slide_11_second.jpg"),
        ("slide_11_third", "Picture 27", "temp_slide_11_third.jpg"),
        ("slide_11_fourth", "Picture 31", "temp_slide_11_fourth.jpg")
    ]
    for img_key, shape_name, temp_path in slide11_img_configs:
        if images and img_key in images and images[img_key] is not None:
            img_bytes = images[img_key].read()
            with open(temp_path, "wb") as f:
                f.write(img_bytes)
            for shape in slide.shapes:
                if shape.name == shape_name:
                    left, top, width, height = shape.left, shape.top, shape.width, shape.height
                    slide.shapes._spTree.remove(shape._element)
                    slide.shapes.add_picture(temp_path, left, top, width=width, height=height)
                    break

    # Social Posts & Stories
    social_posts_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Number of Posts With Stories":
                social_posts_value = row["Unnamed: 11"]
                break



    organic_views_impressions = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Organic (Views)":
                organic_views_impressions = row["Unnamed: 11"]
                break



    organic_reach_impressions = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Organic (Reach)":
                organic_reach_impressions = row["Unnamed: 11"]
                break

    impressions_paid = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Paid":
                impressions_paid = row["Unnamed: 11"]
                break



    # Engagements
    engagements_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Engagements":
                engagements_value = row["Unnamed: 11"]
                break

    # Impressions
    impressions_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            cell_value = str(row["Organic & Total"]).strip()
            if cell_value in ("Total", "Total Impressions"):
                impressions_value = row["Unnamed: 11"]
                break

    # Engagement Rate
    engagement_rate_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Program ER":
                engagement_rate_value = row["Unnamed: 11"]
                engagement_rate_value = float(engagement_rate_value) * 100
                engagement_rate_value = str(engagement_rate_value)
                if engagement_rate_value.startswith("0."):
                    engagement_rate_value = engagement_rate_value[1:]
                dot_idx = engagement_rate_value.find(".")
                if dot_idx != -1:
                    engagement_rate_value = engagement_rate_value[:dot_idx + 3]
                break

    # Engagements & Impressions % INCREASE
    engagements_increase = ""
    impressions_increase = ""
    try:
        engagement_val = excel_df.at[5, "Unnamed: 15"]
        impression_val = excel_df.at[4, "Unnamed: 15"]
        if pd.notna(engagement_val):
            engagements_increase = f"{float(engagement_val) * 100:.1f}%"
        if pd.notna(impression_val):
            impressions_increase = f"{float(impression_val) * 100:.1f}%"
    except Exception as e:
        print("could not find value")

    print("Engagements % increase:", engagements_increase)
    print("Impressions % increase:", impressions_increase)

    # Organic Likes
    organic_likes = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Likes":
                organic_likes = row["Unnamed: 11"]
                break

    # Organic Comments
    organic_comments = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Comments":
                organic_comments = row["Unnamed: 11"]
                break

    # Organic Shares
    organic_shares = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Shares":
                organic_shares = row["Unnamed: 11"]
                break

    # Organic Saves
    organic_saves = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Saves":
                organic_saves = row["Unnamed: 11"]
                break
    
    paid_likes = ""
    if "Unnamed: 14" in excel_df.columns and "Dates" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Reactions":
                paid_likes = row["Dates"]
                break

    paid_comments = ""
    if "Unnamed: 14" in excel_df.columns and "Dates" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Comments":
                paid_comments = row["Dates"]
                break
    
    paid_shares = ""
    if "Unnamed: 14" in excel_df.columns and "Dates" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Shares":
                paid_shares = row["Dates"]
                break

    paid_saves = ""
    if "Unnamed: 14" in excel_df.columns and "Dates" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "Saves":
                paid_saves = row["Dates"]
                break 
    
    paid_threesec = ""
    if "Unnamed: 14" in excel_df.columns and "Dates" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Unnamed: 14"]).strip() == "3 sec vid views":
                paid_threesec = row["Dates"]
                break
    
    total_post_engagements = (
    int(organic_likes) + int(organic_comments) + int(organic_shares) + int(organic_saves)
    + int(paid_likes) + int(paid_comments) + int(paid_shares) + int(paid_saves) + int(paid_threesec)
)
    
    story_engagements = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Story Engagements":
                story_engagements = row["Unnamed: 11"]
                break

    
    # Fill TextBox 2 (Proposed Metrics)
    slide = prs.slides[3]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 2":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if "Proposed Influencers" in text:
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Influencers", "")))
                elif "Proposed Engagements" in text:
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Engagements", "")))
                elif "Proposed Impressions" in text:
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Impressions", "")))
            break

    # Fill TextBox 15 (Program Overview)
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 15":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                # Social Posts & Stories
                if "Social Posts & Stories" in text:
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(social_posts_value))
                # Engagement Rate
                elif "Engagement Rate" in text:
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(engagement_rate_value))
                # Engagements (main)
                elif "Engagements" in text and "#% increase" in text:
                    main_done = False
                    percent_done = False
                    for run in para.runs:
                        if "#" in run.text and not main_done and "% increase" not in run.text:
                            run.text = run.text.replace("#", str(engagements_value), 1)
                            main_done = True
                        if "#% increase" in run.text and not percent_done:
                            run.text = run.text.replace("#", str(engagements_increase), 1)
                            percent_done = True
                # Impressions (main)
                elif "Impressions" in text and "#% increase" in text:
                    main_done = False
                    percent_done = False
                    for run in para.runs:
                        if "#" in run.text and not main_done and "% increase" not in run.text:
                            run.text = run.text.replace("#", str(impressions_value), 1)
                            main_done = True
                        if "#% increase" in run.text and not percent_done:
                            run.text = run.text.replace("#", str(impressions_increase), 1)
                            percent_done = True

    # Fill TextBox 19 (Slide 9 vertical fields)
    slide = prs.slides[8]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 19":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                for run in para.runs:
                    if "10" in run.text and "K" in run.text:
                        run.text = run.text.replace("10", str(organic_likes))
                        run.text = run.text.replace("K", "")
                    if "20" in run.text and "K" in run.text:
                        run.text = run.text.replace("20", str(organic_comments))
                        run.text = run.text.replace("K", "")
                    if "30" in run.text and "K" in run.text:
                        run.text = run.text.replace("30", str(organic_shares))
                        run.text = run.text.replace("K", "")
                    if "40" in run.text and "K" in run.text:
                        run.text = run.text.replace("40", str(organic_saves))
                        run.text = run.text.replace("K", "")

# Paid – TextBox 11
    slide = prs.slides[8]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 11":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                for run in para.runs:
                    if "10" in run.text and "K" in run.text:
                        run.text = run.text.replace("10", str(paid_likes))
                        run.text = run.text.replace("K", "")
                    if "20" in run.text and "K" in run.text:
                        run.text = run.text.replace("20", str(paid_comments))
                        run.text = run.text.replace("K", "")
                    if "30" in run.text and "K" in run.text:
                        run.text = run.text.replace("30", str(paid_shares))
                        run.text = run.text.replace("K", "")
                    if "40" in run.text and "K" in run.text:
                        run.text = run.text.replace("40", str(paid_saves))
                        run.text = run.text.replace("K", "")
                    if "##" in run.text and "K" in run.text:
                        run.text = run.text.replace("##", str(paid_threesec))
                        run.text = run.text.replace("K", "")
    
    #slide 9 last box
    slide = prs.slides[8]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 34":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                for run in para.runs:
                    if "100" in run.text and "K" in run.text:
                        run.text = run.text.replace("100", str(total_post_engagements))
                        run.text = run.text.replace("K", "")
                    if "200" in run.text and "K" in run.text:
                        run.text = run.text.replace("200", str(story_engagements))
                        run.text = run.text.replace("K", "")
    


    #slide 10
    slide = prs.slides[9]
    for shape in slide.shapes:
            if shape.has_text_frame and shape.name == "TextBox 18":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(organic_reach_impressions))

            elif shape.has_text_frame and shape.name == "TextBox 19":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(impressions_paid))
            
            elif shape.has_text_frame and shape.name == "TextBox 21":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(organic_views_impressions))
           
            elif shape.has_text_frame and shape.name == "TextBox 29":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(impressions_value))
    

    #slide 11
    slide = prs.slides[10]
    for shape in slide.shapes:
            if shape.has_text_frame and shape.name == "TextBox 18":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(organic_reach_impressions))

            elif shape.has_text_frame and shape.name == "TextBox 19":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(impressions_paid))
            
            elif shape.has_text_frame and shape.name == "TextBox 21":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(organic_views_impressions))
           
            elif shape.has_text_frame and shape.name == "TextBox 29":
                for para in shape.text_frame.paragraphs:
                    text = para.text.strip()
                    for run in para.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(impressions_value))
    
    
    #slide 6 text
    slide = prs.slides[5]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 9":
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                for run in para.runs:
                    if "influencerhandle" in run.text:
                        run.text = run.text.replace("influencerhandle", handle_slide_6)

    


    prs.save(output_path)
                        
      

# ─────────────────────────────────────────────────────────────────────────────
# CLI Entrypoint (optional, for testing)
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Populate PowerPoint deck from Excel data")
    parser.add_argument("input_file", help="CSV or Excel input")
    parser.add_argument("pptx_template", help="PowerPoint template file")
    parser.add_argument("--output", default="recap_deck.pptx", help="Output PPTX file")
    args = parser.parse_args()

    df = load_dataframe(args.input_file)
    populate_pptx_from_excel(df, args.pptx_template, args.output)
    print(f"Wrote {args.output}")
