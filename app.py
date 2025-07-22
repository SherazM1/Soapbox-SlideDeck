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
def populate_pptx_from_excel(excel_df, pptx_template_path, output_path):
    prs = Presentation(pptx_template_path)

    # ---------- Extract Proposed Metrics Block (TextBox 2) ----------
    try:
        metrics = extract_proposed_metrics_anywhere(excel_df)
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    # ---------- Extract all other values needed for TextBox 15 ----------
    print("COLUMNS:", list(excel_df.columns))

    # Social Posts & Stories
    social_posts_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Number of Posts With Stories":
                social_posts_value = row["Unnamed: 11"]
                break


    engagements_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Engagements":
                engagements_value = row["Unnamed: 11"]
                break

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
# Truncate to two decimal places without rounding:
            dot_idx = engagement_rate_value.find(".")
            if dot_idx != -1:
                engagement_rate_value = engagement_rate_value[:dot_idx + 3]  # Keep 2 decimal digits after dot
                break

    # ---------- NEW: Engagements & Impressions % INCREASE ----------
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

    organic_likes = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Likes":
                organic_likes = row["Unnamed: 11"]
                break
    
    organic_comments = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Comments":
                organic_comments = row["Unnamed: 11"]
                break
    
    organic_shares = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Shares":
                organic_shares = row["Unnamed: 11"]
                break

    organic_saves = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Saves":
                organic_saves = row["Unnamed: 11"]
                break


    # ---------- Fill TextBox 2 (Proposed Metrics) ----------
    slide = prs.slides[3]  # Slide 4 (0-indexed)
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

    # ---------- Fill TextBox 15 (Program Overview) ----------
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
        # Replace ONLY the first '#' (main value)
                        if "#" in run.text and not main_done and "% increase" not in run.text:
                            run.text = run.text.replace("#", str(engagements_value), 1)
                            main_done = True
        # Replace ONLY the '#' before '% increase' (percent value)
                        if "#% increase" in run.text and not percent_done:
                            run.text = run.text.replace("#", str(engagements_increase), 1)
                            percent_done = True
                # Impressions (main)
                elif "Impressions" in text and "#% increase" in text:
                    main_done = False
                    percent_done = False
                    for run in para.runs:
        # Replace ONLY the first '#' (main value)
                        if "#" in run.text and not main_done and "% increase" not in run.text:
                            run.text = run.text.replace("#", str(impressions_value), 1)
                            main_done = True
        # Replace ONLY the '#' before '% increase' (percent value)
                        if "#% increase" in run.text and not percent_done:
                            run.text = run.text.replace("#", str(impressions_increase), 1)
                            percent_done = True
    
    slide = prs.slides[8]
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == "TextBox 19":
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "10" in run.text and "K" in run.text:
                        run.text = run.text.replace("10", str(organic_likes))
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "20" in run.text and "K" in run.text:
                        run.text = run.text.replace("20", str(organic_comments))
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "30" in run.text and "K" in run.text:
                        run.text = run.text.replace("30", str(organic_shares))
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if "40" in run.text and "K" in run.text:
                        run.text = run.text.replace("40", str(organic_saves))

                        
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
