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


# PowerPoint Deck Generation (to be implemented per slide mapping)
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
    metrics = dict(zip(names, values))
    return metrics

def populate_pptx_from_excel(excel_df, pptx_template_path, output_path):
    from pptx import Presentation

    prs = Presentation(pptx_template_path)

    # ---------- ORIGINAL LOGIC: TextBox 2 (Proposed Metrics) ----------
    try:
        metrics = extract_proposed_metrics_anywhere(excel_df)
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    bullet_box_name = "TextBox 2"
    slide = prs.slides[3]  # Slide 4 (0-indexed)
    found = False

    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == bullet_box_name:
            for paragraph in shape.text_frame.paragraphs:
                full_text = paragraph.text.strip()
                if "Proposed Influencers" in full_text and "#" in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Influencers", "")))
                elif "Proposed Engagements" in full_text and "#" in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Engagements", "")))
                elif "Proposed Impressions" in full_text and "#" in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Impressions", "")))
            found = True
            break

    if not found:
        print(f"[ERROR] Could not find shape named '{bullet_box_name}' on Slide 4.")
        for shape in slide.shapes:
            if shape.has_text_frame:
                print(f"- {shape.name}")

    # ========== NEW LOGIC: TextBox 15 (Slide 4, index 3) ==========
    new_box_name = "TextBox 15"
    found = False

    # --- Social Posts & Stories ---
    social_posts_value = ""
    if "Organic & Total" in excel_df.columns:
        for idx, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip().lower() == "total number of posts with stories":
                social_posts_value = row.iloc[1]
                break

    # --- Engagement Rate ---
    engagement_rate_value = ""
    for idx, row in excel_df.iterrows():
        if str(row.iloc[0]).strip().lower() == "program er":
            raw_val = row.iloc[1]
            if isinstance(raw_val, str) and raw_val.startswith("#"):
                engagement_rate_value = ""
            else:
                engagement_rate_value = raw_val
            break

    # --- Engagements Value ---
    engagements_value = ""
    if "Organic & Total" in excel_df.columns:
        for idx, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip().lower() == "total engagements":
                engagements_value = row.iloc[1]
                break

    # --- Engagements Percentage Increase ---
    engagements_increase = ""
    if "Proposed Metrics" in excel_df.columns and "Percentage Increase" in excel_df.columns:
        for idx, row in excel_df.iterrows():
            if str(row["Proposed Metrics"]).strip().lower() == "engagements":
                engagements_increase = row["Percentage Increase"]
                break

    # --- Impressions Value (w/ fallback) ---
    impressions_value = ""
    for idx, row in excel_df.iterrows():
        first_col_val = str(row.iloc[0]).strip().lower()
        if first_col_val == "total impressions":
            impressions_value = row.iloc[1]
            break
        elif first_col_val == "total":
            impressions_value = row.iloc[1]  # fallback

    # --- Impressions Percentage Increase ---
    impressions_increase = ""
    if "Proposed Metrics" in excel_df.columns and "Percentage Increase" in excel_df.columns:
        for idx, row in excel_df.iterrows():
            if str(row["Proposed Metrics"]).strip().lower() == "impressions":
                impressions_increase = row["Percentage Increase"]
                break

    # --- Update TextBox 15 bullets ---
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == new_box_name:
            for paragraph in shape.text_frame.paragraphs:
                full_text = paragraph.text.strip()
                # Social Posts & Stories
                if "Social Posts & Stories" in full_text and "#" in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(social_posts_value))
                # Engagement Rate
                elif "Engagement Rate" in full_text and "#" in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(engagement_rate_value))
                # Engagements (main number)
                elif "Engagements" in full_text and "#" in full_text and "% increase" not in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(engagements_value))
                # Engagements Percentage Increase
                elif "Engagements" in full_text and "% increase" in full_text:
                    for run in paragraph.runs:
                        if "% increase" in run.text:
                            run.text = run.text.replace("% increase", f"{engagements_increase}% increase")
                # Impressions (main number)
                elif "Impressions" in full_text and "#" in full_text and "% increase" not in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(impressions_value))
                # Impressions Percentage Increase
                elif "Impressions" in full_text and "% increase" in full_text:
                    for run in paragraph.runs:
                        if "% increase" in run.text:
                            run.text = run.text.replace("% increase", f"{impressions_increase}% increase")
            found = True
            break

    if not found:
        print(f"[ERROR] Could not find shape named '{new_box_name}' on Slide 4.")
        for shape in slide.shapes:
            if shape.has_text_frame:
                print(f"- {shape.name}")

    prs.save(output_path)
    return output_path
# ─────────────────────────────────────────────────────────────────────────────
# CLI Entrypoint (optional, for testing automation)
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
