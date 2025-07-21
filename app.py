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
# PowerPoint Deck Generation
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
    prs = Presentation(pptx_template_path)

    # ---------- Extract Proposed Metrics Block (TextBox 2) ----------
    try:
        metrics = extract_proposed_metrics_anywhere(excel_df)
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    # ---------- Extract all other values needed for TextBox 15 ----------
    print("COLUMNS:", list(excel_df.columns))

    # Extraction logic for Social Posts & Stories
    social_posts_value = ""
    if "Organic & Total" in excel_df.columns and "Unnamed: 11" in excel_df.columns:
        for idx, row in excel_df.iterrows():
            if str(row["Organic & Total"]).strip() == "Total Number of Posts With Stories":
                social_posts_value = row["Unnamed: 11"]
                break

    # Engagement Rate
    engagement_rate_value = ""
    for idx, row in excel_df.iterrows():
        if str(row.iloc[0]).strip().lower() == "program er":
            engagement_rate_value = row.iloc[1]
            break

    # ---------- Updated: % INCREASE Extraction ----------
    # pick the right column names
    prop_col = "Unnamed: 14"  # where “Engagements” / “Impressions” lives
    perc_col = "Unnamed: 15"  # where the raw decimal % lives

    # ENGAGEMENTS % INCREASE
    engagements_increase = ""
    if prop_col in excel_df.columns and perc_col in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row[prop_col]).strip().lower() == "engagements":
                raw = row[perc_col]
                # only format if it's not NaN
                if pd.notna(raw):
                    num = pd.to_numeric(raw, errors="coerce")
                    if pd.notna(num):
                        engagements_increase = f"{num * 100:.1f}%"
                break

    # IMPRESSIONS % INCREASE
    impressions_increase = ""
    if prop_col in excel_df.columns and perc_col in excel_df.columns:
        for _, row in excel_df.iterrows():
            if str(row[prop_col]).strip().lower() == "impressions":
                raw = row[perc_col]
                if pd.notna(raw):
                    num = pd.to_numeric(raw, errors="coerce")
                    if pd.notna(num):
                        impressions_increase = f"{num * 100:.1f}%"
                break

    print("Engagements % increase:", engagements_increase)   # → “255.6%”
    print("Impressions % increase:", impressions_increase)   # → “429.4%”

    # ---------- Fill TextBox 2 (Proposed Metrics) ----------
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

    # ---------- Fill TextBox 15 (Program Overview, same slide) ----------
    new_box_name = "TextBox 15"
    found = False
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
                            run.text = run.text.replace("#", str(metrics.get("Engagements", "")))
                # Engagements Percentage Increase
                elif "Engagements" in full_text and "% increase" in full_text:
                    for run in paragraph.runs:
                        if "% increase" in run.text:
                            run.text = run.text.replace("% increase", f"{engagements_increase} increase")
                # Impressions (main number)
                elif "Impressions" in full_text and "#" in full_text and "% increase" not in full_text:
                    for run in paragraph.runs:
                        if "#" in run.text:
                            run.text = run.text.replace("#", str(metrics.get("Impressions", "")))
                # Impressions Percentage Increase
                elif "Impressions" in full_text and "% increase" in full_text:
                    for run in paragraph.runs:
                        if "% increase" in run.text:
                            run.text = run.text.replace("% increase", f"{impressions_increase} increase")
            found = True
            break

    if not found:
        print(f"[ERROR] Could not find shape named '{new_box_name}' on Slide 4.")
        for shape in slide.shapes:
            if shape.has_text_frame:
                print(f"- {shape.name}")

    prs.save(output_path)

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
