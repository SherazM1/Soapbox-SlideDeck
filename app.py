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
    prs = Presentation(pptx_template_path)

    try:
        metrics = {
            "Influencers": excel_df.get("Influencers", [""])[0],
            "Engagements": excel_df.get("Engagements", [""])[0],
            "Impressions": excel_df.get("Impressions", [""])[0]
        }
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    bullet_box_name = "TextBox 2"
    slide = prs.slides[3]
    found = False

    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == bullet_box_name:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if '#' in run.text:
                        if "Proposed Influencers" in paragraph.text:
                            run.text = run.text.replace("#", str(metrics.get("Influencers", "")))
                        elif "Proposed Engagements" in paragraph.text:
                            run.text = run.text.replace("#", str(metrics.get("Engagements", "")))
                        elif "Proposed Impressions" in paragraph.text:
                            run.text = run.text.replace("#", str(metrics.get("Impressions", "")))
            found = True
            break

    if not found:
        print(f"[ERROR] Could not find shape named '{bullet_box_name}' on Slide 4.")

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
    # mapping_config = ...  # Define your mapping per slide!
    mapping_config = {}  # Placeholder

    populate_pptx_from_excel(df, args.pptx_template, args.output, mapping_config)
    print(f"Wrote {args.output}")
