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

def populate_pptx_from_excel(excel_df, pptx_template_path, output_path, mapping_config, user_inputs=None):
    """
    Populate a PowerPoint template from Excel data using a mapping config.
    - excel_df: DataFrame with your data.
    - pptx_template_path: Path to the PowerPoint template to use.
    - output_path: Where to save the populated PPTX.
    - mapping_config: Dict describing how to map Excel columns to slide fields.
    - user_inputs: Optional dict for extra custom fields/images per slide.
    """
    prs = Presentation(pptx_template_path)

    # ---- Slide 1: Proposed Program Details ----
    try:
        metrics = extract_proposed_metrics_anywhere(excel_df)
    except Exception as e:
        metrics = {"Impressions": "", "Engagements": "", "Influencers": ""}
        print(f"Warning: Could not extract Proposed Metrics from Excel: {e}")

    # All three bullet points are in a single text box named 'Proposed_Program_Details'
    proposed_box_name = "Proposed_Program_Details"  # Name your text box this in PowerPoint

    bullet_lines = [
        f"• Proposed Influencers: {metrics.get('Influencers', '')}",
        f"• Proposed Engagements: {metrics.get('Engagements', '')}",
        f"• Proposed Impressions: {metrics.get('Impressions', '')}"
    ]
    text_to_fill = "\n".join(bullet_lines)

    slide = prs.slides[0]  # First slide
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == proposed_box_name:
            shape.text = text_to_fill
            break  # Found and filled, exit loop

    # ---- Future: add logic for additional slides below as needed ----

    # Optionally, apply user_inputs logic for custom fields/images here

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
