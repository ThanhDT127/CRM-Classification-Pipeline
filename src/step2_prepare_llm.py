"""
CRM Classification Pipeline - Step 2: Prepare LLM Input
=========================================================
Read CRM_classified.xlsx, identify rows/columns still missing labels,
and export llm_input.json in the v5 prompt format for Gemini.
"""

import json
import pandas as pd
from pathlib import Path
from typing import Any, Dict, List

from config import (
    PATH_OUTPUT_REGEX,
    LLM_INPUT_JSON,
    INPUT_TEXT_COLUMNS,
    LLM_TARGET_COLS,
    COL_CURRENT_STATUS,
    load_keywords,
    build_keyword_index,
    build_canonical_maps,
    COL_BRANDS,
)


def _is_missing(v: Any) -> bool:
    """Check if a cell value is blank or == 'mơ hồ'."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "mơ hồ"


def _serialize_val(v: Any) -> Any:
    """Safely convert a value for JSON serialization."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    s = str(v).strip()
    return s if s else None


def main():
    # 1. Load classified DataFrame
    print(f"Loading {PATH_OUTPUT_REGEX.name} ...")
    df = pd.read_excel(PATH_OUTPUT_REGEX)
    print(f"  Rows: {len(df)}")

    # 2. Load keywords to get allowed labels per column
    from config import PATH_KW_JSON
    kw = json.loads(PATH_KW_JSON.read_text(encoding="utf-8"))
    kw_index, brands = build_keyword_index(kw)
    canonical_by_col, canonical_lower_map = build_canonical_maps(kw_index)

    # Add brands as allowed values for COL_BRANDS
    if brands:
        canonical_by_col[COL_BRANDS] = brands + ["Hãng cạnh tranh"]

    # 3. Build items in v5 prompt format
    items: List[Dict[str, Any]] = []
    skipped = 0

    # Column mapping to short keys (v5 prompt format)
    # R = "Ghi chú/Nội dung làm việc" (first text col, also = "Tình hình hiện tại")
    # S = implied: progress/status info (same source)
    # T = main content (same merged source)
    # U = "Kế hoạch lần tới" (second text col)
    # V = "Thông tin KH/ Ý kiến KH" (third text col)

    for _, row in df.iterrows():
        row_idx = int(row.get("row_idx", 0))

        # Check if any text exists
        has_text = False
        for col in INPUT_TEXT_COLUMNS:
            v = row.get(col)
            if pd.notna(v) and str(v).strip():
                has_text = True
                break

        if not has_text:
            skipped += 1
            continue

        # Existing labels (locked)
        locked_labels: Dict[str, str] = {}
        for col in LLM_TARGET_COLS:
            v = row.get(col)
            if not _is_missing(v):
                locked_labels[col] = str(v).strip()

        # Missing columns
        missing_cols: List[str] = []
        for col in LLM_TARGET_COLS:
            if _is_missing(row.get(col)):
                missing_cols.append(col)

        if not missing_cols:
            continue

        # Build allowed map: only for missing_cols
        allowed: Dict[str, List[str]] = {}
        for col in missing_cols:
            if col in canonical_by_col:
                allowed[col] = canonical_by_col[col]

        # Build item in v5 format
        item: Dict[str, Any] = {
            "row_idx": row_idx,
            "texts": {},
        }

        # Add text columns
        for col in INPUT_TEXT_COLUMNS:
            v = row.get(col)
            if pd.notna(v) and str(v).strip():
                item["texts"][col] = str(v).strip()

        if locked_labels:
            item["existing"] = locked_labels
        item["missing_cols"] = missing_cols
        item["allowed"] = allowed

        items.append(item)

    # 4. Export
    LLM_INPUT_JSON.write_text(
        json.dumps(items, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"\n✓ LLM input prepared!")
    print(f"  Rows needing LLM fills: {len(items)}")
    print(f"  Rows with no text (skipped): {skipped}")
    print(f"  Rows fully classified: {len(df) - len(items) - skipped}")
    print(f"  Output: {LLM_INPUT_JSON.name}")

    if items:
        avg = sum(len(it["missing_cols"]) for it in items) / len(items)
        print(f"  Avg missing cols per row: {avg:.1f}")


if __name__ == "__main__":
    main()
