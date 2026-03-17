"""
CRM Classification Pipeline - Step 1: Regex Classification
============================================================
Read CRM_TDCTDA.xlsx, apply keyword-based regex classification,
and export CRM_classified.xlsx.
"""

import re
import sys
import pandas as pd
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import (
    PATH_INPUT,
    PATH_KW_JSON,
    PATH_OUTPUT_REGEX,
    INPUT_TEXT_COLUMNS,
    OUTPUT_COLUMNS,
    COL_CURRENT_STATUS,
    COL_CURRENT_STATUS_SOURCE,
    COL_BRANDS,
    load_keywords,
    build_keyword_index,
)


# ─── Regex helpers ──────────────────────────────────────────────────────────
def _build_pattern(kws: List[str]) -> Optional[re.Pattern]:
    """Build a compiled regex from a list of keyword strings."""
    if not kws:
        return None
    escaped = [re.escape(k) for k in kws if k]
    if not escaped:
        return None
    return re.compile("|".join(escaped), re.IGNORECASE)


def classify_row(
    txt: str,
    kw_index: Dict[str, Dict[str, List[str]]],
    brand_pattern: Optional[re.Pattern],
) -> Dict[str, Optional[str]]:
    """Given the merged text of a row, return {col: label_or_None}."""
    result: Dict[str, Optional[str]] = {}

    for col, labels_dict in kw_index.items():
        match_found = None
        for label, keywords in labels_dict.items():
            pat = _build_pattern(keywords)
            if pat and pat.search(txt):
                match_found = label
                break
        result[col] = match_found

    # Brands column: find all matching brands
    if brand_pattern:
        found = brand_pattern.findall(txt)
        if found:
            unique = []
            seen_lower = set()
            for b in found:
                if b.lower() not in seen_lower:
                    unique.append(b)
                    seen_lower.add(b.lower())
            result[COL_BRANDS] = "; ".join(unique)
        else:
            result[COL_BRANDS] = None

    return result


def main():
    # 1. Load keywords
    print(f"Loading keywords from {PATH_KW_JSON.name} ...")
    kw = load_keywords()
    kw_index, brands = build_keyword_index(kw)
    print(f"  Columns indexed: {len(kw_index)}")
    print(f"  Brands: {len(brands)}")

    brand_pattern = _build_pattern(brands) if brands else None

    # 2. Read input Excel
    print(f"\nReading {PATH_INPUT.name} ...")
    df = pd.read_excel(PATH_INPUT)
    print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")

    # 3. Add row_idx (1-based, matching notebook convention)
    df["row_idx"] = range(2, len(df) + 2)

    # 4. Merge text from source columns
    def _merge_text(row: pd.Series) -> str:
        parts = []
        for col in INPUT_TEXT_COLUMNS:
            v = row.get(col)
            if pd.notna(v):
                txt = str(v).strip()
                if txt:
                    parts.append(txt)
        return " ".join(parts)

    df["_merged_text"] = df.apply(_merge_text, axis=1)

    # 5. Initialise output columns
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # 6. COL_CURRENT_STATUS = direct copy from 'Tình trạng hiện tại'
    if COL_CURRENT_STATUS_SOURCE in df.columns:
        df[COL_CURRENT_STATUS] = df[COL_CURRENT_STATUS_SOURCE]
    else:
        print(f"  ⚠️ Source column '{COL_CURRENT_STATUS_SOURCE}' not found in input")

    # 7. Apply regex classification
    print("\nRunning regex classification ...")
    classified = 0
    for idx, row in df.iterrows():
        txt = row.get("_merged_text", "")
        if not txt:
            continue
        labels = classify_row(txt, kw_index, brand_pattern)
        for col, val in labels.items():
            if val is not None and col in df.columns:
                df.at[idx, col] = val
                classified += 1

    # 8. Drop helper column
    df.drop(columns=["_merged_text"], inplace=True)

    # 9. Export
    df.to_excel(PATH_OUTPUT_REGEX, index=False)
    print(f"\n✓ Regex classification done!")
    print(f"  Total label assignments: {classified}")
    print(f"  Output: {PATH_OUTPUT_REGEX.name}")

    # 10. Stats: how many cells are still empty per output column
    print("\n  Empty cells per classification column:")
    for col in OUTPUT_COLUMNS:
        if col in df.columns:
            n_empty = df[col].isna().sum() + (df[col].astype(str).str.strip() == "").sum()
            pct = n_empty / len(df) * 100
            print(f"    {col}: {n_empty} ({pct:.1f}%)")

    return df


if __name__ == "__main__":
    main()
