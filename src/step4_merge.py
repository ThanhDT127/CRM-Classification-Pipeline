"""
CRM Classification Pipeline - Step 4: Merge LLM Fills
======================================================
Merge llm_fills.json results back into CRM_classified.xlsx
to produce CRM_classified_with_LLM.xlsx.
"""

import json
import pandas as pd
from pathlib import Path
from typing import Any, Dict

from config import (
    PATH_OUTPUT_REGEX,
    PATH_OUTPUT_LLM,
    OUT_LLM_JSON,
    OUTPUT_COLUMNS,
    LLM_TARGET_COLS,
    COL_CURRENT_STATUS,
    INPUT_TEXT_COLUMNS,
)


def main():
    # 1. Load classified Excel
    print(f"Loading {PATH_OUTPUT_REGEX.name} ...")
    df = pd.read_excel(PATH_OUTPUT_REGEX)
    print(f"  Rows: {len(df)}")

    # 2. Load LLM fills
    if not OUT_LLM_JSON.exists():
        print(f"❌ LLM fills not found: {OUT_LLM_JSON}. Run step3 first.")
        return

    fills_list = json.loads(OUT_LLM_JSON.read_text(encoding='utf-8'))
    print(f"  LLM fill items: {len(fills_list)}")

    # Build lookup by row_idx
    fills_by_idx: Dict[int, Dict[str, Any]] = {}
    for item in fills_list:
        rid = int(item.get('row_idx', 0))
        fills = item.get('fills', {})
        if isinstance(fills, dict):
            fills_by_idx[rid] = fills

    # 3. Ensure output columns exist
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # 4. Merge fills
    merged_count = 0
    for idx, row in df.iterrows():
        rid = int(row.get('row_idx', 0))
        fills = fills_by_idx.get(rid)
        if not fills:
            continue
        for col, val in fills.items():
            if col not in LLM_TARGET_COLS:
                continue
            if val is None:
                continue
            s = str(val).strip()
            if not s:
                continue
            # Only fill if cell is currently empty/mơ hồ
            cur = row.get(col)
            is_empty = pd.isna(cur) or str(cur).strip() == '' or str(cur).strip().lower() == 'mơ hồ'
            if is_empty:
                df.at[idx, col] = s
                merged_count += 1

    # 5. Export
    df.to_excel(PATH_OUTPUT_LLM, index=False)
    print(f"\n✓ Merge done!")
    print(f"  Total cells filled by LLM: {merged_count}")
    print(f"  Output: {PATH_OUTPUT_LLM.name}")

    # 6. Summary stats
    print(f"\n📊 Final fill rates per classification column:")
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            continue
        filled = df[col].notna().sum()
        filled -= (df[col].astype(str).str.strip() == '').sum()
        pct = filled / len(df) * 100
        print(f"    {col}: {filled}/{len(df)} ({pct:.1f}%)")


if __name__ == '__main__':
    main()
