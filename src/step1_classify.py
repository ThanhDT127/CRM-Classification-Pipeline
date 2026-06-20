"""
CRM Classification Pipeline - Step 1: Regex Classification
============================================================
Read CRM_merge.xlsx, apply keyword-based regex classification,
and export CRM_classified.xlsx.

Rules (matching notebook logic):
- 'Tình hình hiện tại' copies directly from column R.
- Hoạt Động CRM, Đối Thủ Cạnh Tranh, AETT, Khách Hàng → use S+T only.
- Kế Hoạch → use U+V only. Đề xuất → V only.
- If keywords match >=2 tags for same column → 'mơ hồ'.
- AETT 'Nội dung làm việc': price indicator → 'Tư vấn khảo sát'.
- AETT 'Nội dung làm việc': no match but S/T has data → 'mơ hồ'.
"""

import re
import pandas as pd
from typing import Any, Dict, List, Optional, Tuple
from unidecode import unidecode

from config import (
    PATH_INPUT,
    PATH_KW_JSON,
    PATH_OUTPUT_REGEX,
    OUTPUT_COLUMNS,
    COL_CURRENT_STATUS,
    COL_CURRENT_STATUS_SOURCE,
    COL_BRANDS,
    COL_WORK_CRM,
    COL_PROGRESS,
    COL_PICKUP,
    COL_COMP_WORK,
    COL_COMP_SUBJECT,
    COL_ADVANTAGE,
    COL_SUBJECT_AETT,
    COL_MARKETING,
    COL_OPINION,
    COL_REVIEW,
    COL_PLAN_NEXT,
    COL_DATE_PLAN,
    COL_PROPOSAL,
    load_keywords,
    build_keyword_index,
)

# ─── Source column names ────────────────────────────────────────────────────
COL_R = "Tình trạng hiện tại"
COL_S = "Tình hình tiến độ công trình"
COL_T = "Nội dung làm việc, yêu cầu KH & đánh giá"
COL_U = "Kế hoạch lần tới"
COL_V = "Đề xuất"

# Columns that read from S+T
ST_COLUMNS = [
    COL_PROGRESS, COL_WORK_CRM, COL_MARKETING, COL_SUBJECT_AETT,
    COL_OPINION, COL_REVIEW,
    COL_COMP_WORK, COL_COMP_SUBJECT, COL_ADVANTAGE,
]

# Price indicator keywords (for AETT fallback)
PRICE_KEYWORDS = ["triệu", "tỷ", "vnđ", "vnd", "usd", "đô", "đồng"]


# ─── Text helpers ───────────────────────────────────────────────────────────
def _clean(v: Any) -> str:
    """Clean a cell value to trimmed string, or '' for blanks."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "") else s


def _has_price(text: str) -> bool:
    low = text.lower()
    return any(kw in low for kw in PRICE_KEYWORDS)


# ─── Regex helpers ──────────────────────────────────────────────────────────
def _compile_patterns(kws: List[str]) -> List[re.Pattern]:
    """Build word-boundary regex patterns, including unidecoded variants."""
    pats: List[re.Pattern] = []
    for kw in kws:
        kw = kw.strip()
        if not kw:
            continue
        pats.append(re.compile(rf"(?<!\w){re.escape(kw)}(?!\w)", re.IGNORECASE))
        ascii_kw = unidecode(kw)
        if ascii_kw != kw:
            pats.append(re.compile(rf"(?<!\w){re.escape(ascii_kw)}(?!\w)", re.IGNORECASE))
    return pats


def _matches(text: str, pats: List[re.Pattern]) -> bool:
    """Check if text (or its ASCII form) matches any pattern."""
    if not text or not pats:
        return False
    ascii_text = unidecode(text)
    return any(p.search(text) or p.search(ascii_text) for p in pats)


def _find_labels(text: str, label_pats: Dict[str, List[re.Pattern]]) -> List[str]:
    """Return all labels whose keywords appear in text (stable order)."""
    return [lbl for lbl, pats in label_pats.items() if _matches(text, pats)]


def _pick(labels: List[str]) -> Optional[str]:
    """1 label → return it; ≥2 → 'mơ hồ'; 0 → None."""
    if len(labels) == 1:
        return labels[0]
    if len(labels) >= 2:
        return "mơ hồ"
    return None


def _find_brands(text: str, brand_pats: List[Tuple[str, List[re.Pattern]]]) -> List[str]:
    """Return all brand names found in text (deduplicated, title-cased)."""
    if not text:
        return []
    ascii_text = unidecode(text)
    found, seen = [], set()
    for brand_name, pats in brand_pats:
        for p in pats:
            m = p.search(text) or p.search(ascii_text)
            if m:
                surface = m.group().strip().title()
                key = surface.lower()
                if key not in seen:
                    found.append(surface)
                    seen.add(key)
                break
    return found


# ─── Main ───────────────────────────────────────────────────────────────────
def main():
    # 1. Load keywords & build compiled patterns
    print(f"Loading keywords from {PATH_KW_JSON.name} ...")
    kw = load_keywords()
    kw_index, brand_list = build_keyword_index(kw)
    print(f"  Columns indexed: {len(kw_index)}")
    print(f"  Brands: {len(brand_list)}")

    col_pats: Dict[str, Dict[str, List[re.Pattern]]] = {}
    for col, labels_dict in kw_index.items():
        col_pats[col] = {lbl: _compile_patterns(kws) for lbl, kws in labels_dict.items()}

    brand_pats = [(b, _compile_patterns([b])) for b in brand_list]

    # 2. Read input
    print(f"\nReading {PATH_INPUT.name} ...")
    df = pd.read_excel(PATH_INPUT)
    print(f"  Rows: {len(df)}, Columns: {len(df.columns)}")

    df["row_idx"] = range(2, len(df) + 2)

    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    # 3. Direct copy: R → COL_CURRENT_STATUS
    if COL_CURRENT_STATUS_SOURCE in df.columns:
        df[COL_CURRENT_STATUS] = df[COL_CURRENT_STATUS_SOURCE]
    else:
        print(f"  ⚠️ Column '{COL_CURRENT_STATUS_SOURCE}' not found")

    # 4. Classify
    print("\nRunning regex classification ...")
    n_classified, n_ambiguous = 0, 0

    for idx, row in df.iterrows():
        text_s = _clean(row.get(COL_S))
        text_t = _clean(row.get(COL_T))
        text_u = _clean(row.get(COL_U))
        text_v = _clean(row.get(COL_V))

        text_st = " | ".join(x for x in [text_s, text_t] if x)
        text_uv = " | ".join(x for x in [text_u, text_v] if x)
        has_st = bool(text_s or text_t)

        # ── S+T columns (CRM, Đối thủ, AETT, Khách hàng) ──
        for col in ST_COLUMNS:
            if col not in col_pats:
                continue
            val = _pick(_find_labels(text_st, col_pats[col]))
            if val is not None:
                df.at[idx, col] = val
                n_classified += 1
                if val == "mơ hồ":
                    n_ambiguous += 1

        # ── AETT Nội dung làm việc: special fallback rules ──
        cur = df.at[idx, COL_WORK_CRM]
        if cur is None or (isinstance(cur, float) and pd.isna(cur)):
            if _has_price(text_st):
                df.at[idx, COL_WORK_CRM] = "Tư vấn khảo sát"
                n_classified += 1
            elif has_st:
                df.at[idx, COL_WORK_CRM] = "mơ hồ"
                n_classified += 1
                n_ambiguous += 1

        # ── Brands from S+T ──
        brands = _find_brands(text_st, brand_pats)
        if brands:
            df.at[idx, COL_BRANDS] = "; ".join(brands)
            n_classified += 1

        # ── U+V columns (Kế hoạch) ──
        if COL_PLAN_NEXT in col_pats:
            val = _pick(_find_labels(text_uv, col_pats[COL_PLAN_NEXT]))
            if val is not None:
                df.at[idx, COL_PLAN_NEXT] = val
                n_classified += 1
                if val == "mơ hồ":
                    n_ambiguous += 1

        if COL_PROPOSAL in col_pats:
            val = _pick(_find_labels(text_v, col_pats[COL_PROPOSAL]))
            if val is not None:
                df.at[idx, COL_PROPOSAL] = val
                n_classified += 1
                if val == "mơ hồ":
                    n_ambiguous += 1

        # Date columns (COL_PICKUP, COL_DATE_PLAN) left for LLM

    # 5. Export
    df.to_excel(PATH_OUTPUT_REGEX, index=False)
    print(f"\n✓ Regex classification done!")
    print(f"  Label assignments: {n_classified}")
    print(f"  Ambiguous ('mơ hồ'): {n_ambiguous}")
    print(f"  Output: {PATH_OUTPUT_REGEX.name}")

    print("\n  Fill rates per classification column:")
    for col in OUTPUT_COLUMNS:
        if col not in df.columns:
            continue
        filled = len(df) - df[col].isna().sum() - (df[col].astype(str).str.strip() == "").sum()
        pct = filled / len(df) * 100
        print(f"    {col}: {filled}/{len(df)} ({pct:.1f}%)")

    return df


if __name__ == "__main__":
    main()
