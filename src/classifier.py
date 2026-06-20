import re
import pandas as pd
from typing import Any, Dict, List, Optional, Tuple
from unidecode import unidecode

from config import (
    OUTPUT_COLUMNS,
    COL_CURRENT_STATUS,
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
    PATH_KW_JSON,
)

# Source column names
COL_S = "Tình hình tiến độ công trình"
COL_T = "Nội dung làm việc, yêu cầu KH & đánh giá"
COL_U = "Kế hoạch lần tới"
COL_V = "Đề xuất"

ST_COLUMNS = [
    COL_PROGRESS, COL_WORK_CRM, COL_MARKETING, COL_SUBJECT_AETT,
    COL_OPINION, COL_REVIEW,
    COL_COMP_WORK, COL_COMP_SUBJECT, COL_ADVANTAGE,
]

PRICE_KEYWORDS = ["triệu", "tỷ", "vnđ", "vnd", "usd", "đô", "đồng"]

def _clean(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "null", "") else s

def _has_price(text: str) -> bool:
    low = text.lower()
    return any(kw in low for kw in PRICE_KEYWORDS)

def _compile_patterns(kws: List[str]) -> List[re.Pattern]:
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
    if not text or not pats:
        return False
    ascii_text = unidecode(text)
    return any(p.search(text) or p.search(ascii_text) for p in pats)

def _find_labels(text: str, label_pats: Dict[str, List[re.Pattern]]) -> List[str]:
    return [lbl for lbl, pats in label_pats.items() if _matches(text, pats)]

def _pick(labels: List[str]) -> Optional[str]:
    if len(labels) == 1:
        return labels[0]
    if len(labels) >= 2:
        return "mơ hồ"
    return None

def _find_brands(text: str, brand_pats: List[Tuple[str, List[re.Pattern]]]) -> List[str]:
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

def classify_row_regex(row: Dict[str, Any], col_pats: Dict[str, Any], brand_pats: List[Any]) -> Dict[str, Any]:
    """Classify a single row dict using regex rules and return the fills."""
    fills = {}
    
    text_s = _clean(row.get(COL_S))
    text_t = _clean(row.get(COL_T))
    text_u = _clean(row.get(COL_U))
    text_v = _clean(row.get(COL_V))

    text_st = " | ".join(x for x in [text_s, text_t] if x)
    text_uv = " | ".join(x for x in [text_u, text_v] if x)
    has_st = bool(text_s or text_t)

    # 1. S+T columns
    for col in ST_COLUMNS:
        if col not in col_pats:
            continue
        val = _pick(_find_labels(text_st, col_pats[col]))
        if val is not None:
            fills[col] = val

    # 2. Special fallback for AETT Nội dung làm việc
    cur = fills.get(COL_WORK_CRM)
    if cur is None or cur == "":
        if _has_price(text_st):
            fills[COL_WORK_CRM] = "Tư vấn khảo sát"
        elif has_st:
            fills[COL_WORK_CRM] = "mơ hồ"

    # 3. Brands
    brands = _find_brands(text_st, brand_pats)
    if brands:
        fills[COL_BRANDS] = "; ".join(brands)

    # 4. Kế hoạch (U+V and V)
    if COL_PLAN_NEXT in col_pats:
        val = _pick(_find_labels(text_uv, col_pats[COL_PLAN_NEXT]))
        if val is not None:
            fills[COL_PLAN_NEXT] = val

    if COL_PROPOSAL in col_pats:
        val = _pick(_find_labels(text_v, col_pats[COL_PROPOSAL]))
        if val is not None:
            fills[COL_PROPOSAL] = val

    return fills
