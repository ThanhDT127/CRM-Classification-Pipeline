"""
CRM Classification Pipeline - Configuration
=============================================
All paths, column definitions, keyword mappings, and constants in one place.
"""

import os
import json
import re
from pathlib import Path
from typing import Dict, List, Any, Tuple, Optional

# ─── Load .env from project root ────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent.parent / ".env")
except ImportError:
    pass  # python-dotenv not installed, rely on system env vars

# ─── Paths ──────────────────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).resolve().parent.parent   # D:\Works\CRM
PATH_INPUT       = PROJECT_ROOT / "data"     / (os.getenv("CRM_INPUT_FILE") or "CRM_merge.xlsx")
PATH_KW_JSON     = PROJECT_ROOT / "keywords" / "keywords_fixed.json"
PATH_PROMPT      = PROJECT_ROOT / "prompts"  / "prompt_CRM_v5.txt"
PATH_OUTPUT_REGEX = PROJECT_ROOT / "output"  / "CRM_classified.xlsx"
PATH_OUTPUT_LLM   = PROJECT_ROOT / "output"  / "CRM_classified_with_LLM.xlsx"
LLM_INPUT_JSON    = PROJECT_ROOT / "output"  / "llm_input.json"
OUT_LLM_JSON      = PROJECT_ROOT / "output"  / "llm_fills.json"
CKPT_JSON         = PROJECT_ROOT / "output"  / "llm_fills_checkpoint.json"

# ─── Gemini API settings ────────────────────────────────────────────────────
API_KEY = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY") or os.getenv("GOOGLE_GENERATIVEAI_API_KEY") or ""
MODEL_NAME = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
MIN_INTERVAL_S = float(os.getenv("GEMINI_MIN_INTERVAL_S") or "3.5")
JITTER_S       = float(os.getenv("GEMINI_JITTER_S") or "0.5")
BATCH_SIZE     = int(os.getenv("GEMINI_BATCH_SIZE") or "40")
MIN_BATCH      = 2

# ─── Column naming convention ───────────────────────────────────────────────
def col_name(major: str, minor: str) -> str:
    """Build '[Major] Minor' column name."""
    return f"[{major}] {minor}"


# Source text columns (from the CRM input file)
# These are the actual column names in CRM_TDCTDA.xlsx
INPUT_TEXT_COLUMNS = [
    "Tình trạng hiện tại",
    "Tình hình tiến độ công trình",
    "Nội dung làm việc, yêu cầu KH & đánh giá",
    "Kế hoạch lần tới",
    "Đề xuất",
]

# Column in the input that directly maps to COL_CURRENT_STATUS
COL_CURRENT_STATUS_SOURCE = "Tình trạng hiện tại"

# ─── Output column definitions (16 classification columns) ──────────────────
COL_CURRENT_STATUS = col_name("Hoạt Động CRM", "Tình hình hiện tại")
COL_PROGRESS       = col_name("Hoạt Động CRM", "Tiến độ")
COL_PICKUP         = col_name("Hoạt Động CRM", "ngày lấy hàng")
COL_WORK_CRM       = col_name("AETT", "Nội dung làm việc")
COL_MARKETING      = col_name("AETT", "Nhận xét tiếp thị")
COL_SUBJECT_AETT   = col_name("AETT", "Đối tượng")
COL_OPINION        = col_name("Khách Hàng", "Ý kiến KH")
COL_REVIEW         = col_name("Khách Hàng", "Nhận xét KH")
COL_PLAN_NEXT      = col_name("Kế Hoạch", "Kế hoạch lần tới")
COL_DATE_PLAN      = col_name("Kế Hoạch", "Ngày làm việc/ giao hàng:")
COL_PROPOSAL       = col_name("Kế Hoạch", "Đề xuất")
COL_COMP_WORK      = col_name("Đối Thủ Cạnh Tranh", "Nội dung làm việc")
COL_COMP_SUBJECT   = col_name("Đối Thủ Cạnh Tranh", "Đối tượng")
COL_ADVANTAGE      = col_name("Đối Thủ Cạnh Tranh", "Lợi thế")
COL_BRANDS         = col_name("Đối Thủ Cạnh Tranh", "Các Hãng đối thủ cạnh tranh")

OUTPUT_COLUMNS = [
    COL_CURRENT_STATUS,
    COL_PROGRESS,
    COL_PICKUP,
    COL_WORK_CRM,
    COL_MARKETING,
    COL_SUBJECT_AETT,
    COL_OPINION,
    COL_REVIEW,
    COL_PLAN_NEXT,
    COL_DATE_PLAN,
    COL_PROPOSAL,
    COL_COMP_WORK,
    COL_COMP_SUBJECT,
    COL_ADVANTAGE,
    COL_BRANDS,
]

# LLM fills all output columns EXCEPT COL_CURRENT_STATUS (which is a direct copy)
LLM_TARGET_COLS = [c for c in OUTPUT_COLUMNS if c != COL_CURRENT_STATUS]


# ─── Keyword loading ────────────────────────────────────────────────────────
def _split_terms(raw: Any) -> List[str]:
    """Split a keyword entry (str or list) by semicolons."""
    if isinstance(raw, list):
        parts = []
        for r in raw:
            parts.extend(_split_terms(r))
        return parts
    if not isinstance(raw, str):
        return []
    return [t.strip() for t in raw.split(";") if t.strip()]


def load_keywords(path: Path = PATH_KW_JSON) -> Dict[str, Any]:
    """Load keywords_fixed.json."""
    return json.loads(path.read_text(encoding="utf-8"))


def build_keyword_index(
    kw: Dict[str, Any],
) -> Tuple[
    Dict[str, Dict[str, List[str]]],   # col -> {label: [keywords]}
    List[str],                          # competitor_brands
]:
    """
    Parse keywords_fixed.json into:
      - kw_index: {col_name: {label: [keyword_strings]}}
      - competitor_brands: list of brand strings
    """
    kw_index: Dict[str, Dict[str, List[str]]] = {}
    competitor_brands: List[str] = []

    for major, minors in (kw or {}).items():
        if not isinstance(minors, dict):
            continue
        for minor, subs in minors.items():
            # Special case: competitor brands
            if major == "Đối Thủ Cạnh Tranh" and minor == "Cạnh Tranh" and isinstance(subs, dict):
                brand_list = subs.get("Các Hãng cụ thể") or subs.get("Các Hãng đối thủ cạnh tranh")
                if isinstance(brand_list, list):
                    for b in brand_list:
                        for term in _split_terms(b):
                            if term:
                                competitor_brands.append(term.strip())
                continue

            if isinstance(subs, dict):
                col = col_name(major, minor)
                col_labels: Dict[str, List[str]] = {}
                for sub, kw_list in subs.items():
                    if not isinstance(sub, str) or not sub.strip():
                        continue
                    terms: List[str] = []
                    if isinstance(kw_list, list):
                        for k in kw_list:
                            terms.extend(_split_terms(k))
                    elif isinstance(kw_list, str):
                        terms.extend(_split_terms(kw_list))
                    col_labels[sub.strip()] = [t for t in terms if t]
                if col_labels:
                    kw_index[col] = col_labels

    # Deduplicate brands
    seen = set()
    unique_brands: List[str] = []
    for b in competitor_brands:
        k = b.lower()
        if k not in seen:
            unique_brands.append(b)
            seen.add(k)

    return kw_index, unique_brands


def build_canonical_maps(
    kw_index: Dict[str, Dict[str, List[str]]],
) -> Tuple[Dict[str, List[str]], Dict[str, Dict[str, str]]]:
    """
    Build:
      - canonical_by_col: {col: [label1, label2, ...]}
      - canonical_lower_map: {col: {label_lower: original_label}}
    """
    canonical_by_col: Dict[str, List[str]] = {}
    canonical_lower_map: Dict[str, Dict[str, str]] = {}

    for col, labels_dict in kw_index.items():
        labels = list(labels_dict.keys())
        if labels:
            canonical_by_col[col] = labels
            canonical_lower_map[col] = {lbl.lower(): lbl for lbl in labels}

    return canonical_by_col, canonical_lower_map
