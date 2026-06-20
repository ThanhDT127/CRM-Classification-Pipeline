"""
CRM Classification Pipeline - Step 3: Call Gemini LLM
======================================================
Read llm_input.json, send batches to Gemini for classification,
save results to llm_fills.json with checkpoint support.

Supports concurrent batch processing for speed.
"""

import json
import os
import httpx
import re
import random
import time
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

from google import genai
from google.genai import types

from config import (
    PROJECT_ROOT,
    API_KEY,
    MODEL_NAME,
    MIN_INTERVAL_S,
    JITTER_S,
    BATCH_SIZE,
    MIN_BATCH,
    LLM_INPUT_JSON,
    OUT_LLM_JSON,
    CKPT_JSON,
    PATH_PROMPT,
    LLM_TARGET_COLS,
    COL_PICKUP,
    COL_DATE_PLAN,
    COL_BRANDS,
    load_keywords,
    build_keyword_index,
    build_canonical_maps,
)

# ─── Concurrency settings ──────────────────────────────────────────────────
CONCURRENT_WORKERS = int(os.getenv("GEMINI_CONCURRENT_WORKERS") or "5")

# ─── Date normalization regex ───────────────────────────────────────────────
_DATE_DMY = re.compile(r'(?<!\d)(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?(?!\d)')
_DATE_YMD = re.compile(r'(?<!\d)(\d{4})-(\d{1,2})-(\d{1,2})(?!\d)')
_MONTH_YEAR = re.compile(r'(?<!\d)(\d{1,2})[/-](\d{4}|\d{2})(?!\d)')
_THANG = re.compile(r'(?:tháng|thang)\s*(\d{1,2})', re.IGNORECASE)


def _norm_ddmmyy(s: str) -> Optional[str]:
    """Normalize date string to dd/mm/yy format."""
    if s is None:
        return None
    raw = str(s).strip()
    if not raw:
        return None

    t = raw.strip()

    m = _DATE_YMD.search(t)
    if m:
        yy = int(m.group(1)) % 100
        mm = int(m.group(2))
        dd = int(m.group(3))
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            return f'{dd:02d}/{mm:02d}/{yy:02d}'

    m = _DATE_DMY.search(t)
    if m:
        dd = int(m.group(1))
        mm = int(m.group(2))
        yy_raw = m.group(3)
        if yy_raw is None:
            yy = 26
        else:
            yy = int(yy_raw) % 100
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            return f'{dd:02d}/{mm:02d}/{yy:02d}'

    m = _MONTH_YEAR.search(t)
    if m:
        mm = int(m.group(1))
        yy = int(m.group(2)) % 100
        if 1 <= mm <= 12:
            return f'01/{mm:02d}/{yy:02d}'

    m = _THANG.search(t)
    if m:
        mm = int(m.group(1))
        if 1 <= mm <= 12:
            return f'01/{mm:02d}/26'

    return None


# ─── JSON extraction from LLM response ─────────────────────────────────────
_JSON_BLOCK_RE = re.compile(r'```(?:json)?\s*(\[.+?\])\s*```', re.DOTALL)


def _extract_json_array(text: str) -> str:
    """Extract JSON array from LLM response text."""
    m = _JSON_BLOCK_RE.search(text)
    if m:
        return m.group(1)
    start = text.find('[')
    end = text.rfind(']')
    if start != -1 and end != -1 and end > start:
        return text[start:end + 1]
    raise ValueError(f'Cannot find JSON array in text: {text[:200]}')


def map_corrupted_key(raw_k: str, target_cols: List[str]) -> str:
    raw_clean = raw_k.lower().strip()
    
    # 1. Exact or suffix match
    for col in target_cols:
        col_clean = col.lower()
        if raw_clean == col_clean or col_clean.endswith(raw_clean) or raw_clean.endswith(col_clean):
            return col
            
    # 2. Check for group names and sub-keywords with prefix resilience
    is_crm = "hoạt động" in raw_clean or "crm" in raw_clean
    is_aett = "aett" in raw_clean
    is_kh = "khách" in raw_clean or "kh" in raw_clean
    is_plan = "kế hoạch" in raw_clean or "hoạch" in raw_clean
    is_comp = "đối thủ" in raw_clean or "cạnh tranh" in raw_clean or "tranh" in raw_clean
    
    if is_crm:
        if any(x in raw_clean for x in ["tiế", "độ"]):
            return "[Hoạt Động CRM] Tiến độ"
        if any(x in raw_clean for x in ["ngày", "lấy", "hàn"]):
            return "[Hoạt Động CRM] ngày lấy hàng"
            
    if is_aett:
        if any(x in raw_clean for x in ["nộ", "dung", "làm", "việc"]):
            return "[AETT] Nội dung làm việc"
        if any(x in raw_clean for x in ["nhậ", "xét", "tiế", "thị"]):
            return "[AETT] Nhận xét tiếp thị"
        if any(x in raw_clean for x in ["đố", "tượ"]):
            return "[AETT] Đối tượng"
            
    if is_kh:
        if any(x in raw_clean for x in ["ý", "kiế"]):
            return "[Khách Hàng] Ý kiến KH"
        if any(x in raw_clean for x in ["nhậ", "xét"]):
            return "[Khách Hàng] Nhận xét KH"
            
    if is_plan:
        if any(x in raw_clean for x in ["lần", "tới"]):
            return "[Kế Hoạch] Kế hoạch lần tới"
        if any(x in raw_clean for x in ["ngày", "giao", "hàn", "làm", "việc"]):
            return "[Kế Hoạch] Ngày làm việc/ giao hàng:"
        if any(x in raw_clean for x in ["đề", "xuấ"]):
            return "[Kế Hoạch] Đề xuất"
            
    if is_comp:
        if any(x in raw_clean for x in ["nộ", "dung", "làm", "việc"]):
            return "[Đối Thủ Cạnh Tranh] Nội dung làm việc"
        if any(x in raw_clean for x in ["đố", "tượ"]):
            return "[Đối Thủ Cạnh Tranh] Đối tượng"
        if any(x in raw_clean for x in ["lợi", "thế"]):
            return "[Đối Thủ Cạnh Tranh] Lợi thế"
        if any(x in raw_clean for x in ["các", "hãn", "thủ"]):
            return "[Đối Thủ Cạnh Tranh] Các Hãng đối thủ cạnh tranh"

    # Fallback to general substring heuristics
    if "tiến độ" in raw_clean:
        return "[Hoạt Động CRM] Tiến độ"
    if "lấy hàng" in raw_clean:
        return "[Hoạt Động CRM] ngày lấy hàng"
    if "tiếp thị" in raw_clean:
        return "[AETT] Nhận xét tiếp thị"
    if "ý kiến kh" in raw_clean:
        return "[Khách Hàng] Ý kiến KH"
    if "nhận xét kh" in raw_clean:
        return "[Khách Hàng] Nhận xét KH"
    if "lần tới" in raw_clean:
        return "[Kế Hoạch] Kế hoạch lần tới"
    if "giao hàng" in raw_clean or "việc/ giao" in raw_clean:
        return "[Kế Hoạch] Ngày làm việc/ giao hàng:"
    if "đề xuất" in raw_clean:
        return "[Kế Hoạch] Đề xuất"
    if "các hãng" in raw_clean or "hãng đối thủ" in raw_clean:
        return "[Đối Thủ Cạnh Tranh] Các Hãng đối thủ cạnh tranh"
    if "lợi thế" in raw_clean:
        return "[Đối Thủ Cạnh Tranh] Lợi thế"
        
    return raw_k


def repair_json_text(text: str, target_cols: List[str]) -> str:
    lines = text.split('\n')
    repaired_lines = []
    for line in lines:
        match = re.match(r'^(\s*)"?([^"]+)"?\s*:\s*(.*)$', line)
        if match:
            indent = match.group(1)
            raw_key = match.group(2).strip()
            val_part = match.group(3)
            
            canonical_key = map_corrupted_key(raw_key, target_cols)
            repaired_lines.append(f'{indent}"{canonical_key}": {val_part}')
        else:
            repaired_lines.append(line)
    return '\n'.join(repaired_lines)


def _parse_llm_json(text: str) -> List[Dict[str, Any]]:
    """Parse and validate LLM JSON response, with automatic repair on failure."""
    try:
        json_text = _extract_json_array(text)
        try:
            data = json.loads(json_text)
        except Exception as json_err:
            print("  [WARN] Standard JSON parsing failed. Attempting automatic repair of keys...")
            try:
                repaired_text = repair_json_text(json_text, LLM_TARGET_COLS)
                data = json.loads(repaired_text)
                print("  [OK] Automatic JSON repair successful!")
            except Exception:
                # Raise original error if repair fails
                raise json_err

        if not isinstance(data, list):
            raise ValueError('Not a JSON array')
        for obj in data:
            if not isinstance(obj, dict):
                raise ValueError('Each item must be an object')
            # Find the row index key resiliently (supports row_idx, idx, _idx, row_index, etc.)
            for k in list(obj.keys()):
                if 'idx' in k.lower() or 'index' in k.lower():
                    obj['row_idx'] = obj[k]
                    break
            if 'row_idx' not in obj:
                raise ValueError('Each item must be an object with row_idx or idx')
            if 'fills' not in obj or not isinstance(obj.get('fills'), dict):
                raise ValueError('Each item must contain fills object')
        return data
    except Exception as e:
        safe_text = text[:500].encode('ascii', errors='replace').decode('ascii')
        print(f"\n[DEBUG] Raw LLM Response that failed parsing:\n{safe_text}\n")
        raise e


# ─── Label canonicalization ─────────────────────────────────────────────────
CANONICAL_LOWER_MAP: Dict[str, Dict[str, str]] = {}


def _setup_canonical():
    """Initialize canonical label maps from keywords."""
    global CANONICAL_LOWER_MAP
    kw = load_keywords()
    kw_index, _ = build_keyword_index(kw)
    _, CANONICAL_LOWER_MAP = build_canonical_maps(kw_index)


def _canonicalize_label(col: str, val: str) -> Optional[str]:
    """Map an LLM-returned label to its canonical form."""
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None

    # Date columns handled separately
    if col in (COL_PICKUP, COL_DATE_PLAN):
        return _norm_ddmmyy(s)

    # Competitor brands: allow any string
    if col == COL_BRANDS:
        return s

    # Other labels: must match canonical set
    lm = CANONICAL_LOWER_MAP.get(col) or {}
    return lm.get(s.lower())


# ─── Rate limiting (thread-safe global rate limiter) ────────────────────────
_last_request_time = 0.0
_rate_limit_lock = threading.Lock()

def wait_for_rate_limit():
    global _last_request_time
    with _rate_limit_lock:
        now = time.time()
        elapsed = now - _last_request_time
        target_interval = MIN_INTERVAL_S + (random.random() * JITTER_S)
        if elapsed < target_interval:
            wait_time = target_interval - elapsed
            time.sleep(wait_time)
        _last_request_time = time.time()


# ─── Checkpoint (thread-safe) ──────────────────────────────────────────────
_ckpt_lock = threading.Lock()


def load_checkpoint() -> Dict[str, Any]:
    if CKPT_JSON.exists():
        return json.loads(CKPT_JSON.read_text(encoding='utf-8'))
    return {'updated_at': None, 'results': {}}


def save_checkpoint(ckpt: Dict[str, Any]) -> None:
    ckpt['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    CKPT_JSON.write_text(json.dumps(ckpt, ensure_ascii=False, indent=2), encoding='utf-8')


# ─── LLM batch call ────────────────────────────────────────────────────────
def call_llm_batch(
    client: genai.Client,
    model_name: str,
    system_prompt: str,
    batch: List[Dict[str, Any]],
    max_retry: int = 10,
) -> List[Dict[str, Any]]:
    """Send a batch to Gemini and parse the JSON response."""
    payload = json.dumps(batch, ensure_ascii=False)
    user_input = 'INPUT_JSON_ARRAY:\n' + payload

    for attempt in range(1, max_retry + 1):
        try:
            # Apply rate limiting to serialise API calls and avoid 429
            wait_for_rate_limit()

            resp = client.models.generate_content(
                model=model_name,
                contents=user_input,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.0,
                    max_output_tokens=65536,
                )
            )
            raw = getattr(resp, 'text', '') or ''
            return _parse_llm_json(raw)
        except Exception as e:
            msg = str(e)
            msg_safe = msg.encode('ascii', errors='replace').decode('ascii')
            low = msg.lower()

            # 429 / rate limit handling
            if ('429' in low) or ('too many request' in low) or ('rate limit' in low) or ('perminute' in low):
                wait = min(120, 15 * attempt) + random.random() * 2
                print(f'[WARN] 429/rate-limit: wait={wait:.1f}s attempt={attempt} err={msg_safe[:120]}')
                time.sleep(wait)
                continue

            # transient resource exhausted
            if ('resource_exhausted' in low) or ('temporarily' in low):
                wait = min(120, 10 * attempt) + random.random() * 2
                print(f'[WARN] transient: wait={wait:.1f}s attempt={attempt} err={msg_safe[:120]}')
                time.sleep(wait)
                continue

            if attempt >= max_retry:
                raise

            wait = (2 ** attempt) + random.random()
            print(f'[WARN] retry={attempt} wait={wait:.1f}s err={msg_safe[:120]}')
            time.sleep(wait)

    raise RuntimeError('LLM_FAILED')


# ─── Worker function ───────────────────────────────────────────────────────
def _process_batch(
    worker_id: int,
    client: genai.Client,
    model_name: str,
    system_prompt: str,
    batch: List[Dict[str, Any]],
    missing_cols_map: Dict[str, set],
    results: Dict[str, Any],
    ckpt: Dict[str, Any],
    batch_idx: int,
    total_batches: int,
) -> int:
    """Process a single batch. Returns number of updated rows."""
    try:
        out = call_llm_batch(client, model_name, system_prompt, batch)
        out_by_id = {str(int(o.get('row_idx'))): o for o in out if isinstance(o, dict) and 'row_idx' in o}

        updated = 0
        batch_results = {}
        for src in batch:
            rid = str(int(src.get('row_idx')))
            allowed = missing_cols_map.get(rid, set())
            o = out_by_id.get(rid)
            if not o:
                continue
            fills = o.get('fills') or {}
            if not isinstance(fills, dict):
                fills = {}

            sanitized: Dict[str, Any] = {}
            for k, v in fills.items():
                if k not in allowed:
                    continue
                if v is None:
                    sanitized[k] = None
                    continue
                if isinstance(v, str) and not v.strip():
                    sanitized[k] = None
                    continue
                if isinstance(v, str):
                    sanitized[k] = _canonicalize_label(k, v)
                else:
                    sanitized[k] = None

            batch_results[rid] = {'row_idx': int(rid), 'fills': sanitized}
            updated += 1

        # Thread-safe update
        with _ckpt_lock:
            results.update(batch_results)
            ckpt['results'] = results
            save_checkpoint(ckpt)

        print(f'  W{worker_id} batch {batch_idx}/{total_batches} [OK] updated={updated}', flush=True)
        return updated

    except Exception as e:
        err = str(e)
        print(f'  W{worker_id} batch {batch_idx}/{total_batches} [FAIL] {err[:140]}', flush=True)
        raise


# ─── Main ───────────────────────────────────────────────────────────────────
def main():
    # 1. Setup
    USE_VERTEX = os.getenv("USE_VERTEX", "True").lower() in ("true", "1", "yes")
    VERTEX_PROJECT = os.getenv("VERTEX_PROJECT")
    VERTEX_LOCATION = os.getenv("VERTEX_LOCATION", "us-central1")
    SA_KEY_PATH = PROJECT_ROOT / "sa-key.json"

    if USE_VERTEX and SA_KEY_PATH.exists():
        print("Using Vertex AI client for Gemini...")
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = str(SA_KEY_PATH)
        if not VERTEX_PROJECT:
            try:
                with open(SA_KEY_PATH, "r", encoding="utf-8") as f:
                    VERTEX_PROJECT = json.load(f).get("project_id")
            except Exception:
                pass
        client = genai.Client(
            vertexai=True,
            project=VERTEX_PROJECT,
            location=VERTEX_LOCATION
        )
        client._api_client._httpx_client.timeout = httpx.Timeout(120.0)
        model_name = MODEL_NAME
        if model_name.startswith("models/"):
            model_name = model_name[len("models/"):]
    else:
        print("Using Google AI Studio client for Gemini...")
        if not API_KEY:
            print("  [ERROR] No API key found. Set GEMINI_API_KEY environment variable.")
            return
        client = genai.Client(api_key=API_KEY)
        client._api_client._httpx_client.timeout = httpx.Timeout(120.0)
        model_name = MODEL_NAME

    _setup_canonical()

    # 2. Load prompt
    if not PATH_PROMPT.exists():
        print(f"  [ERROR] Prompt file not found: {PATH_PROMPT}")
        return
    system_prompt = PATH_PROMPT.read_text(encoding='utf-8')
    print(f"  [OK] Loaded prompt: {PATH_PROMPT.name} ({len(system_prompt)} chars)")

    # 3. Load LLM input
    if not LLM_INPUT_JSON.exists():
        print(f"  [ERROR] LLM input not found: {LLM_INPUT_JSON} (run step2 first)")
        return

    items = json.loads(LLM_INPUT_JSON.read_text(encoding='utf-8'))
    print(f"  [OK] Loaded {len(items)} items from {LLM_INPUT_JSON.name}")

    # 4. Prepare missing_cols lookup
    missing_cols_map: Dict[str, set] = {}
    for it in items:
        rid = str(int(it.get('row_idx')))
        miss = it.get('missing_cols')
        missing_cols_map[rid] = set(miss) if isinstance(miss, list) else set()

    # 5. Load checkpoint
    ckpt = load_checkpoint()
    results: Dict[str, Any] = ckpt.get('results', {}) if isinstance(ckpt.get('results'), dict) else {}

    done_ids = set(results.keys())
    pending = [it for it in items if str(int(it['row_idx'])) not in done_ids]
    print(f"  Checkpoint: done={len(done_ids)}, pending={len(pending)}")

    if not pending:
        print("  [OK] All items already processed!")
    else:
        # 6. Build batches
        batch_size = BATCH_SIZE
        batches = []
        for i in range(0, len(pending), batch_size):
            batches.append(pending[i:i + batch_size])

        total_batches = len(batches)
        workers = min(CONCURRENT_WORKERS, total_batches)
        est_time = total_batches * 80 / workers / 3600  # rough estimate

        print(f'\n  [START] Starting LLM processing:')
        print(f'   Batches: {total_batches} (batch_size={batch_size})')
        print(f'   Workers: {workers} concurrent')
        print(f'   Estimated: ~{est_time:.1f} hours')
        print(f'   Checkpoint: every batch (crash-safe)\n')

        total_updated = 0
        total_errors = 0

        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {}
            for batch_idx, batch in enumerate(batches, 1):
                future = executor.submit(
                    _process_batch,
                    batch_idx % workers,  # worker_id
                    client, model_name, system_prompt,
                    batch, missing_cols_map,
                    results, ckpt,
                    batch_idx, total_batches,
                )
                futures[future] = batch_idx

            for future in as_completed(futures):
                batch_idx = futures[future]
                try:
                    updated = future.result()
                    total_updated += updated
                except Exception as e:
                    total_errors += 1
                    print(f'  [ERROR] Batch {batch_idx} failed: {str(e)[:100]}')

        print(f'\n  Total updated: {total_updated}, errors: {total_errors}')

    # 7. Export final results
    final_list = list(results.values())
    final_list.sort(key=lambda x: int(x.get('row_idx', 0)))
    OUT_LLM_JSON.write_text(json.dumps(final_list, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'\n  [OK] LLM done: {len(final_list)}/{len(items)} rows -> {OUT_LLM_JSON.name}')
    print(f'  [OK] Checkpoint: {CKPT_JSON.name}')


if __name__ == '__main__':
    main()
