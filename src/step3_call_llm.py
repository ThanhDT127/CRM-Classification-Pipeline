"""
CRM Classification Pipeline - Step 3: Call Gemini LLM
======================================================
Read llm_input.json, send batches to Gemini for classification,
save results to llm_fills.json with checkpoint support.
"""

import json
import os
import re
import random
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import google.generativeai as genai

from config import (
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


def _parse_llm_json(text: str) -> List[Dict[str, Any]]:
    """Parse and validate LLM JSON response."""
    json_text = _extract_json_array(text)
    data = json.loads(json_text)
    if not isinstance(data, list):
        raise ValueError('Not a JSON array')
    for obj in data:
        if not isinstance(obj, dict) or 'row_idx' not in obj:
            raise ValueError('Each item must be an object with row_idx')
        if 'fills' not in obj or not isinstance(obj.get('fills'), dict):
            raise ValueError('Each item must contain fills object')
    return data


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


# ─── Throttle ───────────────────────────────────────────────────────────────
_last_call_ts = 0.0


def _throttle():
    """Rate-limit API calls."""
    global _last_call_ts
    now = time.time()
    elapsed = now - _last_call_ts
    need = MIN_INTERVAL_S + random.random() * JITTER_S
    if elapsed < need:
        time.sleep(need - elapsed)
    _last_call_ts = time.time()


# ─── Checkpoint ─────────────────────────────────────────────────────────────
def load_checkpoint() -> Dict[str, Any]:
    if CKPT_JSON.exists():
        return json.loads(CKPT_JSON.read_text(encoding='utf-8'))
    return {'updated_at': None, 'results': {}}


def save_checkpoint(ckpt: Dict[str, Any]) -> None:
    ckpt['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    CKPT_JSON.write_text(json.dumps(ckpt, ensure_ascii=False, indent=2), encoding='utf-8')


# ─── LLM batch call ────────────────────────────────────────────────────────
def call_llm_batch(
    model: Any,
    system_prompt: str,
    batch: List[Dict[str, Any]],
    max_retry: int = 5,
) -> List[Dict[str, Any]]:
    """Send a batch to Gemini and parse the JSON response."""
    payload = json.dumps(batch, ensure_ascii=False)
    user_input = 'INPUT_JSON_ARRAY:\n' + payload

    for attempt in range(1, max_retry + 1):
        try:
            _throttle()
            resp = model.generate_content(
                system_prompt + '\n\n' + user_input,
                generation_config=genai.types.GenerationConfig(
                    temperature=0.0,
                    max_output_tokens=8192,
                ),
            )
            raw = getattr(resp, 'text', '') or ''
            return _parse_llm_json(raw)
        except Exception as e:
            msg = str(e)
            low = msg.lower()

            # 429 / rate limit handling
            if ('429' in low) or ('too many request' in low) or ('rate limit' in low) or ('perminute' in low):
                wait = min(120, 15 * attempt) + random.random() * 2
                print(f'⚠️ 429/rate-limit: wait={wait:.1f}s attempt={attempt} err={msg[:120]}')
                time.sleep(wait)
                continue

            # transient resource exhausted
            if ('resource_exhausted' in low) or ('temporarily' in low):
                wait = min(120, 10 * attempt) + random.random() * 2
                print(f'⚠️ transient: wait={wait:.1f}s attempt={attempt} err={msg[:120]}')
                time.sleep(wait)
                continue

            if attempt >= max_retry:
                raise

            wait = (2 ** attempt) + random.random()
            print(f'⚠️ retry={attempt} wait={wait:.1f}s err={msg[:120]}')
            time.sleep(wait)

    raise RuntimeError('LLM_FAILED')


# ─── Main ───────────────────────────────────────────────────────────────────
def main():
    # 1. Setup
    if not API_KEY:
        print("❌ No API key found. Set GEMINI_API_KEY environment variable.")
        return

    genai.configure(api_key=API_KEY)
    _setup_canonical()

    # 2. Load prompt
    if not PATH_PROMPT.exists():
        print(f"❌ Prompt file not found: {PATH_PROMPT}")
        return
    system_prompt = PATH_PROMPT.read_text(encoding='utf-8')
    print(f"✓ Loaded prompt: {PATH_PROMPT.name} ({len(system_prompt)} chars)")

    # 3. Load LLM input
    if not LLM_INPUT_JSON.exists():
        print(f"❌ LLM input not found: {LLM_INPUT_JSON} (run step2 first)")
        return

    items = json.loads(LLM_INPUT_JSON.read_text(encoding='utf-8'))
    print(f"✓ Loaded {len(items)} items from {LLM_INPUT_JSON.name}")

    # 4. Prepare missing_cols lookup
    missing_cols_map: Dict[str, set] = {}
    for it in items:
        rid = str(int(it.get('row_idx')))
        miss = it.get('missing_cols')
        missing_cols_map[rid] = set(miss) if isinstance(miss, list) else set()

    # 5. Load checkpoint
    model = genai.GenerativeModel(MODEL_NAME)
    ckpt = load_checkpoint()
    results: Dict[str, Any] = ckpt.get('results', {}) if isinstance(ckpt.get('results'), dict) else {}

    done_ids = set(results.keys())
    pending = [it for it in items if str(int(it['row_idx'])) not in done_ids]
    print(f"  Checkpoint: done={len(done_ids)}, pending={len(pending)}")

    if not pending:
        print("✓ All items already processed!")
    else:
        # 6. Process batches
        batch_size = BATCH_SIZE
        idx = 0

        print(f'\nBắt đầu gọi LLM (batch_size={batch_size})...')
        while idx < len(pending):
            cur_batch = pending[idx: idx + batch_size]
            try:
                print(f'  {idx}/{len(pending)} bs={len(cur_batch)}', end=' ')
                out = call_llm_batch(model, system_prompt, cur_batch)

                out_by_id = {str(int(o.get('row_idx'))): o for o in out if isinstance(o, dict) and 'row_idx' in o}

                updated = 0
                for src in cur_batch:
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

                    results[rid] = {'row_idx': int(rid), 'fills': sanitized}
                    updated += 1

                ckpt['results'] = results
                save_checkpoint(ckpt)
                print(f'✓ updated={updated}')

                idx += batch_size
                time.sleep(0.2)

            except Exception as e:
                err = str(e)
                print(f'✗ {err[:140]}')
                if ('not a json' in err.lower()) or ('json' in err.lower()) or ('length' in err.lower()):
                    if batch_size > MIN_BATCH:
                        batch_size = max(MIN_BATCH, batch_size // 2)
                        print(f'  reduce_batch_size={batch_size}')
                        time.sleep(2)
                        continue
                raise

    # 7. Export final results
    final_list = list(results.values())
    final_list.sort(key=lambda x: int(x.get('row_idx', 0)))
    OUT_LLM_JSON.write_text(json.dumps(final_list, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'\n✓ LLM done: {len(final_list)}/{len(items)} rows → {OUT_LLM_JSON.name}')
    print(f'✓ Checkpoint: {CKPT_JSON.name}')


if __name__ == '__main__':
    main()
