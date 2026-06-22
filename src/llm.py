import json
import time
import random
import httpx
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from google import genai
from google.genai import types

import config

def init_llm_client() -> tuple[genai.Client, str]:
    """Initialize client based on environment: Vertex AI or AI Studio."""
    use_vertex = os.getenv("USE_VERTEX", "True").lower() in ("true", "1", "yes")
    sa_key_path = config.PROJECT_ROOT / "sa-key.json"
    
    if use_vertex and sa_key_path.exists():
        print(">>> Using Google Vertex AI client...")
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = str(sa_key_path)
        project_id = os.getenv("VERTEX_PROJECT")
        if not project_id:
            try:
                with open(sa_key_path, "r", encoding="utf-8") as f:
                    project_id = json.load(f).get("project_id")
            except Exception:
                pass
        client = genai.Client(
            vertexai=True,
            project=project_id,
            location=os.getenv("VERTEX_LOCATION", "us-central1")
        )
        client._api_client._httpx_client.timeout = httpx.Timeout(120.0)
        model_name = config.MODEL_NAME
        if model_name.startswith("models/"):
            model_name = model_name[len("models/"):]
    else:
        print(">>> Using Google AI Studio client (Gemini API Key)...")
        api_key = config.API_KEY or os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("No GEMINI_API_KEY environment variable found!")
        client = genai.Client(api_key=api_key)
        client._api_client._httpx_client.timeout = httpx.Timeout(120.0)
        model_name = config.MODEL_NAME
        if not model_name.startswith("models/"):
            model_name = f"models/{model_name}"
            
    return client, model_name

# Rate limiting control
_last_call_time = 0.0

def wait_for_rate_limit():
    global _last_call_time
    now = time.time()
    elapsed = now - _last_call_time
    # Add rate limiting to serialization to prevent 429
    interval = config.MIN_INTERVAL_S + random.random() * config.JITTER_S
    if elapsed < interval:
        time.sleep(interval - elapsed)
    _last_call_time = time.time()

def _parse_llm_json(text: str) -> List[Dict[str, Any]]:
    # Extract JSON Array from prompt output
    start = text.find('[')
    end = text.rfind(']')
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not parse valid JSON array from LLM response: {text[:200]}")

def call_llm_batch(
    client: genai.Client,
    model_name: str,
    system_prompt: str,
    batch: List[Dict[str, Any]],
    max_retry: int = 5,
) -> List[Dict[str, Any]]:
    """Calls Gemini to fill in classification tags for a batch of rows."""
    payload = json.dumps(batch, ensure_ascii=False)
    user_input = "INPUT_JSON_ARRAY:\n" + payload

    for attempt in range(1, max_retry + 1):
        try:
            wait_for_rate_limit()
            resp = client.models.generate_content(
                model=model_name,
                contents=user_input,
                config=types.GenerateContentConfig(
                    system_instruction=system_prompt,
                    temperature=0.0,
                    max_output_tokens=32000,
                )
            )
            raw = getattr(resp, "text", "") or ""
            return _parse_llm_json(raw)
        except Exception as e:
            msg = str(e).encode("ascii", errors="replace").decode("ascii")
            low = msg.lower()
            if "429" in low or "rate limit" in low or "resource_exhausted" in low:
                wait = min(120, 10 * attempt) + random.random() * 2
                print(f"[WARN] API Rate limit/Resource exhausted. Sleeping {wait:.1f}s before retry {attempt}...")
                time.sleep(wait)
                continue
            if attempt < max_retry:
                time.sleep(4.0 * attempt)
                continue
            raise RuntimeError(f"Failed calling Gemini API after {max_retry} retries. Error: {msg}")
    raise RuntimeError("Failed calling Gemini API due to exhausted retries.")

# Helper to import Path and os inside file since they are used in init_llm_client
from pathlib import Path
import os
import re

# Date normalization regex patterns
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
