import json
import logging
import threading
import time
import random
import httpx
import os
from pathlib import Path
from typing import Any, Dict, List, Optional
from google import genai
from google.genai import types

import config

# --- Nhanh `gateway`: goi qua API Gateway noi bo thay vi goi thang Google ------
#
# THEM 09/09/2026. Hai nhanh cu (`vertex`, `apikey`) KHONG sua mot dong nao, va
# mac dinh van la chung -- lui lai chi can bo bien GEMINI_BACKEND.
#
# VI SAO LA MOT LOP DUNG THAY chu khong sua `call_llm_batch`:
# `call_llm_batch` goi `client.models.generate_content(...)` roi doc `resp.text`.
# Nhanh nay tra ve mot doi tuong CO DUNG hinh dang do, nen `call_llm_batch` khong
# phai sua gi -- ke ca vong thu lai va cach nhan 429 cua no.
#
# CAI BAY QUAN TRONG NHAT: `call_llm_batch` nhan ra "qua han muc" bang cach TIM
# CHUOI trong thong bao loi --
#     if "429" in low or "rate limit" in low or "resource_exhausted" in low:
#         wait = min(120, 10 * attempt) + ...     # lui LICH SU: 10, 20, 30 giay
#     if attempt < max_retry:
#         time.sleep(4.0 * attempt)               # lui CHUNG: 4, 8 giay
# -- nen loi tu day PHAI mang chuoi 429. Thieu no thi CRM lui 4-8 giay thay vi
# 10-20-30, tuc la dap vao mot Gateway vua xin no cham lai.
class _GatewayResponse:
    """Chi can mot thuoc tinh `text`, vi `call_llm_batch` chi doc thuoc tinh do."""

    def __init__(self, text: str):
        self.text = text


class _GatewayModels:
    def __init__(self, base_url: str, api_key: str, user: str, timeout_s: float):
        self._base_url = base_url.rstrip("/")
        self._api_key = api_key
        self._user = user
        self._timeout = httpx.Timeout(timeout_s)

    def generate_content(self, model: str, contents, config=None):
        # Tham so ten `config` BUOC phai giu ten do, vi ben goi truyen bang keyword.
        # No che mat module `config` o pham vi ham nay -- khong dung module o day.
        system_instruction = getattr(config, "system_instruction", None) if config else None
        temperature = getattr(config, "temperature", None) if config else None
        max_output_tokens = getattr(config, "max_output_tokens", None) if config else None
        response_mime_type = getattr(config, "response_mime_type", None) if config else None

        # CRM luon truyen `contents` la chuoi. Neu mai kia ai doi no thanh list thi
        # str() se gui repr cua Python len model -- sai IM LANG. Chan o day.
        if not isinstance(contents, str):
            raise RuntimeError(
                "Nhanh gateway chi nhan contents la chuoi, nhan duoc "
                + type(contents).__name__)

        messages = []
        if system_instruction:
            messages.append({"role": "system", "content": str(system_instruction)})
        messages.append({"role": "user", "content": contents})

        # Gateway khai ten nhom KHONG co tien to `models/`, con nhanh apikey thi TU
        # THEM tien to do va README dan dat GEMINI_MODEL=models/gemini-2.5-flash.
        model_name = model[len("models/"):] if model.startswith("models/") else model

        body = {"model": model_name, "messages": messages}
        if temperature is not None:
            body["temperature"] = temperature
        if max_output_tokens:
            body["max_tokens"] = max_output_tokens
        if response_mime_type == "application/json":
            # Gateway dat `drop_params: true`, nen neu nha cung cap khong nhan tham
            # so nay thi LiteLLM BO no ma khong bao. `_parse_llm_json` du phong thu
            # de van chay, nen dung coi "ket qua van dung" la bang chung rang che do
            # JSON da toi noi -- phai do rieng.
            body["response_format"] = {"type": "json_object"}

        headers = {
            "Authorization": "Bearer " + self._api_key,
            "Content-Type": "application/json",
            # Dinh danh DICH VU, khong bia ra mot con nguoi. Cot end_user cua so
            # nhan gia tri nay; dien bua mot cai ten se lam hong chieu nguoi dung
            # cua ca dashboard.
            "X-User": self._user,
        }

        try:
            r = httpx.post(self._base_url + "/v1/chat/completions",
                           json=body, headers=headers, timeout=self._timeout)
        except Exception as e:
            raise RuntimeError(
                "Gateway khong ket noi duoc: " + type(e).__name__ + ": " + str(e)) from e

        if r.status_code == 429:
            # PHAI co chuoi 429 -- xem ghi chu dau khoi nay.
            raise RuntimeError("429 qua han muc tu Gateway: " + r.text[:200])
        if r.status_code >= 400:
            raise RuntimeError(
                "Gateway tra " + str(r.status_code) + ": " + r.text[:300])

        # Bao ve ca buoc phan tich JSON: 200 ma than khong phai JSON thi r.json()
        # nem mot loi kho doc, chang tro vao dau.
        try:
            d = r.json()
            content = d["choices"][0]["message"]["content"]
        except Exception as e:
            raise RuntimeError(
                "Gateway tra 200 nhung than phan hoi ngoai du kien ("
                + type(e).__name__ + "): " + r.text[:300]) from e

        # KHONG doi noi dung rong thanh chuoi rong. `call_llm_batch` doc `.text` roi
        # day sang `_parse_llm_json`, va chuoi rong se ra "Could not parse valid JSON
        # array" -- trieu chung tro vao PROMPT chu khong tro vao cho sai that.
        if not content:
            raise RuntimeError("Gateway tra 200 nhung noi dung tra loi rong")
        return _GatewayResponse(content)


class _GatewayClient:
    def __init__(self, base_url: str, api_key: str, user: str, timeout_s: float):
        self.models = _GatewayModels(base_url, api_key, user, timeout_s)


# --- Tu doi sang duong thang khi Gateway chet -------------------------------
#
# THEM 12/09/2026. `gateway-lb` la DIEM HONG DON: hai instance LiteLLM co du phong
# cho nhau, nhung nginx chi co mot. Dien tap 10/09 do duoc: mat no thi CRM thu ba
# lan trong 21-24 giay roi BO CA LO, va khong ai duoc bao.
#
# VI SAO LAI LA MOT LOP DUNG THAY: giong het ly do cua `_GatewayClient` o tren --
# `call_llm_batch` va `pipeline.py` khong phai sua mot dong nao.
#
# CAI GIA PHAI TRA, GHI RO DE NGUOI SAU KHONG TUONG NHAM: duong du phong dung khoa
# Google that, nen tinh chat "agent khong giu khoa nha cung cap" cua Gateway
# KHONG CON DUNG cho agent nay. Chot boi lead 10/09: he thong on dinh truoc,
# do dac sau.

ERR_UNREACHABLE = "unreachable"
ERR_RATE_LIMIT = "rate_limit"
ERR_REQUEST = "request_error"


def classify_error(e: Exception) -> str:
    """Ba cau loi cua `_GatewayModels` phan biet duoc voi nhau. CHI loai dau doi duong.

    Loi 400 doi duong cung sai y het, chi ton them mot luot goi o noi khac.
    Loi qua han muc giu nhanh lui lich san co: han muc AO cua Gateway va han muc
    THAT cua Google chua phan biet duoc tu phia agent (xem Open Questions).

    Nhan dang bang chuoi la mong manh, nhung chinh repo da dat cuoc vao do tu truoc
    -- xem ghi chu dau file ve chuoi `429`. Phep kiem o `tests/` giu giao keo do.
    """
    msg = str(e).lower()
    if "429" in msg or "rate limit" in msg or "resource_exhausted" in msg:
        return ERR_RATE_LIMIT
    if "khong ket noi duoc" in msg:
        return ERR_UNREACHABLE
    return ERR_REQUEST


class _FallbackModels:
    def __init__(self, outer: "_FallbackClient"):
        self._outer = outer

    def generate_content(self, **kwargs):
        return self._outer._call(**kwargs)


class _FallbackClient:
    """Giu duong dang dung, va tu doi khi Gateway chet N lan LIEN TIEP.

    Dem lien tiep chu khong dem tong: mot luot hong le la chuyen thuong, ba luot
    lien tiep moi la tuyen chet. Bo dem nam duoi mot khoa vi CRM chay 3 worker
    song song dung chung client (`pipeline.py:600,610`) -- thieu khoa thi ba worker
    cung vuot threshold, cung dung client moi va cung gui thu.
    """

    def __init__(self, gateway_client, make_fallback, threshold: int, retry_after_s: float):
        self._gateway = gateway_client
        self._make_fallback = make_fallback      # ham dung client du phong, goi luoi
        self._fallback_client = None
        self._threshold = threshold
        self._retry_after_s = retry_after_s
        self._lock = threading.Lock()
        self._on_fallback = False
        self._consecutive_failures = 0
        self._last_probe_at = 0.0
        self._fallback_calls = 0
        self.models = _FallbackModels(self)

    def _get_fallback_client(self):
        if self._fallback_client is None:
            self._fallback_client = self._make_fallback()
        return self._fallback_client

    def _call(self, **kwargs):
        # Dang di duong thang va da qua han cho -> THAM DO Gateway bang mot luot that.
        # Chi doi trang thai khi luot do THANH CONG; hong thi im lang dung tiep duong
        # thang, khong tinh vao bo dem. Tham do dinh ky chu khong tham do moi luot:
        # moi luot thi moi luot phai tra gia mot lan cho het thoi gian ket noi.
        if self._on_fallback and time.time() - self._last_probe_at >= self._retry_after_s:
            try:
                r = self._gateway.models.generate_content(**kwargs)
                self._back_to_gateway()
                return r
            except Exception:
                self._last_probe_at = time.time()

        if self._on_fallback:
            return self._log_and_call(**kwargs)

        try:
            r = self._gateway.models.generate_content(**kwargs)
        except Exception as e:
            if classify_error(e) == ERR_UNREACHABLE and self._count_failure(e):
                # Vua doi duong xong -> chay NOT luot nay tren duong thang, khong
                # de lo bi bo. Do la toan bo diem cua change nay.
                return self._log_and_call(**kwargs)
            raise
        with self._lock:
            self._consecutive_failures = 0
        return r

    def _log_and_call(self, **kwargs):
        """Ghi lai TUNG luot di duong thang roi moi goi.

        Trong khoang nay Gateway KHONG ghi so. Neu agent cung khong ghi thi khoang
        thoi gian ay chi con biet qua hoa don -- cham mot ngay va khong co chieu
        nguoi dung. Nhat ky da co moc thoi gian san.
        """
        with self._lock:
            self._fallback_calls += 1
            nth = self._fallback_calls
        log_fallback("luot %d di duong thang (ly do: Gateway mat ket noi)" % nth)
        return self._get_fallback_client().models.generate_content(**kwargs)

    def _count_failure(self, exc: Exception = None) -> bool:
        """Tra True DUNG MOT LAN, cho luong lam bo dem cham threshold.

        GHI CA CAU LOI THAT. Dien tap 12/09 cho thay vi sao: fallback lam viec cua no
        qua tot -- 60/60 lo chay xong, 0 lo nem loi ra ngoai -- nen KHONG CO CHO NAO
        con luu lai Gateway da hong kieu gi. Dung container cho loi phan giai ten,
        cong chet cho loi tu choi ket noi; hai kieu do doi hoi hai cach xu khac nhau
        khi len server, va khong ghi thi khong phan biet duoc.
        """
        with self._lock:
            if self._on_fallback:
                return True
            self._consecutive_failures += 1
            if self._consecutive_failures < self._threshold:
                return False
            self._on_fallback = True
            self._last_probe_at = time.time()
            self._outage_started_at = time.time()
        log_fallback("DOI DUONG: Gateway mat ket noi %d lan lien tiep -> Google AI Studio"
                        % self._threshold)
        log_fallback("cau loi cuoi cung tu Gateway: %s" % str(exc)[:300])
        _notify("[CRM Pipeline] Gateway mat - da chuyen sang goi thang",
                     "<p>Gateway khong ket noi duoc <b>%d lan lien tiep</b>.</p>"
                     "<p>Agent da chuyen sang goi thang nha cung cap va <b>van chay tiep</b>.</p>"
                     "<p>Duong du phong: Google AI Studio, dung FALLBACK_API_KEY. "
                     "Cac luot nay khong duoc Gateway ghi SpendLogs.</p>"
                     % self._threshold)
        return True

    def _back_to_gateway(self) -> None:
        with self._lock:
            if not self._on_fallback:
                return
            elapsed = time.time() - getattr(self, "_outage_started_at", time.time())
            self._on_fallback = False
            self._consecutive_failures = 0
        log_fallback("VE DUONG CU: Gateway song lai sau %.0f giay, %d luot da di duong thang"
                        % (elapsed, self._fallback_calls))
        _notify("[CRM Pipeline] Gateway song lai - da quay ve",
                     "<p>Gateway ket noi lai duoc. Agent da quay ve goi qua Gateway.</p>"
                     "<p>Su co keo dai <b>%.0f giay</b>, co <b>%d luot</b> da di duong thang.</p>"
                     % (elapsed, self._fallback_calls))


def log_fallback(cau: str) -> None:
    """Ghi lai moi lan doi duong. Trong khoang do Gateway KHONG ghi so; neu agent
    cung khong ghi thi khoang thoi gian ay chi con biet qua hoa don -- cham mot ngay
    va khong co chieu nguoi dung."""
    logging.getLogger("crm-automation").warning("[FALLBACK] %s", cau)


def _notify(tieu_de: str, than: str) -> None:
    # Import tai cho: `notification` keo theo `msal`, va `llm.py` phai nap duoc ca
    # khi chua cai goi do.
    try:
        from notification import send_alert
        send_alert(tieu_de, than)
    except Exception as e:
        logging.getLogger("crm-automation").error("Khong gui duoc thu canh bao: %s", e)


def _build_ai_studio_client():
    """Dung key Google rieng; GEMINI_API_KEY dang giu virtual key cua Gateway."""
    api_key = (os.getenv("FALLBACK_API_KEY") or "").strip()
    if not api_key:
        raise ValueError("Duong du phong Google AI Studio can FALLBACK_API_KEY")
    client = genai.Client(vertexai=False, api_key=api_key)
    client._api_client._httpx_client.timeout = httpx.Timeout(300.0)
    return client


def init_llm_client() -> tuple[genai.Client | _GatewayClient, str]:
    """Initialize client based on environment: Gateway, Vertex AI, or AI Studio."""
    # Nhanh `gateway` kiem TRUOC va tra ve NGAY, nen hai nhanh cu ben duoi khong
    # doi mot dong nao va mac dinh van la chung.
    if (os.getenv("GEMINI_BACKEND") or "").strip().lower() == "gateway":
        # `GEMINI_API_KEY` o day la VIRTUAL KEY cua Gateway, khong phai khoa Google.
        # Key Google cho duong du phong nam rieng trong FALLBACK_API_KEY.
        api_key = config.API_KEY or os.getenv("GEMINI_API_KEY") or ""
        if not api_key:
            raise ValueError(
                "GEMINI_BACKEND=gateway nhung thieu GEMINI_API_KEY (Virtual Key cua Gateway)")
        base_url = os.getenv("GEMINI_GATEWAY_BASE_URL") or "http://gateway-lb:4000"
        user = os.getenv("GEMINI_GATEWAY_USER") or "svc.crm-feedback"
        print(">>> Using internal API Gateway at " + base_url + " ...")
        # 300 giay -- BANG voi timeout ma hai nhanh cu dat cho SDK cua Google. Mot
        # batch 25 dong voi 8192 token ra khong nhanh; va da quan sat mot luot qua
        # Gateway mat 81 giay (chua tai hien duoc, chua ro nguyen nhan).
        gw = _GatewayClient(base_url, api_key, user, 300.0)
        # Mac dinh TAT. Thieu key AI Studio thi tu tat va keu TO luc khoi dong -- nhung
        # agent VAN khoi dong duoc, vi thieu duong du phong khong phai ly do chan he thong.
        if getattr(config, "FALLBACK_ENABLED", False):
            if (os.getenv("FALLBACK_API_KEY") or "").strip():
                print(">>> Fallback bat: Gateway chet %d lan lien tiep thi goi Google AI Studio"
                      % config.FALLBACK_FAIL_THRESHOLD)
                gw = _FallbackClient(gw, _build_ai_studio_client,
                                     config.FALLBACK_FAIL_THRESHOLD,
                                     config.FALLBACK_RETRY_AFTER_S)
            else:
                logging.getLogger("crm-automation").error(
                    "FALLBACK_ENABLED=True nhung thieu FALLBACK_API_KEY -> TAT fallback. "
                    "Gateway chet se lam bo lo nhu truoc.")
        return gw, config.MODEL_NAME
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
        client._api_client._httpx_client.timeout = httpx.Timeout(300.0)
        model_name = config.MODEL_NAME
        if model_name.startswith("models/"):
            model_name = model_name[len("models/"):]
    else:
        print(">>> Using Google AI Studio client (Gemini API Key)...")
        api_key = config.API_KEY or os.getenv("GEMINI_API_KEY")
        if not api_key:
            raise ValueError("No GEMINI_API_KEY environment variable found!")
        client = genai.Client(api_key=api_key)
        client._api_client._httpx_client.timeout = httpx.Timeout(300.0)
        model_name = config.MODEL_NAME
        if not model_name.startswith("models/"):
            model_name = f"models/{model_name}"
            
    return client, model_name

import threading

# Rate limiting control
_last_call_time = 0.0
_rate_limit_lock = threading.Lock()

def wait_for_rate_limit():
    global _last_call_time
    with _rate_limit_lock:
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
        candidate = text[start:end + 1]
        
        # 1. Try standard JSON parse
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            pass
            
        # 2. Try regex-based trailing comma repair (very common LLM syntax error)
        import re
        repaired = re.sub(r',\s*\]', ']', candidate)
        repaired = re.sub(r',\s*\}', '}', repaired)
        try:
            return json.loads(repaired)
        except json.JSONDecodeError:
            pass
            
        # 3. Try json_repair library fallback (if installed)
        try:
            from json_repair import repair_json
            parsed = repair_json(candidate, return_objects=True)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            pass

        # If all repairs fail, save to log for debugging
        try:
            log_path = config.PATH_OUTPUT / "logs" / "failed_llm_response.txt"
            log_path.parent.mkdir(parents=True, exist_ok=True)
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(text)
        except Exception:
            pass

        # Re-run json.loads to raise original JSONDecodeError description
        try:
            json.loads(candidate)
        except json.JSONDecodeError as jde:
            raise ValueError(f"JSONDecodeError: {jde}. Raw response saved to logs/failed_llm_response.txt. Sample: {text[:200]}")
            
    try:
        log_path = config.PATH_OUTPUT / "logs" / "failed_llm_response.txt"
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception:
        pass
    raise ValueError(f"Could not parse valid JSON array from LLM response (saved to logs/failed_llm_response.txt). Sample: {text[:200]}")


def call_llm_batch(
    client: genai.Client,
    model_name: str,
    system_prompt: str,
    batch: List[Dict[str, Any]],
    max_retry: int = 3,
):
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
                    max_output_tokens=8192,
                    response_mime_type="application/json",
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
