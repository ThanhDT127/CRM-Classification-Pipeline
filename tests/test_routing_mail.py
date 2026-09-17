"""Phep kiem cho nhanh `gateway` (routing) va duong gui thu SMTP.

Chay:  python tests/test_routing_mail.py

Khong dung framework va khong cham mang ngoai: Gateway gia bang `http.server` tren
127.0.0.1, SMTP gia bang mot lop thay `smtplib.SMTP`. Khong ton tien, khong co thu
nao roi vao hop thu nguoi that.

KHONG DAU trong file nay la CO Y, giong `llm.py`: `call_llm_batch` chay cau loi qua
`str(e).encode("ascii", errors="replace")` truoc khi tim chuoi trong do. Chu co dau
se bien thanh `?` va lam hong viec nhan dang loi.
"""
import json
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import llm  # noqa: E402

REPLY_OK = '{"choices":[{"message":{"content":"[]"}}]}'


# --- Gateway gia ------------------------------------------------------------

class _Ghi:
    """Cho ta doc lai request cuoi cung, va dat truoc cau tra loi ke tiep."""
    body = None
    headers = None
    status = 200
    reply = REPLY_OK


class _Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        n = int(self.headers.get("Content-Length") or 0)
        _Ghi.body = json.loads(self.rfile.read(n) or b"{}")
        _Ghi.headers = dict(self.headers)
        self.send_response(_Ghi.status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(_Ghi.reply.encode("utf-8"))

    def log_message(self, *a):
        pass


def _mo_gateway_gia():
    sv = HTTPServer(("127.0.0.1", 0), _Handler)
    threading.Thread(target=sv.serve_forever, daemon=True).start()
    return sv, "http://127.0.0.1:%d" % sv.server_address[1]


class _Cfg:
    """Thay cho `types.GenerateContentConfig` -- chi can dung nhung thuoc tinh do."""

    def __init__(self, **kw):
        self.system_instruction = kw.get("system_instruction")
        self.temperature = kw.get("temperature")
        self.max_output_tokens = kw.get("max_output_tokens")
        self.response_mime_type = kw.get("response_mime_type")


# --- Routing ----------------------------------------------------------------

def test_routing_dung_hinh_dang_gateway():
    """Than request phai dung dang OpenAI, va ten model phai RUNG tien to `models/`.

    Gateway khai ten nhom la `gemini-2.5-flash`, con README cua CRM dan dat
    `models/gemini-2.5-flash`. Khong rung thi ra "model not found".
    """
    sv, url = _mo_gateway_gia()
    try:
        m = llm._GatewayModels(url, "sk-ao", "svc.crm-feedback", 10.0)
        m.generate_content(model="models/gemini-2.5-flash", contents="xin chao",
                           config=_Cfg(system_instruction="he thong", temperature=0.0,
                                       max_output_tokens=8192,
                                       response_mime_type="application/json"))
        b = _Ghi.body
        assert b["model"] == "gemini-2.5-flash", b["model"]
        assert b["messages"][0] == {"role": "system", "content": "he thong"}
        assert b["messages"][1] == {"role": "user", "content": "xin chao"}
        assert b["temperature"] == 0.0
        assert b["max_tokens"] == 8192
        assert b["response_format"] == {"type": "json_object"}
        assert _Ghi.headers["X-User"] == "svc.crm-feedback"
        assert _Ghi.headers["Authorization"] == "Bearer sk-ao"
    finally:
        sv.shutdown()


def test_routing_429_giu_duoc_chuoi_429():
    """Giao keo quan trong nhat: loi qua han muc PHAI mang chuoi `429`.

    `call_llm_batch` nhan ra qua han muc bang cach tim chuoi trong cau loi. Mat chuoi
    do thi no lui 4-8 giay thay vi 10-20-30, tuc la dap vao mot Gateway vua xin
    no cham lai.
    """
    sv, url = _mo_gateway_gia()
    try:
        _Ghi.status, _Ghi.reply = 429, "rate limit"
        m = llm._GatewayModels(url, "sk-ao", "u", 10.0)
        try:
            m.generate_content(model="m", contents="x", config=None)
            raise AssertionError("429 ma khong nem loi")
        except RuntimeError as e:
            # Di qua DUNG cai bo loc ma call_llm_batch dung.
            sach = str(e).encode("ascii", errors="replace").decode("ascii").lower()
            assert "429" in sach, sach
            assert llm.classify_error(e) == llm.ERR_RATE_LIMIT
    finally:
        _Ghi.status, _Ghi.reply = 200, REPLY_OK
        sv.shutdown()


def test_routing_loi_400_khong_bi_coi_la_mat_ket_noi():
    sv, url = _mo_gateway_gia()
    try:
        _Ghi.status, _Ghi.reply = 400, "bad request"
        m = llm._GatewayModels(url, "sk-ao", "u", 10.0)
        try:
            m.generate_content(model="m", contents="x", config=None)
            raise AssertionError("400 ma khong nem loi")
        except RuntimeError as e:
            assert llm.classify_error(e) == llm.ERR_REQUEST
    finally:
        _Ghi.status, _Ghi.reply = 200, REPLY_OK
        sv.shutdown()


def test_routing_noi_dung_rong_thi_keu_len():
    """200 ma noi dung rong PHAI nem loi, khong duoc tra chuoi rong.

    Tra chuoi rong thi `_parse_llm_json` keu "Could not parse valid JSON array" --
    trieu chung tro vao PROMPT chu khong tro vao cho sai that.
    """
    sv, url = _mo_gateway_gia()
    try:
        _Ghi.reply = '{"choices":[{"message":{"content":""}}]}'
        m = llm._GatewayModels(url, "sk-ao", "u", 10.0)
        try:
            m.generate_content(model="m", contents="x", config=None)
            raise AssertionError("noi dung rong ma van tra ve")
        except RuntimeError as e:
            assert "rong" in str(e).lower(), str(e)
    finally:
        _Ghi.reply = REPLY_OK
        sv.shutdown()


def test_routing_than_phan_hoi_la_thu_khac():
    sv, url = _mo_gateway_gia()
    try:
        _Ghi.reply = "khong phai json"
        m = llm._GatewayModels(url, "sk-ao", "u", 10.0)
        try:
            m.generate_content(model="m", contents="x", config=None)
            raise AssertionError("than la rac ma khong nem loi")
        except RuntimeError as e:
            assert llm.classify_error(e) == llm.ERR_REQUEST
    finally:
        _Ghi.reply = REPLY_OK
        sv.shutdown()


def test_routing_contents_khong_phai_chuoi_thi_chan():
    """str() tren mot list se gui repr cua Python len model -- sai IM LANG."""
    m = llm._GatewayModels("http://127.0.0.1:1", "sk-ao", "u", 2.0)
    try:
        m.generate_content(model="m", contents=["a", "b"], config=None)
        raise AssertionError("contents la list ma khong bi chan")
    except RuntimeError as e:
        assert "chuoi" in str(e), str(e)


def test_routing_mat_ket_noi_thi_phan_loai_dung():
    """Cong dong -> phai ra ERR_UNREACHABLE, vi CHI loai nay moi duoc doi duong."""
    m = llm._GatewayModels("http://127.0.0.1:1", "sk-ao", "u", 2.0)
    try:
        m.generate_content(model="m", contents="x", config=None)
        raise AssertionError("cong dong ma khong nem loi")
    except RuntimeError as e:
        assert llm.classify_error(e) == llm.ERR_UNREACHABLE, str(e)


# --- Mail -------------------------------------------------------------------

class _SmtpGia:
    """Thay `smtplib.SMTP`. Ghi lai da lam gi, va co the gia vo hong."""

    ghi = {}
    hong = False

    def __init__(self, host, port, timeout=None):
        if _SmtpGia.hong:
            raise OSError("khong noi duoc toi may chu thu")
        _SmtpGia.ghi = {"host": host, "port": port, "starttls": False,
                        "login": None, "msg": None}

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False

    def starttls(self):
        _SmtpGia.ghi["starttls"] = True

    def login(self, u, p):
        _SmtpGia.ghi["login"] = (u, p)

    def send_message(self, msg):
        _SmtpGia.ghi["msg"] = msg


def _dat_config(nt, **kw):
    """Dat cac bien config cho mot phep kiem, tra ve ham hoan nguyen."""
    cu = {k: getattr(nt.config, k, None) for k in kw}

    def hoan():
        for k, v in cu.items():
            setattr(nt.config, k, v)

    for k, v in kw.items():
        setattr(nt.config, k, v)
    return hoan


def _voi_smtp_gia(nt, **kw):
    import smtplib
    smtp_cu = smtplib.SMTP
    smtplib.SMTP = _SmtpGia
    hoan_config = _dat_config(nt, **kw)

    def thoi():
        smtplib.SMTP = smtp_cu
        hoan_config()

    return thoi


def test_mail_duong_smtp_dung_dia_chi_va_bat_tls():
    import notification as nt
    thoi = _voi_smtp_gia(nt, MAIL_TRANSPORT="smtp", SMTP_HOST="smtp.gmail.com",
                         SMTP_PORT=587, SMTP_USERNAME="bot@rangdong.vn",
                         SMTP_USE_TLS=True, SMTP_PASSWORD="matkhau",
                         NOTIFICATION_SENDER_EMAIL="bot@rangdong.vn",
                         NOTIFICATION_RECIPIENTS=["a@x.vn", "b@x.vn"])
    try:
        _SmtpGia.hong = False
        assert nt.send_alert("tieu de", "<p>than</p>") is True
        g = _SmtpGia.ghi
        assert (g["host"], g["port"]) == ("smtp.gmail.com", 587)
        assert g["starttls"] is True, "SMTP_USE_TLS=True ma khong goi starttls"
        assert g["login"] == ("bot@rangdong.vn", "matkhau")
        # Nhieu nguoi nhan phai noi bang dau phay, khong phai repr cua list.
        assert g["msg"]["To"] == "a@x.vn, b@x.vn", g["msg"]["To"]
        assert g["msg"]["Subject"] == "tieu de"
        assert g["msg"].get_content_type() == "text/html", g["msg"].get_content_type()
    finally:
        thoi()


def test_mail_send_alert_chay_duoc_khi_khong_co_authprovider():
    """Ca ly do ton tai cua `send_alert`: `llm.py` khong co AuthProvider trong tay."""
    import notification as nt
    thoi = _voi_smtp_gia(nt, MAIL_TRANSPORT="smtp", SMTP_HOST="h", SMTP_PORT=25,
                         SMTP_USERNAME="", SMTP_USE_TLS=False, SMTP_PASSWORD="p",
                         NOTIFICATION_SENDER_EMAIL="bot@x.vn",
                         NOTIFICATION_RECIPIENTS=["a@x.vn"])
    try:
        _SmtpGia.hong = False
        assert nt.send_alert("t", "<p>b</p>") is True
        # Thieu SMTP_USERNAME thi dang nhap bang chinh dia chi gui.
        assert _SmtpGia.ghi["login"] == ("bot@x.vn", "p")
        assert _SmtpGia.ghi["starttls"] is False, "SMTP_USE_TLS=False ma van goi starttls"
    finally:
        thoi()


def test_mail_thieu_cau_hinh_thi_bo_qua_chu_khong_sap():
    import notification as nt
    for thieu in ({"SMTP_HOST": ""}, {"SMTP_PASSWORD": ""},
                  {"NOTIFICATION_SENDER_EMAIL": ""}, {"NOTIFICATION_RECIPIENTS": []}):
        day = dict(MAIL_TRANSPORT="smtp", SMTP_HOST="h", SMTP_PORT=25,
                   SMTP_USERNAME="u", SMTP_USE_TLS=False, SMTP_PASSWORD="p",
                   NOTIFICATION_SENDER_EMAIL="bot@x.vn",
                   NOTIFICATION_RECIPIENTS=["a@x.vn"])
        day.update(thieu)
        thoi = _voi_smtp_gia(nt, **day)
        try:
            assert nt.send_alert("t", "<p>b</p>") is False, thieu
        finally:
            thoi()


def test_mail_may_chu_hong_thi_khong_nem_loi_ra_ngoai():
    """Bao dong la viec phu, xu ly du lieu la viec chinh. Thu hong KHONG duoc lam sap lo."""
    import notification as nt
    thoi = _voi_smtp_gia(nt, MAIL_TRANSPORT="smtp", SMTP_HOST="h", SMTP_PORT=25,
                         SMTP_USERNAME="u", SMTP_USE_TLS=False, SMTP_PASSWORD="p",
                         NOTIFICATION_SENDER_EMAIL="bot@x.vn",
                         NOTIFICATION_RECIPIENTS=["a@x.vn"])
    try:
        _SmtpGia.hong = True
        assert nt.send_alert("t", "<p>b</p>") is False
    finally:
        _SmtpGia.hong = False
        thoi()


def test_mail_llm_notify_khong_bao_gio_nem_loi():
    """`_notify` duoc goi tu giua duong doi tuyen. No nem loi la mat ca lo."""
    import notification as nt

    def no_tung(*a, **k):
        raise RuntimeError("no tung")

    goc = nt.send_alert
    nt.send_alert = no_tung
    try:
        llm._notify("t", "<p>b</p>")   # khong duoc nem gi ra ngoai
    finally:
        nt.send_alert = goc


def test_mail_transport_graph_khong_di_duong_smtp():
    """MAIL_TRANSPORT=graph phai di duong Graph. Khong co AuthProvider thi tra False,
    chu khong duoc am tham di duong SMTP."""
    import notification as nt
    thoi = _voi_smtp_gia(nt, MAIL_TRANSPORT="graph", SMTP_HOST="h", SMTP_PORT=25,
                         SMTP_USERNAME="u", SMTP_USE_TLS=False, SMTP_PASSWORD="p",
                         NOTIFICATION_SENDER_EMAIL="bot@x.vn",
                         NOTIFICATION_RECIPIENTS=["a@x.vn"])
    try:
        _SmtpGia.ghi = {}
        assert nt.send_alert("t", "<p>b</p>") is False
        assert _SmtpGia.ghi == {}, "MAIL_TRANSPORT=graph ma van di duong SMTP"
    finally:
        thoi()


def main():
    passed = 0
    for ten, f in sorted(globals().items()):
        if ten.startswith("test_"):
            f()
            print("  dat  ", ten)
            passed += 1
    print("\n%d phep kiem, %d dat, 0 hong" % (passed, passed))


if __name__ == "__main__":
    main()
