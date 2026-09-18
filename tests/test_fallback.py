"""Phep kiem cho viec tu doi sang duong thang khi Gateway chet.

Chay:  python tests/test_fallback.py

Khong dung framework: dung `assert` va mot func `main()`. Cac phep kiem dung client
GIA, nen khong goi mang, khong ton tien, worker trong mili giay.
"""
import sys
import threading
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import llm  # noqa: E402


class _FakeModels:
    def __init__(self, fail_times, name):
        self.remaining_failures = fail_times
        self.name = name
        self.call_count = 0

    def generate_content(self, **kwargs):
        self.call_count += 1
        if self.remaining_failures > 0:
            self.remaining_failures -= 1
            raise RuntimeError("Gateway khong ket noi duoc: ConnectError: test")
        return type("R", (), {"text": self.name})()


class _FakeClient:
    def __init__(self, fail_times=0, name="ok"):
        self.models = _FakeModels(fail_times, name)


def test_three_error_strings_stay_distinct():
    """Ba cau loi cua _GatewayModels PHAI phan biet duoc voi nhau.

    Phep kiem nay ton tai vi viec phan loai dua tren NOI DUNG cau loi: sua mot cau
    chu ma khong ai sent se lam fallback IM LANG ngung hoat dong.
    """
    assert llm.classify_error(RuntimeError(
        "Gateway khong ket noi duoc: ConnectError: [Errno -5]")) == llm.ERR_UNREACHABLE
    assert llm.classify_error(RuntimeError(
        "429 qua han muc tu Gateway: rate limit")) == llm.ERR_RATE_LIMIT
    assert llm.classify_error(RuntimeError(
        "Gateway tra 400: bad request")) == llm.ERR_REQUEST
    assert llm.classify_error(RuntimeError(
        "Gateway tra 503: upstream")) == llm.ERR_REQUEST


def test_switches_at_threshold_and_drops_no_batch():
    """Hong du threshold thi doi duong, VA luot do van worker xong tren duong thang."""
    gw = _FakeClient(fail_times=99, name="gateway")
    c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=3, retry_after_s=999)

    for _ in range(2):
        try:
            c.models.generate_content(model="m", contents="x")
            raise AssertionError("dang le phai nem loi khi chua du threshold")
        except RuntimeError:
            pass
    assert not c._on_fallback, "chua du threshold ma da doi duong"

    r = c.models.generate_content(model="m", contents="x")
    assert c._on_fallback, "du threshold ma khong doi duong"
    assert r.text == "thang", "doi duong roi nhung luot do bi bo, khong worker lai"


def test_single_failure_does_not_switch():
    gw = _FakeClient(fail_times=1, name="gateway")
    c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=3, retry_after_s=999)
    try:
        c.models.generate_content(model="m", contents="x")
    except RuntimeError:
        pass
    assert c.models.generate_content(model="m", contents="x").text == "gateway"
    assert c._consecutive_failures == 0, "luot thanh cong phai dat bo dem ve khong"
    assert not c._on_fallback


def test_request_error_does_not_switch():
    """Loi 400 doi duong cung sai y het, chi ton them mot luot goi o noi khac."""
    class _Hong400:
        def generate_content(self, **kw):
            raise RuntimeError("Gateway tra 400: bad request")

    gw = _FakeClient()
    gw.models = _Hong400()
    c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=1, retry_after_s=999)
    for _ in range(5):
        try:
            c.models.generate_content(model="m", contents="x")
        except RuntimeError:
            pass
    assert not c._on_fallback, "loi 400 khong duoc kich hoat doi duong"


def test_many_threads_switch_once():
    """Ba worker cung gap loi trong cung mot su co -> doi duong DUNG mot lan."""
    # Dem SU KIEN doi duong, khong dem he qua cua no. Truoc day dem so thu gui di;
    # thu chi la dai luong thay the, con dong log `DOI DUONG` MOI la su kien that.
    counter = {"doi": 0}
    original = llm.log_fallback
    llm.log_fallback = lambda cau: counter.__setitem__(
        "doi", counter["doi"] + cau.startswith("DOI DUONG"))
    try:
        gw = _FakeClient(fail_times=999, name="gateway")
        c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=3, retry_after_s=999)

        def worker():
            for _ in range(10):
                try:
                    c.models.generate_content(model="m", contents="x")
                except RuntimeError:
                    pass

        ts = [threading.Thread(target=worker) for _ in range(3)]
        [t.start() for t in ts]
        [t.join() for t in ts]
        assert c._on_fallback
        assert counter["doi"] == 1, f"doi duong {counter['doi']} lan cho mot su co, dang le 1"
    finally:
        llm.log_fallback = original


def test_probe_returns_to_gateway():
    """Qua han cho thi THAM DO Gateway; chi doi trang thai khi tham do THANH CONG."""
    counter = {"moc": 0}
    original = llm.log_fallback
    llm.log_fallback = lambda cau: counter.__setitem__(
        "moc", counter["moc"] + cau.startswith(("DOI DUONG", "VE DUONG CU")))
    try:
        gw = _FakeClient(fail_times=3, name="gateway")
        c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=3, retry_after_s=0.05)
        for _ in range(3):
            try:
                c.models.generate_content(model="m", contents="x")
            except RuntimeError:
                pass
        assert c._on_fallback
        assert counter["moc"] == 1

        time.sleep(0.06)
        r = c.models.generate_content(model="m", contents="x")
        assert r.text == "gateway", "tham do thanh cong ma khong dung ket qua Gateway"
        assert not c._on_fallback, "tham do thanh cong ma khong quay ve"
        assert counter["moc"] == 2, "mot su co phai sinh DUNG hai moc: di va ve"
    finally:
        llm.log_fallback = original


def test_failed_probe_keeps_fallback():
    gw = _FakeClient(fail_times=999, name="gateway")
    c = llm._FallbackClient(gw, lambda: _FakeClient(name="thang"), threshold=1, retry_after_s=0.05)
    try:
        c.models.generate_content(model="m", contents="x")
    except RuntimeError:
        pass
    assert c._on_fallback
    time.sleep(0.06)
    r = c.models.generate_content(model="m", contents="x")
    assert r.text == "thang", "tham do hong ma khong dung duong thang"
    assert c._on_fallback, "tham do hong ma da quay ve Gateway"


def test_nguong_khong_duoc_lon_hon_so_lan_thu_lai():
    """FALLBACK_FAIL_THRESHOLD PHAI <= max_retry cua `call_llm_batch`.

    Lon hon thi lo DAU TIEN gap su co van bi bo -- dung cai ma change nay sinh ra
    de tranh. Khong ai thay duoc dieu do khi chay: lo bi bo trong im lang.
    """
    import inspect
    import config
    max_retry = inspect.signature(llm.call_llm_batch).parameters["max_retry"].default
    assert config.FALLBACK_FAIL_THRESHOLD <= max_retry, (
        "nguong %d > max_retry %d -> lo dau tien se bi bo"
        % (config.FALLBACK_FAIL_THRESHOLD, max_retry))


def main():
    passed = 0
    for name, func in sorted(globals().items()):
        if name.startswith("test_"):
            func()
            print("  dat  ", name)
            passed += 1
    print(f"\n{passed} phep kiem, {passed} dat, 0 hong")


if __name__ == "__main__":
    main()
