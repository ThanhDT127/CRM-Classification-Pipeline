import sys
from pathlib import Path

# Add src to sys.path at the beginning
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pytest
from llm import _norm_ddmmyy
from classifier import _clean, _has_price

def test_norm_ddmmyy():
    # Test YMD format
    assert _norm_ddmmyy("2026-06-20") == "20/06/26"
    assert _norm_ddmmyy("2025-12-05") == "05/12/25"
    
    # Test DMY format
    assert _norm_ddmmyy("25/12/2025") == "25/12/25"
    assert _norm_ddmmyy("05-12-25") == "05/12/25"
    
    # Test Month/Year
    assert _norm_ddmmyy("12/2025") == "01/12/25"
    
    # Test "tháng" string
    assert _norm_ddmmyy("Tháng 6") == "01/06/26"
    assert _norm_ddmmyy("thang 11") == "01/11/26"
    
    # Test invalid date cases
    assert _norm_ddmmyy("không có ngày") is None
    assert _norm_ddmmyy(None) is None
    assert _norm_ddmmyy("") is None

def test_clean_value():
    assert _clean("  Hello World  ") == "Hello World"
    assert _clean("nan") == ""
    assert _clean("None") == ""
    assert _clean("null") == ""
    assert _clean(None) == ""

def test_has_price():
    assert _has_price("Giá là 10 triệu đồng") is True
    assert _has_price("Báo giá 5 tỷ vnđ") is True
    assert _has_price("Chỉ hỏi thăm sức khỏe") is False
