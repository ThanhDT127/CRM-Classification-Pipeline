import sys
import os
import json
import tempfile
from pathlib import Path
from unittest.mock import MagicMock, patch, mock_open
import pytest
import pandas as pd
import openpyxl

# Add src to path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import config
from sharepoint import AuthProvider, SharePointClient
from notification import NotificationService
import pipeline as main

@pytest.fixture
def mock_auth():
    auth = MagicMock(spec=AuthProvider)
    auth.get_access_token.return_value = "mock_token"
    auth.get_headers.return_value = {"Authorization": "Bearer mock_token", "Content-Type": "application/json"}
    return auth

def test_sharepoint_client_download(mock_auth):
    client = SharePointClient(mock_auth)
    
    # Mock requests.Session
    client.session = MagicMock()
    
    # Mock metadata request
    mock_meta_resp = MagicMock()
    mock_meta_resp.status_code = 200
    mock_meta_resp.json.return_value = {
        "@microsoft.graph.downloadUrl": "https://mock.download.url/file.xlsx",
        "name": "CRM_merge.xlsx",
        "size": 1000
    }
    
    # Mock download request
    mock_dl_resp = MagicMock()
    mock_dl_resp.status_code = 200
    mock_dl_resp.iter_content.return_value = [b"mock excel content"]
    
    # Configure session.get side effects
    client.session.get.side_effect = [mock_meta_resp, mock_dl_resp]
    
    with tempfile.TemporaryDirectory() as tmpdir:
        local_path = Path(tmpdir) / "CRM_merge.xlsx"
        result_path = client.download_file("CRM_merge/CRM_merge.xlsx", local_path)
        
        assert result_path == local_path
        assert local_path.exists()
        assert local_path.read_bytes() == b"mock excel content"

def test_sharepoint_client_upload(mock_auth):
    client = SharePointClient(mock_auth)
    client.session = MagicMock()
    
    mock_resp = MagicMock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"id": "mock_file_id", "name": "CRM_merge.xlsx"}
    client.session.put.return_value = mock_resp
    
    with tempfile.TemporaryDirectory() as tmpdir:
        local_path = Path(tmpdir) / "CRM_merge.xlsx"
        local_path.write_bytes(b"mock excel data to upload")
        
        res = client.upload_file(local_path, "CRM_merge/CRM_merge.xlsx")
        assert res["id"] == "mock_file_id"
        client.session.put.assert_called_once()

def test_notification_service_send_email(mock_auth):
    # Temporarily set config values
    config.NOTIFICATION_SENDER_EMAIL = "sender@domain.com"
    config.NOTIFICATION_RECIPIENTS = ["rec@domain.com"]
    
    notifier = NotificationService(mock_auth)
    notifier.session = MagicMock()
    
    mock_resp = MagicMock()
    mock_resp.status_code = 202
    notifier.session.post.return_value = mock_resp
    
    success = notifier._send_email("Test Subject", "<h1>Test Body</h1>")
    assert success is True
    notifier.session.post.assert_called_once()

@patch("pipeline.AuthProvider")
@patch("pipeline.SharePointClient")
@patch("pipeline.NotificationService")
@patch("pipeline.init_llm_client")
@patch("pipeline.call_llm_batch")
def test_full_pipeline_run(mock_call_llm, mock_init_llm, mock_notifier_cls, mock_sp_cls, mock_auth_cls):
    print(">>> main module name:", main.__name__)
    print(">>> sys.modules pipeline keys:", [k for k in sys.modules if "pipeline" in k or "main" in k])
    mock_auth_inst = MagicMock()
    mock_auth_cls.return_value = mock_auth_inst
    
    mock_sp_inst = MagicMock()
    mock_sp_cls.return_value = mock_sp_inst
    
    mock_notifier_inst = MagicMock()
    mock_notifier_cls.return_value = mock_notifier_inst
    
    mock_llm_client = MagicMock()
    mock_init_llm.return_value = (mock_llm_client, "gemini-2.5-flash")
    
    # Create temp files for config mapping
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        
        # Override config paths for testing
        config.PATH_OUTPUT = tmp_path / "output"
        config.PATH_OUTPUT.mkdir(parents=True, exist_ok=True)
        config.PATH_INPUT = tmp_path / "CRM_merge.xlsx"
        config.PATH_BACKUP_DIR = tmp_path / "backups"
        config.DB_JSON_PATH = config.PATH_OUTPUT / "classified_history_db.json"
        config.CKPT_JSON = config.PATH_OUTPUT / "llm_fills_checkpoint.json"
        
        # Create a mock Excel sheet that needs classification
        df = pd.DataFrame([
            {
                "ActivityId": "ACT_001",
                "Tình trạng hiện tại": "Khách hàng muốn mua bóng đèn LED",
                "Tình hình tiến độ công trình": "Đang hoàn thiện phần thô",
                "Nội dung làm việc, yêu cầu KH & đánh giá": "Tư vấn giá và mẫu mã cho anh Thanh",
                "Kế hoạch lần tới": "Gửi báo giá bóng đèn",
                "Đề xuất": "Giảm giá 5%"
            }
        ])
        df.to_excel(config.PATH_INPUT, index=False)
        
        # Mock download_file to write the Excel file when called (since the pipeline unlinks it initially)
        def side_effect_download(remote_path, local_path, *args, **kwargs):
            df.to_excel(local_path, index=False)
            return local_path
        mock_sp_inst.download_file.side_effect = side_effect_download
        mock_sp_inst.check_file_exists.return_value = False
        
        # Mock LLM API response
        mock_call_llm.return_value = [
            {
                "row_idx": "ACT_001",
                "fills": {
                    config.COL_WORK_CRM: "Bóng LED",
                    config.COL_OPINION: "Muốn mua bóng đèn LED",
                    config.COL_PLAN_NEXT: "Gửi báo giá bóng",
                    config.COL_PROPOSAL: "Giảm giá 5%"
                }
            }
        ]
        
        # Run pipeline
        success = main.run_automation_pipeline()
        
        # Assertions
        assert success is True
        assert mock_sp_inst.download_file.call_count == 2
        assert mock_sp_inst.upload_file.call_count == 2
        mock_notifier_inst.send_success.assert_called_once()
        
        # Check that backup was created
        backups = list(config.PATH_BACKUP_DIR.glob("*.xlsx"))
        assert len(backups) == 1
        
        # Check history DB was updated
        assert config.DB_JSON_PATH.exists()
        with open(config.DB_JSON_PATH, "r", encoding="utf-8") as f:
            history = json.load(f)
            assert "ACT_001" in history
            assert history["ACT_001"][config.COL_CURRENT_STATUS] == "Khách hàng muốn mua bóng đèn LED"
