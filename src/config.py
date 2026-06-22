import os
from pathlib import Path
from dotenv import load_dotenv
import pandas as pd

# --- Container Paths ---
PROJECT_ROOT = Path(__file__).resolve().parent.parent  # /app

# Load environment variables from .env if present
load_dotenv(PROJECT_ROOT / ".env")

# Folders mapped via Docker volumes
PATH_DATA = PROJECT_ROOT / "data"      # /app/data
PATH_OUTPUT = PROJECT_ROOT / "output"  # /app/output
PATH_CONFIG = PROJECT_ROOT / "config"  # /app/config

# Core Files
PATH_INPUT = PATH_DATA / "CRM_merge.xlsx"
PATH_KW_JSON = PATH_CONFIG / "keywords_fixed.json"
PATH_PROMPT = PATH_CONFIG / "prompt_CRM_v5.txt"

# Output files
PATH_BACKUP_DIR = PATH_DATA / "backups"
DB_JSON_PATH = PATH_OUTPUT / "classified_history_db.json"
CKPT_JSON = PATH_OUTPUT / "llm_fills_checkpoint.json"

# API keys & model configuration
API_KEY = os.getenv("GEMINI_API_KEY") or ""
MODEL_NAME = os.getenv("GEMINI_MODEL") or "gemini-2.5-flash"
MIN_INTERVAL_S = float(os.getenv("GEMINI_MIN_INTERVAL_S") or "3.5")
JITTER_S = float(os.getenv("GEMINI_JITTER_S") or "0.5")
BATCH_SIZE = min(20, int(os.getenv("GEMINI_BATCH_SIZE") or "20"))

# Azure AD Configuration
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID") or ""
AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID") or ""
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET") or ""

# SharePoint Source Configuration (Input)
SHAREPOINT_SOURCE_DRIVE_ID = os.getenv("SHAREPOINT_SOURCE_DRIVE_ID") or "b!K2h3C6y-m0-nfhBTKpRGUHVm2ajJ6j5Liiy5S410oj50hc-yscOEQrZMgUgL5YIr"
SHAREPOINT_SOURCE_FILE_PATH = os.getenv("SHAREPOINT_SOURCE_FILE_PATH") or "CRM_merge/CRM_merge.xlsx"

# SharePoint Target Configuration (Output)
SHAREPOINT_TARGET_DRIVE_ID = os.getenv("SHAREPOINT_TARGET_DRIVE_ID") or "b!NtlOjnoIAUGK2TSeHPN7wre_0vBaUZhHrpr9lFi7hWvJ-n1mRHXzSr4E9Qjzat5Q"
SHAREPOINT_TARGET_FILE_PATH = os.getenv("SHAREPOINT_TARGET_FILE_PATH") or "Data Lake to SharePoint/Phan_Tich_CRM/Phân loại dữ liệu công trình dự án CRM V4.xlsx"

# Email Notification Configuration
NOTIFICATION_SENDER_EMAIL = os.getenv("NOTIFICATION_SENDER_EMAIL") or ""
_recipients_str = os.getenv("NOTIFICATION_RECIPIENTS") or ""
NOTIFICATION_RECIPIENTS = [email.strip() for email in _recipients_str.split(",") if email.strip()]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# --- Column naming helper ---
def col_name(major: str, minor: str) -> str:
    return f"[{major}] {minor}"

# 15 Output columns
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

LLM_TARGET_COLS = [c for c in OUTPUT_COLUMNS if c != COL_CURRENT_STATUS]

def normalize_id(val) -> str:
    if pd.isna(val):
        return ""
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    if val_str.lower() in ("nan", "none", ""):
        return ""
    return val_str
