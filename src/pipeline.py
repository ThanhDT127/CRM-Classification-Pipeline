import os
import sys
import json
import time
import shutil
import datetime
import traceback
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import config
from classifier import (
    classify_row_regex,
    _compile_patterns,
    _clean
)
from llm import init_llm_client, call_llm_batch
from sharepoint import AuthProvider, SharePointClient
from notification import NotificationService

# --- SETUP LOGGING ---
config.PATH_OUTPUT.mkdir(parents=True, exist_ok=True)
log_dir = config.PATH_OUTPUT / "logs"
log_dir.mkdir(parents=True, exist_ok=True)
log_file = log_dir / "automation.log"

logger = logging.getLogger("crm-automation")
logger.setLevel(logging.INFO)

# Create formatter
formatter = logging.Formatter("[%(asctime)s] %(levelname)s [%(name)s] %(message)s")

# Stdout handler
stdout_handler = logging.StreamHandler(sys.stdout)
stdout_handler.setFormatter(formatter)
logger.addHandler(stdout_handler)

# Rotating file handler (max 10MB per file, keep 5 files)
file_handler = RotatingFileHandler(log_file, maxBytes=10 * 1024 * 1024, backupCount=5, encoding="utf-8")
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

def load_and_index_keywords():
    with open(config.PATH_KW_JSON, "r", encoding="utf-8") as f:
        kw = json.load(f)
        
    kw_index = {}
    competitor_brands = []
    
    for major, minors in (kw or {}).items():
        if not isinstance(minors, dict):
            continue
        for minor, subs in minors.items():
            if major == "Đối Thủ Cạnh Tranh" and minor == "Cạnh Tranh" and isinstance(subs, dict):
                brand_list = subs.get("Các Hãng cụ thể") or subs.get("Các Hãng đối thủ cạnh tranh")
                if isinstance(brand_list, list):
                    for b in brand_list:
                        for term in str(b).split(";"):
                            if term.strip():
                                competitor_brands.append(term.strip())
                continue
            if isinstance(subs, dict):
                col = config.col_name(major, minor)
                col_labels = {}
                for sub, kw_list in subs.items():
                    if not isinstance(sub, str) or not sub.strip():
                        continue
                    terms = []
                    if isinstance(kw_list, list):
                        for k in kw_list:
                            for t in str(k).split(";"):
                                if t.strip():
                                    terms.append(t.strip())
                    elif isinstance(kw_list, str):
                        terms = [t.strip() for t in kw_list.split(";") if t.strip()]
                    if terms:
                        col_labels[sub] = terms
                if col_labels:
                    kw_index[col] = col_labels
                    
    return kw_index, competitor_brands

def apply_excel_styling(file_path: Path):
    logger.info("Applying premium double-header styling to Excel...")
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    group_colors = {
        "Hoạt Động CRM": "FFF2CC",
        "AETT": "E2EFDA",
        "Khách Hàng": "E8D5F5",
        "Kế Hoạch": "FCE4CC",
        "Đối Thủ Cạnh Tranh": "D6EAF8"
    }

    group_data_colors = {
        "Hoạt Động CRM": "FFFBEA",
        "AETT": "F2F9EE",
        "Khách Hàng": "F5EEFA",
        "Kế Hoạch": "FEF3EA",
        "Đối Thủ Cạnh Tranh": "EBF5FB"
    }

    font_major = Font(name="Segoe UI", size=11, bold=True)
    font_minor = Font(name="Segoe UI", size=10, bold=True)
    font_data = Font(name="Segoe UI", size=10)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    thin_border = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )

    ws.insert_rows(1, 1)
    
    col_info = {}
    for col_idx in range(1, ws.max_column + 1):
        cell_val = str(ws.cell(row=2, column=col_idx).value or "")
        if cell_val.startswith("[") and "]" in cell_val:
            parts = cell_val.split("] ", 1)
            major = parts[0][1:]
            minor = parts[1]
            col_info[col_idx] = (major, minor)
            ws.cell(row=1, column=col_idx, value=major)
            ws.cell(row=2, column=col_idx, value=minor)
        else:
            col_info[col_idx] = (None, None)

    for col_idx in range(1, ws.max_column + 1):
        major, minor = col_info.get(col_idx, (None, None))
        
        if major:
            fill_color = group_colors.get(major, "FFFFFF")
            fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
            
            c1 = ws.cell(row=1, column=col_idx)
            c1.fill = fill
            c1.font = font_major
            c1.alignment = align_center
            c1.border = thin_border
            
            c2 = ws.cell(row=2, column=col_idx)
            c2.fill = fill
            c2.font = font_minor
            c2.alignment = align_center
            c2.border = thin_border
        else:
            val = ws.cell(row=2, column=col_idx).value
            ws.cell(row=1, column=col_idx, value=val)
            ws.cell(row=2, column=col_idx, value="")
            
            fill_neutral = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            for r in (1, 2):
                c = ws.cell(row=r, column=col_idx)
                c.fill = fill_neutral
                c.font = font_minor
                c.alignment = align_center
                c.border = thin_border
                
            ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)

        if major:
            data_color = group_data_colors.get(major, "FFFFFF")
            data_fill = PatternFill(start_color=data_color, end_color=data_color, fill_type="solid")
        else:
            data_fill = None

        for row_idx in range(3, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = font_data
            cell.border = thin_border
            if data_fill and major:
                cell.fill = data_fill

    start_col = None
    current_major = None
    for col_idx in range(1, ws.max_column + 2):
        major = col_info.get(col_idx, (None, None))[0]
        if major != current_major:
            if current_major is not None and start_col is not None:
                end_col = col_idx - 1
                if end_col > start_col:
                    ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
            if major is not None:
                start_col = col_idx
                current_major = major
            else:
                start_col = None
                current_major = None

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24

    for col_idx in range(1, ws.max_column + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row_idx in range(1, min(ws.max_row + 1, 100)):
            val_str = str(ws.cell(row=row_idx, column=col_idx).value or "")
            lines = val_str.split("\n")
            for line in lines:
                if len(line) > max_len:
                    max_len = len(line)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(file_path)
    logger.info("[OK] Excel styling successfully applied.")

def run_automation_pipeline() -> bool:
    logger.info("============================================================")
    logger.info("CRM Automated Pipeline Run Started")
    logger.info("============================================================")
    
    t0 = time.time()
    auth_provider = AuthProvider()
    sp_client = SharePointClient(auth_provider)
    notifier = NotificationService(auth_provider)
    
    local_excel_path = config.PATH_INPUT
    
    # Clean up local input directory if leftover
    if local_excel_path.exists():
        try:
            local_excel_path.unlink()
        except Exception:
            pass

    try:
        # 1. Download file from SharePoint Source site
        sp_client.download_file(
            config.SHAREPOINT_SOURCE_FILE_PATH, 
            local_excel_path, 
            drive_id=config.SHAREPOINT_SOURCE_DRIVE_ID
        )
        
        # 2. Schema Validation
        df = pd.read_excel(local_excel_path)
        required_cols = ["ActivityId", "Tình trạng hiện tại", "Tình hình tiến độ công trình", 
                         "Nội dung làm việc, yêu cầu KH & đánh giá", "Kế hoạch lần tới", "Đề xuất"]
        for col in required_cols:
            if col not in df.columns:
                raise ValueError(f"Schema mismatch: missing column '{col}'")

        # 3. Create Backup of the input file
        config.PATH_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = config.PATH_BACKUP_DIR / f"CRM_merge_backup_{timestamp}.xlsx"
        shutil.copy(local_excel_path, backup_path)
        logger.info("[OK] Backup created successfully: %s", backup_path.name)

        # 4. Load History DB
        history_db = {}
        if config.DB_JSON_PATH.exists():
            try:
                with open(config.DB_JSON_PATH, "r", encoding="utf-8") as f:
                    history_db = json.load(f)
                logger.info("Loaded history DB: %d items", len(history_db))
            except Exception as e:
                logger.warning("Failed to load history DB JSON: %s. Starting fresh.", e)

        # 5. Extract Keywords Patterns & Filter Delta Rows
        kw_index, brand_list = load_and_index_keywords()
        col_pats = {col: {lbl: _compile_patterns(kws) for lbl, kws in labels_dict.items()}
                    for col, labels_dict in kw_index.items()}
        brand_pats = [(b, _compile_patterns([b])) for b in brand_list]

        pending_rows = []
        for idx, row in df.iterrows():
            act_id = str(row["ActivityId"]).strip()
            if not act_id or act_id.lower() in ("nan", "none", ""):
                continue
            if act_id not in history_db:
                pending_rows.append({
                    "ActivityId": act_id,
                    "row_idx": idx,
                    "Tình trạng hiện tại": row.get("Tình trạng hiện tại"),
                    "Tình hình tiến độ công trình": row.get("Tình hình tiến độ công trình"),
                    "Nội dung làm việc, yêu cầu KH & đánh giá": row.get("Nội dung làm việc, yêu cầu KH & đánh giá"),
                    "Kế hoạch lần tới": row.get("Kế hoạch lần tới"),
                    "Đề xuất": row.get("Đề xuất")
                })

        logger.info("Total rows in source file: %d | New delta rows to classify: %d", len(df), len(pending_rows))

        cells_filled_count = 0
        
        if pending_rows:
            # 5.1 Run Regex Classifier
            new_fills = {}
            llm_input_payload = []
            
            for item in pending_rows:
                act_id = item["ActivityId"]
                r_val = item.get("Tình trạng hiện tại")
                r_cleaned = _clean(r_val)
                row_fills = {config.COL_CURRENT_STATUS: r_cleaned if r_cleaned else None}
                
                # Regex step
                regex_fills = classify_row_regex(item, col_pats, brand_pats)
                row_fills.update(regex_fills)
                
                # Check for LLM backup columns
                missing_cols = []
                for col in config.LLM_TARGET_COLS:
                    val = row_fills.get(col)
                    if val is None or val == "mơ hồ":
                        missing_cols.append(col)
                        
                if missing_cols:
                    llm_input_payload.append({
                        "row_idx": act_id, # Use ActivityId
                        "Tình trạng hiện tại": item.get("Tình trạng hiện tại") or "",
                        "Tình hình tiến độ công trình": item.get("Tình hình tiến độ công trình") or "",
                        "Nội dung làm việc, yêu cầu KH & đánh giá": item.get("Nội dung làm việc, yêu cầu KH & đánh giá") or "",
                        "Kế hoạch lần tới": item.get("Kế hoạch lần tới") or "",
                        "Đề xuất": item.get("Đề xuất") or "",
                        "missing_cols": missing_cols
                    })
                new_fills[act_id] = row_fills

            # 5.2 Call Gemini Client for empty/ambiguous columns
            if llm_input_payload:
                logger.info("Calling Gemini API for %d items...", len(llm_input_payload))
                client, model_name = init_llm_client()
                system_prompt = Path(config.PATH_PROMPT).read_text(encoding="utf-8")
                
                batch_size = config.BATCH_SIZE
                llm_results = {}
                for i in range(0, len(llm_input_payload), batch_size):
                    batch = llm_input_payload[i:i + batch_size]
                    res = call_llm_batch(client, model_name, system_prompt, batch)
                    for item in res:
                        rid = str(item.get("row_idx"))
                        llm_results[rid] = item.get("fills") or {}
                
                # Merge LLM fills back into regex records
                for act_id, fills in new_fills.items():
                    if act_id in llm_results:
                        llm_fills = llm_results[act_id]
                        for col, val in llm_fills.items():
                            if col in config.LLM_TARGET_COLS:
                                if val is None or str(val).strip() == "":
                                    fills[col] = None
                                else:
                                    fills[col] = str(val).strip()
            
            # Save new delta results to local JSON DB
            history_db.update(new_fills)
            with open(config.DB_JSON_PATH, "w", encoding="utf-8") as f:
                json.dump(history_db, f, ensure_ascii=False, indent=2)
            logger.info("JSON History database successfully updated.")

        # 6. Tải file đích từ SharePoint Target site (hoặc khởi tạo từ df nếu chưa tồn tại)
        target_excel_path = config.PATH_OUTPUT / "CRM_classified.xlsx"
        if target_excel_path.exists():
            try:
                target_excel_path.unlink()
            except Exception:
                pass

        target_exists = sp_client.check_file_exists(
            config.SHAREPOINT_TARGET_FILE_PATH, 
            drive_id=config.SHAREPOINT_TARGET_DRIVE_ID
        )
        
        if target_exists:
            logger.info("Downloading existing target file from SharePoint Target site...")
            sp_client.download_file(
                config.SHAREPOINT_TARGET_FILE_PATH, 
                target_excel_path, 
                drive_id=config.SHAREPOINT_TARGET_DRIVE_ID
            )
            target_df = pd.read_excel(target_excel_path)
        else:
            logger.info("Target file does not exist on SharePoint. Initializing from source schema...")
            target_df = df.copy()

        # 6.1 Gộp kết quả an toàn bằng Index (Merge safely by ActivityId index)
        for col in config.OUTPUT_COLUMNS:
            if col not in target_df.columns:
                target_df[col] = None
            target_df[col] = target_df[col].astype(object)

        target_df.set_index("ActivityId", inplace=True)
        
        for act_id, fills in history_db.items():
            if act_id in target_df.index:
                for col in config.OUTPUT_COLUMNS:
                    val = fills.get(col)
                    current_val = target_df.at[act_id, col]
                    is_empty_or_mo_ho = pd.isna(current_val) or str(current_val).strip() == '' or str(current_val).strip().lower() == 'mơ hồ'
                    
                    if is_empty_or_mo_ho and val is not None:
                        target_df.at[act_id, col] = val
                        cells_filled_count += 1
                    elif is_empty_or_mo_ho and val is None:
                        target_df.at[act_id, col] = None

        target_df.reset_index(inplace=True)

        # 7. Ghi file đích cục bộ, áp dụng styles
        target_df.to_excel(target_excel_path, index=False)
        apply_excel_styling(target_excel_path)
        
        # 8. Upload file đích ngược lại SharePoint Target site
        sp_client.upload_file(
            target_excel_path, 
            config.SHAREPOINT_TARGET_FILE_PATH, 
            drive_id=config.SHAREPOINT_TARGET_DRIVE_ID
        )
        
        elapsed = time.time() - t0
        logger.info("[SUCCESS] Pipeline execution finished in %.1fs!", elapsed)
        
        # 9. Gửi email báo cáo thành công
        notifier.send_success(elapsed, len(pending_rows), cells_filled_count)
        
        # 10. Dọn dẹp các file Excel tạm cục bộ
        for p in (local_excel_path, target_excel_path):
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
            
        return True

    except Exception as e:
        err_msg = traceback.format_exc()
        logger.error("Pipeline crashed! Exception traceback:\n%s", err_msg)
        
        # Gửi email báo cáo lỗi
        try:
            notifier.send_error(err_msg)
        except Exception as mail_err:
            logger.error("Failed to send error notification email: %s", mail_err)
            
        # Dọn dẹp tệp tạm
        for p in (local_excel_path, config.PATH_OUTPUT / "CRM_classified.xlsx"):
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
                
        return False

def get_seconds_until_next_run(target_hour=3, target_minute=30) -> float:
    now = datetime.datetime.now()
    target_today = now.replace(hour=target_hour, minute=target_minute, second=0, microsecond=0)
    
    if now < target_today:
        diff = target_today - now
    else:
        target_tomorrow = target_today + datetime.timedelta(days=1)
        diff = target_tomorrow - now
        
    return diff.total_seconds()

if __name__ == "__main__":
    run_as_daemon = os.getenv("RUN_AS_DAEMON", "false").lower() == "true"
    
    if not run_as_daemon:
        success = run_automation_pipeline()
        sys.exit(0 if success else 1)
    else:
        logger.info("Starting pipeline in Daemon Mode (RUN_AS_DAEMON=True)...")
        # Run immediately on startup
        run_automation_pipeline()
        
        while True:
            sec = get_seconds_until_next_run(3, 30)
            hours = sec / 3600.0
            logger.info("Daemon Mode: Sleeping for %.2f hours until next run at 03:30 AM...", hours)
            time.sleep(sec)
            
            logger.info("Daemon Mode: Waking up to run scheduled pipeline...")
            run_automation_pipeline()
