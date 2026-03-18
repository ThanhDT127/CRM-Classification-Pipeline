# 📋 TÀI LIỆU BÀN GIAO — CRM Classification Pipeline

> **Phiên bản:** 1.0  
> **Ngày tạo:** 18/03/2026  
> **Tác giả:** AI-assisted  
> **Dự án:** Phân loại dữ liệu CRM tự động (Regex + Gemini LLM)

---

## Mục Lục

1. [Tổng Quan Dự Án](#1-tổng-quan-dự-án)
2. [Kiến Trúc Hệ Thống](#2-kiến-trúc-hệ-thống)
3. [Cấu Trúc Thư Mục](#3-cấu-trúc-thư-mục)
4. [Luồng Xử Lý (Pipeline Flow)](#4-luồng-xử-lý-pipeline-flow)
5. [Chi Tiết Từng Bước](#5-chi-tiết-từng-bước)
6. [Hướng Dẫn Cài Đặt & Thiết Lập](#6-hướng-dẫn-cài-đặt--thiết-lập)
7. [Hướng Dẫn Vận Hành](#7-hướng-dẫn-vận-hành)
8. [Cấu Hình Hệ Thống](#8-cấu-hình-hệ-thống)
9. [Taxonom & Keyword Mapping](#9-taxonomy--keyword-mapping)
10. [Prompt LLM (Gemini)](#10-prompt-llm-gemini)
11. [Xử Lý Lỗi & Troubleshooting](#11-xử-lý-lỗi--troubleshooting)
12. [Bảo Trì & Mở Rộng](#12-bảo-trì--mở-rộng)
13. [Câu Hỏi Thường Gặp (FAQ)](#13-câu-hỏi-thường-gặp-faq)
14. [Liên Hệ & Tham Khảo](#14-liên-hệ--tham-khảo)

---

## 1. Tổng Quan Dự Án

### 1.1 Mục Đích

Hệ thống tự động phân loại dữ liệu CRM từ file Excel (`CRM_TDCTDA.xlsx`) cho công ty sản xuất **thiết bị điện / chiếu sáng**. Pipeline thực hiện:

- **Đọc** dữ liệu CRM thô từ Excel
- **Phân loại** tự động bằng **regex keyword matching** (nhanh, chính xác cao)
- **Bổ sung** các trường còn thiếu bằng **Gemini LLM** (AI suy luận ngữ cảnh)
- **Xuất** kết quả phân loại hoàn chỉnh ra file Excel

### 1.2 Bài Toán Kinh Doanh

| Hạng mục | Mô tả |
|----------|-------|
| **Input** | Báo cáo CRM hàng tuần/tháng từ sales team (Excel) |
| **Output** | File Excel đã phân loại **14 cột** theo taxonomy chuẩn |
| **Đối tượng sử dụng** | Bộ phận Marketing, Quản lý bán hàng, Ban lãnh đạo |
| **Tần suất chạy** | Khi có batch CRM mới (tuần/tháng) |

### 1.3 Công Nghệ Sử Dụng

| Công nghệ | Phiên bản | Mục đích |
|-----------|-----------|----------|
| Python | ≥ 3.10 | Ngôn ngữ chính |
| pandas | ≥ 2.0 | Xử lý dữ liệu Excel |
| openpyxl | ≥ 3.1 | Đọc/ghi file .xlsx |
| requests | ≥ 2.31 | HTTP client |
| python-dotenv | ≥ 1.0 | Quản lý biến môi trường |
| google-generativeai | — | Gọi Gemini API |
| Google Gemini | gemini-2.0-flash | LLM phân loại |

---

## 2. Kiến Trúc Hệ Thống

### 2.1 Sơ Đồ Tổng Quan

```
┌──────────────┐
│  CRM_TDCTDA  │  ← File Excel đầu vào (sales reports)
│    .xlsx     │
└──────┬───────┘
       │
       ▼
┌──────────────────────────────────────────────────────────┐
│                   PIPELINE ENGINE                        │
│                                                          │
│  ┌──────────┐   ┌──────────┐   ┌──────────┐   ┌──────┐ │
│  │  Step 1  │──▶│  Step 2  │──▶│  Step 3  │──▶│Step 4│ │
│  │  Regex   │   │ Prepare  │   │ Call LLM │   │Merge │ │
│  │Classify  │   │LLM Input │   │ (Gemini) │   │      │ │
│  └──────────┘   └──────────┘   └──────────┘   └──────┘ │
│       │              │              │              │     │
│       ▼              ▼              ▼              ▼     │
│  CRM_classi-    llm_input      llm_fills       CRM_     │
│  fied.xlsx       .json          .json        classified  │
│                                             _with_LLM    │
│                                               .xlsx      │
└──────────────────────────────────────────────────────────┘
       │                              │
       ▼                              ▼
┌──────────────┐             ┌────────────────┐
│  keywords    │             │   Gemini API   │
│ _fixed.json  │             │  (Google AI)   │
└──────────────┘             └────────────────┘
```

### 2.2 Luồng Dữ Liệu Chi Tiết

```
Input Excel                    Keyword JSON             Prompt Template
(CRM_TDCTDA.xlsx)        (keywords_fixed.json)       (prompt_CRM_v5.txt)
       │                         │                          │
       ▼                         ▼                          │
  ┌─────────────────────────────────────┐                   │
  │     STEP 1: Regex Classification    │                   │
  │  • Đọc Excel → DataFrame           │                   │
  │  • Merge 5 cột text → merged_text   │                   │
  │  • Match regex từ keyword JSON      │                   │
  │  • Gán label cho 14 cột output      │                   │
  │  • Copy "Tình trạng hiện tại"       │                   │
  └──────────────┬──────────────────────┘                   │
                 │                                          │
                 ▼                                          │
        CRM_classified.xlsx                                 │
        (kết quả regex, có cells trống)                     │
                 │                                          │
                 ▼                                          │
  ┌─────────────────────────────────────┐                   │
  │  STEP 2: Prepare LLM Input          │                   │
  │  • Đọc CRM_classified.xlsx          │                   │
  │  • Tìm cells trống/mơ hồ            │                   │
  │  • Tạo JSON items cho LLM           │                   │
  │  • Gắn allowed values per column    │                   │
  └──────────────┬──────────────────────┘                   │
                 │                                          │
                 ▼                                          │
          llm_input.json                                    │
          (items cần LLM fill)                              │
                 │                                          │
                 ▼                                          ▼
  ┌─────────────────────────────────────────────────────────────┐
  │          STEP 3: Call Gemini LLM                             │
  │  • Chia items → batches (mặc định 20 items/batch)          │
  │  • Gọi Gemini API + system prompt                          │
  │  • Parse JSON response → fills                             │
  │  • Canonicalize labels (đối chiếu allowed)                 │
  │  • Normalize dates (dd/mm/yy)                              │
  │  • Retry on 429/rate-limit (15s → 120s backoff)            │
  │  • Checkpoint mỗi batch (resume nếu bị crash)             │
  └──────────────┬──────────────────────────────────────────────┘
                 │
                 ▼
          llm_fills.json
          llm_fills_checkpoint.json
                 │
                 ▼
  ┌─────────────────────────────────────┐
  │     STEP 4: Merge Results           │
  │  • Đọc CRM_classified.xlsx         │
  │  • Đọc llm_fills.json              │
  │  • Fill CHỈ cells trống/mơ hồ       │
  │  • Giữ nguyên cells đã có regex     │
  │  • Thống kê fill rate per column    │
  └──────────────┬──────────────────────┘
                 │
                 ▼
        CRM_classified_with_LLM.xlsx
        (KẾT QUẢ CUỐI CÙNG ✓)
```

---

## 3. Cấu Trúc Thư Mục

```
D:\Works\CRM\
│
├── .env                          # API keys (KHÔNG commit lên git)
├── .gitignore                    # Ignore data/, output/, docs/, .env
├── README.md                     # Tài liệu tổng quan ngắn
├── requirements.txt              # Dependencies Python
│
├── src/                          # ★ SOURCE CODE CHÍNH
│   ├── config.py                 # Cấu hình: paths, columns, keywords, API
│   ├── run_pipeline.py           # Entry point - chạy pipeline
│   ├── step1_classify.py         # Bước 1: Regex classification
│   ├── step2_prepare_llm.py      # Bước 2: Chuẩn bị LLM input
│   ├── step3_call_llm.py         # Bước 3: Gọi Gemini API
│   ├── step4_merge.py            # Bước 4: Gộp kết quả
│   └── extract_keywords_from_spec_NEW.py  # Utility: trích keywords từ Excel spec
│
├── data/                         # ★ DỮ LIỆU ĐẦU VÀO (gitignored)
│   ├── CRM_TDCTDA.xlsx           # File CRM chính (input)
│   ├── CRM raw.xlsx              # File CRM gốc (backup)
│   └── Đánh tay gốc.xlsx        # File đánh tay tham khảo
│
├── output/                       # ★ KẾT QUẢ (gitignored)
│   ├── CRM_classified.xlsx       # Output Step 1 (regex only)
│   ├── CRM_classified_with_LLM.xlsx  # Output Step 4 (final)
│   ├── llm_input.json            # Output Step 2 (LLM input)
│   ├── llm_fills.json            # Output Step 3 (LLM results)
│   ├── llm_fills_checkpoint.json # Checkpoint Step 3 (resume)
│   └── archive/                  # Kết quả cũ
│
├── prompts/                      # ★ PROMPT TEMPLATES
│   └── prompt_CRM_v5.txt         # System prompt cho Gemini (v5)
│
├── keywords/                     # ★ KEYWORD MAPPING
│   └── keywords_fixed.json       # JSON keyword → label mapping
│
├── notebooks/                    # Jupyter notebooks tham khảo
│   ├── CRM_Classification.ipynb
│   └── Phan_Loai_CRM_IMPROVED.ipynb
│
└── docs/                         # Tài liệu tham khảo (gitignored)
    ├── Check list tính năng BOT Khai thác CRM.pdf
    ├── tu_khoa_cho_cac_cot_cong_trinh.xlsx
    └── Định nghĩa tiến độ công trình_12.7.25 (1).xlsx
```

### 3.1 File Quan Trọng Nhất

| File | Tầm quan trọng | Mô tả |
|------|:---:|-------|
| `src/config.py` | ⭐⭐⭐ | Trung tâm cấu hình - thay đổi paths, columns tại đây |
| `keywords/keywords_fixed.json` | ⭐⭐⭐ | Bộ keyword quyết định chất lượng regex |
| `prompts/prompt_CRM_v5.txt` | ⭐⭐⭐ | Prompt quyết định chất lượng LLM |
| `src/step3_call_llm.py` | ⭐⭐ | Logic gọi API phức tạp nhất |
| `.env` | ⭐⭐ | API key — BẮT BUỘC có để chạy step 3 |

---

## 4. Luồng Xử Lý (Pipeline Flow)

### 4.1 Flowchart Tổng Quan

```
                        ┌─────────────┐
                        │   BẮT ĐẦU   │
                        └──────┬──────┘
                               │
                    ┌──────────▼──────────┐
                    │ Có file CRM_TDCTDA  │──No──▶ ⛔ LỖI: Thiếu input
                    │     .xlsx ?         │
                    └──────────┬──────────┘
                          Yes  │
                               ▼
                    ┌─────────────────────┐
                    │  STEP 1: Regex      │
                    │  Phân loại nhanh    │
                    │  bằng keyword       │
                    └──────────┬──────────┘
                               │
                          Xuất CRM_classified.xlsx
                               │
                    ┌──────────▼──────────┐
                    │  STEP 2: Prepare    │
                    │  Tìm cells trống    │
                    │  Tạo LLM input      │
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │ Có cells trống?     │──No──▶ ✅ XONG (chỉ regex đủ)
                    └──────────┬──────────┘
                          Yes  │
                               ▼
                    ┌──────────▼──────────┐
                    │ Có GEMINI_API_KEY?  │──No──▶ ⛔ LỖI: Thiếu API key
                    └──────────┬──────────┘
                          Yes  │
                               ▼
                    ┌─────────────────────┐
                    │  STEP 3: Call LLM   │
                    │  Gọi Gemini batch   │◀────── Retry/Resume
                    │  + checkpoint       │         (nếu lỗi)
                    └──────────┬──────────┘
                               │
                    ┌──────────▼──────────┐
                    │  STEP 4: Merge      │
                    │  Gộp LLM fills vào  │
                    │  Excel cuối cùng    │
                    └──────────┬──────────┘
                               │
                        ┌──────▼──────┐
                        │   KẾT THÚC   │
                        │     ✅       │
                        └─────────────┘
```

### 4.2 Thời Gian Ước Tính

| Bước | Thời gian | Phụ thuộc |
|------|-----------|-----------|
| Step 1 (Regex) | 5–30 giây | Kích thước file input |
| Step 2 (Prepare) | 2–10 giây | Số rows |
| Step 3 (LLM) | **5–60 phút** | Số cells trống × batch size × rate limit |
| Step 4 (Merge) | 2–10 giây | Kích thước file |
| **Tổng** | **~10–65 phút** | Chủ yếu do Step 3 |

---

## 5. Chi Tiết Từng Bước

### 5.1 Step 1: Regex Classification (`step1_classify.py`)

**Input:** `data/CRM_TDCTDA.xlsx`  
**Output:** `output/CRM_classified.xlsx`

**Quy trình:**

1. Đọc file `keywords_fixed.json` → build keyword index (col → label → keywords)
2. Đọc file Excel input → DataFrame
3. Thêm cột `row_idx` (2-based, matching Excel row numbers)
4. **Merge** 5 cột text nguồn thành 1 chuỗi `_merged_text`:
   - `Tình trạng hiện tại`
   - `Tình hình tiến độ công trình`
   - `Nội dung làm việc, yêu cầu KH & đánh giá`
   - `Kế hoạch lần tới`
   - `Đề xuất`
5. **Copy trực tiếp**: `[Hoạt Động CRM] Tình hình hiện tại` ← `Tình trạng hiện tại`
6. **Regex matching**: mỗi row, duyệt qua tất cả cột phân loại → match keyword → gán label
7. **Xử lý hãng đối thủ**: tìm tất cả brand names trong text → join bởi `"; "`
8. Xuất ra `CRM_classified.xlsx` + thống kê cells trống

**Cơ chế regex:**
- Sử dụng `re.IGNORECASE` cho tất cả pattern
- Keywords được `re.escape()` trước khi build pattern
- Match **first-match-wins** (label đầu tiên match được sử dụng)

---

### 5.2 Step 2: Prepare LLM Input (`step2_prepare_llm.py`)

**Input:** `output/CRM_classified.xlsx`  
**Output:** `output/llm_input.json`

**Quy trình:**

1. Đọc `CRM_classified.xlsx` → DataFrame
2. Build canonical maps (allowed values per column)
3. Duyệt từng row:
   - Skip nếu không có text nào
   - Tìm `locked_labels` (đã có giá trị từ regex)
   - Tìm `missing_cols` (cells trống hoặc == `"mơ hồ"`)
   - Skip nếu không có missing cols
4. Tạo JSON item format:

```json
{
  "row_idx": 123,
  "texts": {
    "Tình trạng hiện tại": "...",
    "Nội dung làm việc, yêu cầu KH & đánh giá": "..."
  },
  "existing": {"[AETT] Đối tượng": "Chủ đầu tư"},
  "missing_cols": ["[AETT] Nội dung làm việc", ...],
  "allowed": {
    "[AETT] Nội dung làm việc": ["Tiếp cận ban đầu", "Tư vấn khảo sát", ...]
  }
}
```

5. Xuất toàn bộ items → `llm_input.json`

---

### 5.3 Step 3: Call Gemini LLM (`step3_call_llm.py`)

**Input:** `output/llm_input.json` + `prompts/prompt_CRM_v5.txt`  
**Output:** `output/llm_fills.json` + `output/llm_fills_checkpoint.json`

**Quy trình:**

1. Load API key từ `.env`
2. Load system prompt từ `prompt_CRM_v5.txt`
3. Load items từ `llm_input.json`
4. Load checkpoint (nếu có — để resume)
5. **Batching**: chia items thành batches (mặc định 20 items/batch)
6. Mỗi batch:
   - Gọi `model.generate_content()` với prompt + JSON payload
   - Parse JSON response
   - **Canonicalize** labels: đối chiếu với allowed values
   - **Normalize dates**: thống nhất format `dd/mm/yy`
   - **Save checkpoint** ngay sau mỗi batch (resume-safe)
7. Export final `llm_fills.json`

**Cơ chế Resilience:**

| Tính năng | Chi tiết |
|-----------|----------|
| **Rate limiting** | Min 2.5s delay + 0.5s jitter giữa các lần gọi |
| **429 handler** | Backoff: 15s × attempt, max 120s |
| **Transient error** | Backoff: 10s × attempt |
| **Generic retry** | Exponential: 2^attempt + random |
| **Max retries** | 5 lần per batch |
| **Checkpoint** | Save sau mỗi batch → resume nếu crash |
| **Batch reduction** | Nếu lỗi JSON parse → giảm batch_size / 2 (min = 2) |

**Label canonicalization:**
- Tất cả labels LLM trả về đều được đối chiếu `.lower()` với canonical map
- Nếu không match → giá trị bị loại (set `None`)
- Cột ngày: normalize sang format `dd/mm/yy`
- Cột hãng đối thủ: allow any string (không filter)

---

### 5.4 Step 4: Merge Results (`step4_merge.py`)

**Input:** `output/CRM_classified.xlsx` + `output/llm_fills.json`  
**Output:** `output/CRM_classified_with_LLM.xlsx`

**Quy trình:**

1. Đọc `CRM_classified.xlsx` (kết quả regex)
2. Đọc `llm_fills.json` → build lookup by `row_idx`
3. Merge: chỉ fill cells **hiện đang trống/mơ hồ** 
4. **KHÔNG ghi đè** cells đã có giá trị từ regex
5. Export → `CRM_classified_with_LLM.xlsx`
6. In thống kê fill rate per column

**Quy tắc merge:**
```
is_empty = pd.isna(cur) 
         OR str(cur).strip() == '' 
         OR str(cur).strip().lower() == 'mơ hồ'

if is_empty AND fills[col] is not None:
    df[idx, col] = fills[col]
```

---

## 6. Hướng Dẫn Cài Đặt & Thiết Lập

### 6.1 Yêu Cầu Hệ Thống

- **OS:** Windows 10/11, Linux, macOS
- **Python:** ≥ 3.10
- **RAM:** ≥ 4GB (tùy kích thước file Excel)
- **Internet:** Cần cho Step 3 (gọi Gemini API)

### 6.2 Cài Đặt Từ Đầu

```bash
# 1. Clone repository
git clone <repo-url> D:\Works\CRM
cd D:\Works\CRM

# 2. Tạo virtual environment (khuyến nghị)
python -m venv .venv

# Windows:
.venv\Scripts\activate

# Linux/Mac:
# source .venv/bin/activate

# 3. Cài dependencies
pip install -r requirements.txt

# 4. Cài thêm google-generativeai (nếu chưa có)
pip install google-generativeai
```

### 6.3 Cấu Hình API Key

1. Truy cập [Google AI Studio](https://aistudio.google.com/apikey) → tạo API key
2. Sửa file `.env` tại thư mục gốc:

```env
# ─── CRM Classification Pipeline ───
GEMINI_API_KEY=AIzaSy____________________________
# GEMINI_MODEL=models/gemini-2.0-flash       # (tùy chọn)
# GEMINI_BATCH_SIZE=20                        # (tùy chọn)
```

> ⚠️ **QUAN TRỌNG:** File `.env` đã được gitignore. KHÔNG BAO GIỜ commit API key lên git.

### 6.4 Chuẩn Bị Dữ Liệu

1. Đặt file CRM đầu vào tại: `data/CRM_TDCTDA.xlsx`
2. Đảm bảo file Excel có các cột nguồn:
   - `Tình trạng hiện tại`
   - `Tình hình tiến độ công trình`
   - `Nội dung làm việc, yêu cầu KH & đánh giá`
   - `Kế hoạch lần tới`
   - `Đề xuất`

### 6.5 Verify Cài Đặt

```bash
cd D:\Works\CRM

# Test import
python -c "from src.config import *; print('Config OK:', PROJECT_ROOT)"

# Test API key
python -c "import os; from dotenv import load_dotenv; load_dotenv(); print('Key:', os.getenv('GEMINI_API_KEY', 'MISSING')[:10] + '...')"
```

---

## 7. Hướng Dẫn Vận Hành

### 7.1 Chạy Pipeline Đầy Đủ

```bash
cd D:\Works\CRM\src

# Chạy toàn bộ 4 bước
python run_pipeline.py
```

**Kết quả mong đợi:**
```
============================================================
CRM Classification Pipeline
Steps to run: [1, 2, 3, 4]
============================================================

────────────────────────────────────────
STEP 1: Regex Classification
────────────────────────────────────────
Loading keywords from keywords_fixed.json ...
  Columns indexed: 12
  Brands: 25
Reading CRM_TDCTDA.xlsx ...
  Rows: 500, Columns: 40
Running regex classification ...
✓ Regex classification done!
  Total label assignments: 3200

────────────────────────────────────────
STEP 2: Prepare LLM Input
────────────────────────────────────────
✓ LLM input prepared!
  Rows needing LLM fills: 180
  Rows with no text (skipped): 20
  Rows fully classified: 300

────────────────────────────────────────
STEP 3: Call LLM (Gemini)
────────────────────────────────────────
✓ Loaded prompt: prompt_CRM_v5.txt (12015 chars)
✓ Loaded 180 items
  Checkpoint: done=0, pending=180
Bắt đầu gọi LLM (batch_size=20)...
  0/180 bs=20 ✓ updated=20
  20/180 bs=20 ✓ updated=20
  ...
✓ LLM done: 180/180 rows → llm_fills.json

────────────────────────────────────────
STEP 4: Merge LLM Results
────────────────────────────────────────
✓ Merge done!
  Total cells filled by LLM: 850

============================================================
✓ Pipeline finished!
============================================================
```

### 7.2 Chạy Từng Bước

```bash
cd D:\Works\CRM\src

# Chỉ bước 1 (regex) — nhanh, không cần API
python run_pipeline.py 1

# Bước 1+2 (regex + prepare LLM)
python run_pipeline.py 1 2

# Chỉ bước 3+4 (LLM + merge — khi đã có output bước 1,2)
python run_pipeline.py 3 4

# Resume bước 3 (nếu bị crash giữa chừng)
python run_pipeline.py 3
```

### 7.3 Quy Trình Vận Hành Hàng Tuần

```
1. Nhận file CRM mới từ Sales team
                │
2. Backup file cũ:
   │  copy output\CRM_classified_with_LLM.xlsx → output\archive\
   │
3. Thay thế file input:
   │  copy <file_mới> → data\CRM_TDCTDA.xlsx
   │
4. Xóa checkpoint cũ (nếu muốn chạy mới hoàn toàn):
   │  del output\llm_fills_checkpoint.json
   │
5. Chạy pipeline:
   │  cd D:\Works\CRM\src
   │  python run_pipeline.py
   │
6. Kiểm tra kết quả:
   │  Mở output\CRM_classified_with_LLM.xlsx
   │  Kiểm tra fill rate per column
   │
7. Gửi kết quả cho bộ phận sử dụng
```

### 7.4 Lấy Kết Quả

| File | Mô tả | Khi nào dùng |
|------|-------|--------------|
| `output/CRM_classified.xlsx` | Chỉ regex, không LLM | Khi cần kết quả nhanh, không cần AI |
| `output/CRM_classified_with_LLM.xlsx` | **Kết quả cuối cùng** | Dùng file này cho báo cáo |

---

## 8. Cấu Hình Hệ Thống

### 8.1 Biến Môi Trường (`.env`)

| Biến | Bắt buộc | Mặc định | Mô tả |
|------|:---:|---------|-------|
| `GEMINI_API_KEY` | ✅ | — | API key Google AI Studio |
| `GEMINI_MODEL` | ❌ | `models/gemini-2.0-flash` | Model Gemini sử dụng |
| `GEMINI_BATCH_SIZE` | ❌ | `20` | Số items per batch |
| `GEMINI_MIN_INTERVAL_S` | ❌ | `2.5` | Delay tối thiểu giữa các API calls (giây) |
| `GEMINI_JITTER_S` | ❌ | `0.5` | Random jitter thêm vào delay |

### 8.2 Paths Cấu Hình (`config.py`)

| Constant | Đường dẫn mặc định | Thay đổi khi |
|----------|-------------------|--------------|
| `PATH_INPUT` | `data/CRM_TDCTDA.xlsx` | Đổi tên file input |
| `PATH_KW_JSON` | `keywords/keywords_fixed.json` | Đổi file keywords |
| `PATH_PROMPT` | `prompts/prompt_CRM_v5.txt` | Cập nhật prompt |
| `PATH_OUTPUT_REGEX` | `output/CRM_classified.xlsx` | Đổi tên output |
| `PATH_OUTPUT_LLM` | `output/CRM_classified_with_LLM.xlsx` | Đổi tên output |

### 8.3 Cột Phân Loại (14 cột output)

| # | Tên Cột | Nhóm | Nguồn |
|---|---------|-------|-------|
| 1 | `[Hoạt Động CRM] Tình hình hiện tại` | Hoạt Động CRM | Copy trực tiếp |
| 2 | `[Hoạt Động CRM] Tiến độ` | Hoạt Động CRM | Regex + LLM |
| 3 | `[Hoạt Động CRM] ngày lấy hàng` | Hoạt Động CRM | LLM (date) |
| 4 | `[AETT] Nội dung làm việc` | AETT | Regex + LLM |
| 5 | `[AETT] Nhận xét tiếp thị` | AETT | Regex + LLM |
| 6 | `[AETT] Đối tượng` | AETT | Regex + LLM |
| 7 | `[Khách Hàng] Ý kiến KH` | Khách Hàng | Regex + LLM |
| 8 | `[Khách Hàng] Nhận xét KH` | Khách Hàng | Regex + LLM |
| 9 | `[Kế Hoạch] Kế hoạch lần tới` | Kế Hoạch | Regex + LLM |
| 10 | `[Kế Hoạch] Ngày làm việc/ giao hàng:` | Kế Hoạch | LLM (date) |
| 11 | `[Kế Hoạch] Đề xuất` | Kế Hoạch | Regex + LLM |
| 12 | `[Đối Thủ Cạnh Tranh] Nội dung làm việc` | Đối Thủ | Regex + LLM |
| 13 | `[Đối Thủ Cạnh Tranh] Đối tượng` | Đối Thủ | Regex + LLM |
| 14 | `[Đối Thủ Cạnh Tranh] Các Hãng đối thủ cạnh tranh` | Đối Thủ | Regex + LLM |

---

## 9. Taxonomy & Keyword Mapping

### 9.1 Cấu Trúc `keywords_fixed.json`

```json
{
  "MAJOR_GROUP": {
    "MINOR_COL": {
      "LABEL_1": ["keyword1", "keyword2", "keyword3"],
      "LABEL_2": ["keyword4", "keyword5"]
    }
  }
}
```

**Ví dụ thực tế:**
```json
{
  "AETT": {
    "Nội dung làm việc": {
      "Tiếp cận ban đầu": ["gặp lần đầu", "gửi catalogue", "giới thiệu"],
      "Tư vấn khảo sát": ["khảo sát", "tư vấn", "báo giá"],
      "Chốt giá": ["chốt giá", "thống nhất giá", "chốt đơn"]
    }
  }
}
```

### 9.2 Cách Cập Nhật Keywords

1. **Từ file Excel spec** (cách chính):
   ```bash
   cd D:\Works\CRM\src
   python extract_keywords_from_spec_NEW.py
   ```
   - Input: `docs/tu_khoa_cho_cac_cot_cong_trinh.xlsx`
   - Output: `keywords/keywords_fixed.json`

2. **Sửa trực tiếp** file `keywords_fixed.json`:
   - Thêm keyword mới vào array tương ứng
   - Thêm label mới: tạo key mới trong dict
   - **Lưu ý:** phải update cả prompt nếu thêm label mới

### 9.3 Danh Sách Tags Theo Nhóm

#### [Hoạt Động CRM] Tình hình hiện tại (12 tags)
| Priority | Tag | Trigger keywords |
|:---:|-----|-----------------|
| 1 | Tiền khả thi | Nghiên cứu khả thi, lập đề án |
| 2 | Phê duyệt chủ trương đầu tư | Chờ phê duyệt dự án |
| 3 | Thiết kế | Giai đoạn thiết kế chi tiết |
| 4 | Đấu thầu | Đang đấu thầu, dự thầu |
| 5 | Khởi công | Mới bắt đầu thi công |
| 6 | Xây thô | Đang xây phần thô |
| 7 | Thi công điện nước | Lắp hệ thống M&E |
| 8 | Hoàn thiện nội thất | Hoàn thiện bên trong |
| 9 | Đang hoàn thiện bên ngoài | Hoàn thiện mặt ngoài |
| 10 | Hoàn công, nghiệm thu | Bàn giao, nghiệm thu |
| 11 | Đang bảo trì, sửa chữa | Sửa chữa, bảo trì |
| 12 | Tạm dừng | Dừng do vốn/pháp lý |

#### [AETT] Nội dung làm việc (10 tags — priority ascending)
| Priority | Tag | Trigger keywords |
|:---:|-----|-----------------|
| 1 | Tiếp cận ban đầu | gặp lần đầu, gửi catalogue |
| 2 | Tư vấn khảo sát | khảo sát, tư vấn, báo giá |
| 3 | Thiết kế | thiết kế, bản vẽ, 3D |
| 4 | Tư vấn chuyên sâu | tư vấn chi tiết, dò giá |
| 5 | Chốt giá | chốt giá, thống nhất giá |
| 6 | Thương thảo hợp đồng | thương thảo, đàm phán |
| 7 | Ký hợp đồng | ký HĐ, đặt cọc |
| 8 | Cấp hàng | cấp hàng, giao hàng, ETA |
| 9 | Lắp đặt | lắp đặt, thi công, nghiệm thu |
| 10 | Bảo hành | bảo hành, bảo trì, warranty |

#### Các nhóm khác
- **[Hoạt Động CRM] Tiến độ**: Chậm, Nhanh, Tạm dừng
- **[AETT] Nhận xét tiếp thị**: Khả năng thành công, Cơ hội tiềm năng, Khả năng trượt, Không thành công
- **[AETT] Đối tượng** (11 tags): C1 CTDA, C2 CTDA, Chủ đầu tư, Nhà Thầu, Ban chỉ huy, chủ nhà, Đại lý uỷ quyền, Tư vấn giám sát, Tư vấn thiết kế, Vệ tinh
- **[Khách Hàng] Ý kiến KH**: Đánh giá cao, Thích, Đồng ý phương án, Đổi phương án
- **[Khách Hàng] Nhận xét KH**: Giá cao, So sánh giá, Ưu tiên giá rẻ, Sản phẩm lỗi, Đề xuất giảm giá, Thắc mắc sản phẩm
- **[Kế Hoạch] Đề xuất**: Sản phẩm, Cơ chế, Dịch vụ, Quan hệ, Tài liệu

---

## 10. Prompt LLM (Gemini)

### 10.1 File Prompt

- **Đường dẫn:** `prompts/prompt_CRM_v5.txt`
- **Phiên bản:** v5 (Optimized)
- **Kích thước:** ~12 KB

### 10.2 Cấu Trúc Prompt

| Section | Nội dung |
|---------|----------|
| **MỤC TIÊU** | Vai trò: hệ thống phân loại CRM công trình |
| **INPUT** | JSON array format, mô tả các fields |
| **OUTPUT** | JSON array response format (fills + evidence) |
| **QUY TẮC** | Chỉ fill missing_cols, chỉ dùng allowed values |
| **TAXONOMY** | Danh sách đầy đủ 50+ tags & priority rules |
| **NGÀY THÁNG** | Quy tắc normalize date format |
| **FEW-SHOT** | 6 ví dụ cụ thể cover conflict, brands, dates, feedback |

### 10.3 Quy Tắc Quan Trọng Trong Prompt

1. **CHỈ điền `missing_cols`** — không thay đổi `locked_labels`
2. **CHỈ chọn từ `allowed`** — không tự bịa tag mới
3. **Priority rule**: nếu nhiều signal → chọn tag sâu hơn (priority cao hơn)
4. **Giá tiền** → tối thiểu "Tư vấn khảo sát"
5. **Hãng đối thủ**: nhiều hãng ngăn bởi `"; "`, chung chung → `"Hãng cạnh tranh"`
6. **Không đủ tín hiệu** → trả `null` (KHÔNG trả `"mơ hồ"`)

### 10.4 Hướng Dẫn Cập Nhật Prompt

Khi thêm/đổi tags:
1. Sửa `keywords_fixed.json` (thêm keywords)
2. Sửa `prompt_CRM_v5.txt` (thêm tag vào taxonomy section)
3. Chạy test với 1 batch nhỏ: `python run_pipeline.py 2 3`
4. Kiểm tra `llm_fills.json` → xem LLM có trả đúng tag mới không

---

## 11. Xử Lý Lỗi & Troubleshooting

### 11.1 Lỗi Thường Gặp

| Lỗi | Nguyên nhân | Cách xử lý |
|-----|------------|------------|
| `❌ No API key found` | Thiếu `GEMINI_API_KEY` trong `.env` | Thêm API key vào `.env` |
| `❌ Prompt file not found` | Thiếu `prompts/prompt_CRM_v5.txt` | Kiểm tra file prompt |
| `❌ LLM input not found` | Chưa chạy Step 2 | Chạy `python run_pipeline.py 2` |
| `❌ LLM fills not found` | Chưa chạy Step 3 | Chạy `python run_pipeline.py 3` |
| `⚠️ 429/rate-limit` | Quá nhiều request | Tự retry, tăng `GEMINI_MIN_INTERVAL_S` |
| `Cannot find JSON array` | LLM trả response không đúng format | Tự retry, giảm batch_size |
| `Source column not found` | Cột input đổi tên | Sửa `INPUT_TEXT_COLUMNS` trong `config.py` |

### 11.2 Step 3 Bị Crash Giữa Chừng

**Không mất dữ liệu!** Hệ thống có checkpoint:

```bash
# Kiểm tra checkpoint
type output\llm_fills_checkpoint.json
# → Xem "updated_at" và số results đã xử lý

# Resume từ checkpoint
python run_pipeline.py 3

# Nếu muốn chạy lại từ đầu
del output\llm_fills_checkpoint.json
python run_pipeline.py 3
```

### 11.3 Chất Lượng Kết Quả Kém

| Triệu chứng | Nguyên nhân có thể | Giải pháp |
|--------------|-------------------|-----------|
| Nhiều cells vẫn trống | Keywords thiếu hoặc text input ngắn | Bổ sung keywords vào JSON |
| Gán sai label | Keyword trùng giữa các labels | Ưu tiên keyword cụ thể hơn |
| LLM trả null nhiều | Text quá mơ hồ hoặc prompt chưa tốt | Cải thiện prompt / few-shot examples |
| Date format sai | Regex normalize chưa cover edge case | Sửa regex trong `step3_call_llm.py` |

### 11.4 Performance Tuning

| Tham số | Mặc định | Tăng khi | Giảm khi |
|---------|----------|----------|----------|
| `BATCH_SIZE` | 20 | Ít row cần xử lý | Bị rate limit hoặc JSON parse fail |
| `MIN_INTERVAL_S` | 2.5 | Bị 429 liên tục | Có API key paid tier |
| `JITTER_S` | 0.5 | — | — |

---

## 12. Bảo Trì & Mở Rộng

### 12.1 Bảo Trì Định Kỳ

| Tần suất | Hành động |
|----------|-----------|
| Hàng tuần | Kiểm tra fill rate → cải thiện keywords/prompt |
| Hàng tháng | Review kết quả sample → phát hiện sai label |
| Khi có hãng mới | Thêm vào `keywords_fixed.json` → section đối thủ |
| Khi đổi cột CRM | Cập nhật `INPUT_TEXT_COLUMNS` trong `config.py` |

### 12.2 Thêm Cột Phân Loại Mới

1. Thêm definition trong `config.py`:
   ```python
   COL_NEW = col_name("Nhóm", "Tên cột mới")
   ```
2. Thêm vào `OUTPUT_COLUMNS` list
3. Thêm keywords vào `keywords_fixed.json`
4. Thêm taxonomy tags vào `prompt_CRM_v5.txt`
5. Test: `python run_pipeline.py`

### 12.3 Đổi LLM Model

Sửa `.env`:
```env
GEMINI_MODEL=models/gemini-2.0-pro
# hoặc
GEMINI_MODEL=models/gemini-1.5-flash
```

### 12.4 Nâng Cấp Keywords Từ Excel Spec

```bash
# Khi có file spec mới
cd D:\Works\CRM\src
python extract_keywords_from_spec_NEW.py
# → Xuất ra keywords_fixed.json mới
```

---

## 13. Câu Hỏi Thường Gặp (FAQ)

### Q1: Chạy mất bao lâu?
> Step 1+2+4 chạy trong vài giây. Step 3 (LLM) mất **5–60 phút** tùy số rows cần xử lý và rate limit.

### Q2: Có cần internet không?
> Chỉ Step 3 cần internet (gọi Gemini API). Step 1, 2, 4 chạy offline.

### Q3: File input phải đúng format gì?
> File `.xlsx` với ít nhất 5 cột text nguồn (xem mục 6.4). Tên cột **phải trùng khớp** với `INPUT_TEXT_COLUMNS` trong `config.py`.

### Q4: Bị rate limit (429) thì sao?
> Tự động retry với backoff 15s → 120s. Nếu liên tục bị → tăng `GEMINI_MIN_INTERVAL_S` trong `.env`, hoặc dùng API key paid tier.

### Q5: Kết quả regex đủ tốt, không cần LLM?
> Hoàn toàn OK! Chạy `python run_pipeline.py 1` → dùng `CRM_classified.xlsx`.

### Q6: Muốn thay đổi prompt thì sửa ở đâu?
> Sửa file `prompts/prompt_CRM_v5.txt`. Có thể tạo version mới `prompt_CRM_v6.txt` rồi update path trong `config.py`.

### Q7: Thêm hãng đối thủ mới?
> Thêm vào `keywords_fixed.json` → section `Đối Thủ Cạnh Tranh` → `Cạnh Tranh` → `Các Hãng đối thủ cạnh tranh`.

### Q8: Có thể chạy trên server/cloud không?
> Có. Chỉ cần Python + dependencies + `.env` file. Có thể chạy headless (không cần GUI).

---

## 14. Liên Hệ & Tham Khảo

### Tài Liệu Liên Quan

| Tài liệu | Vị trí |
|-----------|--------|
| Checklist tính năng BOT | `docs/Check list tính năng BOT Khai thác CRM.pdf` |
| Spec keywords (Excel) | `docs/tu_khoa_cho_cac_cot_cong_trinh.xlsx` |
| Định nghĩa tiến độ | `docs/Định nghĩa tiến độ công trình_12.7.25 (1).xlsx` |
| Notebooks tham khảo | `notebooks/CRM_Classification.ipynb` |

### Git Repository

```bash
# Xem lịch sử thay đổi
cd D:\Works\CRM
git log --oneline -20

# Kiểm tra trạng thái
git status
```

### API Reference

- [Google AI Studio - API Key](https://aistudio.google.com/apikey)
- [Gemini API Docs](https://ai.google.dev/gemini-api/docs)
- [Rate Limits](https://ai.google.dev/gemini-api/docs/rate-limits)

---

> 📝 **Ghi chú bàn giao:** Toàn bộ source code nằm trong `src/`. File quan trọng nhất cần bảo trì là `keywords_fixed.json` (cập nhật keywords) và `prompt_CRM_v5.txt` (cải thiện LLM quality). Pipeline đã có checkpoint support nên an toàn khi chạy trên dữ liệu lớn.
