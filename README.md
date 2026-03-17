# CRM Classification Pipeline

Phân loại dữ liệu CRM tự động bằng regex + Gemini LLM.

## Cấu trúc thư mục

```
CRM/
├── src/                  # Source code
│   ├── config.py         # Cấu hình chung (paths, columns, keywords)
│   ├── step1_classify.py # Phân loại bằng regex
│   ├── step2_prepare_llm.py  # Chuẩn bị input cho LLM
│   ├── step3_call_llm.py     # Gọi Gemini API (batching, retry, checkpoint)
│   ├── step4_merge.py        # Gộp kết quả LLM vào Excel
│   └── run_pipeline.py       # Chạy toàn bộ hoặc từng bước
├── data/                 # Dữ liệu đầu vào (gitignored)
├── output/               # Kết quả phân loại (gitignored)
├── prompts/              # Prompt templates cho LLM
├── keywords/             # File JSON keyword mapping
├── notebooks/            # Jupyter notebooks gốc (tham khảo)
├── docs/                 # Tài liệu tham khảo (gitignored)
├── .env                  # API keys (gitignored)
└── requirements.txt
```

## Cài đặt

```bash
pip install -r requirements.txt
```

## Sử dụng

### 1. Cấu hình API key

Sửa file `.env`:
```
GEMINI_API_KEY=your-key-here
```

### 2. Chạy pipeline

```bash
# Chạy toàn bộ 4 bước
python src/run_pipeline.py

# Hoặc chạy từng bước
python src/run_pipeline.py 1      # Regex classification
python src/run_pipeline.py 2      # Prepare LLM input
python src/run_pipeline.py 3      # Call Gemini API
python src/run_pipeline.py 4      # Merge results
python src/run_pipeline.py 1 2    # Steps 1+2 only
```

### 3. Output

- `output/CRM_classified.xlsx` — kết quả sau bước regex
- `output/CRM_classified_with_LLM.xlsx` — kết quả cuối cùng (regex + LLM)

## Pipeline

| Bước | File | Mô tả |
|------|------|-------|
| 1 | `step1_classify.py` | Đọc `CRM_TDCTDA.xlsx`, match keyword regex → 14 cột phân loại |
| 2 | `step2_prepare_llm.py` | Tìm cells còn trống, tạo `llm_input.json` |
| 3 | `step3_call_llm.py` | Gọi Gemini API với batching + checkpoint (resume được) |
| 4 | `step4_merge.py` | Gộp kết quả LLM vào Excel, chỉ fill cells trống |
