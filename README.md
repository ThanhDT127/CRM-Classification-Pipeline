# CRM Classification Pipeline

[![Python CI Pipeline](https://github.com/ThanhDT127/CRM-Classification-Pipeline/actions/workflows/ci.yml/badge.svg)](https://github.com/ThanhDT127/CRM-Classification-Pipeline/actions/workflows/ci.yml)
[![Python Version](https://img.shields.io/badge/python-3.11-blue.svg)](https://python.org)
[![Docker Support](https://img.shields.io/badge/docker-enabled-blue.svg)](https://www.docker.com)
[![Testing](https://img.shields.io/badge/tests-pytest-green.svg)](https://docs.pytest.org)

Hệ thống phân loại tự động dữ liệu phản hồi CRM quy mô lớn (>27,000 dòng) dành cho ngành thiết bị điện chiếu sáng bằng phương pháp kết hợp **Regex (Keyword Matching)** và **LLM (Gemini 2.5 Flash trên Vertex AI)**.

---

## 🌟 Tính Năng Nổi Bật (Key Features)

1. **Xử lý lai (Hybrid Regex + LLM Classifier):** Lọc nhanh qua bộ từ khóa phân loại cứng bằng Regex trước để giảm tải, chỉ gửi các ô trống hoặc bị phân loại "mơ hồ" sang LLM (Gemini 2.5 Flash).
2. **Xử lý gia tăng (Incremental Delta Load):** Tự động phát hiện các dòng dữ liệu mới được nối thêm vào file nguồn, chỉ phân loại những dòng mới để tiết kiệm tài nguyên và chi phí gọi API.
3. **An toàn dữ liệu tuyệt đối (Zero Row-Shifting):** Gộp kết quả ngược lại file Excel thông qua ánh xạ khóa chính (`ActivityId`) thay vì chỉ mục dòng vật lý. Loại bỏ hoàn toàn nguy cơ xô lệch dữ liệu khi người dùng chèn, xóa hoặc sắp xếp lại dòng trên SharePoint.
4. **Khả năng tự phục hồi (Crash-Safe Checkpointing):** Lưu tiến trình xử lý liên tục. Nếu gặp lỗi kết nối hoặc cạn kiệt tài nguyên (Rate Limit 429), hệ thống có thể tiếp tục từ mốc dừng cuối cùng mà không cần chạy lại từ đầu.
5. **Kháng lỗi đa luồng & timeout:** Chạy bất đồng bộ đa luồng (`ThreadPoolExecutor`) kèm cấu hình Rate-limit thích ứng. Sửa đổi giới hạn timeout SDK lên 120 giây và lọc ký tự không ASCII để tránh crash bảng điều khiển Windows.
6. **Định dạng bảng biểu cao cấp (Premium Excel Formatting):** Tự động áp dụng cấu trúc tiêu đề kép (Double-row Header), tô màu sắc trực quan theo từng nhóm chuyên môn (CRM, AETT, Khách hàng, Kế hoạch, Đối thủ) và tự động kéo rộng độ rộng cột (Auto-fit).
7. **Đóng gói Docker sẵn sàng:** Tách biệt môi trường Phát triển và Triển khai tự động chạy ngầm trên Máy ảo (VM).

---

## 📂 Cấu Trúc Thư Mục (Repository Structure)

Dự án được phân tách rõ ràng thành hai phân vùng:

```text
CRM-Classification-Pipeline/
├── .github/workflows/          # Cấu hình kiểm thử tự động (CI/CD GitHub Actions)
│   └── ci.yml
├── automation/                  # MÔI TRƯỜNG TRIỂN KHAI TỰ ĐỘNG (DOCKER / VM / PRODUCTION)
│   ├── config/                  # Keywords và Prompt templates tĩnh
│   ├── src/                     # Code chính cho automation (Delta, ID matching)
│   ├── Dockerfile
│   ├── docker-compose.yml
│   └── requirements.txt
├── src/                         # MÔI TRƯỜNG PHÁT TRIỂN / R&D (DEVELOPMENT & RESEARCH)
│   ├── config.py                # Cấu hình chung
│   ├── step1_classify.py        # Phân loại bằng regex
│   ├── step2_prepare_llm.py     # Chuẩn bị payload LLM
│   ├── step3_call_llm.py        # Gọi API Gemini song song có checkpoint
│   ├── step4_merge.py           # Gộp kết quả và định dạng Excel
│   └── run_pipeline.py          # Script chạy test từng bước
├── tests/                       # Thư mục kiểm thử phần mềm (Testing)
│   └── test_pipeline.py         # Unit tests cho logic cốt lõi
├── prompts/ / keywords/         # Cấu hình tĩnh cho môi trường phát triển
├── notebooks/                   # File Jupyter Notebook gốc
├── .env.example                 # File cấu hình mẫu môi trường
├── sa-key.json.example          # File key dịch vụ GCP mẫu
├── .gitignore
├── requirements.txt             # Thư viện cho môi trường phát triển
└── README.md
```

---

## 🚀 Cài Đặt & Chạy Môi Trường Phát Triển (Local Setup)

> [!IMPORTANT]
> **Bảo mật dữ liệu (Data Confidentiality):**
> File dữ liệu gốc chứa thông tin khách hàng nhạy cảm đã được lược bỏ khỏi repository công khai. 
> Vui lòng chạy lệnh sinh dữ liệu giả lập (`make sample`) để tạo file kiểm thử nhanh tại `sample_data/CRM_merge_sample.xlsx` trước khi chạy pipeline.

### 1. Cài đặt thư viện & Dữ liệu mẫu
Dự án hỗ trợ `Makefile` để đơn giản hóa các thao tác:
```bash
# Cài đặt thư viện
make setup

# Tạo dữ liệu giả lập kiểm thử nhanh
make sample
```

### 2. Thiết lập cấu hình biến môi trường
Tạo file `.env` từ file `.env.example`:
```bash
cp .env.example .env
```
Điền key API của bạn vào `.env`:
* Dùng Google AI Studio: Điền `GEMINI_API_KEY=AIzaSy...`
* Dùng Vertex AI: Đặt `USE_VERTEX=True` và lưu file key dịch vụ GCP của bạn tại `./sa-key.json`.

### 3. Chạy Pipeline kiểm thử từng bước
Bạn có thể sử dụng các lệnh shortcut qua `Makefile`:
```bash
# Chạy toàn bộ pipeline (Regex + LLM)
make run

# Hoặc chạy thủ công từng bước cụ thể bằng script
python src/run_pipeline.py 1      # Chỉ chạy phân loại Regex
python src/run_pipeline.py 3 4    # Chỉ chạy gọi LLM và gộp kết quả vào Excel
```

---

## 📦 Triển Khai Chạy Tự Động Với Docker (Production / VM / SharePoint)

Thư mục `./automation` được thiết kế độc lập, đóng gói gọn nhẹ để deploy lên máy ảo chạy ngầm hằng ngày (ví dụ: 3h30 sáng).

### 1. Cách thức hoạt động
1. SharePoint hoặc OneDrive đồng bộ file Excel nguồn về máy ảo Host tại thư mục `/mnt/sharepoint/Input/CRM_merge.xlsx`.
2. Máy ảo cài đặt **Cron Job** tự động kích hoạt container lúc 3h30 sáng hằng ngày.
3. Container khởi chạy quét Delta dòng mới dựa theo `ActivityId`, gọi LLM, gộp kết quả an toàn, áp dụng định dạng Excel và lưu đè trực tiếp lên SharePoint.

### 2. Khởi chạy bằng Docker Compose
Di chuyển vào thư mục `automation`, điền các tham số API Key hoặc file `sa-key.json` tương ứng rồi chạy:
```bash
cd automation
docker compose up --build
```

---

## 🧪 Kiểm Thử Phần Mềm (Testing)

Dự án sử dụng `pytest` để kiểm thử tự động các chức năng chuẩn hóa dữ liệu đầu vào.

Chạy test suite cục bộ:
```bash
pytest
```

Các test case sẽ tự động được chạy thông qua **GitHub Actions CI Pipeline** trên mỗi lượt Push hoặc Pull Request để đảm bảo code luôn hoạt động ổn định trước khi tích hợp.

---

## 💡 Quyết Định Thiết Kế Kỹ Thuật (Design Decisions & Trade-offs)

* **Tại sao dùng ActivityId làm Index gộp dữ liệu?** 
  * *Vấn đề:* Người dùng SharePoint thường xuyên thao tác chèn dòng, xóa dòng, hoặc Sort lại dữ liệu làm dịch chuyển số thứ tự dòng vật lý (`row_idx`).
  * *Giải pháp:* Ghi nhận kết quả phân loại ánh xạ theo `ActivityId`. Dù vị trí dòng trong file Excel thay đổi, code vẫn đối chiếu và cập nhật chính xác tuyệt đối.
* **Xử lý Rate Limit 429 và Timeout 5s của SDK:**
  * *Vấn đề:* Gọi API hàng chục nghìn dòng với tốc độ cao dễ gây lỗi nghẽn cổ chai (Quota Exhausted) hoặc lỗi timeout do model sinh văn bản tiếng Việt dài.
  * *Giải pháp:* Tăng timeout của client lên 120s, chia sub-batch nhỏ (20-40 dòng/lượt), và cấu hình hàm delay thích ứng (exponential backoff with jitter) để tự động ngủ đông khi gặp lỗi 429 trước khi thử lại.
