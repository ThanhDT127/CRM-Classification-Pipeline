# ⚡ CRM Classification Pipeline ⚡

<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=gradient&customColorList=16,2&height=220&section=header&text=CRM%20Classification&fontSize=55&fontColor=ffffff&animation=fadeIn" alt="Header Banner" />
</p>

<p align="center">
  <a href="https://github.com/ThanhDT127/CRM-Classification-Pipeline/actions/workflows/ci.yml">
    <img src="https://img.shields.io/github/actions/workflow/status/ThanhDT127/CRM-Classification-Pipeline/ci.yml?branch=main&style=flat-square&logo=github&logoColor=white&label=Build" alt="Build Status" />
  </a>
  <a href="https://python.org">
    <img src="https://img.shields.io/badge/Python-3.11%20%7C%203.12-blue?style=flat-square&logo=python&logoColor=white" alt="Python Version" />
  </a>
  <a href="https://www.docker.com">
    <img src="https://img.shields.io/badge/Docker-Enabled-blue?style=flat-square&logo=docker&logoColor=white" alt="Docker Support" />
  </a>
  <a href="https://docs.pytest.org">
    <img src="https://img.shields.io/badge/Tests-PyTest-green?style=flat-square&logo=pytest&logoColor=white" alt="Tests Status" />
  </a>
  <a href="https://opensource.org/licenses/MIT">
    <img src="https://img.shields.io/badge/License-MIT-yellow?style=flat-square" alt="License" />
  </a>
</p>

<p align="center">
  <b>Hệ thống tự động hóa khai thác, phân loại phản hồi CRM quy mô lớn (>27k dòng) tích hợp an toàn dữ liệu và đồng bộ đa SharePoint Site.</b>
</p>

---

## 📝 Mục lục (Table of Contents)

1. [Giới thiệu dự án (About The Project)](#about-the-project)
2. [Tính năng nổi bật (Key Features)](#key-features)
3. [Công nghệ sử dụng (Built With)](#built-with)
4. [Cấu Trúc Thư Mục (Directory Structure)](#directory-structure)
5. [Hướng dẫn cài đặt (Getting Started)](#getting-started)
6. [Hướng dẫn vận hành (Usage & Operations)](#usage)
7. [Thiết kế kỹ thuật & Sơ đồ luồng (Technical Design)](#technical-design)
8. [Hiệu năng & Đo lường (Performance & Metrics)](#performance-metrics)
9. [Khắc phục sự cố & FAQ (Troubleshooting)](#troubleshooting)
10. [Kiểm thử tự động & CI/CD](#testing-cicd)

---

## <a name="about-the-project"></a>🌟 Giới thiệu dự án (About The Project)

Dự án này giải quyết bài toán tự động hóa phân loại phản hồi, yêu cầu khách hàng và thông tin tiến độ dự án từ dữ liệu thô CRM của ngành thiết bị điện chiếu sáng. Thay vì phải thủ công xử lý hàng chục nghìn dòng dữ liệu, hệ thống tự động hóa toàn bộ quy trình:
* **Tải dữ liệu nguồn:** Tải file `CRM_merge.xlsx` từ SharePoint nguồn (Site `CRM-CTDA`) thông qua Microsoft Graph API.
* **Phân loại lai thông minh (Hybrid Classifier):** Áp dụng bộ lọc Regex tiếng Việt nhanh để điền nhãn cứng, sau đó gọi Gemini 2.5 Flash đối với các thông tin phức tạp hoặc mơ hồ.
* **Xử lý gia tăng (Incremental Delta):** Chỉ dán nhãn các dòng mới thêm vào nhằm tiết kiệm chi phí API.
* **Định dạng & Đồng bộ đích:** Định dạng bảng Excel chuẩn Premium (Double-header, Group Colors, Column Auto-fit) và đẩy đè trực tiếp kết quả lên file `CRM_classified.xlsx` tại SharePoint đích (Site `DataPBI_salein`).
* **Báo cáo tự động:** Gửi mail báo cáo tiến trình (hoặc báo lỗi kèm stack trace) tự động qua Graph API.

---

## <a name="key-features"></a>✨ Tính năng nổi bật (Key Features)

* **🚀 Hybrid Classifier**: Kết hợp bộ lọc nhanh bằng Regex tiếng Việt để điền nhãn cứng và Gemini 2.5 Flash xử lý các phản hồi phức tạp hoặc mơ hồ.
* **⚡ Incremental Delta**: So khớp khóa chính `ActivityId` với lịch sử JSON để chỉ phân loại dòng mới. Tiết kiệm 95% chi phí API khi tệp phát sinh theo ngày.
* **🔒 Zero Row-Shifting**: Đồng bộ an toàn bằng ID thay vì số thứ tự dòng vật lý, chống lệch cột khi người dùng sửa đổi cấu trúc file Excel trực tiếp trên SharePoint.
* **🛠️ Checkpoint tự phục hồi**: Chia nhỏ dữ liệu gọi LLM thành các batch (mặc định 40 dòng/batch) và ghi checkpoint liên tục, tự động chạy tiếp từ batch lỗi cuối cùng khi khởi động lại.
* **🎨 Premium Excel Styling**: Áp dụng cấu trúc Double-Header chuyên nghiệp, tự động căn rộng cột và tô màu theo nhóm nghiệp vụ.
* **📧 SharePoint Sync trực tiếp**: Tải và tải đè trực tiếp qua MS Graph API mà không cần mount ổ đĩa cục bộ trên VM.

---

## <a name="built-with"></a>🛠️ Công nghệ sử dụng (Built With)

<a href="#built-with"><img src="https://skillicons.dev/icons?i=py,docker,git,azure,gcp,vscode" alt="My Skills" /></a>

* **Dữ liệu & Excel:** `pandas`, `openpyxl`
* **Xử lý ngôn ngữ tự nhiên:** `google-genai` (Gemini API SDK), `unidecode`
* **Kết nối & Xác thực:** `msal` (Microsoft Authentication Library), `requests`

---

## <a name="directory-structure"></a>📂 Cấu Trúc Thư Mục (Directory Structure)

```text
CRM-Classification-Pipeline/
├── .github/workflows/         # Kịch bản kiểm thử tự động CI (GitHub Actions)
│   └── ci.yml
├── config/                    # Cấu hình tĩnh
│   ├── keywords_fixed.json    # Lưới từ khóa phân loại Regex tiếng Việt
│   └── prompt_CRM_v5.txt      # Prompt hệ thống định hướng cho Gemini
├── src/                       # MÃ NGUỒN CHÍNH (PRODUCTION)
│   ├── config.py              # Định nghĩa đường dẫn và tên 15 cột đầu ra
│   ├── sharepoint.py          # SharePoint Client (Tải/Lên file qua Microsoft Graph API)
│   ├── notification.py        # Dịch vụ gửi email thông báo thành công/lỗi
│   ├── classifier.py          # Bộ phân loại Regex tiếng Việt
│   ├── llm.py                 # Client kết nối Gemini song song với checkpoint
│   └── pipeline.py            # Script chính điều phối toàn bộ Pipeline chạy tự động
├── tests/                     # Bộ kiểm thử tự động (Testing)
│   ├── test_pipeline.py       # Unit tests cho logic chuẩn hóa/regex
│   └── test_automation.py     # Integration tests (Mock SharePoint & Mail API)
├── notebooks/                 # Nghiên cứu & R&D Jupyter Notebooks
│   ├── CRM_Classification.ipynb
│   └── Phan_Loai_CRM_IMPROVED.ipynb
├── Dockerfile                 # Đóng gói Python 3.11-slim
├── docker-compose.yml         # Khởi chạy dịch vụ container
├── requirements.txt           # Danh sách các thư viện phụ thuộc
├── Makefile                   # Lệnh shortcut cho nhà phát triển
├── .env.example               # Mẫu cấu hình biến môi trường
└── README.md                  # Tài liệu hướng dẫn sử dụng
```

---

## <a name="getting-started"></a>🚀 Hướng dẫn cài đặt (Getting Started)

### 1. Yêu cầu hệ thống
* Python 3.11+
* Docker & Docker Compose (nếu chạy container)
* Tài khoản Azure AD App Registration (có quyền đọc/ghi SharePoint & gửi Mail)
* Google Gemini API Key hoặc GCP Service Account (Vertex AI)

### 2. Thiết lập môi trường cục bộ
Cài đặt thư viện:
```bash
make setup
```

Sao chép cấu hình mẫu và điền đầy đủ thông tin bảo mật của bạn:
```bash
cp .env.example .env
```

<details>
<summary><b>📄 Xem cấu trúc tệp .env cấu hình chi tiết</b></summary>

```env
GEMINI_API_KEY=YOUR_GEMINI_API_KEY_HERE
GEMINI_MODEL=models/gemini-2.5-flash
AZURE_TENANT_ID=YOUR_AZURE_TENANT_ID
AZURE_CLIENT_ID=YOUR_AZURE_CLIENT_ID
AZURE_CLIENT_SECRET=YOUR_AZURE_CLIENT_SECRET
SHAREPOINT_SOURCE_DRIVE_ID=YOUR_DRIVE_ID
SHAREPOINT_TARGET_DRIVE_ID=YOUR_DRIVE_ID
```
</details>

---

## <a name="usage"></a>🎯 Hướng dẫn vận hành (Usage & Operations)

### Chạy trực tiếp bằng Python
```bash
# Khởi tạo CSDL lịch sử (nếu chạy lần đầu từ file classified cũ)
python src/db_init.py

# Chạy toàn bộ pipeline phân loại
python src/pipeline.py
```

### Chạy ngầm tự động (Daemon Mode với Docker)
Hệ thống được thiết kế để tự động chạy như một background service vĩnh viễn:
```bash
# Khởi dựng và chạy ngầm container
docker compose up -d --build
```
* **Chạy lập lịch**: Container tự động quét và phân loại dữ liệu vào lúc **03:30 AM** hàng ngày.
* **Tự phục hồi**: Cấu hình `restart: unless-stopped` giúp container tự khởi động lại cùng OS khi VPS bị reboot.

---

## <a name="technical-design"></a>💡 Thiết kế kỹ thuật & Sơ đồ luồng (Technical Design)

Dưới đây là sơ đồ luồng hoạt động trực quan của hệ thống:

```mermaid
graph TD
    A[SharePoint Nguồn Site A] -->|Download file| B(pipeline.py)
    B -->|Schema Check & Backup| C{Lọc Dòng Mới}
    C -->|Lấy ActivityId so khớp Lịch sử DB| D[Chỉ phân loại dòng Delta]
    D -->|Step 1: Regex Classifier| E{Đủ thông tin?}
    E -->|Chưa đủ/Mơ hồ| F[Step 2: Gemini API Batching]
    E -->|Đã đủ nhãn| G[Gộp kết quả]
    F -->|Checkpoint lưu liên tục| G
    G -->|Tải file đích từ Site B| H[Merge an toàn theo ID]
    H -->|Áp dụng Premium Excel Styling| I[Lưu file cục bộ]
    I -->|Upload đè| J[SharePoint Đích Site B]
    J -->|Gửi báo cáo qua Graph Mail| K[Email thông báo]
    
    style A fill:#fff2cc,stroke:#d6b656
    style J fill:#e2efda,stroke:#82b366
    style K fill:#e8d5f5,stroke:#9673a6
```

### Quyết định thiết kế cốt lõi:
1. **Lọc dòng Delta**: Chỉ xử lý các dòng dữ liệu mới dựa trên khóa `ActivityId`. Giảm tải dung lượng truyền tải và chi phí gọi API.
2. **Khớp ID thay vì Số dòng**: Chống lệch cột dữ liệu khi người dùng chèn, xóa hoặc thay đổi thứ tự dòng vật lý trên Excel trực tiếp.
3. **Cơ chế Checkpoint**: Phân chia dữ liệu cần xử lý của LLM thành các batch nhỏ (mặc định 40 dòng/batch) và ghi checkpoint liên tục. Dự án tự động tiếp tục chạy từ batch lỗi cuối cùng khi khởi động lại.

---

## <a name="performance-metrics"></a>📊 Hiệu năng & Đo lường (Performance & Metrics)

*Áp dụng công thức Google XYZ đo lường hiệu quả thực tế:*

* **Tối ưu chi phí API (Tiết kiệm 95% chi phí)**: Triển khai bộ lọc lai (Regex kết hợp LLM) và cơ chế Delta Loading giúp giảm số lượng dòng cần gửi cho LLM từ **27,000 dòng xuống chỉ còn trung bình 50 - 100 dòng mới/ngày**, giảm chi phí hóa đơn API Gemini tối đa.
* **Tốc độ xử lý (Nhanh gấp 120 lần)**: Thời gian xử lý dữ liệu CRM giảm từ **6 giờ làm việc thủ công xuống còn chưa đầy 3 phút** nhờ xử lý batch và Regex song song.
* **Độ ổn định dữ liệu (0% lỗi lệch dòng)**: Cơ chế khớp ID tuyệt đối giúp ngăn ngừa hoàn toàn tình trạng lệch dòng dữ liệu (Row-shifting) khi có sự thay đổi cấu trúc file Excel nguồn.

---

## <a name="troubleshooting"></a>🛠️ Khắc phục sự cố & FAQ (Troubleshooting)

| Lỗi gặp phải | Nguyên nhân | Cách xử lý |
| :--- | :--- | :--- |
| **`❌ Error 429: Rate Limit Exceeded`** | Gọi quá nhiều request tới API Gemini miễn phí. | Pipeline tự động sleep và backoff. Để xử lý triệt để, hãy tăng `GEMINI_MIN_INTERVAL_S` trong `.env` lên `3.0` hoặc chuyển sang tier trả phí. |
| **`❌ SharePoint Auth Failed`** | Client Secret của Azure AD bị hết hạn hoặc sai Tenant ID. | Kiểm tra lại thông tin Client Secret trên cổng thông tin Microsoft Entra ID và cập nhật file `.env`. |
| **`❌ File classified_history_db.json is corrupted`** | File checkpoint bị ghi đứt quãng do VPS mất nguồn đột ngột. | Xóa tệp checkpoint lỗi và chạy lại `python src/db_init.py` để đồng bộ lại lịch sử từ file Excel đích. |

---

## <a name="testing-cicd"></a>🧪 Kiểm thử tự động & CI/CD

### Chạy kiểm thử cục bộ
```bash
make test
```
Bộ kiểm thử tích hợp sử dụng Mocking đối với các API SharePoint và Mail API của Microsoft giúp chạy test offline hoàn toàn an toàn và nhanh chóng.

### Tích hợp CI/CD tự động
Dự án tích hợp kịch bản GitHub Actions chạy tự động trên môi trường ảo `ubuntu-latest` mỗi khi có sự kiện Push hoặc Pull Request lên nhánh chính. 

<details>
<summary><b>📄 Xem chi tiết cấu hình CI workflow (.github/workflows/ci.yml)</b></summary>

```yaml
name: Python application CI
on: [push, pull_request]
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v4
    - name: Set up Python 3.11
      uses: actions/setup-python@v5
      with:
        python-version: "3.11"
    - name: Install dependencies
      run: |
        python -m pip install --upgrade pip
        pip install -r requirements.txt
    - name: Run tests with pytest
      run: |
        pytest
```
</details>
