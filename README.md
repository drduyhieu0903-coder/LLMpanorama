# 🦷 Dental AI Evaluation Tool – MLLM Research

Hệ thống đánh giá và đối chuẩn năng lực các mô hình ngôn ngữ thị giác lớn (Multimodal LLMs: **GEMINI**, **CLAUDE**, **CHATGPT**) trong phân loại răng khôn hàm dưới lệch ngầm theo phân loại **Pell & Gregory** và **Winter**. Nghiên cứu gồm **129 ca bệnh** với 4 thành viên nhập liệu (**GIANG**, **HOÀNG**, **MINH**, **HÂN**).

---

## 🌟 Tính Năng Nổi Bật

1. **Giao diện Web Streamlit chuyên nghiệp:**
   - Thiết kế chuẩn phong cách Medical Dark UI (`#0f172a`, `#00f2fe`).
   - Tùy chọn 4 thành viên nhập liệu chuẩn (**GIANG**, **HOÀNG**, **MINH**, **HÂN**) và 3 mô hình AI (**GEMINI**, **CLAUDE**, **CHATGPT**).

2. **Bóc tách phản hồi AI tự động (Smart Auto-Parser):**
   - Tự động trích xuất các trường: *P&G Class (I, II, III)*, *P&G Position (A, B, C)*, *Winter's Class*, *Pederson Level*, *Độ tin cậy Confidence (%)*.
   - Hỗ trợ cả định dạng cấu trúc JSON lẫn báo cáo giải phẫu văn bản tự nhiên.

3. **Bảng tính Google Sheets chuyên biệt (Interactive Spreadsheet View):**
   - Sắp xếp thứ tự ca chuẩn tự nhiên (**129 ca**: `Case_001` → `Case_129`).
   - Bộ lọc theo mô hình AI, tìm kiếm mã ca, xóa ca linh hoạt, xóa toàn bộ bảng tính an toàn.
   - Chỉnh sửa trực tiếp trên từng ô dữ liệu.

4. **An toàn dữ liệu & Auto-Save đa tầng:**
   - **Tự động lưu sau 5 phút (Auto-Save 5m):** Định kỳ sao lưu dữ liệu và tự động tạo Snapshot dự phòng.
   - **Đồng bộ file Excel trực tiếp trên máy:** Tự động ghi ra file `ai_evaluation_results.xlsx` (2 sheet chuẩn có màu sắc, định dạng và biểu đồ tóm tắt).
   - Tải về tức thì dưới định dạng Excel (.xlsx), CSV (UTF-8 BOM Tiếng Việt) và ZIP chứa toàn bộ JSON phản hồi AI thô.

---

## 🚀 Hướng Dẫn Deploy Lên Streamlit Community Cloud (Miễn Phí 100%)

Nền tảng **Streamlit Community Cloud** (`share.streamlit.io`) là nơi tối ưu nhất để đưa ứng dụng này lên mạng Internet hoạt động 24/7.

### Bước 1: Đẩy mã nguồn lên GitHub
1. Tạo một repository mới trên [GitHub](https://github.com/new) (ví dụ đặt tên: `dental-ai-evaluator`), chọn chế độ **Public** hoặc **Private**.
2. Đưa các file sau vào repository:
   - `app.py`
   - `data_manager.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `.gitignore`
   - `ai_evaluation_results.csv`
   - `session_progress.json`
   - `README.md`

### Bước 2: Triển khai trên Streamlit Cloud
1. Truy cập [share.streamlit.io](https://share.streamlit.io/) và đăng nhập bằng tài khoản GitHub.
2. Nhấn nút **"New app"**.
3. Điền thông tin:
   - **Repository:** Chọn repository vừa tạo (VD: `your-username/dental-ai-evaluator`).
   - **Branch:** `main` (hoặc `master`).
   - **Main file path:** `app.py`.
4. Nhấn **"Deploy!"**.
5. Hệ thống sẽ tự động cài đặt các thư viện trong `requirements.txt` và khởi chạy web trong 1–2 phút. Bạn sẽ nhận được đường link công khai (dạng `https://ten-ung-dung.streamlit.app`).

---

## 💻 Hướng Dẫn Chạy Cục Bộ (Localhost)

1. **Cài đặt thư viện:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Khởi chạy ứng dụng:**
   ```bash
   streamlit run app.py
   ```
   Hoặc trên Windows, bạn chỉ cần click đúp vào file `run_web.bat`.

---

## 📁 Cấu Trúc Thư Mục Chuẩn

```
Dental_AI_Tool/
├── .streamlit/
│   └── config.toml             # Cấu hình giao diện Dark theme & máy chủ
├── ai_responses/               # Thư mục lưu các file JSON phản hồi AI thô
├── backups/                    # Thư mục lưu các bản sao lưu Snapshot tự động
├── ai_evaluation_results.csv   # Cơ sở dữ liệu CSV chính (16 cột chuẩn)
├── ai_evaluation_results.xlsx  # File Excel tự động đồng bộ
├── app.py                      # Giao diện chính Streamlit Web
├── data_manager.py             # Module quản lý dữ liệu, parser & xuất Excel
├── requirements.txt            # Danh sách thư viện Python cần thiết
├── session_progress.json       # Tiến độ làm việc của người dùng
├── .gitignore                  # Cấu hình bỏ qua file rác khi đẩy lên Git
└── README.md                   # Tài liệu hướng dẫn sử dụng & triển khai
```
