"""
========================================================
  Dental AI Tool – Data Manager Module
  Quản lý dữ liệu, sao lưu tự động, xuất Excel & JSON
========================================================
"""

import os
import csv
import json
import re
import shutil
import zipfile
import io
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Cấu hình đường dẫn ──────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV = os.path.join(BASE_DIR, "ai_evaluation_results.csv")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "ai_evaluation_results.xlsx")
PROGRESS_JSON = os.path.join(BASE_DIR, "session_progress.json")
RESPONSES_DIR = os.path.join(BASE_DIR, "ai_responses")
BACKUPS_DIR = os.path.join(BASE_DIR, "backups")

# 16 cột chuẩn của nghiên cứu
CSV_HEADERS = [
    "Case_ID", "AI_Model", "Prompt_Type", "Time_Seconds",
    "Pell_Gregory_Class", "Pell_Gregory_Position", "Pell_Gregory_Full",
    "Winter_Class", "Confidence", "Pederson_Level",
    "Hallucination_Flag", "Reasoning_Quality", "Operator",
    "Timestamp", "Session_Notes", "AI_Raw_Response"
]

# 4 người nhập số liệu chính thức
OPERATORS_LIST = [
    "GIANG",
    "HOÀNG",
    "MINH",
    "HÂN"
]

# 3 mô hình AI đánh giá
MODELS_LIST = [
    "GEMINI",
    "CLAUDE",
    "CHATGPT"
]

# Tổng số ca nghiên cứu (129 ca: Case_001 -> Case_129)
TOTAL_CASES = 129
ALL_CASES = [f"Case_{str(i).zfill(3)}" for i in range(1, TOTAL_CASES + 1)]

PG_CLASSES = ["I", "II", "III", "Không xác định"]
PG_POSITIONS = ["A", "B", "C", "Không xác định"]

WINTER_VALUES = [
    "Mesioangular", "Vertical", "Horizontal",
    "Distoangular", "Buccolingual", "Others"
]

PEDERSON_LEVELS = ["Nhẹ (3–4)", "Trung bình (5–6)", "Khó (7–10)"]

REASONING_LEVELS = [
    "0 – Không có lý luận",
    "1 – Lý luận mơ hồ",
    "2 – Có lý luận cơ sở",
    "3 – Lý luận chuẩn xác"
]

HALLUCINATION_LEVELS = ["Không", "Có – Nhẹ", "Có – Rõ ràng"]

PROMPTS = {
    "P1 – Basic": (
        "Please classify this impacted mandibular third molar based on the panoramic "
        "X-ray using Pell & Gregory and Winter classification systems."
    ),
    "P2 – Structured Medical": (
        "You are an oral and maxillofacial surgery specialist. Analyze the provided panoramic "
        "radiograph step by step:\n\n"
        "Step 1 – Identify anatomical landmarks: Locate the anterior border of the ramus, "
        "the occlusal plane of the second molar, and the alveolar crest.\n\n"
        "Step 2 – Pell & Gregory Classification:\n"
        "  Class I: Crown fully anterior to ramus\n"
        "  Class II: Crown partially under ramus\n"
        "  Class III: Crown fully within ramus\n"
        "  Position A: Crown at/above occlusal plane of M2\n"
        "  Position B: Crown between occlusal plane and cervical line of M2\n"
        "  Position C: Crown below cervical line of M2\n\n"
        "Step 3 – Winter Classification: Determine angulation "
        "(Mesioangular / Vertical / Horizontal / Distoangular / Buccolingual / Others)\n\n"
        "Step 4 – Confidence: State your confidence level (0–100%).\n\n"
        'Output format (JSON):\n{\n  "P&G_class": "",\n  "P&G_position": "",\n'
        '  "Winter": "",\n  "confidence": 0,\n  "reasoning": ""\n}'
    ),
    "P3 – Feedback-assisted": (
        "You are an oral and maxillofacial surgery specialist. "
        "[Apply all steps from the Structured Medical Prompt above]\n\n"
        "COMMON ERRORS TO AVOID:\n"
        "  1. Do NOT confuse the retromolar pad with the ramus border\n"
        "  2. Verify tooth angulation against the mandibular plane, NOT the image border\n"
        "  3. If the crown is fully covered by bone → default Position C unless landmarks "
        "clearly indicate otherwise\n"
        "  4. Do not hallucinate anatomical structures not visible on the film\n\n"
        "FEW-SHOT EXAMPLE:\n"
        "  Image A → Class II, Position B, Mesioangular, Confidence 82%\n"
        "  Image B → Class I, Position A, Vertical, Confidence 90%\n\n"
        "Now classify the provided image following all guidelines above."
    )
}

PROMPT_TAGS = {
    "P1 – Basic": "P1: Prompt tối giản – mô phỏng người dùng thông thường",
    "P2 – Structured Medical": "P2: Cấu trúc y khoa đầy đủ + Chain-of-Thought + JSON output",
    "P3 – Feedback-assisted": "P3: Few-shot examples + hướng dẫn tránh lỗi thường gặp"
}


def ensure_directories():
    """Đảm bảo các thư mục lưu trữ cần thiết tồn tại."""
    os.makedirs(RESPONSES_DIR, exist_ok=True)
    os.makedirs(BACKUPS_DIR, exist_ok=True)
    if not os.path.exists(OUTPUT_CSV):
        with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerow(CSV_HEADERS)


def natural_sort_key(s: str):
    """
    Hàm tách số tự nhiên để sắp xếp đúng thứ tự (VD: Case_1, Case_2, Case_10).
    """
    match = re.search(r'\d+', str(s))
    return (int(match.group()), str(s)) if match else (999999, str(s))


def load_dataset(sort_by_case: bool = True) -> pd.DataFrame:
    """
    Đọc dữ liệu từ file CSV chính, chuẩn hóa cột theo CSV_HEADERS
    và sắp xếp theo thứ tự tự nhiên của Mã ca (Case_ID).
    """
    ensure_directories()
    if not os.path.exists(OUTPUT_CSV) or os.path.getsize(OUTPUT_CSV) == 0:
        return pd.DataFrame(columns=CSV_HEADERS)

    try:
        # Đọc linh hoạt bỏ qua số cột bất thường nếu có
        df = pd.read_csv(OUTPUT_CSV, dtype=str, on_bad_lines="skip")
        
        # Đảm bảo có đủ 16 cột
        for col in CSV_HEADERS:
            if col not in df.columns:
                df[col] = ""
        
        # Chỉ giữ đúng 16 cột chuẩn theo thứ tự
        df = df[CSV_HEADERS].fillna("")

        # Chuyển đổi kiểu dữ liệu cơ bản cho bảng tính
        df["Time_Seconds"] = pd.to_numeric(df["Time_Seconds"], errors="coerce").fillna(0).astype(int)
        df["Confidence"] = pd.to_numeric(df["Confidence"], errors="coerce").fillna(50).astype(int)

        # Sắp xếp theo thứ tự tự nhiên của Case_ID
        if sort_by_case and not df.empty and "Case_ID" in df.columns:
            df["_sort_key"] = df["Case_ID"].apply(lambda x: natural_sort_key(x)[0])
            df = df.sort_values(by="_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)

        return df
    except Exception as e:
        print(f"Lỗi khi đọc file CSV: {e}")
        return pd.DataFrame(columns=CSV_HEADERS)


def clear_dataset(create_backup: bool = True) -> bool:
    """
    Xóa toàn bộ dữ liệu bảng tính (giữ lại header chuẩn)
    và tự động tạo 1 bản snapshot dự phòng an toàn trước khi xóa.
    """
    ensure_directories()
    if create_backup and os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0:
        create_snapshot_backup(note="before_clear_all")

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        csv.writer(f).writerow(CSV_HEADERS)
    return True


def delete_cases(case_ids: list, create_backup: bool = True) -> bool:
    """
    Xóa một hoặc nhiều ca bệnh cụ thể theo Case_ID khỏi cơ sở dữ liệu.
    """
    ensure_directories()
    df = load_dataset(sort_by_case=False)
    if not df.empty and case_ids:
        filtered_df = df[~df["Case_ID"].isin(case_ids)]
        save_dataset(filtered_df, create_backup=create_backup)
        return True
    return False


def get_case_record(case_id: str) -> dict:
    """
    Lấy thông tin bản ghi đánh giá mới nhất của một Case_ID cụ thể.
    """
    df = load_dataset(sort_by_case=False)
    if not df.empty and "Case_ID" in df.columns:
        matches = df[df["Case_ID"] == case_id]
        if not matches.empty:
            return matches.iloc[-1].to_dict()
    return {}


def upsert_evaluation_record(record: dict) -> bool:
    """
    Thêm mới hoặc cập nhật bản ghi đánh giá ca bệnh nếu Case_ID đã tồn tại.
    """
    ensure_directories()
    case_id = record.get("Case_ID", "")
    df = load_dataset(sort_by_case=False)

    if not df.empty and case_id in df["Case_ID"].values:
        # Cập nhật bản ghi đã có
        idx = df[df["Case_ID"] == case_id].index[-1]
        for col in CSV_HEADERS:
            if col in record:
                df.at[idx, col] = record[col]
        save_dataset(df, create_backup=True)
    else:
        # Thêm mới
        append_evaluation_record(record)

    # Lưu thêm file JSON metadata
    safe_model = re.sub(r'[^a-zA-Z0-9_-]', '_', str(record.get("AI_Model", "model")))
    safe_prompt = re.sub(r'[^a-zA-Z0-9_-]', '_', str(record.get("Prompt_Type", "prompt")))
    ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = f"{case_id}_{safe_model}_{safe_prompt}_{ts_str}.json"
    json_path = os.path.join(RESPONSES_DIR, json_filename)
    try:
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(record, jf, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Lỗi khi lưu JSON: {e}")

    return True


def sync_excel_file() -> bool:
    """
    Tự động ghi và đồng bộ trực tiếp ra file Excel (ai_evaluation_results.xlsx)
    trên máy tính mỗi khi có thao tác lưu/cập nhật dữ liệu.
    """
    try:
        df = load_dataset(sort_by_case=True)
        excel_bytes = export_to_excel_bytes(df)
        with open(OUTPUT_EXCEL, "wb") as f:
            f.write(excel_bytes)
        return True
    except Exception as e:
        print(f"Lỗi khi đồng bộ file Excel: {e}")
        return False


def save_dataset(df: pd.DataFrame, create_backup: bool = True):
    """
    Lưu DataFrame trở lại file CSV chính, tự động tạo snapshot backup
    và tự động đồng bộ file Excel trên máy.
    """
    ensure_directories()
    if create_backup and os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0:
        create_snapshot_backup(note="auto_before_save")

    # Chuẩn hóa thứ tự cột
    save_df = df.copy()
    for col in CSV_HEADERS:
        if col not in save_df.columns:
            save_df[col] = ""
    save_df = save_df[CSV_HEADERS]

    save_df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")
    sync_excel_file()


def append_evaluation_record(record: dict) -> bool:
    """
    Ghi thêm 1 bản ghi ca bệnh vào CSV, tạo file JSON metadata, tạo snapshot backup
    và tự động cập nhật file Excel trên máy.
    """
    ensure_directories()
    
    # 1. Tự động tạo snapshot backup
    create_snapshot_backup(note=f"case_{record.get('Case_ID', 'unknown')}")

    # 2. Ghi vào file CSV
    row = [record.get(col, "") for col in CSV_HEADERS]
    file_exists = os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0
    with open(OUTPUT_CSV, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(CSV_HEADERS)
        writer.writerow(row)

    # 3. Lưu file JSON riêng biệt chứa metadata và toàn văn phản hồi AI
    case_id = record.get("Case_ID", "Case_000")
    safe_model = re.sub(r'[^a-zA-Z0-9_-]', '_', str(record.get("AI_Model", "model")))
    safe_prompt = re.sub(r'[^a-zA-Z0-9_-]', '_', str(record.get("Prompt_Type", "prompt")))
    ts_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = f"{case_id}_{safe_model}_{safe_prompt}_{ts_str}.json"
    json_path = os.path.join(RESPONSES_DIR, json_filename)

    try:
        with open(json_path, "w", encoding="utf-8") as jf:
            json.dump(record, jf, ensure_ascii=False, indent=4)
    except Exception as e:
        print(f"Lỗi khi lưu JSON metadata: {e}")

    # 4. Tự động đồng bộ file Excel trên máy
    sync_excel_file()

    return True


def create_snapshot_backup(note: str = "") -> str:
    """
    Tạo bản sao lưu snapshot file CSV vào thư mục backups/
    """
    ensure_directories()
    if not os.path.exists(OUTPUT_CSV) or os.path.getsize(OUTPUT_CSV) == 0:
        return ""

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_note = re.sub(r'[^a-zA-Z0-9_-]', '_', note) if note else "manual"
    backup_filename = f"backup_{timestamp}_{safe_note}.csv"
    backup_path = os.path.join(BACKUPS_DIR, backup_filename)

    shutil.copy2(OUTPUT_CSV, backup_path)

    # Giới hạn giữ tối đa 30 bản sao lưu gần nhất
    clean_old_backups(max_keep=30)
    return backup_filename


def clean_old_backups(max_keep: int = 30):
    """Xóa các bản sao lưu cũ quá số lượng tối đa."""
    if not os.path.exists(BACKUPS_DIR):
        return
    backups = [
        os.path.join(BACKUPS_DIR, f) for f in os.listdir(BACKUPS_DIR)
        if f.startswith("backup_") and f.endswith(".csv")
    ]
    if len(backups) > max_keep:
        backups.sort(key=os.path.getmtime)
        for old_file in backups[:-max_keep]:
            try:
                os.remove(old_file)
            except OSError:
                pass


def list_backups() -> list:
    """
    Trả về danh sách thông tin các bản sao lưu có sẵn.
    """
    ensure_directories()
    if not os.path.exists(BACKUPS_DIR):
        return []

    files = [f for f in os.listdir(BACKUPS_DIR) if f.startswith("backup_") and f.endswith(".csv")]
    files.sort(key=lambda x: os.path.getmtime(os.path.join(BACKUPS_DIR, x)), reverse=True)

    backup_info = []
    for f in files:
        full_path = os.path.join(BACKUPS_DIR, f)
        mtime = datetime.fromtimestamp(os.path.getmtime(full_path)).strftime("%Y-%m-%d %H:%M:%S")
        size_kb = round(os.path.getsize(full_path) / 1024, 2)
        try:
            with open(full_path, "r", encoding="utf-8-sig", errors="ignore") as csv_f:
                row_count = sum(1 for _ in csv_f) - 1  # trừ header
                row_count = max(0, row_count)
        except Exception:
            row_count = "?"

        backup_info.append({
            "filename": f,
            "timestamp": mtime,
            "size_kb": size_kb,
            "row_count": row_count,
            "path": full_path
        })
    return backup_info


def restore_backup(backup_filename: str) -> bool:
    """
    Khôi phục dữ liệu từ một bản sao lưu snapshot.
    """
    ensure_directories()
    backup_path = os.path.join(BACKUPS_DIR, backup_filename)
    if not os.path.exists(backup_path):
        return False

    # Sao lưu trạng thái hiện tại trước khi khôi phục
    if os.path.exists(OUTPUT_CSV) and os.path.getsize(OUTPUT_CSV) > 0:
        create_snapshot_backup(note="pre_restore_safety")

    shutil.copy2(backup_path, OUTPUT_CSV)
    return True


def export_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Xuất dữ liệu thành file Excel (.xlsx) chuẩn nghiên cứu y khoa với 2 sheet:
    - Sheet 1: Dữ liệu Đánh giá (đầy đủ các dòng và format tiêu đề)
    - Sheet 2: Báo cáo Thống kê (Tổng quan chỉ số nghiên cứu)
    """
    wb = openpyxl.Workbook()
    # Sheet 1: Data
    ws_data = wb.active
    ws_data.title = "Dữ liệu Đánh giá"

    # Style tiêu đề
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    # Ghi header
    cols = CSV_HEADERS
    for col_idx, col_name in enumerate(cols, start=1):
        cell = ws_data.cell(row=1, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = align_center

    # Ghi data
    for row_idx, (_, row) in enumerate(df.iterrows()):
        for col_idx, col_name in enumerate(cols, start=1):
            val = row.get(col_name, "")
            cell = ws_data.cell(row=row_idx + 2, column=col_idx, value=val)
            cell.border = border_thin
            if col_name in ["Time_Seconds", "Confidence"]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left

    # Tự động điều chỉnh độ rộng cột
    for col in ws_data.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > 50:
                val_str = val_str[:50]
            max_len = max(max_len, len(val_str))
        ws_data.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # Sheet 2: Summary Report
    ws_summary = wb.create_sheet(title="Thống kê Nghiên cứu")
    ws_summary.column_dimensions['A'].width = 32
    ws_summary.column_dimensions['B'].width = 22

    summary_title_fill = PatternFill(start_color="203764", end_color="203764", fill_type="solid")
    title_cell = ws_summary.cell(row=1, column=1, value="BÁO CÁO TỔNG QUAN ĐÁNH GIÁ AI")
    title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    title_cell.fill = summary_title_fill
    ws_summary.merge_cells('A1:B1')

    # Tính toán chỉ số an toàn
    avg_time = "—"
    avg_conf = "—"
    if not df.empty:
        try:
            t_series = pd.to_numeric(df["Time_Seconds"], errors="coerce")
            avg_time = f"{round(t_series.mean(), 1)} giây"
        except Exception:
            pass
        try:
            c_series = pd.to_numeric(df["Confidence"], errors="coerce")
            avg_conf = f"{round(c_series.mean(), 1)}%"
        except Exception:
            pass

    metrics = [
        ("Thời gian xuất báo cáo", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("Tổng số ca đã đánh giá", len(df)),
        ("Thời gian trung bình mỗi ca", avg_time),
        ("Độ tin cậy AI trung bình", avg_conf),
    ]

    r = 3
    for label, val in metrics:
        ws_summary.cell(row=r, column=1, value=label).font = Font(bold=True)
        ws_summary.cell(row=r, column=2, value=val)
        r += 1

    r += 1
    ws_summary.cell(row=r, column=1, value="PHÂN BỔ THEO MÔ HÌNH AI").font = Font(bold=True, color="1F4E79")
    r += 1
    if not df.empty and "AI_Model" in df.columns:
        model_counts = df["AI_Model"].value_counts()
        for m_name, count in model_counts.items():
            ws_summary.cell(row=r, column=1, value=str(m_name))
            ws_summary.cell(row=r, column=2, value=int(count))
            r += 1

    r += 1
    ws_summary.cell(row=r, column=1, value="PHÂN BỔ PELL & GREGORY").font = Font(bold=True, color="1F4E79")
    r += 1
    if not df.empty and "Pell_Gregory_Full" in df.columns:
        pg_counts = df["Pell_Gregory_Full"].value_counts()
        for pg_name, count in pg_counts.items():
            ws_summary.cell(row=r, column=1, value=str(pg_name))
            ws_summary.cell(row=r, column=2, value=int(count))
            r += 1

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def create_ai_responses_zip() -> bytes:
    """
    Nén toàn bộ các file JSON trong thư mục ai_responses thành file ZIP bytes.
    """
    ensure_directories()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        if os.path.exists(RESPONSES_DIR):
            for root, _, files in os.walk(RESPONSES_DIR):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, RESPONSES_DIR)
                    zf.write(file_path, arcname=arcname)
    buf.seek(0)
    return buf.getvalue()


def load_session_progress() -> dict:
    """
    Tải tiến độ phiên làm việc từ session_progress.json.
    """
    if os.path.exists(PROGRESS_JSON):
        try:
            with open(PROGRESS_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass

    # Mặc định danh sách 129 ca (Case_001 -> Case_129)
    cases = list(ALL_CASES)
    return {
        "cases_remaining": cases,
        "cases_done": [],
        "operator": OPERATORS_LIST[0],
        "model": MODELS_LIST[0],
        "prompt": "P2 – Structured Medical"
    }


def save_session_progress(progress_data: dict):
    """
    Lưu tiến độ phiên làm việc vào session_progress.json.
    """
    try:
        with open(PROGRESS_JSON, "w", encoding="utf-8") as f:
            json.dump(progress_data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Lỗi khi lưu tiến độ phiên: {e}")


def parse_ai_response(raw_text: str) -> dict:
    """
    Tự động phân tích và trích xuất các trường phân loại y khoa từ phản hồi AI
    (hỗ trợ cả định dạng JSON lẫn văn bản tự nhiên).
    """
    if not raw_text or len(raw_text.strip()) < 5:
        return {}

    text = raw_text.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
    result = {}

    # ── 1. Thử bóc tách theo khối JSON nếu có ───────────
    json_match = re.search(r'\{[^{}]*"P&G[^{}]*\}', text, re.DOTALL | re.IGNORECASE)
    if not json_match:
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
    
    if json_match:
        try:
            data = json.loads(json_match.group(0))
            for k, v in data.items():
                k_lower = str(k).lower().replace("_", "").replace(" ", "")
                v_str = str(v).strip()
                if "class" in k_lower and "position" not in k_lower and "winter" not in k_lower:
                    c_m = re.search(r'\b(III|II|I|3|2|1)\b', v_str, re.I)
                    if c_m:
                        val = c_m.group(1).upper()
                        result["Pell_Gregory_Class"] = "I" if val == "1" else ("II" if val == "2" else ("III" if val == "3" else val))
                elif "position" in k_lower:
                    p_m = re.search(r'\b([ABC])\b', v_str, re.I)
                    if p_m:
                        result["Pell_Gregory_Position"] = p_m.group(1).upper()
                elif "winter" in k_lower:
                    w_m = re.search(r'\b(mesioangular|horizontal|vertical|distoangular|buccolingual|others?)\b', v_str, re.I)
                    if w_m:
                        w_cap = w_m.group(1).capitalize()
                        result["Winter_Class"] = "Others" if w_cap.startswith("Other") else w_cap
                elif "confidence" in k_lower:
                    try:
                        num = int(float(v))
                        if 0 <= num <= 100:
                            result["Confidence"] = num
                    except (ValueError, TypeError):
                        pass
        except Exception:
            pass

    # ── 2. Trích xuất bằng Regex cho văn bản tự nhiên / bổ sung ──
    # Winter's Classification
    if "Winter_Class" not in result:
        w_match = re.search(r'\b(mesioangular|horizontal|vertical|distoangular|buccolingual|others?)\b', text, re.I)
        if w_match:
            w_val = w_match.group(1).capitalize()
            if w_val.startswith("Other"):
                w_val = "Others"
            result["Winter_Class"] = w_val

    # Pell & Gregory Class
    if "Pell_Gregory_Class" not in result:
        c_match = re.search(r'(?:class|p&g_class|p&g class)[\"\'\s:]*(?:class\s*)?(iii|ii|i|3|2|1)\b', text, re.I)
        if c_match:
            c_val = c_match.group(1).upper()
            if c_val == "1": c_val = "I"
            elif c_val == "2": c_val = "II"
            elif c_val == "3": c_val = "III"
            result["Pell_Gregory_Class"] = c_val

    # Pell & Gregory Position
    if "Pell_Gregory_Position" not in result:
        p_match = re.search(r'(?:position|p&g_position|p&g position)[\"\'\s:]*(?:position\s*)?([abc])\b', text, re.I)
        if p_match:
            result["Pell_Gregory_Position"] = p_match.group(1).upper()

    # Dạng kết hợp P&G (VD: II-B hoặc Class II, Position B)
    if "Pell_Gregory_Class" not in result or "Pell_Gregory_Position" not in result:
        combo = re.search(r'\b(I|II|III)\s*[-–/]\s*([A-C])\b', text, re.I)
        if combo:
            result["Pell_Gregory_Class"] = combo.group(1).upper()
            result["Pell_Gregory_Position"] = combo.group(2).upper()

    # Confidence (Độ tin cậy)
    if "Confidence" not in result:
        # Ưu tiên tìm kèm từ khóa confidence/%
        conf_match = re.search(r'(?:confidence|độ tin cậy)[^\n\r]*?(\d{1,3})\s*%', text, re.I)
        if not conf_match:
            conf_match = re.search(r'[\"\']?confidence[\"\']?\s*[:=]\s*(\d{1,3})', text, re.I)
        if not conf_match:
            conf_match = re.search(r'\(?\b([1-9]\d?|100)\s*%\)?', text)

        if conf_match:
            val = int(conf_match.group(1))
            if 0 <= val <= 100:
                result["Confidence"] = val

    # Pederson Level
    if "Pederson_Level" not in result:
        if re.search(r'\b(nhẹ|mild|3[-–]4)\b', text, re.I):
            result["Pederson_Level"] = "Nhẹ (3–4)"
        elif re.search(r'\b(trung bình|moderate|5[-–]6)\b', text, re.I):
            result["Pederson_Level"] = "Trung bình (5–6)"
        elif re.search(r'\b(khó|difficult|high|7[-–]10)\b', text, re.I):
            result["Pederson_Level"] = "Khó (7–10)"

    return result
