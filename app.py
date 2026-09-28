"""
========================================================
  🦷 Dental AI Evaluation Tool – MLLM Research (Web Streamlit)
  Phiên bản Web hoàn thiện với:
  - Tùy chọn chuyển đổi ca bệnh đang làm linh hoạt
  - Sắp xếp thứ tự ca chuẩn tự nhiên (Case_001 -> Case_129)
  - Bảng tính tương tác Google Sheets (Sửa, Thêm, Xóa ca, Xóa bảng tính)
  - Đồng hồ bấm giờ thủ công, Sao lưu snapshot & Xuất Excel
========================================================
"""

import os
import time
import json
import re
import importlib
from datetime import datetime
import pandas as pd
import streamlit as st

import data_manager as dm

# Reload module data_manager để đảm bảo các hàm mới luôn có hiệu lực
try:
    importlib.reload(dm)
except Exception:
    pass


# ── Hàm Bóc Tách Phản Hồi AI Trực Tiếp (Auto-Parser) ─────
def parse_ai_response(raw_text: str) -> dict:
    """
    Tự động phân tích và trích xuất các trường phân loại y khoa từ phản hồi AI.
    Hỗ trợ cả định dạng JSON lẫn văn bản báo cáo tự nhiên.
    """
    if hasattr(dm, "parse_ai_response"):
        try:
            res = dm.parse_ai_response(raw_text)
            if res:
                return res
        except Exception:
            pass

    if not raw_text or len(raw_text.strip()) < 5:
        return {}

    text = raw_text.replace('“', '"').replace('”', '"').replace('‘', "'").replace('’', "'")
    result = {}

    # 1. Thử bóc tách JSON
    json_matches = re.findall(r'\{[^{}]*\}', text, re.DOTALL)
    for j_str in json_matches:
        try:
            data = json.loads(j_str)
            for k, v in data.items():
                k_lower = str(k).lower().replace("_", "").replace(" ", "").replace("-", "")
                v_str = str(v).strip()
                if "class" in k_lower and "position" not in k_lower and "winter" not in k_lower:
                    c_m = re.search(r'\b(III|II|I|3|2|1)\b', v_str, re.I)
                    if c_m:
                        val = c_m.group(1).upper()
                        result["Pell_Gregory_Class"] = "I" if val == "1" else ("II" if val == "2" else ("III" if val == "3" else val))
                elif "position" in k_lower or "pos" in k_lower:
                    p_m = re.search(r'\b([ABC])\b', v_str, re.I)
                    if p_m:
                        result["Pell_Gregory_Position"] = p_m.group(1).upper()
                elif "winter" in k_lower or "angulation" in k_lower:
                    w_m = re.search(r'\b(mesioangular|horizontal|vertical|distoangular|buccolingual|others?)\b', v_str, re.I)
                    if w_m:
                        w_cap = w_m.group(1).capitalize()
                        result["Winter_Class"] = "Others" if w_cap.startswith("Other") else w_cap
                elif "confidence" in k_lower or "conf" in k_lower:
                    try:
                        num = int(float(v_str.replace('%', '')))
                        if 0 <= num <= 100:
                            result["Confidence"] = num
                    except (ValueError, TypeError):
                        pass
        except Exception:
            pass

    # 2. Regex fallback / bổ sung
    if "Winter_Class" not in result:
        w_match = re.search(r'\b(mesioangular|horizontal|vertical|distoangular|buccolingual|others?)\b', text, re.I)
        if w_match:
            w_val = w_match.group(1).capitalize()
            if w_val.startswith("Other"):
                w_val = "Others"
            result["Winter_Class"] = w_val

    if "Pell_Gregory_Class" not in result:
        c_match = re.search(r'(?:class|p&g_class|p&g class)[\"\'\s:]*(?:class\s*)?(iii|ii|i|3|2|1)\b', text, re.I)
        if c_match:
            c_val = c_match.group(1).upper()
            if c_val == "1": c_val = "I"
            elif c_val == "2": c_val = "II"
            elif c_val == "3": c_val = "III"
            result["Pell_Gregory_Class"] = c_val

    if "Pell_Gregory_Position" not in result:
        p_match = re.search(r'(?:position|p&g_position|p&g position)[\"\'\s:]*(?:position\s*)?([abc])\b', text, re.I)
        if p_match:
            result["Pell_Gregory_Position"] = p_match.group(1).upper()

    if "Pell_Gregory_Class" not in result or "Pell_Gregory_Position" not in result:
        combo = re.search(r'\b(I|II|III)\s*[-–/]\s*([A-C])\b', text, re.I)
        if combo:
            result["Pell_Gregory_Class"] = combo.group(1).upper()
            result["Pell_Gregory_Position"] = combo.group(2).upper()

    if "Confidence" not in result:
        conf_match = re.search(r'(?:confidence|độ tin cậy)[^\n\r]*?(\d{1,3})\s*%', text, re.I)
        if not conf_match:
            conf_match = re.search(r'[\"\']?confidence[\"\']?\s*[:=]\s*(\d{1,3})', text, re.I)
        if not conf_match:
            conf_match = re.search(r'\(?\b([1-9]\d?|100)\s*%\)?', text)

        if conf_match:
            val = int(conf_match.group(1))
            if 0 <= val <= 100:
                result["Confidence"] = val

    if "Pederson_Level" not in result:
        if re.search(r'\b(nhẹ|mild|3[-–]4)\b', text, re.I):
            result["Pederson_Level"] = "Nhẹ (3–4)"
        elif re.search(r'\b(trung bình|moderate|5[-–]6)\b', text, re.I):
            result["Pederson_Level"] = "Trung bình (5–6)"
        elif re.search(r'\b(khó|difficult|high|7[-–]10)\b', text, re.I):
            result["Pederson_Level"] = "Khó (7–10)"

    return result

# ── Cấu hình trang Streamlit ────────────────────────────
st.set_page_config(
    page_title="Dental AI Evaluator – MLLM Research",
    page_icon="🦷",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS Giao diện Chuyên nghiệp Y khoa ────────────
st.markdown("""
<style>
    /* Google Fonts & Base Styling */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    }

    /* Modern Card container */
    .card-box {
        background-color: #111827;
        border: 1px solid #1f2937;
        border-radius: 12px;
        padding: 20px;
        margin-bottom: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2);
    }
    .card-header {
        font-size: 1.05rem;
        font-weight: 700;
        color: #f1f5f9;
        margin-bottom: 14px;
        display: flex;
        align-items: center;
        gap: 8px;
        border-bottom: 1px solid #1f2937;
        padding-bottom: 10px;
        letter-spacing: 0.3px;
    }

    /* Section Subheadings */
    .section-title {
        color: #38bdf8;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        margin-top: 14px;
        margin-bottom: 10px;
        display: flex;
        align-items: center;
        gap: 6px;
    }

    /* Badges & Chips */
    .badge-chip {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: #1e293b;
        color: #94a3b8;
        border: 1px solid #334155;
        border-radius: 20px;
        padding: 4px 12px;
        font-size: 0.82rem;
        font-weight: 500;
    }
    .badge-done {
        background: rgba(16, 185, 129, 0.15);
        color: #34d399;
        border: 1px solid rgba(16, 185, 129, 0.4);
        border-radius: 20px;
        padding: 4px 12px;
        font-size: 0.8rem;
        font-weight: 700;
    }
    .badge-pending {
        background: rgba(148, 163, 184, 0.12);
        color: #94a3b8;
        border: 1px solid rgba(148, 163, 184, 0.3);
        border-radius: 20px;
        padding: 4px 12px;
        font-size: 0.8rem;
        font-weight: 600;
    }

    /* Digital Timer Clock Display */
    .timer-display {
        font-family: 'JetBrains Mono', 'Courier New', monospace;
        font-size: 2.2rem;
        font-weight: 700;
        padding: 8px 14px;
        border-radius: 10px;
        text-align: center;
        letter-spacing: 3px;
        background-color: #0b0f19;
        border: 1px solid #1e293b;
        color: #38bdf8;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.4);
    }
    .timer-warning {
        color: #f59e0b !important;
        border-color: rgba(245, 158, 11, 0.4) !important;
    }
    .timer-danger {
        color: #f43f5e !important;
        border-color: rgba(244, 63, 94, 0.4) !important;
    }

    /* Streamlit Components Polish */
    div[data-testid="stMetricValue"] {
        font-family: 'JetBrains Mono', monospace;
        font-weight: 700;
        color: #38bdf8;
    }

    /* Primary Buttons Glow */
    button[kind="primary"] {
        border-radius: 8px !important;
        font-weight: 600 !important;
        letter-spacing: 0.3px !important;
    }

    /* Code blocks */
    .stCodeBlock {
        border-radius: 8px !important;
        border: 1px solid #1f2937 !important;
    }
</style>
""", unsafe_allow_html=True)


# ── Khởi tạo Danh sách 129 Ca Chuẩn Sắp Xếp Tự Nhiên ─────
def get_all_ordered_cases(df: pd.DataFrame) -> list:
    """
    Tạo danh sách các ca từ Case_001 -> Case_129 và bổ sung bất kỳ mã ca nào
    có trong CSV, sắp xếp chuẩn số học tự nhiên.
    """
    base_cases = [f"Case_{str(i).zfill(3)}" for i in range(1, dm.TOTAL_CASES + 1)]
    if not df.empty and "Case_ID" in df.columns:
        existing_cases = df["Case_ID"].dropna().astype(str).tolist()
        all_unique = list(set(base_cases + existing_cases))
    else:
        all_unique = base_cases
    all_unique.sort(key=dm.natural_sort_key)
    return all_unique


# ── Khởi tạo Session State ──────────────────────────────
def init_state():
    if "session_progress" not in st.session_state:
        st.session_state.session_progress = dm.load_session_progress()
    
    prog = st.session_state.session_progress
    if "operator" not in st.session_state:
        st.session_state.operator = prog.get("operator", dm.OPERATORS_LIST[0])
    # Tự động chuẩn hóa nếu operator cũ không nằm trong danh sách 4 người
    if not st.session_state.operator:
        st.session_state.operator = dm.OPERATORS_LIST[0]
    elif st.session_state.operator not in dm.OPERATORS_LIST and "Giang" in str(st.session_state.operator):
        st.session_state.operator = "GIANG"

    if "ai_model" not in st.session_state:
        st.session_state.ai_model = prog.get("model", dm.MODELS_LIST[0])
    # Tự động chuẩn hóa mô hình cũ về 3 mô hình chuẩn: GEMINI, CLAUDE, CHATGPT
    if st.session_state.ai_model not in dm.MODELS_LIST:
        m_up = str(st.session_state.ai_model).upper()
        if "GEMINI" in m_up:
            st.session_state.ai_model = "GEMINI"
        elif "CLAUDE" in m_up:
            st.session_state.ai_model = "CLAUDE"
        elif "CHATGPT" in m_up or "GPT" in m_up:
            st.session_state.ai_model = "CHATGPT"
        else:
            st.session_state.ai_model = dm.MODELS_LIST[0]

    if "prompt_type" not in st.session_state:
        st.session_state.prompt_type = prog.get("prompt", "P2 – Structured Medical")
        
    df_curr = dm.load_dataset(sort_by_case=True)
    all_cases = get_all_ordered_cases(df_curr)
    
    if "current_case" not in st.session_state:
        remaining = prog.get("cases_remaining", [])
        if remaining:
            st.session_state.current_case = remaining[0]
        else:
            st.session_state.current_case = all_cases[0] if all_cases else "Case_001"
        
    # Stopwatch state
    if "timer_running" not in st.session_state:
        st.session_state.timer_running = False
    if "timer_start_time" not in st.session_state:
        st.session_state.timer_start_time = 0.0
    if "timer_elapsed_sec" not in st.session_state:
        st.session_state.timer_elapsed_sec = 0

    if "reload_data_trigger" not in st.session_state:
        st.session_state.reload_data_trigger = 0

    if "autosave_enabled" not in st.session_state:
        st.session_state.autosave_enabled = True
    if "last_autosave_timestamp" not in st.session_state:
        st.session_state.last_autosave_timestamp = time.time()

init_state()

# Đọc dữ liệu hiện tại
df_dataset = dm.load_dataset(sort_by_case=True)
all_cases_list = get_all_ordered_cases(df_dataset)


# ── Tính thời gian trôi qua hiện tại ────────────────────
def get_current_elapsed() -> int:
    if st.session_state.timer_running:
        return int(time.time() - st.session_state.timer_start_time)
    return st.session_state.timer_elapsed_sec


def format_time(seconds: int) -> str:
    m, s = divmod(seconds, 60)
    return f"{m:02d}:{s:02d}"


# ── Lấy danh sách các ca đã đánh giá trong CSV ──────────
evaluated_cases_dict = {}
if not df_dataset.empty and "Case_ID" in df_dataset.columns:
    for _, r in df_dataset.iterrows():
        c_id = str(r["Case_ID"])
        evaluated_cases_dict[c_id] = {
            "Pell_Gregory_Full": r.get("Pell_Gregory_Full", ""),
            "AI_Model": r.get("AI_Model", ""),
            "Winter_Class": r.get("Winter_Class", ""),
            "Timestamp": r.get("Timestamp", "")
        }


# ========================================================
# SIDEBAR: THÔNG TIN PHIÊN & TÙY CHỌN CA ĐANG LÀM
# ========================================================
with st.sidebar:
    st.markdown("### 🦷 Dental AI Evaluator")
    st.caption("Nghiên cứu MLLMs phân loại răng khôn hàm dưới v2.0")
    st.divider()

    # 1. Tùy chọn Ca bệnh đang làm
    st.markdown("#### 🎯 Ca bệnh đang làm")
    
    # Tạo nhãn đẹp hiển thị trạng thái từng ca
    def get_case_display_label(c_id: str) -> str:
        if c_id in evaluated_cases_dict:
            pg = evaluated_cases_dict[c_id]["Pell_Gregory_Full"]
            return f"✅ {c_id} ({pg if pg else 'Đã làm'})"
        return f"⏳ {c_id} (Chưa làm)"

    curr_case_idx = all_cases_list.index(st.session_state.current_case) if st.session_state.current_case in all_cases_list else 0

    chosen_case = st.selectbox(
        "Tùy chọn Ca đang làm:",
        options=all_cases_list,
        index=curr_case_idx,
        format_func=get_case_display_label,
        key="sb_case_selector"
    )

    if chosen_case != st.session_state.current_case:
        st.session_state.current_case = chosen_case
        st.session_state.timer_running = False
        st.session_state.timer_elapsed_sec = 0
        st.rerun()

    # Nút chuyển ca nhanh
    nav_c1, nav_c2, nav_c3 = st.columns(3)
    with nav_c1:
        if st.button("⬅ Trước", use_container_width=True, help="Chuyển về ca liền trước"):
            if curr_case_idx > 0:
                st.session_state.current_case = all_cases_list[curr_case_idx - 1]
                st.session_state.timer_running = False
                st.session_state.timer_elapsed_sec = 0
                st.rerun()
    with nav_c2:
        if st.button("Sau ➡", use_container_width=True, help="Chuyển sang ca liền sau"):
            if curr_case_idx < len(all_cases_list) - 1:
                st.session_state.current_case = all_cases_list[curr_case_idx + 1]
                st.session_state.timer_running = False
                st.session_state.timer_elapsed_sec = 0
                st.rerun()
    with nav_c3:
        if st.button("⏭ Chưa làm", use_container_width=True, help="Chuyển ngay tới ca chưa làm gần nhất"):
            pending = [c for c in all_cases_list if c not in evaluated_cases_dict]
            if pending:
                st.session_state.current_case = pending[0]
                st.session_state.timer_running = False
                st.session_state.timer_elapsed_sec = 0
                st.rerun()
            else:
                st.toast("🎉 Bạn đã hoàn thành tất cả các ca!")

    st.divider()

    # 2. Thông tin người nhập số liệu & Mô hình
    st.markdown("#### 👤 Người nhập số liệu")
    op_options = dm.OPERATORS_LIST + ["Khác (Tùy chỉnh)..."]
    current_op = st.session_state.operator
    if current_op in dm.OPERATORS_LIST:
        op_idx = dm.OPERATORS_LIST.index(current_op)
    elif current_op and current_op in op_options:
        op_idx = op_options.index(current_op)
    elif current_op:
        op_idx = len(op_options) - 1
    else:
        op_idx = 0

    selected_op = st.selectbox(
        "Chọn người nhập (4 thành viên):",
        options=op_options,
        index=op_idx,
        key="sb_operator_select"
    )
    if selected_op == "Khác (Tùy chỉnh)...":
        custom_op = st.text_input("Nhập tên người nhập:", value=current_op if current_op not in dm.OPERATORS_LIST else "", key="sb_custom_op_input")
        final_op = custom_op.strip() if custom_op.strip() else current_op
    else:
        final_op = selected_op

    if final_op != st.session_state.operator:
        st.session_state.operator = final_op
        st.session_state.session_progress["operator"] = final_op
        dm.save_session_progress(st.session_state.session_progress)

    st.markdown("#### 🤖 Mô hình AI đánh giá")
    model_options = dm.MODELS_LIST + ["Khác (Tùy chỉnh)..."]
    current_m = st.session_state.ai_model
    if current_m in dm.MODELS_LIST:
        current_m_idx = dm.MODELS_LIST.index(current_m)
    elif current_m and current_m in model_options:
        current_m_idx = model_options.index(current_m)
    elif current_m:
        current_m_idx = len(model_options) - 1
    else:
        current_m_idx = 0

    selected_model = st.selectbox(
        "Chọn mô hình AI (GEMINI / CLAUDE / CHATGPT):",
        options=model_options,
        index=current_m_idx,
        key="sb_model_select"
    )
    if selected_model == "Khác (Tùy chỉnh)...":
        custom_model = st.text_input("Nhập tên mô hình:", value=current_m if current_m not in dm.MODELS_LIST else "", key="sb_custom_model_input")
        final_model = custom_model.strip() if custom_model.strip() else current_m
    else:
        final_model = selected_model

    if final_model != st.session_state.ai_model:
        st.session_state.ai_model = final_model
        st.session_state.session_progress["model"] = final_model
        dm.save_session_progress(st.session_state.session_progress)

    st.divider()

    # 3. Tiến độ phiên làm việc
    total_count = len(all_cases_list)
    done_count = len([c for c in all_cases_list if c in evaluated_cases_dict])
    prog_ratio = done_count / total_count if total_count > 0 else 0.0

    st.markdown("#### 📊 Tiến độ nghiên cứu")
    st.progress(prog_ratio)
    st.write(f"**Đã đánh giá:** `{done_count} / {total_count}` ca ({int(prog_ratio * 100)}%)")

    if not df_dataset.empty and "Time_Seconds" in df_dataset.columns:
        valid_times = pd.to_numeric(df_dataset["Time_Seconds"], errors="coerce").dropna()
        if len(valid_times) > 0:
            avg_sec = int(valid_times.mean())
            st.metric("⏱ Thời gian TB mỗi ca", f"{avg_sec // 60:02d}:{avg_sec % 60:02d}")
        else:
            st.metric("⏱ Thời gian TB mỗi ca", "—")
    else:
        st.metric("⏱ Thời gian TB mỗi ca", "—")

    st.divider()
    st.markdown("#### 🛡️ An toàn dữ liệu & Auto-Save")
    st.session_state.autosave_enabled = st.toggle(
        "Tự động lưu sau mỗi 5 phút",
        value=st.session_state.autosave_enabled,
        help="Hệ thống tự động lưu ca bệnh, đồng bộ file Excel trên máy và tạo Snapshot an toàn mỗi 5 phút."
    )
    if st.session_state.autosave_enabled:
        st.caption("🟢 Auto-Save 5p & Excel Sync: **BẬT**")

    st.divider()
    
    # Nút Reset phiên làm việc
    with st.popover("🔄 Reset phiên làm việc mới", use_container_width=True):
        st.warning("Hành động này sẽ tạo lại danh sách 129 ca theo thứ tự từ Case_001 -> Case_129. Dữ liệu trong CSV đã lưu KHÔNG bị xóa.")
        if st.button("Xác nhận Reset Phiên", type="primary", use_container_width=True):
            cases_new = list(dm.ALL_CASES)
            st.session_state.session_progress = {
                "cases_remaining": cases_new,
                "cases_done": [],
                "operator": st.session_state.operator,
                "model": st.session_state.ai_model,
                "prompt": st.session_state.prompt_type
            }
            dm.save_session_progress(st.session_state.session_progress)
            st.session_state.current_case = cases_new[0]
            st.session_state.timer_running = False
            st.session_state.timer_elapsed_sec = 0
            st.success("Đã reset phiên làm việc thành công với 129 ca!")
            st.rerun()


# ── MAIN TABS ───────────────────────────────────────────
tabs = st.tabs([
    "📋 Đánh giá từng ca (Form & Timer)",
    "📊 Bảng tính Google Sheets (Chuyên biệt)",
    "💾 Sao lưu & Xuất dữ liệu"
])


# ========================================================
# TAB 1: ĐÁNH GIÁ TỪNG CA (FORM & TIMER VIEW)
# ========================================================
with tabs[0]:
    # Lấy dữ liệu ca hiện tại nếu đã từng làm
    saved_record = dm.get_case_record(st.session_state.current_case)
    is_already_done = bool(saved_record)

    # ── CƠ CHẾ TỰ ĐỘNG LƯU SAU 5 PHÚT (AUTO-SAVE & EXCEL SYNC) ──
    if st.session_state.get("autosave_enabled", True):
        now_ts = time.time()
        elapsed_autosave = now_ts - st.session_state.get("last_autosave_timestamp", now_ts)
        cur_elapsed_time = get_current_elapsed()

        # Kích hoạt khi quá 300s kể từ lần autosave trước hoặc thời gian ca hiện tại >= 300s
        should_autosave = False
        autosave_reason = ""
        cur_c = st.session_state.current_case

        if elapsed_autosave >= 300:
            should_autosave = True
            autosave_reason = "định kỳ 5 phút"
        elif cur_elapsed_time >= 300 and not st.session_state.get(f"autosaved_case_{cur_c}", False):
            should_autosave = True
            st.session_state[f"autosaved_case_{cur_c}"] = True
            autosave_reason = f"ca {cur_c} đạt 5 phút"

        if should_autosave:
            st.session_state.last_autosave_timestamp = now_ts
            # 1. Tạo snapshot backup an toàn
            dm.create_snapshot_backup(note=f"autosave_5m_{cur_c}")

            # 2. Nếu có dữ liệu phản hồi AI hoặc ca đã làm, tự động cập nhật bản ghi
            cur_resp = st.session_state.get(f"form_ai_response_{cur_c}", "").strip()
            if st.session_state.operator.strip() and (cur_resp or is_already_done):
                cls_v = st.session_state.get(f"form_pg_class_{cur_c}", dm.PG_CLASSES[0])
                pos_v = st.session_state.get(f"form_pg_pos_{cur_c}", dm.PG_POSITIONS[0])
                f_pg = "Không xác định" if (cls_v == "Không xác định" or pos_v == "Không xác định") else f"{cls_v}-{pos_v}"

                auto_rec = {
                    "Case_ID": cur_c,
                    "AI_Model": st.session_state.ai_model,
                    "Prompt_Type": st.session_state.prompt_type,
                    "Time_Seconds": cur_elapsed_time if cur_elapsed_time > 0 else int(saved_record.get("Time_Seconds", 0) if saved_record else 0),
                    "Pell_Gregory_Class": cls_v,
                    "Pell_Gregory_Position": pos_v,
                    "Pell_Gregory_Full": f_pg,
                    "Winter_Class": st.session_state.get(f"form_winter_{cur_c}", dm.WINTER_VALUES[0]),
                    "Confidence": st.session_state.get(f"form_conf_{cur_c}", 50),
                    "Pederson_Level": st.session_state.get(f"form_pederson_{cur_c}", dm.PEDERSON_LEVELS[0]),
                    "Hallucination_Flag": st.session_state.get(f"form_halluc_{cur_c}", dm.HALLUCINATION_LEVELS[0]),
                    "Reasoning_Quality": str(st.session_state.get(f"form_reason_{cur_c}", "2"))[0],
                    "Operator": st.session_state.operator,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Session_Notes": st.session_state.get(f"form_notes_{cur_c}", ""),
                    "AI_Raw_Response": cur_resp
                }
                dm.upsert_evaluation_record(auto_rec)
                st.toast(f"💾 [Auto-Save {autosave_reason}] Đã tự động lưu & đồng bộ Excel trên máy cho {cur_c}!", icon="⏱")
            else:
                dm.sync_excel_file()
                st.toast(f"🛡 [Auto-Save {autosave_reason}] Đã tạo snapshot dự phòng & đồng bộ file Excel an toàn!", icon="💾")

    # ── HEADER CỦA CA HIỆN TẠI & ĐỒNG HỒ BẤM GIỜ ─────────
    top_col1, top_col2 = st.columns([1.35, 1.65], gap="medium")

    with top_col1:
        with st.container(border=True):
            h_c1, h_c2 = st.columns([1.2, 1.4])
            with h_c1:
                st.markdown("<span style='color: #94a3b8; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;'>🎯 CA BỆNH ĐANG LÀM</span>", unsafe_allow_html=True)
            with h_c2:
                if is_already_done:
                    st.markdown(f"<span class='badge-done'>✅ ĐÃ LÀM: {saved_record.get('Pell_Gregory_Full', '')}</span>", unsafe_allow_html=True)
                else:
                    st.markdown("<span class='badge-pending'>⏳ CHƯA ĐÁNH GIÁ</span>", unsafe_allow_html=True)

            st.markdown(f"<div style='font-size: 2.2rem; font-weight: 800; color: #38bdf8; margin: 2px 0 8px 0; letter-spacing: 0.5px;'>{st.session_state.current_case}</div>", unsafe_allow_html=True)
            st.markdown(f"""
            <div style="display: flex; gap: 8px; flex-wrap: wrap;">
                <span class="badge-chip">🤖 Mô hình: <b style="color: #f1f5f9;">{st.session_state.ai_model}</b></span>
                <span class="badge-chip">👤 Người nhập: <b style="color: #f1f5f9;">{st.session_state.operator or 'Chưa đặt'}</b></span>
            </div>
            """, unsafe_allow_html=True)

    with top_col2:
        with st.container(border=True):
            st.markdown("<span style='color: #94a3b8; font-size: 0.8rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;'>⏱ ĐỒNG HỒ BẤM GIỜ THỦ CÔNG</span>", unsafe_allow_html=True)
            cur_elapsed = get_current_elapsed()
            time_str = format_time(cur_elapsed)

            timer_cls = "timer-display"
            if cur_elapsed > 300:
                timer_cls += " timer-danger"
            elif cur_elapsed > 180:
                timer_cls += " timer-warning"

            t_c1, t_c2 = st.columns([1.1, 2.3], gap="small")
            with t_c1:
                st.markdown(f"""<div class="{timer_cls}">{time_str}</div>""", unsafe_allow_html=True)
            with t_c2:
                btn_c1, btn_c2, btn_c3 = st.columns(3)
                with btn_c1:
                    if not st.session_state.timer_running:
                        if st.button("▶ Bắt đầu", type="primary", use_container_width=True, key="btn_timer_start"):
                            st.session_state.timer_running = True
                            st.session_state.timer_start_time = time.time() - st.session_state.timer_elapsed_sec
                            st.rerun()
                    else:
                        if st.button("⏸ Tạm dừng", type="secondary", use_container_width=True, key="btn_timer_pause"):
                            st.session_state.timer_elapsed_sec = int(time.time() - st.session_state.timer_start_time)
                            st.session_state.timer_running = False
                            st.rerun()

                with btn_c2:
                    if st.session_state.timer_running:
                        if st.button("⏹ Kết thúc", use_container_width=True, key="btn_timer_stop"):
                            st.session_state.timer_elapsed_sec = int(time.time() - st.session_state.timer_start_time)
                            st.session_state.timer_running = False
                            st.rerun()
                    else:
                        if st.session_state.timer_elapsed_sec > 0:
                            if st.button("▶ Tiếp tục", use_container_width=True, key="btn_timer_resume"):
                                st.session_state.timer_start_time = time.time() - st.session_state.timer_elapsed_sec
                                st.session_state.timer_running = True
                                st.rerun()
                        else:
                            st.button("⏹ Kết thúc", disabled=True, use_container_width=True, key="btn_timer_stop_dis")

                with btn_c3:
                    if st.button("🔄 Đặt lại", use_container_width=True, key="btn_timer_reset"):
                        st.session_state.timer_running = False
                        st.session_state.timer_elapsed_sec = 0
                        st.rerun()

    # Báo nhắc nếu ca đã được lưu trước đó
    if is_already_done:
        st.info(f"ℹ️ Ca **{st.session_state.current_case}** đã được đánh giá lúc `{saved_record.get('Timestamp', '')}`. Dữ liệu bên dưới đã được điền sẵn từ bản ghi trước, bạn có thể chỉnh sửa và nhấn **Lưu / Cập nhật**.")

    st.write("")

    # ── BODY: PROMPT (TRÁI) & KẾT QUẢ AI + PHÂN LOẠI (PHẢI) ──
    col_left, col_right = st.columns([1, 1.25], gap="large")

    # ── BƯỚC 1: PROMPT NGHIÊN CỨU ──────────────────────
    with col_left:
        st.markdown("""
        <div class="card-header">
            <span>📋 BƯỚC 1: PROMPT NGHIÊN CỨU</span>
        </div>
        """, unsafe_allow_html=True)

        selected_prompt = st.selectbox(
            "Chọn loại Prompt chuẩn:",
            options=list(dm.PROMPTS.keys()),
            index=list(dm.PROMPTS.keys()).index(st.session_state.prompt_type) if st.session_state.prompt_type in dm.PROMPTS else 1,
            key="prompt_select"
        )
        st.session_state.prompt_type = selected_prompt
        st.session_state.session_progress["prompt"] = selected_prompt

        st.info(f"💡 {dm.PROMPT_TAGS.get(selected_prompt, '')}")

        prompt_content = dm.PROMPTS[selected_prompt]
        st.code(prompt_content, language="markdown")
        st.caption("📋 Nhấp biểu tượng góc trên bên phải khung code để sao chép nhanh sang ChatGPT / Claude / Gemini.")

    # ── BƯỚC 2: KẾT QUẢ AI & PHÂN LOẠI ──────────────────
    with col_right:
        st.markdown("""
        <div class="card-header">
            <span>🤖 BƯỚC 2: KẾT QUẢ AI & PHÂN LOẠI Y KHOA</span>
        </div>
        """, unsafe_allow_html=True)

        # Khởi tạo các khóa session_state riêng cho ca hiện tại
        cur_c = st.session_state.current_case
        k_resp = f"form_ai_response_{cur_c}"
        k_pg_cls = f"form_pg_class_{cur_c}"
        k_pg_pos = f"form_pg_pos_{cur_c}"
        k_winter = f"form_winter_{cur_c}"
        k_pederson = f"form_pederson_{cur_c}"
        k_conf = f"form_conf_{cur_c}"
        k_halluc = f"form_halluc_{cur_c}"
        k_reason = f"form_reason_{cur_c}"
        k_notes = f"form_notes_{cur_c}"

        # Hàm tự động bóc tách khi nội dung thay đổi hoặc bấm nút
        def trigger_auto_parse(c_id: str):
            raw_text = st.session_state.get(f"form_ai_response_{c_id}", "")
            parsed = parse_ai_response(raw_text)
            if parsed:
                if "Pell_Gregory_Class" in parsed:
                    st.session_state[f"form_pg_class_{c_id}"] = parsed["Pell_Gregory_Class"]
                if "Pell_Gregory_Position" in parsed:
                    st.session_state[f"form_pg_pos_{c_id}"] = parsed["Pell_Gregory_Position"]
                if "Winter_Class" in parsed:
                    st.session_state[f"form_winter_{c_id}"] = parsed["Winter_Class"]
                if "Confidence" in parsed:
                    st.session_state[f"form_conf_{c_id}"] = parsed["Confidence"]
                if "Pederson_Level" in parsed:
                    st.session_state[f"form_pederson_{c_id}"] = parsed["Pederson_Level"]
                st.session_state[f"auto_parsed_alert_{c_id}"] = parsed
                return True
            return False

        # Tự động nhận diện khi nội dung ô dán có thay đổi mới
        last_parsed_key = f"last_auto_parsed_text_{cur_c}"
        current_ai_text = st.session_state.get(k_resp, "")
        if current_ai_text and current_ai_text.strip() != st.session_state.get(last_parsed_key, ""):
            st.session_state[last_parsed_key] = current_ai_text.strip()
            trigger_auto_parse(cur_c)

        # Khởi tạo giá trị mặc định cho từng widget nếu chưa tồn tại trong session_state
        if k_resp not in st.session_state:
            st.session_state[k_resp] = saved_record.get("AI_Raw_Response", "") if is_already_done else ""
        if k_pg_cls not in st.session_state:
            st.session_state[k_pg_cls] = saved_record.get("Pell_Gregory_Class", dm.PG_CLASSES[0]) if is_already_done else dm.PG_CLASSES[0]
            if st.session_state[k_pg_cls] not in dm.PG_CLASSES: st.session_state[k_pg_cls] = dm.PG_CLASSES[0]
        if k_pg_pos not in st.session_state:
            st.session_state[k_pg_pos] = saved_record.get("Pell_Gregory_Position", dm.PG_POSITIONS[0]) if is_already_done else dm.PG_POSITIONS[0]
            if st.session_state[k_pg_pos] not in dm.PG_POSITIONS: st.session_state[k_pg_pos] = dm.PG_POSITIONS[0]
        if k_winter not in st.session_state:
            st.session_state[k_winter] = saved_record.get("Winter_Class", dm.WINTER_VALUES[0]) if is_already_done else dm.WINTER_VALUES[0]
            if st.session_state[k_winter] not in dm.WINTER_VALUES: st.session_state[k_winter] = dm.WINTER_VALUES[0]
        if k_pederson not in st.session_state:
            st.session_state[k_pederson] = saved_record.get("Pederson_Level", dm.PEDERSON_LEVELS[0]) if is_already_done else dm.PEDERSON_LEVELS[0]
            if st.session_state[k_pederson] not in dm.PEDERSON_LEVELS: st.session_state[k_pederson] = dm.PEDERSON_LEVELS[0]
        if k_conf not in st.session_state:
            try:
                st.session_state[k_conf] = int(saved_record.get("Confidence", 50)) if is_already_done else 50
            except Exception:
                st.session_state[k_conf] = 50
        if k_halluc not in st.session_state:
            st.session_state[k_halluc] = saved_record.get("Hallucination_Flag", dm.HALLUCINATION_LEVELS[0]) if is_already_done else dm.HALLUCINATION_LEVELS[0]
            if st.session_state[k_halluc] not in dm.HALLUCINATION_LEVELS: st.session_state[k_halluc] = dm.HALLUCINATION_LEVELS[0]
        if k_reason not in st.session_state:
            st.session_state[k_reason] = dm.REASONING_LEVELS[2]
            if is_already_done:
                r_str = str(saved_record.get("Reasoning_Quality", "2"))
                for r_opt in dm.REASONING_LEVELS:
                    if r_opt.startswith(r_str[0]):
                        st.session_state[k_reason] = r_opt
                        break
        if k_notes not in st.session_state:
            st.session_state[k_notes] = saved_record.get("Session_Notes", "") if is_already_done else ""

        def on_ai_text_change():
            trigger_auto_parse(cur_c)

        # Ô nhập toàn bộ câu trả lời AI
        ai_response_text = st.text_area(
            "Dán toàn bộ phản hồi từ AI vào đây:",
            placeholder="Dán toàn văn câu trả lời từ GEMINI / CLAUDE / CHATGPT vào đây để hệ thống tự động nhận diện và điền kết quả...",
            height=150,
            key=k_resp,
            on_change=on_ai_text_change
        )

        # Thanh công cụ bóc tách tự động
        btn_parse_col1, btn_parse_col2 = st.columns([2, 1], gap="small")
        with btn_parse_col1:
            if st.button("⚡ TỰ ĐỘNG TRÍCH XUẤT & ĐIỀN KẾT QUẢ", type="primary", use_container_width=True, key=f"btn_parse_{cur_c}"):
                trigger_auto_parse(cur_c)
                st.rerun()
        with btn_parse_col2:
            if st.button("🗑 Xóa ô nhập", use_container_width=True, key=f"btn_clear_{cur_c}"):
                st.session_state[k_resp] = ""
                st.session_state[last_parsed_key] = ""
                st.session_state[f"auto_parsed_alert_{cur_c}"] = None
                st.rerun()

        # Hiển thị thông báo kết quả tự động trích xuất
        alert_info = st.session_state.get(f"auto_parsed_alert_{cur_c}")
        if alert_info:
            tags = []
            if "Pell_Gregory_Class" in alert_info: tags.append(f"Class: **{alert_info['Pell_Gregory_Class']}**")
            if "Pell_Gregory_Position" in alert_info: tags.append(f"Position: **{alert_info['Pell_Gregory_Position']}**")
            if "Winter_Class" in alert_info: tags.append(f"Winter: **{alert_info['Winter_Class']}**")
            if "Confidence" in alert_info: tags.append(f"Độ tin cậy: **{alert_info['Confidence']}%**")
            if "Pederson_Level" in alert_info: tags.append(f"Pederson: **{alert_info['Pederson_Level']}**")
            st.success(f"🎯 **Đã tự động trích xuất & cập nhật:** {' · '.join(tags)}")

        st.markdown('<div class="section-title">🦴 PHÂN LOẠI GIẢI PHẪU (PELL & GREGORY VÀ WINTER)</div>', unsafe_allow_html=True)

        # 1. Pell & Gregory: Tách riêng Class & Position
        pg_c1, pg_c2, pg_c3 = st.columns([1.2, 1.2, 1.6])
        with pg_c1:
            val_pg_class = st.selectbox(
                "P&G Class:",
                options=dm.PG_CLASSES,
                key=k_pg_cls
            )
        with pg_c2:
            val_pg_pos = st.selectbox(
                "P&G Position:",
                options=dm.PG_POSITIONS,
                key=k_pg_pos
            )
        with pg_c3:
            if val_pg_class == "Không xác định" or val_pg_pos == "Không xác định":
                full_pg = "Không xác định"
            else:
                full_pg = f"{val_pg_class}-{val_pg_pos}"
            st.text_input("P&G Kết hợp:", value=full_pg, disabled=True, key=f"form_pg_full_{cur_c}")

        # 2. Winter & Pederson
        w_c1, w_c2 = st.columns(2)
        with w_c1:
            val_winter = st.selectbox(
                "Winter's Class (Thế răng):",
                options=dm.WINTER_VALUES,
                key=k_winter
            )
        with w_c2:
            val_pederson = st.selectbox(
                "Pederson Level (Độ khó):",
                options=dm.PEDERSON_LEVELS,
                key=k_pederson
            )

        st.markdown('<div class="section-title">🎯 ĐÁNH GIÁ ĐỘ TIN CẬY & LẬP LUẬN</div>', unsafe_allow_html=True)

        # 3. Confidence Slider & Hallucination
        conf_c1, conf_c2 = st.columns([1.5, 1])
        with conf_c1:
            val_confidence = st.slider(
                "Độ tin cậy của AI (Confidence %):",
                min_value=0,
                max_value=100,
                step=1,
                key=k_conf
            )
        with conf_c2:
            val_halluc = st.selectbox(
                "Ảo giác (Hallucination):",
                options=dm.HALLUCINATION_LEVELS,
                key=k_halluc
            )

        # 4. Reasoning Quality
        val_reasoning = st.selectbox(
            "Chất lượng lập luận (Reasoning Quality):",
            options=dm.REASONING_LEVELS,
            key=k_reason
        )

        st.markdown('<div class="section-title">📝 GHI CHÚ CA BỆNH</div>', unsafe_allow_html=True)

        # 5. Ghi chú phiên này
        val_notes = st.text_input(
            "Ghi chú ca bệnh (tùy chọn):",
            placeholder="Ghi chú thêm về giải phẫu, trường hợp đặc biệt...",
            key=k_notes
        )

        st.write("")

        # ── NÚT LƯU DỮ LIỆU & CHUYỂN CA ─────────────────
        button_label = "💾 CẬP NHẬT CA & CHUYỂN TIẾP →" if is_already_done else "💾 LƯU DỮ LIỆU & CHUYỂN CA TIẾP THEO →"
        save_btn = st.button(
            button_label,
            type="primary",
            use_container_width=True,
            key="btn_save_and_next"
        )

        if save_btn:
            if not st.session_state.operator.strip():
                st.error("⚠️ Vui lòng nhập Tên / Mã thành viên ở thanh bên trái trước khi lưu!")
            else:
                final_time = get_current_elapsed()
                if is_already_done and final_time == 0 and "Time_Seconds" in saved_record:
                    final_time = int(saved_record.get("Time_Seconds", 0))

                record = {
                    "Case_ID": st.session_state.current_case,
                    "AI_Model": st.session_state.ai_model,
                    "Prompt_Type": st.session_state.prompt_type,
                    "Time_Seconds": final_time,
                    "Pell_Gregory_Class": val_pg_class,
                    "Pell_Gregory_Position": val_pg_pos,
                    "Pell_Gregory_Full": full_pg,
                    "Winter_Class": val_winter,
                    "Confidence": val_confidence,
                    "Pederson_Level": val_pederson,
                    "Hallucination_Flag": val_halluc,
                    "Reasoning_Quality": val_reasoning[0],
                    "Operator": st.session_state.operator,
                    "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Session_Notes": val_notes.strip(),
                    "AI_Raw_Response": ai_response_text.strip()
                }

                # Ghi / Cập nhật vào CSV và lưu JSON
                dm.upsert_evaluation_record(record)

                cur_case = st.session_state.current_case
                if cur_case in st.session_state.session_progress["cases_remaining"]:
                    st.session_state.session_progress["cases_remaining"].remove(cur_case)
                if cur_case not in st.session_state.session_progress["cases_done"]:
                    st.session_state.session_progress["cases_done"].append(cur_case)
                
                dm.save_session_progress(st.session_state.session_progress)

                # Tìm ca tiếp theo chưa làm hoặc ca kế tiếp theo thứ tự
                cur_idx = all_cases_list.index(cur_case) if cur_case in all_cases_list else 0
                next_case = None
                
                # Ưu tiên ca chưa làm gần nhất
                remaining_pending = [c for c in all_cases_list if c != cur_case and c not in evaluated_cases_dict]
                if remaining_pending:
                    next_case = remaining_pending[0]
                elif cur_idx < len(all_cases_list) - 1:
                    next_case = all_cases_list[cur_idx + 1]

                if next_case:
                    st.session_state.current_case = next_case
                    st.session_state.timer_running = False
                    st.session_state.timer_elapsed_sec = 0
                    st.success(f"✅ Đã lưu {cur_case}! Đã chuyển sang {next_case}.")
                else:
                    st.balloons()
                    st.success(f"🎉 Chúc mừng! Đã hoàn thành toàn bộ danh sách ca nghiên cứu!")

                st.session_state.reload_data_trigger += 1
                time.sleep(1)
                st.rerun()


# ========================================================
# TAB 2: BẢNG TÍNH GOOGLE SHEETS (CHUYÊN BIỆT)
# ========================================================
with tabs[1]:
    st.markdown("""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <div class="card-header">📊 BẢNG TÍNH DỮ LIỆU ĐÁNH GIÁ (GOOGLE SHEETS VIEW)</div>
        <div style="color: #64ffda; font-size: 0.9rem;">Sắp xếp theo thứ tự ca chuẩn · Chỉnh sửa trực tiếp · Thêm & Xóa ca linh hoạt</div>
    </div>
    """, unsafe_allow_html=True)

    df_sheet = dm.load_dataset(sort_by_case=True)

    # ── THANH CÔNG CỤ SẮP XẾP, TÌM KIẾM & BỘ LỌC ─────────
    tb_c1, tb_c2, tb_c3 = st.columns([1.5, 1, 1.5])
    with tb_c1:
        sort_mode = st.selectbox(
            "📌 Sắp xếp bảng tính theo:",
            options=[
                "Mã ca (Tăng dần Case_001 -> Case_129)",
                "Mã ca (Giảm dần Case_129 -> Case_001)",
                "Thời gian lưu (Mới nhất trước)",
                "Thời gian lưu (Cũ nhất trước)"
            ],
            index=0,
            key="sheet_sort_mode"
        )
    with tb_c2:
        search_case_sheet = st.text_input("🔍 Tìm Mã ca:", placeholder="VD: Case_005", key="sheet_search_case")
    with tb_c3:
        filter_model_sheet = st.multiselect(
            "🤖 Lọc theo Mô hình AI:",
            options=list(df_sheet["AI_Model"].unique()) if not df_sheet.empty else [],
            key="sheet_filter_model"
        )

    # Áp dụng sắp xếp
    sorted_df = df_sheet.copy()
    if not sorted_df.empty:
        if sort_mode == "Mã ca (Tăng dần Case_001 -> Case_129)":
            sorted_df["_sort"] = sorted_df["Case_ID"].apply(lambda x: dm.natural_sort_key(x)[0])
            sorted_df = sorted_df.sort_values(by="_sort").drop(columns=["_sort"]).reset_index(drop=True)
        elif sort_mode == "Mã ca (Giảm dần Case_129 -> Case_001)":
            sorted_df["_sort"] = sorted_df["Case_ID"].apply(lambda x: dm.natural_sort_key(x)[0])
            sorted_df = sorted_df.sort_values(by="_sort", ascending=False).drop(columns=["_sort"]).reset_index(drop=True)
        elif sort_mode == "Thời gian lưu (Mới nhất trước)" and "Timestamp" in sorted_df.columns:
            sorted_df = sorted_df.sort_values(by="Timestamp", ascending=False).reset_index(drop=True)
        elif sort_mode == "Thời gian lưu (Cũ nhất trước)" and "Timestamp" in sorted_df.columns:
            sorted_df = sorted_df.sort_values(by="Timestamp", ascending=True).reset_index(drop=True)

    # Áp dụng tìm kiếm & lọc
    if search_case_sheet.strip():
        sorted_df = sorted_df[sorted_df["Case_ID"].str.contains(search_case_sheet.strip(), case=False, na=False)]
    if filter_model_sheet:
        sorted_df = sorted_df[sorted_df["AI_Model"].isin(filter_model_sheet)]

    # ── THANH HÀNH ĐỘNG BẢNG TÍNH (LƯU, THÊM, XÓA) ───────
    act_col1, act_col2, act_col3, act_col4 = st.columns([1.5, 1, 1.2, 1.2], gap="small")

    with act_col1:
        save_sheet_changes = st.button("💾 Lưu tất cả thay đổi trên Bảng tính", type="primary", use_container_width=True)

    with act_col2:
        if st.button("🔄 Tải lại dữ liệu", use_container_width=True):
            st.rerun()

    with act_col3:
        # Xóa ca được chọn
        with st.popover("🗑 Xóa ca được chọn...", use_container_width=True):
            st.markdown("##### Chọn ca muốn xóa khỏi bảng:")
            if not sorted_df.empty:
                cases_to_del = st.multiselect("Chọn các Case_ID:", options=list(sorted_df["Case_ID"].unique()))
                if st.button("Xác nhận xóa ca đã chọn", type="primary", use_container_width=True):
                    if cases_to_del:
                        dm.delete_cases(cases_to_del, create_backup=True)
                        st.success(f"Đã xóa {len(cases_to_del)} ca thành công!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.warning("Vui lòng chọn ít nhất 1 ca để xóa.")
            else:
                st.info("Bảng tính hiện không có dữ liệu.")

    with act_col4:
        # XÓA TOÀN BỘ BẢNG TÍNH (CÓ XÁC NHẬN BẢO VỆ)
        with st.popover("⚠️ Xóa toàn bộ bảng tính", use_container_width=True):
            st.error("Hành động này sẽ XÓA TOÀN BỘ dữ liệu đánh giá hiện tại để bạn làm lại từ đầu.")
            st.info("💡 An tâm: Hệ thống sẽ tự động tạo 1 bản Snapshot sao lưu trước khi xóa. Bạn có thể khôi phục lại bất kỳ lúc nào tại Tab 'Sao lưu & Xuất dữ liệu'.")
            confirm_clear = st.checkbox("Tôi hiểu và chắc chắn muốn xóa toàn bộ bảng tính", key="confirm_clear_chk")
            if st.button("Xác nhận Xóa Bảng Tính Ngay", type="primary", disabled=not confirm_clear, use_container_width=True):
                dm.clear_dataset(create_backup=True)
                # Reset tiến độ
                cases_new = list(dm.ALL_CASES)
                st.session_state.session_progress = {
                    "cases_remaining": cases_new,
                    "cases_done": [],
                    "operator": st.session_state.operator,
                    "model": st.session_state.ai_model,
                    "prompt": st.session_state.prompt_type
                }
                dm.save_session_progress(st.session_state.session_progress)
                st.session_state.current_case = cases_new[0]
                st.success("Đã xóa toàn bộ bảng tính và tạo bản sao lưu snapshot an toàn!")
                time.sleep(1)
                st.rerun()

    # Cấu hình từng cột trong st.data_editor
    column_config = {
        "Case_ID": st.column_config.TextColumn("Mã ca", help="Mã ca bệnh (VD: Case_001)", required=True),
        "AI_Model": st.column_config.SelectboxColumn("Mô hình AI", options=dm.MODELS_LIST + ["Khác"], required=True),
        "Prompt_Type": st.column_config.SelectboxColumn("Loại Prompt", options=list(dm.PROMPTS.keys())),
        "Time_Seconds": st.column_config.NumberColumn("Thời gian (s)", min_value=0, step=1, format="%d s"),
        "Pell_Gregory_Class": st.column_config.SelectboxColumn("P&G Class", options=dm.PG_CLASSES),
        "Pell_Gregory_Position": st.column_config.SelectboxColumn("P&G Position", options=dm.PG_POSITIONS),
        "Pell_Gregory_Full": st.column_config.TextColumn("P&G Full", help="Tự động tính từ Class & Position"),
        "Winter_Class": st.column_config.SelectboxColumn("Winter's Class", options=dm.WINTER_VALUES),
        "Confidence": st.column_config.NumberColumn("Confidence (%)", min_value=0, max_value=100, step=1, format="%d%%"),
        "Pederson_Level": st.column_config.SelectboxColumn("Pederson", options=dm.PEDERSON_LEVELS),
        "Hallucination_Flag": st.column_config.SelectboxColumn("Ảo giác", options=dm.HALLUCINATION_LEVELS),
        "Reasoning_Quality": st.column_config.SelectboxColumn("Lập luận", options=["0", "1", "2", "3"]),
        "Operator": st.column_config.SelectboxColumn("Người nhập", options=dm.OPERATORS_LIST + ["Khác"]),
        "Timestamp": st.column_config.TextColumn("Thời gian lưu"),
        "Session_Notes": st.column_config.TextColumn("Ghi chú"),
        "AI_Raw_Response": st.column_config.TextColumn("Phản hồi AI thô", width="medium"),
    }

    # Bảng tính tương tác st.data_editor
    edited_df = st.data_editor(
        sorted_df,
        column_config=column_config,
        num_rows="dynamic",
        use_container_width=True,
        height=480,
        key="spreadsheet_data_editor"
    )

    # Xử lý khi nhấn nút Lưu thay đổi trên bảng tính
    if save_sheet_changes:
        save_copy = edited_df.copy()
        # Tự động đồng bộ Pell_Gregory_Full nếu có chỉnh sửa Class hoặc Position
        for idx, r in save_copy.iterrows():
            cls = str(r.get("Pell_Gregory_Class", "")).strip()
            pos = str(r.get("Pell_Gregory_Position", "")).strip()
            if cls == "Không xác định" or pos == "Không xác định":
                save_copy.at[idx, "Pell_Gregory_Full"] = "Không xác định"
            elif cls and pos:
                save_copy.at[idx, "Pell_Gregory_Full"] = f"{cls}-{pos}"

        # Sắp xếp lại thứ tự tự nhiên của Case_ID trước khi lưu
        if "Case_ID" in save_copy.columns:
            save_copy["_sort"] = save_copy["Case_ID"].apply(lambda x: dm.natural_sort_key(x)[0])
            save_copy = save_copy.sort_values(by="_sort").drop(columns=["_sort"]).reset_index(drop=True)

        dm.save_dataset(save_copy, create_backup=True)
        st.success(f"✅ Đã lưu toàn bộ {len(save_copy)} dòng vào file dữ liệu thành công! Bản sao lưu snapshot tự động đã được ghi nhận.")
        time.sleep(1)
        st.rerun()

    # Thống kê nhanh dưới bảng
    st.write("")
    st.markdown("##### 📈 Tóm tắt dữ liệu hiện tại")
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        st.metric("Tổng số ca trong bảng", len(edited_df))
    with m_col2:
        if not edited_df.empty and "Time_Seconds" in edited_df.columns:
            avg_t = pd.to_numeric(edited_df["Time_Seconds"], errors="coerce").mean()
            st.metric("Thời gian TB", f"{round(avg_t, 1)} s" if not pd.isna(avg_t) else "—")
        else:
            st.metric("Thời gian TB", "—")
    with m_col3:
        if not edited_df.empty and "Confidence" in edited_df.columns:
            avg_c = pd.to_numeric(edited_df["Confidence"], errors="coerce").mean()
            st.metric("Độ tin cậy TB", f"{round(avg_c, 1)}%" if not pd.isna(avg_c) else "—")
        else:
            st.metric("Độ tin cậy TB", "—")
    with m_col4:
        if not edited_df.empty and "Hallucination_Flag" in edited_df.columns:
            halluc_count = (edited_df["Hallucination_Flag"].str.startswith("Có", na=False)).sum()
            rate = round((halluc_count / len(edited_df)) * 100, 1)
            st.metric("Tỷ lệ Ảo giác", f"{rate}% ({halluc_count} ca)")
        else:
            st.metric("Tỷ lệ Ảo giác", "—")


# ========================================================
# TAB 3: SAO LƯU & XUẤT DỮ LIỆU (BACKUP & EXPORT)
# ========================================================
with tabs[2]:
    st.markdown("""
    <div class="card-header">💾 QUẢN LÝ SAO LƯU, XUẤT EXCEL & KHÔI PHỤC DỮ LIỆU</div>
    """, unsafe_allow_html=True)

    df_export = dm.load_dataset(sort_by_case=True)

    exp_col1, exp_col2 = st.columns(2, gap="large")

    # ── CỘT 1: XUẤT DỮ LIỆU (EXCEL, CSV, ZIP) ───────────
    with exp_col1:
        st.markdown("""
        <div class="card-box">
            <h4 style="color: #64ffda; margin-top: 0;">📥 Tải xuống dữ liệu nghiên cứu</h4>
            <p style="color: #a8b2d1; font-size: 0.9rem;">
                Xuất toàn bộ cơ sở dữ liệu đánh giá ra các định dạng chuẩn để báo cáo khoa học hoặc xử lý thống kê.
            </p>
        """, unsafe_allow_html=True)

        # Trạng thái file Excel tự động trên máy tính
        if os.path.exists(dm.OUTPUT_EXCEL):
            mod_time = datetime.fromtimestamp(os.path.getmtime(dm.OUTPUT_EXCEL)).strftime("%H:%M:%S %d/%m/%Y")
            file_size_kb = round(os.path.getsize(dm.OUTPUT_EXCEL) / 1024, 1)
            st.info(f"📁 **File Excel tự động đồng bộ trên máy:** `{os.path.basename(dm.OUTPUT_EXCEL)}` ({file_size_kb} KB)\n\n"
                    f"📍 Đường dẫn: `{os.path.abspath(dm.OUTPUT_EXCEL)}`\n\n"
                    f"⏱ Cập nhật lần cuối: `{mod_time}` *(Tự động cập nhật mỗi khi Lưu hoặc sau 5 phút)*")

        # 1. Tải Excel .xlsx
        excel_data = dm.export_to_excel_bytes(df_export)
        st.download_button(
            label="📊 Tải xuống file Excel (.xlsx) [Đa Sheet & Định dạng chuẩn]",
            data=excel_data,
            file_name=f"Dental_AI_Evaluation_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

        st.write("")

        # 2. Tải CSV .csv
        csv_data = df_export.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button(
            label="📄 Tải xuống file CSV (.csv) [UTF-8 BOM Tiếng Việt]",
            data=csv_data,
            file_name=f"ai_evaluation_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            use_container_width=True
        )

        st.write("")

        # 3. Tải ZIP phản hồi AI thô
        zip_data = dm.create_ai_responses_zip()
        st.download_button(
            label="📦 Tải xuống toàn bộ phản hồi AI thô (.zip JSONs)",
            data=zip_data,
            file_name=f"ai_responses_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            use_container_width=True
        )

        # 4. Tải gói mã nguồn Deploy Streamlit Cloud (.zip)
        deploy_zip_path = os.path.join(dm.BASE_DIR, "Dental_AI_Streamlit_Deploy.zip")
        if os.path.exists(deploy_zip_path):
            st.write("")
            with open(deploy_zip_path, "rb") as zf:
                st.download_button(
                    label="🚀 Tải gói mã nguồn Deploy Streamlit Cloud (.zip)",
                    data=zf.read(),
                    file_name="Dental_AI_Streamlit_Deploy.zip",
                    mime="application/zip",
                    help="Gói zip đã được chuẩn hóa đầy đủ requirements.txt, .streamlit/config.toml, .gitignore, README.md sẵn sàng đẩy lên GitHub để deploy Streamlit Cloud.",
                    use_container_width=True
                )

        st.markdown("</div>", unsafe_allow_html=True)

        # Nhập dữ liệu từ ngoài
        st.markdown("""
        <div class="card-box">
            <h4 style="color: #64ffda; margin-top: 0;">📤 Nhập dữ liệu từ bên ngoài (Import)</h4>
            <p style="color: #a8b2d1; font-size: 0.9rem;">Tải lên file CSV hoặc Excel từ máy tính để bổ sung hoặc khôi phục dữ liệu.</p>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader("Chọn file CSV hoặc Excel:", type=["csv", "xlsx"], key="import_file_uploader")
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith(".csv"):
                    import_df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    import_df = pd.read_excel(uploaded_file, dtype=str)
                
                st.write(f"Đã phát hiện **{len(import_df)} dòng** dữ liệu trong file.")
                
                imp_c1, imp_c2 = st.columns(2)
                with imp_c1:
                    if st.button("➕ Nối tiếp vào dữ liệu hiện có", use_container_width=True):
                        combined = pd.concat([df_export, import_df], ignore_index=True)
                        dm.save_dataset(combined, create_backup=True)
                        st.success("Đã nối tiếp dữ liệu thành công!")
                        time.sleep(1)
                        st.rerun()
                with imp_c2:
                    if st.button("⚠️ Ghi đè toàn bộ dữ liệu", type="secondary", use_container_width=True):
                        dm.save_dataset(import_df, create_backup=True)
                        st.success("Đã ghi đè toàn bộ dữ liệu thành công!")
                        time.sleep(1)
                        st.rerun()
            except Exception as e:
                st.error(f"Lỗi khi đọc file tải lên: {e}")

        st.markdown("</div>", unsafe_allow_html=True)

    # ── CỘT 2: QUẢN LÝ BẢN SAO LƯU SNAPSHOT (BACKUPS) ───
    with exp_col2:
        st.markdown("""
        <div class="card-box">
            <h4 style="color: #64ffda; margin-top: 0;">📸 Quản lý Bản sao lưu Snapshot tự động</h4>
            <p style="color: #a8b2d1; font-size: 0.9rem;">
                Hệ thống tự động tạo snapshot trước mỗi lần lưu hoặc xóa. Bạn cũng có thể chủ động tạo bản sao lưu ngay lập tức bất cứ lúc nào.
            </p>
        """, unsafe_allow_html=True)

        # Nút tạo snapshot thủ công
        if st.button("📸 Tạo 1 bản sao lưu Snapshot tức thì", type="primary", use_container_width=True):
            bf = dm.create_snapshot_backup(note="manual_snapshot")
            if bf:
                st.success(f"✅ Đã tạo bản sao lưu thành công: `{bf}`")
                time.sleep(1)
                st.rerun()
            else:
                st.warning("Chưa có dữ liệu để tạo bản sao lưu.")

        st.write("")
        st.markdown("##### 📂 Danh sách các bản sao lưu hiện có:")
        backups_list = dm.list_backups()

        if backups_list:
            b_df = pd.DataFrame(backups_list)[["filename", "timestamp", "size_kb", "row_count"]]
            b_df.columns = ["Tên file", "Thời gian tạo", "Dung lượng (KB)", "Số ca (dòng)"]
            st.dataframe(b_df, use_container_width=True, height=220)

            st.write("")
            st.markdown("##### 🔄 Khôi phục từ bản sao lưu:")
            backup_names = [b["filename"] for b in backups_list]
            selected_backup = st.selectbox("Chọn bản sao lưu muốn khôi phục:", options=backup_names)

            with st.popover("⚠️ Xác nhận khôi phục bản sao lưu này", use_container_width=True):
                st.error(f"Hành động này sẽ thay thế dữ liệu hiện tại bằng `{selected_backup}`. Hệ thống sẽ tự động sao lưu an toàn dữ liệu hiện tại trước khi khôi phục.")
                if st.button("Xác nhận Khôi phục ngay", type="primary", use_container_width=True):
                    ok = dm.restore_backup(selected_backup)
                    if ok:
                        st.success(f"✅ Đã khôi phục thành công từ bản `{selected_backup}`!")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error("Không thể khôi phục file sao lưu đã chọn.")
        else:
            st.info("Chưa có bản sao lưu snapshot nào trong thư mục backups/.")

        st.markdown("</div>", unsafe_allow_html=True)
