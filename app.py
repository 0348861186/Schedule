from pathlib import Path

code = r'''# app.py
# ============================================================
# CODE 4 - HỆ THỐNG ĐỐI CHIẾU CHẤM CÔNG NÂNG CẤP
# Kiến trúc:
#   70% CODE 2: engine xác định chấm công bằng Python
#   20% CODE 1: Smart Excel Loader + giao diện dễ dùng
#   10% CODE 3: dashboard song ngữ + báo cáo bất thường
#
# Nguyên tắc:
#   - Python quyết định giờ vào/ra, số giờ, BV/BR, vắng, trễ, về sớm.
#   - Gemini CHỈ hỗ trợ nhận diện tên cột và tóm tắt bất thường.
#   - Có fallback không cần Gemini nếu cấu hình API chưa có.
# ============================================================

import io
import re
import time
import datetime as dt
from typing import Optional, Dict, List, Tuple

import pandas as pd
import streamlit as st
import plotly.express as px

try:
    from google import genai
except ImportError:
    genai = None


# ============================================================
# 0. PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Dashboard Chấm Công / 考勤 Dashboard",
    page_icon="📊",
    layout="wide",
)

# ============================================================
# 1. NGÔN NGỮ
# ============================================================

LANG = {
    "vi": {
        "title": "📊 HỆ THỐNG ĐỐI CHIẾU & PHÂN TÍCH CHẤM CÔNG",
        "subtitle": "Dashboard Thống Kê Chấm Công Nhân Sự",
        "switch": "切换至中文 (Chuyển sang tiếng Trung)",
        "upload": "📁 1. Tải dữ liệu",
        "finger": "1) File bấm vân tay",
        "schedule": "2) File lịch xếp ca công nhân",
        "vp": "3) Danh sách nhân viên Văn phòng",
        "cn": "4) Danh sách Công nhân",
        "date": "📅 2. Chọn ngày kiểm tra",
        "run": "🚀 Chạy đối chiếu chấm công",
        "settings": "⚙️ Cấu hình",
        "gemini_key": "Gemini API Key (tùy chọn)",
        "gemini_help": "Gemini chỉ hỗ trợ nhận diện cột và tóm tắt; không quyết định kết quả chấm công.",
        "dashboard": "📈 3. Tổng quan",
        "abnormal": "🚨 4. Các trường hợp bất thường",
        "export": "💾 5. Xuất báo cáo",
        "total": "Tổng nhân sự dự kiến",
        "normal": "Bình thường",
        "abnormal_count": "Bất thường",
        "absent": "Vắng",
        "off": "Nghỉ",
        "all": "Tất cả",
        "none": "Không có dữ liệu",
        "excel": "📥 Tải Excel kết quả",
        "summary": "Tóm tắt bất thường",
        "no_abnormal": "✅ Không phát hiện bất thường.",
        "mapping": "Ánh xạ cột",
        "engine": "Kết quả xử lý",
        "show_raw": "Xem dữ liệu chuẩn hóa",
    },
    "zh": {
        "title": "📊 员工考勤核对与分析系统",
        "subtitle": "考勤统计 Dashboard",
        "switch": "Chuyển sang tiếng Việt (切换至越南语)",
        "upload": "📁 1. 上传数据",
        "finger": "1) 指纹打卡文件",
        "schedule": "2) 工人排班文件",
        "vp": "3) 办公室员工名单",
        "cn": "4) 工人名单",
        "date": "📅 2. 选择检查日期",
        "run": "🚀 开始考勤核对",
        "settings": "⚙️ 系统配置",
        "gemini_key": "Gemini API Key（可选）",
        "gemini_help": "Gemini 仅用于列名识别和异常摘要，不负责决定考勤结果。",
        "dashboard": "📈 3. 总览",
        "abnormal": "🚨 4. 异常情况",
        "export": "💾 5. 导出报告",
        "total": "预计人员",
        "normal": "正常",
        "abnormal_count": "异常",
        "absent": "缺勤",
        "off": "休息",
        "all": "全部",
        "none": "暂无数据",
        "excel": "📥 下载 Excel 结果",
        "summary": "异常摘要",
        "no_abnormal": "✅ 未发现异常。",
        "mapping": "列名映射",
        "engine": "处理结果",
        "show_raw": "查看标准化数据",
    },
}

if "lang" not in st.session_state:
    st.session_state.lang = "vi"

L = LANG[st.session_state.lang]

if st.sidebar.button(L["switch"]):
    st.session_state.lang = "zh" if st.session_state.lang == "vi" else "vi"
    st.rerun()

st.title(L["title"])
st.caption(L["subtitle"])


# ============================================================
# 2. CẤU HÌNH
# ============================================================

with st.sidebar:
    st.header(L["settings"])

    gemini_key = st.text_input(
        L["gemini_key"],
        type="password",
        help=L["gemini_help"],
    )

    st.caption(
        "Có thể dùng GEMINI_API_KEY trong Streamlit Secrets để không nhập mỗi lần."
    )

    # Các quy tắc ca đặc biệt được đưa ra UI để tránh hard-code khó sửa.
    st.markdown("### Quy tắc VP")
    vp_default_in = st.text_input("VP giờ vào mặc định", "08:00")
    vp_default_out = st.text_input("VP giờ ra mặc định", "17:00")
    late_grace = st.number_input("Cho phép trễ (phút)", min_value=0, max_value=60, value=5)

    st.markdown("### Nhân sự đặc biệt")
    st.caption("Có thể sửa giờ trực tiếp khi doanh nghiệp thay đổi quy định.")

    special_575_in = st.text_input("Mã 575 - vào", "07:00")
    special_575_out = st.text_input("Mã 575 - ra", "15:00")
    special_575_expected = st.number_input(
        "Mã 575 - giờ công chuẩn",
        min_value=0.0,
        max_value=24.0,
        value=8.0,
        step=0.5,
    )

    special_749_in = st.text_input("Mã 749/949 - vào", "07:00")
    special_749_out = st.text_input("Mã 749/949 - ra", "19:00")
    special_749_expected = st.number_input(
        "Mã 749/949 - giờ công chuẩn",
        min_value=0.0,
        max_value=24.0,
        value=12.0,
        step=0.5,
    )


# ============================================================
# 3. SMART EXCEL LOADER - lấy từ tinh thần CODE 1, nhưng an toàn hơn
# ============================================================

def normalize_col_name(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def load_smart_excel(uploaded_file) -> Optional[pd.DataFrame]:
    """Đọc Excel/CSV và tự tìm dòng header hợp lý."""
    if uploaded_file is None:
        return None

    try:
        name = uploaded_file.name.lower()

        if name.endswith(".csv"):
            # thử UTF-8 rồi fallback cp1258
            uploaded_file.seek(0)
            try:
                return pd.read_csv(uploaded_file)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding="cp1258")

        uploaded_file.seek(0)
        raw = pd.read_excel(uploaded_file, header=None)

        if raw.empty:
            return pd.DataFrame()

        keywords = [
            "mã", "ma", "id", "nhân viên", "nhan vien",
            "ngày", "ngay", "date", "thời gian", "thoi gian",
            "time", "giờ", "gio"
        ]

        best_row = 0
        best_score = -1

        for idx, row in raw.iterrows():
            text = " | ".join(str(v).strip().lower() for v in row.tolist())
            score = sum(1 for k in keywords if k in text)
            if score > best_score:
                best_score = score
                best_row = idx

        df = pd.read_excel(uploaded_file, header=best_row)
        df.columns = [normalize_col_name(c) for c in df.columns]

        # bỏ cột Unnamed hoàn toàn rỗng
        keep = []
        for c in df.columns:
            if str(c).lower().startswith("unnamed"):
                if df[c].notna().sum() == 0:
                    continue
            keep.append(c)

        return df[keep].copy()

    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None


# ============================================================
# 4. CHUẨN HÓA GIÁ TRỊ
# ============================================================

def normalize_id(value) -> str:
    """00123, 123, 123.0 -> 123; giữ mã chữ nếu có."""
    if pd.isna(value):
        return ""

    s = str(value).strip()

    if s.lower() in {"nan", "none", "nat", ""}:
        return ""

    # 123.0 -> 123
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".")[0]

    # 00123 -> 123 (phù hợp phần lớn máy chấm công)
    if s.isdigit():
        s = str(int(s))

    return s


def clean_name(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def parse_datetime_value(value) -> pd.Timestamp:
    if pd.isna(value):
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def parse_time_value(value) -> Optional[dt.time]:
    if pd.isna(value):
        return None

    if isinstance(value, dt.datetime):
        return value.time()

    if isinstance(value, dt.time):
        return value

    s = str(value).strip()

    if s.lower() in {"", "nan", "none", "nat", "-", "0:00:00"}:
        return None

    # Excel có thể trả 07:30 hoặc 07:30:00
    parsed = pd.to_datetime(s, errors="coerce")
    if pd.isna(parsed):
        return None

    return parsed.time()


def combine_date_time(date_value, time_value) -> pd.Timestamp:
    t = parse_time_value(time_value)
    if t is None:
        return pd.NaT

    if isinstance(date_value, dt.datetime):
        d = date_value.date()
    elif isinstance(date_value, dt.date):
        d = date_value
    else:
        d = pd.to_datetime(date_value, errors="coerce")
        if pd.isna(d):
            return pd.NaT
        d = d.date()

    return pd.Timestamp(dt.datetime.combine(d, t))


# ============================================================
# 5. NHẬN DIỆN CỘT
# ============================================================

SYNONYMS = {
    "employee_id": [
        "mã nhân viên", "ma nhan vien", "mã nv", "ma nv",
        "mã công nhân", "ma cong nhan", "employee id", "employee",
        "staff id", "id", "mã", "ma", "card", "thẻ", "the"
    ],
    "name": [
        "họ và tên", "ho va ten", "họ tên", "ho ten",
        "tên nhân viên", "ten nhan vien", "name", "employee name",
        "nhân viên", "nhan vien"
    ],
    "datetime": [
        "thời gian", "thoi gian", "datetime", "date time",
        "timestamp", "ngày giờ", "ngay gio", "thời điểm",
        "thoi diem", "time"
    ],
    "date": [
        "ngày", "ngay", "date", "ngày chấm công", "ngay cham cong"
    ],
    "time": [
        "giờ", "gio", "time", "thời gian", "thoi gian"
    ],
    "shift": [
        "ca", "shift", "lịch", "lich", "xếp ca", "xep ca", "schedule"
    ],
}


def local_column_mapping(columns: List[str], concept: str) -> Optional[str]:
    """Ưu tiên fuzzy đơn giản trước khi gọi Gemini."""
    cols = [str(c).strip() for c in columns]
    normalized = {c: re.sub(r"[^a-z0-9à-ỹ ]", " ", c.lower()) for c in cols}

    for c, n in normalized.items():
        for syn in SYNONYMS.get(concept, []):
            if syn in n:
                return c

    return None


def gemini_column_mapping(
    columns: List[str],
    target_concept: str,
    client=None,
) -> Optional[str]:
    """Gemini chỉ nhận danh sách cột, không gửi toàn bộ dữ liệu nhân viên."""
    if client is None or not columns:
        return None

    prompt = f"""
Danh sách tên cột Excel:
{columns}

Khái niệm cần tìm:
{target_concept}

Chỉ trả về đúng MỘT tên cột lấy nguyên văn từ danh sách.
Nếu không có, trả về NONE.
Không giải thích.
"""

    try:
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )
        answer = response.text.strip().strip("`").strip()

        for c in columns:
            if answer.lower() == str(c).strip().lower():
                return c

        return None
    except Exception:
        return None


def map_column(
    df: pd.DataFrame,
    concept: str,
    client=None,
) -> Optional[str]:
    local = local_column_mapping(list(df.columns), concept)
    if local:
        return local

    return gemini_column_mapping(list(df.columns), concept, client)


# ============================================================
# 6. CHUẨN HÓA FILE VÂN TAY
# ============================================================

def prepare_fingerprint(
    df: pd.DataFrame,
    client=None,
) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:

    df = df.copy()
    df.columns = [normalize_col_name(c) for c in df.columns]

    id_col = map_column(df, "employee_id", client)
    datetime_col = map_column(df, "datetime", client)

    # Nếu không có datetime duy nhất, tìm date + time
    date_col = map_column(df, "date", client)
    time_col = map_column(df, "time", client)

    if not id_col:
        raise ValueError("Không tìm thấy cột Mã nhân viên/thẻ trong file vân tay.")

    if not datetime_col and not (date_col and time_col):
        raise ValueError(
            "Không tìm thấy cột ngày giờ hoặc cặp cột ngày + giờ trong file vân tay."
        )

    out = df.copy()
    out["_EmployeeID"] = out[id_col].apply(normalize_id)

    if datetime_col:
        out["_DateTime"] = out[datetime_col].apply(parse_datetime_value)
    else:
        out["_DateTime"] = [
            combine_date_time(d, t)
            for d, t in zip(out[date_col], out[time_col])
        ]

    out["_Date"] = out["_DateTime"].dt.date

    mapping = {
        "employee_id": id_col,
        "datetime": datetime_col,
        "date": date_col,
        "time": time_col,
    }

    return out, mapping


# ============================================================
# 7. CHUẨN HÓA MASTER
# ============================================================

def prepare_master(
    df: pd.DataFrame,
    kind: str,
    client=None,
) -> Tuple[pd.DataFrame, Dict[str, Optional[str]]]:

    df = df.copy()
    df.columns = [normalize_col_name(c) for c in df.columns]

    id_col = map_column(df, "employee_id", client)
    name_col = map_column(df, "name", client)

    if not id_col:
        raise ValueError(f"Không tìm thấy cột mã nhân viên trong danh sách {kind}.")
    if not name_col:
        raise ValueError(f"Không tìm thấy cột họ tên trong danh sách {kind}.")

    out = df.copy()
    out["_EmployeeID"] = out[id_col].apply(normalize_id)
    out["_Name"] = out[name_col].apply(clean_name)

    return out, {
        "employee_id": id_col,
        "name": name_col,
    }


# ============================================================
# 8. CHUẨN HÓA LỊCH CA
# ============================================================

def prepare_schedule(df: pd.DataFrame, client=None):
    df = df.copy()
    df.columns = [normalize_col_name(c) for c in df.columns]

    id_col = map_column(df, "employee_id", client)
    if not id_col:
        raise ValueError("Không tìm thấy cột mã nhân viên trong lịch ca.")

    out = df.copy()
    out["_EmployeeID"] = out[id_col].apply(normalize_id)

    return out, {"employee_id": id_col}


def get_schedule_shift(schedule_df, employee_id: str, target_date: dt.date):
    """Hỗ trợ dạng cột ngày = 1,2,3... hoặc ngày thực tế."""
    if schedule_df is None or schedule_df.empty:
        return "OFF"

    rows = schedule_df[schedule_df["_EmployeeID"] == employee_id]
    if rows.empty:
        return "OFF"

    row = rows.iloc[0]

    # 1) cột đúng số ngày trong tháng
    candidates = [
        str(target_date.day),
        f"{target_date.day:02d}",
    ]

    # 2) cột ngày dạng dd/mm/yyyy, yyyy-mm-dd...
    candidates += [
        target_date.strftime("%d/%m/%Y"),
        target_date.strftime("%Y-%m-%d"),
        target_date.strftime("%d-%m-%Y"),
    ]

    for c in candidates:
        if c in schedule_df.columns:
            value = row[c]
            if pd.isna(value):
                return "OFF"
            return str(value).strip()

    return "OFF"


def normalize_shift(value: str) -> str:
    s = str(value).strip().upper()

    mapping = {
        "N": "N",
        "CA NGÀY": "N",
        "CA NGAY": "N",
        "DAY": "N",
        "D": "D",
        "Đ": "D",
        "CA ĐÊM": "D",
        "CA DEM": "D",
        "NIGHT": "D",
        "OFF": "OFF",
        "NGHỈ": "OFF",
        "NGHI": "OFF",
        "-": "OFF",
        "NAN": "OFF",
    }

    return mapping.get(s, s)


# ============================================================
# 9. ENGINE CHẤM CÔNG
# ============================================================

def fmt_time(ts) -> str:
    if ts is None or pd.isna(ts):
        return "-"
    return pd.Timestamp(ts).strftime("%H:%M")


def hours_between(t_in, t_out) -> float:
    if t_in is None or t_out is None or pd.isna(t_in) or pd.isna(t_out):
        return 0.0

    sec = (pd.Timestamp(t_out) - pd.Timestamp(t_in)).total_seconds()

    # nếu dữ liệu có tình huống qua ngày
    if sec < 0:
        sec += 24 * 3600

    return round(sec / 3600, 2)


def get_first_last(logs: pd.DataFrame):
    if logs.empty:
        return None, None

    logs = logs.sort_values("_DateTime")
    t_in = logs["_DateTime"].iloc[0]
    t_out = logs["_DateTime"].iloc[-1] if len(logs) >= 2 else None
    return t_in, t_out


def evaluate_pair(
    t_in,
    t_out,
    expected_in: dt.time,
    expected_out: dt.time,
    expected_hours: float,
    target_date: dt.date,
    grace_minutes: int,
):
    status_codes = []
    notes = []

    if t_in is None or pd.isna(t_in):
        status_codes.append("MISSING_IN")
        notes.append("BV")
    if t_out is None or pd.isna(t_out):
        status_codes.append("MISSING_OUT")
        notes.append("BR")

    if t_in is None or pd.isna(t_in):
        if t_out is None or pd.isna(t_out):
            status_codes = ["ABSENT"]
            notes = ["Vắng"]
        return 0.0, status_codes, notes

    actual_hours = hours_between(t_in, t_out)

    limit_in = pd.Timestamp(dt.datetime.combine(target_date, expected_in))
    limit_out = pd.Timestamp(dt.datetime.combine(target_date, expected_out))

    # Cho phép giờ ra nằm ngày sau nếu cần
    if pd.Timestamp(t_out) < pd.Timestamp(t_in):
        limit_out += pd.Timedelta(days=1)

    if pd.Timestamp(t_in) > limit_in + pd.Timedelta(minutes=grace_minutes):
        status_codes.append("LATE")
        notes.append("Đi trễ")

    if t_out is not None and not pd.isna(t_out):
        if pd.Timestamp(t_out) < limit_out:
            status_codes.append("EARLY")
            notes.append("Về sớm")

    if (
        t_out is not None
        and not pd.isna(t_out)
        and expected_hours > 0
        and actual_hours + 1e-9 < expected_hours
        and "EARLY" not in status_codes
    ):
        status_codes.append("SHORT_HOURS")
        notes.append(f"Thiếu giờ ({actual_hours}h)")

    if not status_codes:
        status_codes = ["NORMAL"]
        notes = ["Đủ công"]

    return actual_hours, status_codes, notes


def process_vp(
    master_vp,
    finger,
    target_date,
    special_rules,
    grace_minutes,
):
    rows = []

    for _, emp in master_vp.iterrows():
        eid = emp["_EmployeeID"]
        name = emp["_Name"]

        logs = finger[
            (finger["_EmployeeID"] == eid)
            & (finger["_Date"] == target_date)
        ].sort_values("_DateTime")

        # Chủ nhật nghỉ VP
        if target_date.weekday() == 6:
            rows.append({
                "Mã NV": eid,
                "Họ và Tên": name,
                "Khối": "VP",
                "Ca": "OFF",
                "Giờ vào": "-",
                "Giờ ra": "-",
                "Giờ làm thực tế": 0.0,
                "Status_Code": "OFF",
                "Ghi chú": "Nghỉ tuần",
            })
            continue

        if eid == "575":
            in_s, out_s, expected_h = special_rules["575"]
        elif eid in {"749", "949"}:
            in_s, out_s, expected_h = special_rules["749_949"]
        else:
            in_s, out_s, expected_h = special_rules["DEFAULT"]

        expected_in = dt.time.fromisoformat(in_s)
        expected_out = dt.time.fromisoformat(out_s)

        t_in, t_out = get_first_last(logs)

        actual, codes, notes = evaluate_pair(
            t_in,
            t_out,
            expected_in,
            expected_out,
            expected_h,
            target_date,
            grace_minutes,
        )

        rows.append({
            "Mã NV": eid,
            "Họ và Tên": name,
            "Khối": "VP",
            "Ca": "VP",
            "Giờ vào": fmt_time(t_in),
            "Giờ ra": fmt_time(t_out),
            "Giờ làm thực tế": actual,
            "Status_Code": "|".join(codes),
            "Ghi chú": ", ".join(notes),
        })

    return rows


def process_cn(
    master_cn,
    schedule,
    finger,
    target_date,
    grace_minutes,
):
    rows = []

    for _, emp in master_cn.iterrows():
        eid = emp["_EmployeeID"]
        name = emp["_Name"]

        raw_shift = get_schedule_shift(schedule, eid, target_date)
        shift = normalize_shift(raw_shift)

        if shift == "OFF":
            rows.append({
                "Mã NV": eid,
                "Họ và Tên": name,
                "Khối": "CN",
                "Ca": "OFF",
                "Giờ vào": "-",
                "Giờ ra": "-",
                "Giờ làm thực tế": 0.0,
                "Status_Code": "OFF",
                "Ghi chú": "Nghỉ ca",
            })
            continue

        # ==========================
        # CA NGÀY: 07:00 -> 19:00
        # ==========================
        if shift == "N":
            logs = finger[
                (finger["_EmployeeID"] == eid)
                & (finger["_Date"] == target_date)
            ].sort_values("_DateTime")

            t_in, t_out = get_first_last(logs)

            actual, codes, notes = evaluate_pair(
                t_in,
                t_out,
                dt.time(7, 0),
                dt.time(19, 0),
                12.0,
                target_date,
                grace_minutes,
            )

            rows.append({
                "Mã NV": eid,
                "Họ và Tên": name,
                "Khối": "CN",
                "Ca": "N",
                "Giờ vào": fmt_time(t_in),
                "Giờ ra": fmt_time(t_out),
                "Giờ làm thực tế": actual,
                "Status_Code": "|".join(codes),
                "Ghi chú": ", ".join(notes),
            })

        # ==========================
        # CA ĐÊM: 19:00 -> 07:00 hôm sau
        # ==========================
        elif shift == "D":
            today_logs = finger[
                (finger["_EmployeeID"] == eid)
                & (finger["_Date"] == target_date)
            ].sort_values("_DateTime")

            next_date = target_date + dt.timedelta(days=1)

            next_logs = finger[
                (finger["_EmployeeID"] == eid)
                & (finger["_Date"] == next_date)
            ].sort_values("_DateTime")

            # vào: từ 18:00 trở đi để có biên an toàn
            in_candidates = today_logs[
                today_logs["_DateTime"].dt.time >= dt.time(18, 0)
            ]

            # ra: ngày hôm sau trước 08:00
            out_candidates = next_logs[
                next_logs["_DateTime"].dt.time <= dt.time(8, 0)
            ]

            t_in = in_candidates["_DateTime"].min() if not in_candidates.empty else None
            t_out = out_candidates["_DateTime"].max() if not out_candidates.empty else None

            status_codes = []
            notes = []

            if t_in is None:
                status_codes.append("MISSING_IN")
                notes.append("BV")

            if t_out is None:
                status_codes.append("MISSING_OUT")
                notes.append("BR")

            if t_in is None and t_out is None:
                status_codes = ["ABSENT"]
                notes = ["Vắng"]
                actual = 0.0
            else:
                actual = hours_between(t_in, t_out)

                if t_in is not None:
                    expected_in_dt = pd.Timestamp(
                        dt.datetime.combine(target_date, dt.time(19, 0))
                    )
                    if t_in > expected_in_dt + pd.Timedelta(minutes=grace_minutes):
                        status_codes.append("LATE")
                        notes.append("Đi trễ")

                if t_out is not None:
                    expected_out_dt = pd.Timestamp(
                        dt.datetime.combine(next_date, dt.time(7, 0))
                    )
                    if t_out < expected_out_dt:
                        status_codes.append("EARLY")
                        notes.append("Về sớm")

                if t_in is not None and t_out is not None and not status_codes:
                    status_codes = ["NORMAL"]
                    notes = ["Đủ công"]

            rows.append({
                "Mã NV": eid,
                "Họ và Tên": name,
                "Khối": "CN",
                "Ca": "D",
                "Giờ vào": fmt_time(t_in),
                "Giờ ra": fmt_time(t_out),
                "Giờ làm thực tế": actual,
                "Status_Code": "|".join(status_codes),
                "Ghi chú": ", ".join(notes),
            })

        # Ca lạ -> báo lỗi thay vì tự đoán
        else:
            rows.append({
                "Mã NV": eid,
                "Họ và Tên": name,
                "Khối": "CN",
                "Ca": shift,
                "Giờ vào": "-",
                "Giờ ra": "-",
                "Giờ làm thực tế": 0.0,
                "Status_Code": "UNKNOWN_SHIFT",
                "Ghi chú": f"Không nhận diện ca: {raw_shift}",
            })

    return rows


def run_attendance_engine(
    finger,
    master_vp,
    master_cn,
    schedule,
    target_date,
    special_rules,
    grace_minutes,
):
    vp_rows = process_vp(
        master_vp,
        finger,
        target_date,
        special_rules,
        grace_minutes,
    )

    cn_rows = process_cn(
        master_cn,
        schedule,
        finger,
        target_date,
        grace_minutes,
    )

    result = pd.DataFrame(vp_rows + cn_rows)

    if not result.empty:
        result["Ngày"] = target_date.isoformat()

        # thứ tự cột
        cols = [
            "Ngày", "Mã NV", "Họ và Tên", "Khối", "Ca",
            "Giờ vào", "Giờ ra", "Giờ làm thực tế",
            "Status_Code", "Ghi chú"
        ]
        result = result[[c for c in cols if c in result.columns]]

    return result


# ============================================================
# 10. GEMINI - CHỈ TÓM TẮT KẾT QUẢ, KHÔNG TÍNH CHẤM CÔNG
# ============================================================

def make_gemini_client(api_key: str):
    if genai is None:
        return None

    key = api_key.strip() if api_key else ""

    if not key:
        try:
            key = st.secrets.get("GEMINI_API_KEY", "")
        except Exception:
            key = ""

    if not key:
        return None

    try:
        return genai.Client(api_key=key)
    except Exception:
        return None


def gemini_summary(df_abnormal: pd.DataFrame, client) -> str:
    if client is None or df_abnormal.empty:
        return ""

    safe = df_abnormal[
        [
            "Mã NV", "Họ và Tên", "Khối", "Ca",
            "Giờ vào", "Giờ ra", "Giờ làm thực tế",
            "Status_Code", "Ghi chú"
        ]
    ].copy()

    prompt = f"""
Bạn là trợ lý lập báo cáo chấm công.

Kết quả dưới đây ĐÃ được hệ thống Python tính toán.
Bạn KHÔNG được thay đổi, suy diễn hoặc sửa Status_Code.

Hãy viết tối đa 5 dòng tóm tắt:
- Có bao nhiêu trường hợp bất thường.
- Nhóm lỗi nổi bật.
- Nếu có, nêu các mã nhân viên cần chú ý nhất.
- Viết song ngữ Việt - Trung.
- Không dùng bảng.
- Không nói về cách hệ thống tính toán.

DỮ LIỆU:
{safe.to_string(index=False)}
"""

    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
            )
            return response.text.strip()
        except Exception as e:
            if "503" in str(e) and attempt < 2:
                time.sleep(2)
                continue
            return ""

    return ""


# ============================================================
# 11. DASHBOARD
# ============================================================

st.header(L["upload"])

c1, c2 = st.columns(2)
c3, c4 = st.columns(2)

with c1:
    f_finger = st.file_uploader(
        L["finger"],
        type=["xlsx", "xls", "csv"],
        key="finger",
    )

with c2:
    f_schedule = st.file_uploader(
        L["schedule"],
        type=["xlsx", "xls", "csv"],
        key="schedule",
    )

with c3:
    f_vp = st.file_uploader(
        L["vp"],
        type=["xlsx", "xls", "csv"],
        key="vp",
    )

with c4:
    f_cn = st.file_uploader(
        L["cn"],
        type=["xlsx", "xls", "csv"],
        key="cn",
    )


# ============================================================
# 12. LOAD FILE
# ============================================================

df_finger_raw = load_smart_excel(f_finger)
df_schedule_raw = load_smart_excel(f_schedule)
df_vp_raw = load_smart_excel(f_vp)
df_cn_raw = load_smart_excel(f_cn)

finger_prepared = None
schedule_prepared = None
vp_prepared = None
cn_prepared = None

mapping_info = {}


if df_finger_raw is not None:
    try:
        temp_client = make_gemini_client(gemini_key)
        finger_prepared, mapping_info["fingerprint"] = prepare_fingerprint(
            df_finger_raw,
            temp_client,
        )
    except Exception as e:
        st.error(f"File vân tay: {e}")


if df_schedule_raw is not None:
    try:
        temp_client = make_gemini_client(gemini_key)
        schedule_prepared, mapping_info["schedule"] = prepare_schedule(
            df_schedule_raw,
            temp_client,
        )
    except Exception as e:
        st.error(f"File lịch ca: {e}")


if df_vp_raw is not None:
    try:
        temp_client = make_gemini_client(gemini_key)
        vp_prepared, mapping_info["vp"] = prepare_master(
            df_vp_raw,
            "VP",
            temp_client,
        )
    except Exception as e:
        st.error(f"File VP: {e}")


if df_cn_raw is not None:
    try:
        temp_client = make_gemini_client(gemini_key)
        cn_prepared, mapping_info["cn"] = prepare_master(
            df_cn_raw,
            "CN",
            temp_client,
        )
    except Exception as e:
        st.error(f"File CN: {e}")


# ============================================================
# 13. CHỌN NGÀY
# ============================================================

st.header(L["date"])

available_dates = []

if finger_prepared is not None and not finger_prepared.empty:
    available_dates = sorted(
        d for d in finger_prepared["_Date"].dropna().unique()
    )

if available_dates:
    default_index = len(available_dates) - 1
    target_date = st.selectbox(
        "Ngày / 日期",
        available_dates,
        index=default_index,
        format_func=lambda x: x.strftime("%d/%m/%Y"),
    )
else:
    target_date = st.date_input(
        "Ngày / 日期",
        value=dt.date.today(),
    )


# ============================================================
# 14. CHẠY ENGINE
# ============================================================

all_ready = all([
    finger_prepared is not None,
    vp_prepared is not None,
    cn_prepared is not None,
    schedule_prepared is not None,
])

if not all_ready:
    st.info(
        "Vui lòng tải đủ 4 file: vân tay + lịch ca + VP + CN."
    )

if st.button(L["run"], type="primary", disabled=not all_ready):

    special_rules = {
        "DEFAULT": (
            vp_default_in.strip() + (":00" if len(vp_default_in.strip().split(":")) == 2 else ""),
            vp_default_out.strip() + (":00" if len(vp_default_out.strip().split(":")) == 2 else ""),
            8.0,
        ),
        "575": (
            special_575_in.strip() + (":00" if len(special_575_in.strip().split(":")) == 2 else ""),
            special_575_out.strip() + (":00" if len(special_575_out.strip().split(":")) == 2 else ""),
            float(special_575_expected),
        ),
        "749_949": (
            special_749_in.strip() + (":00" if len(special_749_in.strip().split(":")) == 2 else ""),
            special_749_out.strip() + (":00" if len(special_749_out.strip().split(":")) == 2 else ""),
            float(special_749_expected),
        ),
    }

    with st.spinner("Đang đối chiếu dữ liệu bằng Python..."):
        try:
            result = run_attendance_engine(
                finger_prepared,
                vp_prepared,
                cn_prepared,
                schedule_prepared,
                target_date,
                special_rules,
                int(late_grace),
            )

            st.session_state["final_report"] = result
            st.session_state["result_date"] = target_date

            # Gemini chỉ tóm tắt phần bất thường
            client = make_gemini_client(gemini_key)

            abnormal_mask = ~result["Status_Code"].isin(["NORMAL", "OFF"])
            df_abnormal = result[abnormal_mask].copy()

            summary = gemini_summary(df_abnormal, client)
            st.session_state["gemini_summary"] = summary

            st.success("✅ Đối chiếu hoàn tất.")

        except Exception as e:
            st.error(f"Lỗi xử lý: {e}")


# ============================================================
# 15. HIỂN THỊ KẾT QUẢ
# ============================================================

if "final_report" in st.session_state:
    result = st.session_state["final_report"].copy()

    if not result.empty:
        st.header(L["dashboard"])

        total = len(result)
        normal = int((result["Status_Code"] == "NORMAL").sum())
        absent = int(result["Status_Code"].str.contains("ABSENT", na=False).sum())
        off = int((result["Status_Code"] == "OFF").sum())
        abnormal = total - normal - off

        m1, m2, m3, m4 = st.columns(4)
        m1.metric(L["total"], total)
        m2.metric(L["normal"], normal)
        m3.metric(L["abnormal_count"], abnormal)
        m4.metric(L["absent"], absent)

        chart_df = pd.DataFrame({
            "Trạng thái": [
                "Bình thường",
                "Bất thường",
                "Vắng",
                "Nghỉ",
            ],
            "Số lượng": [
                normal,
                abnormal,
                absent,
                off,
            ],
        })

        fig = px.pie(
            chart_df,
            names="Trạng thái",
            values="Số lượng",
            hole=0.45,
        )
        st.plotly_chart(fig, use_container_width=True)

        # -------------------------------
        # Bất thường
        # -------------------------------
        st.header(L["abnormal"])

        abnormal_df = result[
            ~result["Status_Code"].isin(["NORMAL", "OFF"])
        ].copy()

        if abnormal_df.empty:
            st.success(L["no_abnormal"])
        else:
            st.dataframe(
                abnormal_df,
                use_container_width=True,
                hide_index=True,
            )

            summary = st.session_state.get("gemini_summary", "")
            if summary:
                st.info(summary)

        # -------------------------------
        # Toàn bộ kết quả
        # -------------------------------
        with st.expander(L["engine"]):
            st.dataframe(
                result,
                use_container_width=True,
                hide_index=True,
            )

        # -------------------------------
        # Mapping
        # -------------------------------
        with st.expander(L["mapping"]):
            for source, mapping in mapping_info.items():
                st.markdown(f"**{source}**")
                st.json(mapping)

        # -------------------------------
        # Raw standardized data
        # -------------------------------
        with st.expander(L["show_raw"]):
            if finger_prepared is not None:
                st.dataframe(
                    finger_prepared[
                        ["_EmployeeID", "_DateTime", "_Date"]
                    ].sort_values("_DateTime"),
                    use_container_width=True,
                    hide_index=True,
                )

        # ====================================================
        # 16. EXCEL EXPORT
        # ====================================================

        st.header(L["export"])

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl",
        ) as writer:

            result.to_excel(
                writer,
                index=False,
                sheet_name="KetQua_ChamCong",
            )

            abnormal_df.to_excel(
                writer,
                index=False,
                sheet_name="BatThuong",
            )

            # Mapping
            mapping_rows = []
            for source, mapping in mapping_info.items():
                for concept, col in mapping.items():
                    mapping_rows.append({
                        "Nguồn": source,
                        "Khái niệm": concept,
                        "Cột được chọn": col or "NONE",
                    })

            pd.DataFrame(mapping_rows).to_excel(
                writer,
                index=False,
                sheet_name="Mapping",
            )

            # Thống kê
            chart_df.to_excel(
                writer,
                index=False,
                sheet_name="ThongKe",
            )

        st.download_button(
            label=L["excel"],
            data=output.getvalue(),
            file_name=f"Bao_Cao_Cham_Cong_{target_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption(
            "PDF: dùng Ctrl+P → Save as PDF để lưu nguyên giao diện Dashboard."
        )

else:
    st.info("Sau khi chạy đối chiếu, kết quả sẽ xuất hiện tại đây.")
'''

