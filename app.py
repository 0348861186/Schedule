from pathlib import Path

file_path = Path(path)
# Tạo thư mục cha nếu chưa tồn tại (parents=True giúp tạo cả thư mục cha lồng nhau)
file_path.parent.mkdir(parents=True, exist_ok=True)

# Sau đó mới tiến hành ghi tệp
file_path.write_text(code, encoding="utf-8")

code = r'''import io
import os
import re
import json
import time
import hashlib
import unicodedata
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
from google import genai
from google.genai import types
import streamlit.components.v1 as components


# ============================================================
# CẤU HÌNH TRANG STREAMLIT
# ============================================================
st.set_page_config(
    page_title="考勤统计仪表盘 / Dashboard Thống Kê Chấm Công",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.header("⚙️ 系统配置 / Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập API Key Gemini", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key


# ============================================================
# TIÊU ĐỀ DASHBOARD - GIỮ GIAO DIỆN GỐC
# ============================================================
st.title("📊 员工考勤与统计仪表盘")
st.markdown("### Dashboard Thống Kê & Phân Tích Chấm Công Nhân Sự")


# ============================================================
# CÁC HÀM TIỆN ÍCH
# ============================================================
def normalize_text(value):
    """
    Chuẩn hóa text để Python có thể kiểm tra hỗ trợ,
    nhưng quyết định cuối cùng về ý nghĩa cột do Gemini thực hiện.
    """
    if value is None:
        return ""

    text = str(value).strip().lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(
        ch for ch in text
        if unicodedata.category(ch) != "Mn"
    )
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def dataframe_to_sample(df, max_rows=20, max_cols=60):
    """
    Lấy mẫu nhỏ gửi Gemini:
    - tên cột
    - tối đa 20 dòng dữ liệu
    Không gửi toàn bộ file để tiết kiệm token.
    """
    sample = df.iloc[:max_rows, :max_cols].copy()

    # Chuyển NaN/NaT thành chuỗi dễ hiểu
    sample = sample.fillna("")

    rows = []
    for _, row in sample.iterrows():
        rows.append([str(v)[:150] for v in row.tolist()])

    return rows


def get_file_bytes(uploaded_file):
    """
    Đọc bytes mà không làm mất trạng thái file uploader.
    """
    if uploaded_file is None:
        return None

    uploaded_file.seek(0)
    data = uploaded_file.getvalue()
    uploaded_file.seek(0)
    return data


def read_raw_dataframe(file_bytes, filename):
    """
    Đọc Excel/CSV KHÔNG giả định header nằm ở dòng 0.
    Gemini sẽ xác định dòng header thật.
    """
    ext = Path(filename).suffix.lower()

    if ext == ".csv":
        # Thử UTF-8 trước, sau đó các encoding phổ biến.
        last_error = None
        for encoding in ["utf-8-sig", "utf-8", "cp1258", "gb18030"]:
            try:
                return pd.read_csv(
                    io.BytesIO(file_bytes),
                    header=None,
                    encoding=encoding
                )
            except Exception as exc:
                last_error = exc

        raise last_error

    if ext in [".xlsx", ".xls"]:
        return pd.read_excel(
            io.BytesIO(file_bytes),
            header=None
        )

    raise ValueError(f"Định dạng file chưa hỗ trợ: {ext}")


def clean_duplicate_columns(columns):
    """
    Đảm bảo tên cột sau khi AI xác định không bị trùng.
    """
    result = []
    seen = {}

    for col in columns:
        col = str(col).strip()

        if not col:
            col = "Unnamed"

        if col not in seen:
            seen[col] = 0
            result.append(col)
        else:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")

    return result


def safe_json_loads(text):
    """
    Gemini được yêu cầu trả JSON.
    Hàm này vẫn xử lý trường hợp model lỡ bọc JSON bằng ```json.
    """
    text = (text or "").strip()

    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)

    return json.loads(text)


# ============================================================
# GEMINI - NHẬN DIỆN CẤU TRÚC FILE
# ============================================================
COLUMN_SCHEMA = {
    "header_row": 0,
    "columns": [],
    "confidence": 0,
    "warnings": []
}


def build_header_detection_prompt(file_role, raw_df):
    """
    Gemini nhận:
      1. Toàn bộ tên cột nếu đã có
      2. Mẫu 20 dòng đầu
    Gemini xác định:
      - header_row
      - vai trò của từng cột
      - confidence
    """

    # Không giới hạn cứng số cột của file thật,
    # nhưng gửi tối đa 80 cột đầu tiên để tránh prompt quá lớn.
    preview = raw_df.iloc[:20, :80].fillna("")

    matrix = []
    for _, row in preview.iterrows():
        matrix.append([
            str(v)[:180] if v is not None else ""
            for v in row.tolist()
        ])

    prompt = f"""
Bạn là chuyên gia phân tích cấu trúc bảng Excel chấm công nhân sự.

Loại file:
{file_role}

Nhiệm vụ:
1. Xác định dòng tiêu đề (header) thật sự.
2. Xác định ý nghĩa của từng cột.
3. Không được đoán chỉ dựa trên tên cột. Hãy kết hợp tên cột
   với dữ liệu mẫu bên dưới.
4. Hỗ trợ tiếng Việt có dấu/không dấu, tiếng Trung, tiếng Anh,
   viết tắt và tên cột không chuẩn.
5. Nếu một cột không chắc chắn, confidence phải thấp.
6. Không tự tạo cột không tồn tại.

Các vai trò chuẩn cần nhận diện nếu có:
- employee_id: Mã nhân viên
- employee_name: Họ và tên
- date: Ngày chấm công
- check_in: Giờ vào
- check_out: Giờ ra
- department: Bộ phận / phòng ban
- shift: Ca làm
- total_hours: Tổng giờ làm
- status: Trạng thái
- other: Cột khác

QUY TẮC:
- header_row là số thứ tự dòng bắt đầu từ 0 trong dữ liệu thô.
- column_index là số thứ tự cột bắt đầu từ 0.
- column_name là tên tiêu đề thật sự ở dòng header.
- confidence từ 0 đến 100.
- Nếu không xác định được vai trò thì dùng "other".
- Có thể có nhiều cột "other".
- Nếu header có ô trống thì vẫn giữ cột đó với tên "Unnamed".
- Nếu bảng có nhiều dòng tiêu đề, chọn dòng chứa các tên cột nghiệp vụ
  rõ ràng nhất và ghi cảnh báo.
- Nếu có merged-cell hoặc tiêu đề 2 tầng nhưng không thể xác định hoàn toàn,
  hãy chọn dòng thực tế được sử dụng làm header pandas và ghi warnings.

TRẢ VỀ DUY NHẤT JSON THEO CẤU TRÚC:

{{
  "header_row": 0,
  "columns": [
    {{
      "column_index": 0,
      "column_name": "Tên cột",
      "role": "employee_id",
      "confidence": 95,
      "reason": "Giải thích ngắn"
    }}
  ],
  "overall_confidence": 95,
  "warnings": []
}}

DỮ LIỆU THÔ:

{matrix}
"""
    return prompt


@st.cache_data(show_spinner=False, ttl=3600)
def detect_excel_structure_cached(
    file_bytes,
    filename,
    file_role,
    api_key_signature,
    _api_key
):
    """
    Gemini nhận diện cấu trúc.
    api_key_signature dùng để phân biệt cache nhưng không lưu raw API key
    trong cache key.
    """

    raw_df = read_raw_dataframe(file_bytes, filename)

    client = genai.Client(api_key=_api_key)

    prompt = build_header_detection_prompt(file_role, raw_df)

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0,
            response_mime_type="application/json"
        )
    )

    result = safe_json_loads(response.text)

    # Bảo vệ kết quả AI
    if not isinstance(result, dict):
        raise ValueError("Gemini trả về cấu trúc JSON không hợp lệ.")

    header_row = result.get("header_row", 0)

    try:
        header_row = int(header_row)
    except Exception:
        header_row = 0

    header_row = max(0, min(header_row, max(0, len(raw_df) - 1)))

    columns = result.get("columns", [])
    if not isinstance(columns, list):
        columns = []

    # Chuẩn hóa column metadata
    normalized_columns = []

    for item in columns:
        if not isinstance(item, dict):
            continue

        try:
            col_idx = int(item.get("column_index", -1))
        except Exception:
            col_idx = -1

        if col_idx < 0 or col_idx >= raw_df.shape[1]:
            continue

        role = str(item.get("role", "other")).strip()

        try:
            confidence = float(item.get("confidence", 0))
        except Exception:
            confidence = 0

        normalized_columns.append({
            "column_index": col_idx,
            "column_name": str(
                item.get("column_name", raw_df.iloc[header_row, col_idx])
            ),
            "role": role,
            "confidence": max(0, min(100, confidence)),
            "reason": str(item.get("reason", ""))
        })

    result["header_row"] = header_row
    result["columns"] = normalized_columns
    result["overall_confidence"] = max(
        0,
        min(100, float(result.get("overall_confidence", 0)))
    )
    result["warnings"] = result.get("warnings", [])

    return result


def apply_ai_structure(raw_df, detection):
    """
    Sau khi Gemini xác định header:
    Python dùng kết quả đó để tạo DataFrame chuẩn.
    """

    header_row = int(detection["header_row"])

    header_values = raw_df.iloc[header_row].tolist()

    columns = []
    for value in header_values:
        text = "" if pd.isna(value) else str(value).strip()
        columns.append(text if text else "Unnamed")

    columns = clean_duplicate_columns(columns)

    data = raw_df.iloc[header_row + 1:].copy()
    data.columns = columns
    data = data.reset_index(drop=True)

    # Xóa các dòng hoàn toàn trống
    data = data.dropna(how="all").reset_index(drop=True)

    # Loại bỏ cột Unnamed hoàn toàn rỗng
    empty_cols = []
    for col in data.columns:
        if str(col).lower().startswith("unnamed"):
            non_empty = data[col].astype(str).str.strip().replace(
                {"nan": "", "None": ""}
            ).ne("").any()

            if not non_empty:
                empty_cols.append(col)

    if empty_cols:
        data = data.drop(columns=empty_cols)

    return data


def role_column_map(detection):
    """
    Chuyển kết quả Gemini thành:
    role -> tên cột
    Chỉ nhận các cột có confidence tương đối cao.
    """
    mapping = {}

    priority = {
        "employee_id": 100,
        "employee_name": 95,
        "date": 100,
        "check_in": 100,
        "check_out": 100,
        "department": 80,
        "shift": 75,
        "total_hours": 70,
        "status": 60,
        "other": 0
    }

    for item in detection.get("columns", []):
        role = item.get("role", "other")
        confidence = float(item.get("confidence", 0))

        if role == "other":
            continue

        if confidence < 65:
            continue

        if role not in mapping:
            mapping[role] = item
        else:
            old_conf = float(mapping[role].get("confidence", 0))

            if confidence > old_conf:
                mapping[role] = item
            elif confidence == old_conf:
                if priority.get(role, 0) > 0:
                    mapping[role] = item

    return {
        role: item["column_name"]
        for role, item in mapping.items()
    }


def find_column_by_role(df, detection, role):
    """
    Lấy tên cột theo vai trò Gemini đã nhận diện.
    Không fallback bằng từ khóa ở đây để tránh quay lại cơ chế cũ.
    """

    mapping = role_column_map(detection)
    column_name = mapping.get(role)

    if column_name in df.columns:
        return column_name

    # Một số Excel có tên header bị thay đổi nhẹ khi pandas đọc.
    target = normalize_text(column_name) if column_name else ""

    if target:
        for col in df.columns:
            if normalize_text(col) == target:
                return col

    return None


def display_detection_result(title, detection):
    """
    Hiển thị compact để người dùng kiểm tra.
    """
    confidence = float(detection.get("overall_confidence", 0))

    if confidence >= 90:
        badge = "🟢"
    elif confidence >= 70:
        badge = "🟡"
    else:
        badge = "🔴"

    st.markdown(
        f"**{title}** — Header dòng `{int(detection['header_row']) + 1}` "
        f"{badge} Confidence: **{confidence:.0f}%**"
    )

    rows = []

    for item in detection.get("columns", []):
        rows.append({
            "Cột": item.get("column_name", ""),
            "Vai trò": item.get("role", "other"),
            "Độ tin cậy": f"{float(item.get('confidence', 0)):.0f}%",
            "Lý do": item.get("reason", "")
        })

    if rows:
        st.dataframe(
            pd.DataFrame(rows),
            use_container_width=True,
            hide_index=True
        )

    warnings = detection.get("warnings", [])
    if warnings:
        for warning in warnings:
            st.warning(f"⚠️ {warning}")


# ============================================================
# HÀM XỬ LÝ FILE THÔNG QUA GEMINI
# ============================================================
def process_uploaded_file(uploaded_file, file_role):
    """
    Trả về:
      df
      detection
    """
    if uploaded_file is None:
        return None, None

    if not api_key:
        return None, None

    file_bytes = get_file_bytes(uploaded_file)

    api_key_signature = hashlib.sha256(
        api_key.encode("utf-8")
    ).hexdigest()[:16]

    detection = detect_excel_structure_cached(
        file_bytes=file_bytes,
        filename=uploaded_file.name,
        file_role=file_role,
        api_key_signature=api_key_signature,
        _api_key=api_key
    )

    raw_df = read_raw_dataframe(
        file_bytes=file_bytes,
        filename=uploaded_file.name
    )

    df = apply_ai_structure(
        raw_df=raw_df,
        detection=detection
    )

    return df, detection


# ============================================================
# PHẦN 1: TẢI FILE
# ============================================================
st.subheader("📁 1. 数据文件上传 / Tải lên tệp dữ liệu")

col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader(
        "1) 上传指纹打卡 Excel 文件 / Tải file Excel bấm vân tay",
        type=["xlsx", "xls", "csv"],
        key="fingerprint_file"
    )

    uploaded_schedule = st.file_uploader(
        "2) 上传排班表文件 (仅限工人) / Tải file lịch xếp ca (cho công nhân)",
        type=["xlsx", "xls", "csv"],
        key="schedule_file"
    )

with col2:
    uploaded_staff_vp = st.file_uploader(
        "3) 上传办公室员工名单 (NV) / Tải file danh sách nhân viên VP",
        type=["xlsx", "xls", "csv"],
        key="staff_vp_file"
    )

    uploaded_staff_cn = st.file_uploader(
        "4) 上传工人名单 (CN) / Tải file danh sách công nhân",
        type=["xlsx", "xls", "csv"],
        key="staff_cn_file"
    )


# ============================================================
# AI STRUCTURE DETECTION
# ============================================================
if api_key:
    try:
        with st.spinner(
            "🤖 Gemini đang đọc cấu trúc các file và xác định tiêu đề cột..."
        ):
            df_fp, det_fp = process_uploaded_file(
                uploaded_fingerprint,
                "File bấm vân tay / dữ liệu chấm công"
            )

            df_sc, det_sc = process_uploaded_file(
                uploaded_schedule,
                "File lịch xếp ca công nhân"
            )

            df_vp, det_vp = process_uploaded_file(
                uploaded_staff_vp,
                "Danh sách nhân viên văn phòng"
            )

            df_cn, det_cn = process_uploaded_file(
                uploaded_staff_cn,
                "Danh sách công nhân"
            )

        if any(x is not None for x in [det_fp, det_sc, det_vp, det_cn]):
            with st.expander(
                "🔍 AI nhận diện cấu trúc file / Kiểm tra tiêu đề cột",
                expanded=False
            ):
                if det_fp:
                    display_detection_result(
                        "1. File bấm vân tay",
                        det_fp
                    )

                if det_sc:
                    display_detection_result(
                        "2. File lịch xếp ca",
                        det_sc
                    )

                if det_vp:
                    display_detection_result(
                        "3. Danh sách NV Văn Phòng",
                        det_vp
                    )

                if det_cn:
                    display_detection_result(
                        "4. Danh sách Công Nhân",
                        det_cn
                    )

    except Exception as e:
        st.error(
            f"❌ Gemini không thể nhận diện cấu trúc file: {e}"
        )

        # Không để app sập hoàn toàn.
        # Cho phép người dùng biết cần kiểm tra API/file.
        df_fp = None
        df_sc = None
        df_vp = None
        df_cn = None
        det_fp = None
        det_sc = None
        det_vp = None
        det_cn = None

else:
    df_fp = None
    df_sc = None
    df_vp = None
    df_cn = None
    det_fp = None
    det_sc = None
    det_vp = None
    det_cn = None

    if any([
        uploaded_fingerprint is not None,
        uploaded_schedule is not None,
        uploaded_staff_vp is not None,
        uploaded_staff_cn is not None
    ]):
        st.info(
            "🔑 Vui lòng nhập Gemini API Key ở thanh bên để AI nhận diện "
            "cấu trúc và tiêu đề cột."
        )


# ============================================================
# PHẦN 2: CHỌN NGÀY VÀ PHÂN TÍCH THEO NGÀY
# ============================================================
st.subheader("📅 2. 选择考勤核对日期 / Chọn ngày kiểm tra từ file vân tay")

selected_date_str = ""
df_fp_filtered = None

col_vao = None
col_ra = None
col_dept = None
col_id = None
col_name = None
col_total = None

if df_fp is not None and det_fp is not None:

    # --------------------------------------------------------
    # Python dùng kết quả Gemini để lấy đúng cột.
    # Không tự dò "ngày/date/vao/ra" bằng từ khóa nữa.
    # --------------------------------------------------------
    date_col = find_column_by_role(
        df_fp,
        det_fp,
        "date"
    )

    col_vao = find_column_by_role(
        df_fp,
        det_fp,
        "check_in"
    )

    col_ra = find_column_by_role(
        df_fp,
        det_fp,
        "check_out"
    )

    col_dept = find_column_by_role(
        df_fp,
        det_fp,
        "department"
    )

    col_id = find_column_by_role(
        df_fp,
        det_fp,
        "employee_id"
    )

    col_name = find_column_by_role(
        df_fp,
        det_fp,
        "employee_name"
    )

    col_total = find_column_by_role(
        df_fp,
        det_fp,
        "total_hours"
    )

    if date_col:

        # ----------------------------------------------------
        # Chuyển ngày
        # ----------------------------------------------------
        df_fp["Ngày_Clean"] = pd.to_datetime(
            df_fp[date_col],
            errors="coerce"
        ).dt.date

        available_dates = sorted(
            df_fp["Ngày_Clean"].dropna().unique()
        )

        if available_dates:

            selected_date = st.selectbox(
                "选择文件中的日期 / Chọn ngày có trong file vân tay",
                available_dates
            )

            selected_date_str = str(selected_date)

            df_fp_filtered = df_fp[
                df_fp["Ngày_Clean"] == selected_date
            ].copy()

            # ------------------------------------------------
            # Hàm xác định ô thời gian có dữ liệu
            # ------------------------------------------------
            def has_time(val):
                if pd.isna(val):
                    return False

                if isinstance(val, datetime):
                    return True

                # Time / Timestamp của pandas
                if isinstance(val, (pd.Timestamp,)):
                    return True

                s = str(val).strip().lower()

                empty_values = {
                    "",
                    "nan",
                    "none",
                    "nat",
                    "-",
                    "0:00:00",
                    "00:00:00"
                }

                if s in empty_values:
                    return False

                # Các trường hợp Excel có giá trị số 0
                try:
                    if float(s) == 0:
                        return False
                except Exception:
                    pass

                return True

            df_fp_filtered["Co_Vao"] = (
                df_fp_filtered[col_vao].apply(has_time)
                if col_vao
                else False
            )

            df_fp_filtered["Co_Ra"] = (
                df_fp_filtered[col_ra].apply(has_time)
                if col_ra
                else False
            )

            def classify_attendance(row):
                if row["Co_Vao"] and row["Co_Ra"]:
                    return "Có mặt đủ giờ / 出勤正常"

                elif row["Co_Vao"] and not row["Co_Ra"]:
                    return "Thiếu giờ ra (BR) / 缺下班卡"

                elif not row["Co_Vao"] and row["Co_Ra"]:
                    return "Thiếu giờ vào (BV) / 缺上班卡"

                else:
                    return "Vắng / Không bấm thẻ / 缺勤"

            df_fp_filtered["Trạng Thái"] = (
                df_fp_filtered.apply(
                    classify_attendance,
                    axis=1
                )
            )

            st.success(
                f"✅ 已选择日期 / Đã chọn ngày: **{selected_date_str}**"
            )

        else:
            st.warning(
                "⚠️ Gemini đã nhận diện cột ngày nhưng dữ liệu trong "
                "cột này không chuyển được thành ngày."
            )

    else:
        st.error(
            "❌ Gemini chưa xác định được cột 'Ngày' trong file vân tay. "
            "Hãy kiểm tra phần 'AI nhận diện cấu trúc file'."
        )

elif uploaded_fingerprint is None:
    st.info(
        "ℹ️ Vui lòng tải file Excel bấm vân tay lên ở bước 1."
    )


# ============================================================
# PHẦN 3: THỐNG KÊ BIỂU ĐỒ
# ============================================================
st.subheader(
    "📈 3. 当日考勤数据统计图表 / Biểu đồ thống kê theo ngày chọn"
)

if df_fp_filtered is not None:

    total_day_records = len(df_fp_filtered)

    count_full = (
        df_fp_filtered["Trạng Thái"]
        == "Có mặt đủ giờ / 出勤正常"
    ).sum()

    count_missing = (
        df_fp_filtered["Trạng Thái"].isin([
            "Thiếu giờ ra (BR) / 缺下班卡",
            "Thiếu giờ vào (BV) / 缺上班卡"
        ])
    ).sum()

    count_absent = (
        df_fp_filtered["Trạng Thái"]
        == "Vắng / Không bấm thẻ / 缺勤"
    ).sum()

    count_active = count_full + count_missing

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "当日总人数 / Tổng nhân sự trong ngày",
        f"{total_day_records} 人"
    )

    m2.metric(
        "实际出勤 / Có đi làm trong ngày",
        f"{count_active} 人"
    )

    m3.metric(
        "缺卡人数 / Thiếu giờ (BV/BR)",
        f"{count_missing} 人"
    )

    m4.metric(
        "缺勤人数 / Vắng mặt trong ngày",
        f"{count_absent} 人"
    )

    col_c1, col_c2 = st.columns(2)

    with col_c1:

        st.markdown(
            "#### 考勤状态分布 / Tỷ lệ trạng thái đi làm trong ngày"
        )

        status_df = (
            df_fp_filtered["Trạng Thái"]
            .value_counts()
            .reset_index()
        )

        status_df.columns = [
            "Trạng thái",
            "Số lượng"
        ]

        fig_pie = px.pie(
            status_df,
            values="Số lượng",
            names="Trạng thái",
            hole=0.45
        )

        st.plotly_chart(
            fig_pie,
            use_container_width=True
        )

    with col_c2:

        st.markdown(
            "#### 部门出勤对比 / So sánh đi làm theo bộ phận"
        )

        if col_dept and col_dept in df_fp_filtered.columns:

            dept_series = (
                df_fp_filtered[col_dept]
                .astype(str)
            )

        else:

            dept_series = pd.Series(
                ["Chung"] * len(df_fp_filtered),
                index=df_fp_filtered.index
            )

        df_fp_filtered["Bộ Phận Chuẩn"] = (
            dept_series.apply(
                lambda x:
                "Văn Phòng / 办公室"
                if any(
                    k in normalize_text(x)
                    for k in [
                        "van phong",
                        "vp",
                        "office"
                    ]
                )
                else "Công Nhân / 工人"
            )
        )

        dept_summary = (
            df_fp_filtered
            .groupby(
                ["Bộ Phận Chuẩn", "Trạng Thái"]
            )
            .size()
            .reset_index(name="Số lượng")
        )

        fig_bar = px.bar(
            dept_summary,
            x="Bộ Phận Chuẩn",
            y="Số lượng",
            color="Trạng Thái",
            barmode="group",
            text_auto=True
        )

        st.plotly_chart(
            fig_bar,
            use_container_width=True
        )

else:
    st.info(
        "Vui lòng tải file vân tay và nhập API Key Gemini "
        "để hiển thị biểu đồ thống kê."
    )


# ============================================================
# PHẦN 4: ĐỐI CHIẾU THÔNG MINH
# ============================================================
st.subheader(
    "📋 4. 智能数据核对与分析 / Phân tích & Đối chiếu chuyên sâu"
)

if st.button(
    "🚀 开始考勤核对分析 / Chạy phân tích chấm công",
    type="primary"
):

    if not api_key:

        st.warning(
            "请输入 API Key 以继续 / Vui lòng nhập API Key ở thanh bên."
        )

    elif (
        df_fp_filtered is None
        or df_vp is None
        or df_cn is None
    ):

        st.warning(
            "请完整上传文件并选择有效日期 / "
            "Vui lòng tải đủ file và chọn ngày hợp lệ."
        )

    else:

        with st.spinner(
            "系统正在智能核对排班与考勤数据，请稍候..."
        ):

            try:

                client = genai.Client(
                    api_key=api_key
                )

                # ------------------------------------------------
                # Chỉ gửi dữ liệu cần thiết cho Gemini.
                # Việc nhận diện cột đã được thực hiện trước đó.
                # ------------------------------------------------
                fp_data = df_fp_filtered.to_string(
                    index=False
                )

                vp_data = (
                    df_vp.to_string(index=False)
                    if df_vp is not None
                    else "Không có"
                )

                cn_data = (
                    df_cn.to_string(index=False)
                    if df_cn is not None
                    else "Không có"
                )

                sc_data = (
                    df_sc.to_string(index=False)
                    if df_sc is not None
                    else "Không có"
                )

                prompt = f"""
Bạn là hệ thống đối chiếu nhân sự và chấm công tự động thông minh.

Hãy thực hiện đối chiếu và phân tích dữ liệu CHO ĐÚNG NGÀY:
{selected_date_str}

CÁC CỘT ĐÃ ĐƯỢC PYTHON + GEMINI NHẬN DIỆN:

FILE BẤM VÂN TAY:
- Mã nhân viên: {col_id}
- Họ tên: {col_name}
- Ngày: {find_column_by_role(df_fp, det_fp, "date") if det_fp else None}
- Giờ vào: {col_vao}
- Giờ ra: {col_ra}
- Bộ phận: {col_dept}
- Tổng giờ: {col_total}

DỮ LIỆU ĐẦU VÀO:

1. File bấm vân tay:
{fp_data}

2. Danh sách NV Văn Phòng:
{vp_data}

3. Danh sách Công Nhân:
{cn_data}

4. Lịch xếp ca CN:
{sc_data}

QUY TẮC:
- VP chuẩn 8h-17h, Chủ Nhật nghỉ.
- Mã 575 (7h-15h), 749 & 949 (7h-19h).
- CN theo lịch ca N (7h-19h) hoặc Đ (19h-7h hôm sau).
- Thiếu vào = "BV".
- Thiếu ra = "BR".
- Thiếu cả = "Vắng".
- TRỌNG TÂM: Chỉ đưa ra các trường hợp bất thường:
  "đi trễ", "về sớm", "làm không đúng lịch",
  "vắng", "BV", "BR", "Lễ".

YÊU CẦU:
Chỉ đưa ra các trường hợp bất thường.

TABLE MARKDOWN:
Mã NV | Họ và Tên | Giờ vào | Giờ ra | Giờ làm thực tế | Ghi chú
"""

                max_retries = 3
                response = None

                for attempt in range(max_retries):

                    try:

                        response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=prompt,
                            config=types.GenerateContentConfig(
                                temperature=0
                            )
                        )

                        break

                    except Exception as err:

                        if (
                            "503" in str(err)
                            and attempt < max_retries - 1
                        ):

                            time.sleep(3)

                            continue

                        raise err

                st.session_state["analysis_result"] = (
                    response.text
                    if response
                    else "Không nhận được kết quả từ Gemini."
                )

                st.session_state["result_date"] = (
                    selected_date_str
                )

                st.success(
                    "✅ 考勤核对完成！ / Hoàn tất đối chiếu chấm công!"
                )

            except Exception as e:

                st.error(
                    f"处理出错 / Lỗi xử lý: {e}"
                )


if "analysis_result" in st.session_state:

    res_date = st.session_state.get(
        "result_date",
        selected_date_str
    )

    st.markdown(
        f"### 📋 异常考勤与核对结果 ({res_date})"
    )

    st.markdown(
        st.session_state["analysis_result"]
    )


# ============================================================
# PHẦN 5: XUẤT BÁO CÁO EXCEL VÀ IN PDF
# ============================================================
st.subheader(
    "💾 5. 导出报告 / Xuất báo cáo"
)

col_dl1, col_dl2 = st.columns(2)


with col_dl1:

    if df_fp_filtered is not None:

        output = io.BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            df_export = pd.DataFrame({
                "Mã NV":
                    df_fp_filtered[col_id]
                    if col_id and col_id in df_fp_filtered.columns
                    else "",

                "Họ và Tên":
                    df_fp_filtered[col_name]
                    if col_name and col_name in df_fp_filtered.columns
                    else "",

                "Giờ vào":
                    df_fp_filtered[col_vao]
                    if col_vao and col_vao in df_fp_filtered.columns
                    else "",

                "Giờ ra":
                    df_fp_filtered[col_ra]
                    if col_ra and col_ra in df_fp_filtered.columns
                    else "",

                "Giờ làm thực tế":
                    df_fp_filtered[col_total]
                    if col_total and col_total in df_fp_filtered.columns
                    else "",

                "Ghi chú":
                    df_fp_filtered["Trạng Thái"]
            })

            sheet_name = (
                f"ChamCong_{selected_date_str}"
                if selected_date_str
                else "ChamCong"
            )

            # Excel sheet tối đa 31 ký tự
            sheet_name = sheet_name[:31]

            df_export.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name
            )

        excel_data = output.getvalue()

        st.download_button(
            label="📥 下载 Excel 报告 / Tải xuống file Excel",
            data=excel_data,
            file_name=(
                f"Attendance_Report_{selected_date_str}.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    else:

        st.info(
            "Vui lòng chọn ngày để tải báo cáo Excel."
        )


with col_dl2:

    st.markdown(
        "#### 🖨️ 导出 PDF 界面 / Lưu trang thành PDF giao diện"
    )

    pdf_button_html = """
    <button onclick="window.print()" style="
        background-color: #ff4b4b;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        font-size: 16px;
        font-weight: bold;
        width: 100%;
    ">
        📄 打印 / 保存为 PDF (Lưu y chang giao diện Dashboard)
    </button>
    """

    components.html(
        pdf_button_html,
        height=50
    )
'''

path = "/mnt/data/attendance_dashboard_gemini_smart_columns.py"
Path(path).write_text(code, encoding="utf-8")

# Tạo requirements tương ứng để dễ triển khai Streamlit Cloud
requirements = """streamlit
pandas
openpyxl
xlrd
plotly
google-genai
"""
req_path = "/mnt/data/requirements.txt"
Path(req_path).write_text(requirements, encoding="utf-8")

print(f"Đã tạo: {path}")
print(f"Đã tạo: {req_path}")
