from pathlib import Path
import zipfile, textwrap, os

root = Path("/mnt/data/attendance_streamlit_project")
if root.exists():
    import shutil
    shutil.rmtree(root)
(root / "modules").mkdir(parents=True)
(root / ".streamlit").mkdir(parents=True)

files = {}

files["app.py"] = r'''
import streamlit as st
from datetime import date
from io import BytesIO

from modules.excel_reader import read_attendance, read_shift_schedule, read_employee_list
from modules.attendance_engine import build_daily_report
from modules.gemini_analysis import analyze_anomalies
from modules.excel_export import export_excel
from modules.pdf_export import export_pdf
from modules.i18n import T

st.set_page_config(
    page_title="员工考勤统计 / Thống kê chấm công",
    page_icon="📊",
    layout="wide",
)

# ---------- CSS ----------
st.markdown("""
<style>
.main-title {font-size: 30px; font-weight: 800; margin-bottom: 0;}
.sub-title {font-size: 14px; opacity: .75; margin-bottom: 20px;}
.kpi {padding: 18px; border-radius: 14px; border: 1px solid rgba(128,128,128,.25);
      background: rgba(128,128,128,.06); text-align:center;}
.kpi-value {font-size: 28px; font-weight:800;}
.kpi-label {font-size: 13px; opacity:.75;}
.badge {padding:4px 9px; border-radius:9px; font-weight:700;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">员工考勤统计 / THỐNG KÊ CHẤM CÔNG</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">双语管理面板 · Bảng điều khiển song ngữ</div>', unsafe_allow_html=True)

# ---------- Sidebar ----------
with st.sidebar:
    st.header("⚙️ 设置 / Cài đặt")
    selected_date = st.date_input(
        "日期 / Ngày kiểm tra",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    st.divider()
    st.caption("系统规则 / Quy tắc hệ thống")
    st.write("VP: 08:00 → 17:00")
    st.write("N: 07:00 → 19:00")
    st.write("Đ: 19:00 → 07:00 ngày hôm sau")
    st.write("575: 07:00 → 15:00")
    st.write("749/949: 07:00 → 19:00")

# ---------- Uploads ----------
c1, c2 = st.columns(2)
with c1:
    attendance_file = st.file_uploader(
        "📁 上传考勤/指纹文件 / Tải file chấm công - vân tay",
        type=["xlsx", "xls", "csv"],
        key="attendance",
    )
    shift_file = st.file_uploader(
        "📁 上传排班表 / Tải lịch xếp ca",
        type=["xlsx", "xls", "csv"],
        key="shift",
    )
with c2:
    office_file = st.file_uploader(
        "📁 上传办公室员工名单 / Tải DS nhân viên VP",
        type=["xlsx", "xls", "csv"],
        key="office",
    )
    worker_file = st.file_uploader(
        "📁 上传工人名单 / Tải DS công nhân",
        type=["xlsx", "xls", "csv"],
        key="worker",
    )

if not attendance_file:
    st.info("请先上传考勤/指纹文件。 / Vui lòng tải file chấm công trước.")
    st.stop()

# ---------- Read ----------
try:
    attendance_df, attendance_meta = read_attendance(attendance_file)
except Exception as e:
    st.error(f"Không đọc được file chấm công: {e}")
    st.stop()

shift_df = None
office_df = None
worker_df = None

if shift_file:
    try:
        shift_df, _ = read_shift_schedule(shift_file)
    except Exception as e:
        st.warning(f"Không đọc được lịch xếp ca: {e}")

if office_file:
    try:
        office_df, _ = read_employee_list(office_file)
    except Exception as e:
        st.warning(f"Không đọc được DS nhân viên VP: {e}")

if worker_file:
    try:
        worker_df, _ = read_employee_list(worker_file)
    except Exception as e:
        st.warning(f"Không đọc được DS công nhân: {e}")

st.success(
    f"考勤记录 / Dữ liệu chấm công: {len(attendance_df):,} dòng"
)

if st.button("🔎 开始检查 / KIỂM TRA", type="primary", use_container_width=True):
    with st.spinner("正在计算…… / Đang xử lý…"):
        report = build_daily_report(
            attendance_df=attendance_df,
            shift_df=shift_df,
            office_df=office_df,
            worker_df=worker_df,
            target_date=selected_date,
        )
        st.session_state["report"] = report

if "report" not in st.session_state:
    st.stop()

report = st.session_state["report"]

# ---------- KPIs ----------
total = len(report)
present = int((report["status_group"] == "present").sum())
absent = int((report["status_group"] == "absent").sum())
abnormal = int((report["status_group"] == "abnormal").sum())

k1, k2, k3, k4 = st.columns(4)
for col, value, label in [
    (k1, total, "员工总数 / Tổng NV"),
    (k2, present, "出勤 / Có mặt"),
    (k3, absent, "缺勤 / Vắng"),
    (k4, abnormal, "异常 / Bất thường"),
]:
    col.markdown(
        f'<div class="kpi"><div class="kpi-value">{value:,}</div>'
        f'<div class="kpi-label">{label}</div></div>',
        unsafe_allow_html=True,
    )

st.divider()

# ---------- Charts ----------
chart_col1, chart_col2 = st.columns(2)
with chart_col1:
    st.subheader("出勤状态 / Trạng thái")
    status_counts = report["note"].replace("", "Bình thường").value_counts()
    st.bar_chart(status_counts)

with chart_col2:
    st.subheader("部门类型 / Loại nhân viên")
    type_counts = report["employee_type"].value_counts()
    st.bar_chart(type_counts)

# ---------- Important anomalies ----------
st.subheader("重点异常 / BẤT THƯỜNG TRỌNG TÂM")
anomalies = report[report["status_group"] == "abnormal"].copy()

if len(anomalies):
    st.dataframe(
        anomalies[["employee_id", "name", "actual_in", "actual_out", "actual_hours", "note"]]
        .rename(columns={
            "employee_id": "Mã NV",
            "name": "Họ và tên",
            "actual_in": "Giờ vào",
            "actual_out": "Giờ ra",
            "actual_hours": "Giờ làm thực tế",
            "note": "Ghi chú",
        }),
        use_container_width=True,
        hide_index=True,
    )

    if st.button("📝 生成重点分析 / PHÂN TÍCH TRỌNG TÂM"):
        with st.spinner("正在生成分析…… / Đang phân tích…"):
            analysis = analyze_anomalies(anomalies, selected_date)
            st.session_state["analysis"] = analysis
else:
    st.success("未发现重点异常。 / Không phát hiện bất thường trọng tâm.")

if st.session_state.get("analysis"):
    st.info(st.session_state["analysis"])

# ---------- Full result ----------
st.subheader("详细结果 / KẾT QUẢ CHI TIẾT")
display_df = report[
    ["employee_id", "name", "actual_in", "actual_out", "actual_hours", "note"]
].copy()
display_df.columns = ["Mã NV", "Họ và tên", "Giờ vào", "Giờ ra", "Giờ làm thực tế", "Ghi chú"]
st.dataframe(display_df, use_container_width=True, hide_index=True)

# ---------- Downloads ----------
excel_bytes = export_excel(report, selected_date)
pdf_bytes = export_pdf(
    report=report,
    target_date=selected_date,
    analysis=st.session_state.get("analysis", ""),
)

d1, d2 = st.columns(2)
with d1:
    st.download_button(
        "⬇️ 下载 Excel / TẢI EXCEL",
        data=excel_bytes,
        file_name=f"thong_ke_cham_cong_{selected_date:%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
with d2:
    st.download_button(
        "⬇️ 下载 PDF / TẢI PDF",
        data=pdf_bytes,
        file_name=f"thong_ke_cham_cong_{selected_date:%Y%m%d}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )

with st.expander("🔧 Kiểm tra cấu trúc dữ liệu"):
    st.write("Các cột nhận diện trong file chấm công:", attendance_meta)
    st.write("5 dòng đầu:")
    st.dataframe(attendance_df.head(), use_container_width=True)
'''

files["modules/i18n.py"] = r'''
T = {
    "title": "员工考勤统计 / THỐNG KÊ CHẤM CÔNG",
}
'''

files["modules/excel_reader.py"] = r'''
import pandas as pd
import re
from io import BytesIO

def _read(file):
    name = getattr(file, "name", "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def _norm(x):
    return re.sub(r"[\s_\-:/()（）]+", "", str(x).strip().lower())

def _find_col(df, candidates, required=True):
    normalized = {_norm(c): c for c in df.columns}
    for candidate in candidates:
        nc = _norm(candidate)
        if nc in normalized:
            return normalized[nc]
    # partial match
    for c in df.columns:
        nc = _norm(c)
        if any(_norm(x) in nc or nc in _norm(x) for x in candidates):
            return c
    if required:
        raise ValueError(
            f"Thiếu cột. Cần một trong: {candidates}. "
            f"Các cột hiện có: {list(df.columns)}"
        )
    return None

def read_attendance(file):
    df = _read(file)
    df.columns = [str(c).strip() for c in df.columns]

    id_col = _find_col(df, ["Mã NV", "Ma NV", "Mã nhân viên", "Employee ID", "ID", "工号", "员工编号"])
    name_col = _find_col(df, ["Họ và tên", "Ho va ten", "Tên", "Name", "姓名"], required=False)

    # Supported layouts:
    # A) one row per punch: ID + date + time
    # B) one row per employee/day: ID + date + in + out
    date_col = _find_col(df, ["Ngày", "Date", "日期", "Ngày chấm công"], required=False)
    time_col = _find_col(df, ["Giờ", "Time", "Thời gian", "Datetime", "DateTime", "时间"], required=False)
    in_col = _find_col(df, ["Giờ vào", "Vào", "Check In", "In", "上班", "打卡上班"], required=False)
    out_col = _find_col(df, ["Giờ ra", "Ra", "Check Out", "Out", "下班", "打卡下班"], required=False)
    weekday_col = _find_col(df, ["Thứ", "Thu", "Weekday", "星期"], required=False)

    out = pd.DataFrame()
    out["employee_id"] = df[id_col].astype(str).str.strip()
    out["name"] = df[name_col].astype(str).str.strip() if name_col else ""

    if in_col or out_col:
        if not date_col:
            # Try extracting date from in/out values
            source = df[in_col] if in_col else df[out_col]
            dt = pd.to_datetime(source, errors="coerce", dayfirst=True)
            out["date"] = dt.dt.date
        else:
            out["date"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date

        out["in"] = pd.to_datetime(df[in_col], errors="coerce", dayfirst=True) if in_col else pd.NaT
        out["out"] = pd.to_datetime(df[out_col], errors="coerce", dayfirst=True) if out_col else pd.NaT
    else:
        if not date_col or not time_col:
            raise ValueError(
                "File chấm công cần dạng ID + Ngày + Giờ, hoặc ID + Ngày + Giờ vào + Giờ ra."
            )
        dt = pd.to_datetime(
            df[date_col].astype(str) + " " + df[time_col].astype(str),
            errors="coerce",
            dayfirst=True,
        )
        out["datetime"] = dt
        out["date"] = dt.dt.date
        out["in"] = pd.NaT
        out["out"] = pd.NaT

    out["weekday_source"] = df[weekday_col].astype(str) if weekday_col else ""
    out = out.dropna(subset=["employee_id"])
    return out, {
        "id": id_col, "name": name_col, "date": date_col,
        "time": time_col, "in": in_col, "out": out_col, "weekday": weekday_col
    }

def read_shift_schedule(file):
    df = _read(file)
    df.columns = [str(c).strip() for c in df.columns]
    id_col = _find_col(df, ["Mã NV", "Ma NV", "Mã nhân viên", "Employee ID", "ID", "工号"])
    date_col = _find_col(df, ["Ngày", "Date", "日期"])
    shift_col = _find_col(df, ["Ca", "Shift", "排班", "班次"])
    name_col = _find_col(df, ["Họ và tên", "Tên", "Name", "姓名"], required=False)

    out = pd.DataFrame()
    out["employee_id"] = df[id_col].astype(str).str.strip()
    out["date"] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date
    out["shift"] = df[shift_col].astype(str).str.strip().str.upper()
    out["name"] = df[name_col].astype(str).str.strip() if name_col else ""
    return out, {"id": id_col, "date": date_col, "shift": shift_col, "name": name_col}

def read_employee_list(file):
    df = _read(file)
    df.columns = [str(c).strip() for c in df.columns]
    id_col = _find_col(df, ["Mã NV", "Ma NV", "Mã nhân viên", "Employee ID", "ID", "工号", "员工编号"])
    name_col = _find_col(df, ["Họ và tên", "Ho va ten", "Tên", "Name", "姓名"], required=False)
    out = pd.DataFrame()
    out["employee_id"] = df[id_col].astype(str).str.strip()
    out["name"] = df[name_col].astype(str).str.strip() if name_col else ""
    return out, {"id": id_col, "name": name_col}
'''

files["modules/attendance_engine.py"] = r'''
import pandas as pd
from datetime import datetime, date, time, timedelta

SPECIAL_RULES = {
    "575": (time(7, 0), time(15, 0)),
    "749": (time(7, 0), time(19, 0)),
    "949": (time(7, 0), time(19, 0)),
}

def _combine(d, t):
    if pd.isna(d) or pd.isna(t):
        return None
    if isinstance(t, pd.Timestamp):
        t = t.to_pydatetime().time()
    elif isinstance(t, datetime):
        t = t.time()
    elif isinstance(t, time):
        pass
    else:
        parsed = pd.to_datetime(t, errors="coerce")
        if pd.isna(parsed):
            return None
        t = parsed.to_pydatetime().time()
    return datetime.combine(d, t)

def _fmt(dt):
    return "" if dt is None or pd.isna(dt) else dt.strftime("%d/%m/%Y %H:%M")

def _hours(a, b):
    if not a or not b:
        return None
    return round((b-a).total_seconds()/3600, 2)

def _nearest_pair(punches, start, end):
    if not punches:
        return None, None
    punches = sorted(punches)
    ins = [x for x in punches if x >= start - timedelta(hours=2) and x <= start + timedelta(hours=4)]
    outs = [x for x in punches if x >= end - timedelta(hours=4) and x <= end + timedelta(hours=4)]
    actual_in = min(ins, key=lambda x: abs((x-start).total_seconds())) if ins else None
    actual_out = min(outs, key=lambda x: abs((x-end).total_seconds())) if outs else None
    return actual_in, actual_out

def _make_punch_map(attendance_df):
    # Converts both supported attendance layouts to {employee_id: [datetimes]}
    m = {}
    for _, r in attendance_df.iterrows():
        eid = str(r["employee_id"]).strip()
        m.setdefault(eid, [])
        for col in ["datetime", "in", "out"]:
            if col in r.index and not pd.isna(r[col]):
                val = r[col]
                if isinstance(val, pd.Timestamp):
                    val = val.to_pydatetime()
                elif isinstance(val, datetime):
                    pass
                else:
                    val = pd.to_datetime(val, errors="coerce", dayfirst=True)
                    if pd.isna(val):
                        continue
                    val = val.to_pydatetime()
                m[eid].append(val)
    for eid in m:
        m[eid] = sorted(set(m[eid]))
    return m

def _shift_for(shift_df, eid, target):
    if shift_df is None or shift_df.empty:
        return None
    x = shift_df[(shift_df.employee_id == eid) & (shift_df.date == target)]
    if x.empty:
        return None
    return str(x.iloc[0]["shift"]).strip().upper()

def _employee_dict(df):
    if df is None or df.empty:
        return {}
    return {
        str(r.employee_id).strip(): str(r.name).strip()
        for _, r in df.iterrows()
    }

def _candidate_workers(worker_df, shift_df):
    ids = set()
    if worker_df is not None:
        ids.update(worker_df.employee_id.astype(str).str.strip())
    if shift_df is not None:
        ids.update(shift_df.employee_id.astype(str).str.strip())
    return ids

def build_daily_report(attendance_df, shift_df, office_df, worker_df, target_date):
    target = target_date if isinstance(target_date, date) else pd.to_datetime(target_date).date()
    punches = _make_punch_map(attendance_df)

    office = _employee_dict(office_df)
    workers = _employee_dict(worker_df)

    # If lists are missing, fall back to IDs found in attendance/shift.
    office_ids = set(office)
    worker_ids = set(workers)
    if not office_ids and not worker_ids:
        worker_ids = _candidate_workers(worker_df, shift_df)
        office_ids = set(attendance_df.employee_id.astype(str).str.strip())

    rows = []

    # -------- VP --------
    for eid in sorted(office_ids):
        name = office.get(eid, "")
        if target.weekday() == 6:
            rows.append(_row(eid, name, "VP", None, None, 0, "Nghỉ", "off"))
            continue

        start = datetime.combine(target, time(8,0))
        end = datetime.combine(target, time(17,0))
        actual_in, actual_out = _nearest_pair(
            punches.get(eid, []), start, end
        )

        # For robust real-world punch data, if nearest_pair fails, derive min/max
        # punches inside the day.
        day_punches = [p for p in punches.get(eid, [])
                       if target - timedelta(hours=2) <= p <= target + timedelta(days=1)]
        if actual_in is None and day_punches:
            same_day = [p for p in day_punches if p.date() == target]
            if same_day:
                actual_in = min(same_day)
        if actual_out is None and day_punches:
            same_day = [p for p in day_punches if p.date() == target]
            if len(same_day) >= 2:
                actual_out = max(same_day)

        note = ""
        group = "present"
        if actual_in is None and actual_out is None:
            note, group = "Vắng", "absent"
        elif actual_in is None:
            note, group = "BV", "abnormal"
        elif actual_out is None:
            note, group = "BR", "abnormal"
        else:
            if actual_in > start:
                note = "Đi trễ"
                group = "abnormal"
            hours = _hours(actual_in, actual_out)
            if hours is not None and hours < 8:
                note = "Về sớm" if not note else note + "; Về sớm"
                group = "abnormal"
            if not note:
                note = "Bình thường"

        rows.append(_row(eid, name, "VP", actual_in, actual_out,
                         _hours(actual_in, actual_out) or 0, note, group))

    # -------- CN --------
    for eid in sorted(worker_ids):
        name = workers.get(eid, "")
        shift = _shift_for(shift_df, eid, target)

        # If no schedule, cannot classify as planned work.
        if not shift:
            rows.append(_row(eid, name, "CN", None, None, 0,
                             "Không có lịch", "abnormal"))
            continue

        if shift in {"OFF", "O", "NGHỈ", "REST", "R", "-", "X"}:
            rows.append(_row(eid, name, "CN", None, None, 0,
                             "Nghỉ", "off"))
            continue

        # N = 07:00-19:00 same day
        # Đ = 19:00 target -> 07:00 next day
        if shift == "N":
            start = datetime.combine(target, time(7,0))
            end = datetime.combine(target, time(19,0))
        elif shift in {"Đ", "D", "DEM", "NIGHT"}:
            start = datetime.combine(target, time(19,0))
            end = datetime.combine(target + timedelta(days=1), time(7,0))
        else:
            # Allow shift strings that contain N or Đ.
            if "Đ" in shift or "D" == shift:
                start = datetime.combine(target, time(19,0))
                end = datetime.combine(target + timedelta(days=1), time(7,0))
            elif "N" in shift:
                start = datetime.combine(target, time(7,0))
                end = datetime.combine(target, time(19,0))
            else:
                rows.append(_row(eid, name, "CN", None, None, 0,
                                 "Lịch không xác định", "abnormal"))
                continue

        actual_in, actual_out = _nearest_pair(punches.get(eid, []), start, end)

        # For night shift, out is on the following day.
        note = ""
        group = "present"
        if actual_in is None and actual_out is None:
            note, group = "Vắng", "absent"
        elif actual_in is None:
            note, group = "BV", "abnormal"
        elif actual_out is None:
            note, group = "BR", "abnormal"
        else:
            hours = _hours(actual_in, actual_out)
            if actual_in > start + timedelta(minutes=1):
                note = "Đi trễ"
                group = "abnormal"
            if hours is not None and hours < 12:
                note = ("Thiếu giờ; " + note).strip("; ") if note else "Thiếu giờ"
                group = "abnormal"
            if actual_in < start - timedelta(hours=4) or actual_in > start + timedelta(hours=6):
                note = "Làm không đúng lịch"
                group = "abnormal"
            if not note:
                note = "Bình thường"

        rows.append(_row(eid, name, "CN", actual_in, actual_out,
                         _hours(actual_in, actual_out) or 0, note, group))

    return pd.DataFrame(rows)

def _row(eid, name, employee_type, actual_in, actual_out, hours, note, group):
    return {
        "employee_id": str(eid),
        "name": "" if name == "nan" else name,
        "actual_in": _fmt(actual_in),
        "actual_out": _fmt(actual_out),
        "actual_hours": round(float(hours or 0), 2),
        "note": note,
        "employee_type": employee_type,
        "status_group": group,
    }
'''

files["modules/gemini_analysis.py"] = r'''
import os
import pandas as pd

def analyze_anomalies(anomalies: pd.DataFrame, target_date):
    if anomalies is None or anomalies.empty:
        return "Không có trường hợp bất thường."

    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return _fallback(anomalies)

    try:
        from google import genai
        client = genai.Client(api_key=api_key)

        records = anomalies[
            ["employee_id", "name", "actual_in", "actual_out", "actual_hours", "note", "employee_type"]
        ].to_dict(orient="records")

        prompt = f"""
Bạn là trợ lý phân tích chấm công.
Ngày kiểm tra: {target_date:%d/%m/%Y}.

Hãy phân tích DANH SÁCH BẤT THƯỜNG bên dưới.
Không tự thay đổi kết quả tính toán.
Chỉ nêu các trường hợp đáng chú ý, ngắn gọn, tối đa 8 dòng.
Mỗi dòng: Mã NV - vấn đề - số liệu chính - kết luận ngắn.
Không dùng từ "AI".
Không suy đoán khi thiếu dữ liệu.

Dữ liệu:
{records}
"""
        response = client.models.generate_content(
            model=os.getenv("GEMINI_MODEL", "gemini-2.5-flash"),
            contents=prompt,
        )
        text = getattr(response, "text", None)
        return text.strip() if text else _fallback(anomalies)
    except Exception as e:
        return _fallback(anomalies) + f"\n\n(Ghi chú hệ thống: chưa gọi được dịch vụ phân tích: {e})"

def _fallback(df):
    lines = []
    for _, r in df.head(8).iterrows():
        lines.append(
            f"- {r['employee_id']} - {r['name']}: {r['note']} "
            f"(Vào: {r['actual_in'] or '-'}, Ra: {r['actual_out'] or '-'})"
        )
    return "\n".join(lines)
'''

files["modules/excel_export.py"] = r'''
from io import BytesIO
import pandas as pd

def export_excel(report, target_date):
    out = report[[
        "employee_id", "name", "actual_in", "actual_out",
        "actual_hours", "note"
    ]].copy()
    out.columns = [
        "Mã NV", "Họ và tên", "Giờ vào", "Giờ ra",
        "Giờ làm thực tế", "Ghi chú"
    ]

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        out.to_excel(
            writer,
            sheet_name=f"{target_date:%d-%m-%Y}",
            index=False
        )
        ws = writer.sheets[f"{target_date:%d-%m-%Y}"]
        ws.freeze_panes = "A2"
        widths = [16, 28, 22, 22, 18, 28]
        for i, width in enumerate(widths, start=1):
            ws.column_dimensions[chr(64+i)].width = width

    return bio.getvalue()
'''

files["modules/pdf_export.py"] = r'''
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pathlib import Path

def _register_font():
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
    ]
    for p in candidates:
        if Path(p).exists():
            pdfmetrics.registerFont(TTFont("AppFont", p))
            return "AppFont"
    return "Helvetica"

def export_pdf(report, target_date, analysis=""):
    font = _register_font()
    buf = BytesIO()

    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        rightMargin=22, leftMargin=22, topMargin=22, bottomMargin=22
    )
    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "Title2", parent=styles["Title"], fontName=font,
        alignment=TA_CENTER, fontSize=18, leading=22
    )
    normal = ParagraphStyle(
        "Normal2", parent=styles["Normal"], fontName=font,
        fontSize=8.5, leading=11
    )

    story = [
        Paragraph("员工考勤统计 / THỐNG KÊ CHẤM CÔNG", title),
        Paragraph(
            f"日期 / Ngày: {target_date:%d/%m/%Y} &nbsp;&nbsp; "
            f"Tổng NV: {len(report)} &nbsp;&nbsp; "
            f"Bất thường: {(report.status_group == 'abnormal').sum()}",
            normal
        ),
        Spacer(1, 10),
    ]

    if analysis:
        story += [
            Paragraph("<b>重点异常 / BẤT THƯỜNG TRỌNG TÂM</b>", normal),
            Paragraph(analysis.replace("\n", "<br/>"), normal),
            Spacer(1, 8),
        ]

    data = [[
        "Mã NV", "Họ và tên", "Giờ vào", "Giờ ra",
        "Giờ làm thực tế", "Ghi chú"
    ]]
    for _, r in report.iterrows():
        data.append([
            str(r["employee_id"]),
            str(r["name"]),
            str(r["actual_in"]),
            str(r["actual_out"]),
            f"{r['actual_hours']:.2f}",
            str(r["note"]),
        ])

    table = Table(data, repeatRows=1, colWidths=[65, 150, 105, 105, 80, 150])
    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), font),
        ("FONTSIZE", (0,0), (-1,-1), 7.5),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#E9EEF5")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.black),
        ("GRID", (0,0), (-1,-1), 0.35, colors.grey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(table)
    doc.build(story)
    return buf.getvalue()
'''

files["requirements.txt"] = r'''
streamlit>=1.40
pandas>=2.2
openpyxl>=3.1
xlrd>=2.0
reportlab>=4.2
google-genai>=1.0
'''

files[".streamlit/secrets.toml.example"] = r'''
# Copy this to Streamlit Cloud Secrets:
GEMINI_API_KEY = "YOUR_GEMINI_API_KEY"
GEMINI_MODEL = "gemini-2.5-flash"
'''

files["README.md"] = r'''
# Attendance Streamlit Dashboard

Dashboard song ngữ Trung / Việt để kiểm tra chấm công theo ngày.

## 1. Chức năng

- Tải file chấm công / vân tay.
- Tải lịch xếp ca công nhân.
- Tải danh sách nhân viên VP.
- Tải danh sách công nhân.
- Chọn ngày kiểm tra.
- VP: 08:00-17:00, Chủ nhật nghỉ.
- CN ca N: 07:00-19:00.
- CN ca Đ: 19:00 ngày kiểm tra đến 07:00 ngày hôm sau.
- 575: 07:00-15:00.
- 749/949: 07:00-19:00.
- BV = thiếu giờ vào.
- BR = thiếu giờ ra.
- Thiếu cả vào và ra trong ngày làm việc dự kiến = Vắng.
- Xuất Excel và PDF.
- Phân tích bất thường bằng Gemini.

## 2. Chạy local

```bash
pip install -r requirements.txt
streamlit run app.py
