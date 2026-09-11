import streamlit as st
import pandas as pd
import numpy as np
import datetime
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# Page configuration
st.set_page_config(
    page_title="Hệ thống Quản lý Chấm công | 考勤管理系统",
    page_icon="⏱️",
    layout="wide"
)

# Language dictionary / Từ điển ngôn ngữ
LANG = {
    "vi": {
        "title": "Hệ thống Quản lý Chấm công & Thống kê Nhân sự",
        "subtitle": "Quản lý giờ vào/ra, thống kê đi trễ, về sớm, vắng mặt và kiểm tra lịch làm việc",
        "upload_section": "1. Tải lên các tệp dữ liệu",
        "fp_file": "Tệp bấm vân tay (Excel/CSV)",
        "shift_file": "Tệp lịch xếp ca công nhân (Excel)",
        "vp_file": "Tệp danh sách văn phòng (Excel)",
        "cn_file": "Tệp danh sách công nhân (Excel)",
        "filter_section": "2. Bộ lọc Ngày / Tháng / Năm",
        "select_date": "Chọn ngày kiểm tra",
        "stats_overview": "3. Tổng quan Thống kê Ngày",
        "total_emp": "Tổng số nhân viên",
        "present": "Có mặt",
        "absent": "Vắng mặt",
        "late": "Đi trễ",
        "early": "Về sớm",
        "table_title": "Chi tiết Chấm công Nhân viên",
        "download_excel": "Tải xuống Báo cáo Excel",
        "download_pdf": "Tải xuống Báo cáo PDF",
        "col_manv": "Mã NV",
        "col_hoten": "Họ và Tên",
        "col_loai": "Loại NV",
        "col_vao": "Giờ Vào",
        "col_ra": "Giờ Ra",
        "col_giolam": "Giờ Làm Thực Tế",
        "col_ghichu": "Ghi Chú",
        "chart_status": "Phân tích Trạng thái Chấm công",
        "chart_hours": "Phân phối Giờ làm thực tế",
        "lang_switch": "Chuyển ngữ / 语言",
        "no_data": "Vui lòng tải lên các tệp dữ liệu cần thiết để bắt đầu.",
        "sample_data_btn": "Sử dụng dữ liệu mẫu (Demo)"
    },
    "zh": {
        "title": "考勤与人事统计管理系统",
        "subtitle": "管理进出勤时间、迟到、早退、缺勤统计及排班核对",
        "upload_section": "1. 上传数据文件",
        "fp_file": "指纹打卡文件 (Excel/CSV)",
        "shift_file": "工人排班表文件 (Excel)",
        "vp_file": "办公室人员名单文件 (Excel)",
        "cn_file": "工人名单文件 (Excel)",
        "filter_section": "2. 日期 / 月份 / 年份 筛选",
        "select_date": "选择检查日期",
        "stats_overview": "3. 当日统计概览",
        "total_emp": "员工总数",
        "present": "出勤",
        "absent": "缺勤",
        "late": "迟到",
        "early": "早退",
        "table_title": "员工考勤详细明细",
        "download_excel": "下载 Excel 报告",
        "download_pdf": "下载 PDF 报告",
        "col_manv": "工号",
        "col_hoten": "姓名",
        "col_loai": "员工类型",
        "col_vao": "上班时间",
        "col_ra": "下班时间",
        "col_giolam": "实际工作小时",
        "col_ghichu": "备注",
        "chart_status": "考勤状态分析",
        "chart_hours": "实际工时分布",
        "lang_switch": "语言切换 / Language",
        "no_data": "请上传必要的数据文件以开始使用。",
        "sample_data_btn": "使用演示示例数据"
    }
}

# Sidebar Language Selection
st.sidebar.title("🌐 Settings / Cài đặt")
lang_choice = st.sidebar.radio("Language / Ngôn ngữ", options=["Tiếng Việt", "中文"], index=0)
current_lang = "vi" if lang_choice == "Tiếng Việt" else "zh"
t = LANG[current_lang]

st.title(f"⏱️ {t['title']}")
st.markdown(f"*{t['subtitle']}*")

# Sidebar Uploads
st.sidebar.header(t['upload_section'])
uploaded_fp = st.sidebar.file_uploader(t['fp_file'], type=["xlsx", "xls", "csv"])
uploaded_shift = st.sidebar.file_uploader(t['shift_file'], type=["xlsx", "xls"])
uploaded_vp = st.sidebar.file_uploader(t['vp_file'], type=["xlsx", "xls"])
uploaded_cn = st.sidebar.file_uploader(t['cn_file'], type=["xlsx", "xls"])

use_sample = st.sidebar.button(t['sample_data_btn'])

# Khởi tạo giá trị mặc định tránh lỗi NameError
df_fp, df_shift, df_vp, df_cn = None, None, None, None

if use_sample:
    try:
        df_fp = pd.read_excel("sample_data/bang_cham_cong.xlsx")
        df_shift = pd.read_excel("sample_data/lich_xep_ca.xlsx")
        df_vp = pd.read_excel("sample_data/danh_sach_van_phong.xlsx")
        df_cn = pd.read_excel("sample_data/danh_sach_cong_nhan.xlsx")
        st.sidebar.success("Đã tải dữ liệu mẫu thành công! / 示例数据加载成功！")
    except Exception as e:
        st.sidebar.error(f"Lỗi tải mẫu (Vui lòng upload file thủ công): {e}")
else:
    if uploaded_fp is not None:
        df_fp = pd.read_excel(uploaded_fp) if uploaded_fp.name.endswith(('.xlsx', '.xls')) else pd.read_csv(uploaded_fp)
    if uploaded_shift is not None:
        df_shift = pd.read_excel(uploaded_shift)
    if uploaded_vp is not None:
        df_vp = pd.read_excel(uploaded_vp)
    if uploaded_cn is not None:
        df_cn = pd.read_excel(uploaded_cn)

if df_fp is not None:
    # Ensure datetime format
    df_fp['ThoiGian'] = pd.to_datetime(df_fp['ThoiGian'])
    df_fp['Ngay'] = df_fp['ThoiGian'].dt.date
    df_fp['MaNV'] = df_fp['MaNV'].astype(str).str.strip()

    # Date Filter
    st.sidebar.header(t['filter_section'])
    available_dates = sorted(df_fp['Ngay'].unique())
    selected_date = st.sidebar.selectbox(t['select_date'], options=available_dates)

    # Process lists
    if df_cn is not None:
        df_cn['MaNV'] = df_cn['MaNV'].astype(str).str.strip()
    if df_vp is not None:
        df_vp['MaNV'] = df_vp['MaNV'].astype(str).str.strip()
    if df_shift is not None:
        df_shift['MaNV'] = df_shift['MaNV'].astype(str).str.strip()
        df_shift['Ngay'] = pd.to_datetime(df_shift['Ngay']).dt.date

    # Combine all employees or create master list
    workers_list = df_cn['MaNV'].tolist() if df_cn is not None else []
    vp_list = df_vp['MaNV'].tolist() if df_vp is not None else []
    special_ids = {"575": {"in": "07:00", "out": "15:00"},
                   "749": {"in": "07:00", "out": "19:00"},
                   "949": {"in": "07:00", "out": "19:00"},
                   "673": {"in": "07:00", "out": "15:00"},
                   "A068": {"in": "10:00", "out": "18:00"}}

    # Get names mapping
    name_map = {}
    if df_cn is not None:
        name_map.update(dict(zip(df_cn['MaNV'], df_cn['HoTen'])))
    if df_vp is not None:
        name_map.update(dict(zip(df_vp['MaNV'], df_vp['HoTen'])))

    # Filter logs for selected date (and next day morning for night shifts)
    day_logs = df_fp[df_fp['Ngay'] == selected_date]

    # Evaluate attendance for all known employees
    all_manvs = set(workers_list + vp_list + list(special_ids.keys()) + df_fp['MaNV'].unique().tolist())

    results = []
    for manv in all_manvs:
        name = name_map.get(manv, f"NV {manv}")
        is_worker = manv in workers_list
        is_vp = manv in vp_list
        is_special = manv in special_ids

        # Determine employee type
        emp_type = "Công Nhân" if is_worker else ("Văn Phòng" if is_vp else "Khác")

        # Get punches for this employee around selected date
        emp_punches = df_fp[(df_fp['MaNV'] == manv) & (df_fp['ThoiGian'].dt.date >= selected_date) & (df_fp['ThoiGian'].dt.date <= selected_date + pd.Timedelta(days=1))]
        emp_punches = emp_punches.sort_values('ThoiGian')

        check_in_str = "-"
        check_out_str = "-"
        actual_hours = 0.0
        notes = []

        if is_special:
            std_in = datetime.datetime.combine(selected_date, datetime.datetime.strptime(special_ids[manv]["in"], "%H:%M").time())
            std_out = datetime.datetime.combine(selected_date, datetime.datetime.strptime(special_ids[manv]["out"], "%H:%M").time())
            
            day_punches = emp_punches[emp_punches['ThoiGian'].dt.date == selected_date]
            if not day_punches.empty:
                in_punch = day_punches[day_punches['ThoiGian'] <= std_in + pd.Timedelta(hours=2)]
                out_punch = day_punches[day_punches['ThoiGian'] >= std_in]

                if not in_punch.empty:
                    ci = in_punch.iloc[0]['ThoiGian']
                    check_in_str = ci.strftime("%H:%M")
                    if ci > std_in + pd.Timedelta(minutes=5):
                        notes.append("Đi trễ")
                else:
                    notes.append("BV")

                if not out_punch.empty:
                    co = out_punch.iloc[-1]['ThoiGian']
                    check_out_str = co.strftime("%H:%M")
                    if co < std_out - pd.Timedelta(minutes=5):
                        notes.append("Về sớm")
                else:
                    notes.append("BR")

                if "BV" in notes and "BR" in notes:
                    notes = ["Vắng"]
                elif not notes:
                    notes = ["Đúng giờ"]
                
                if check_in_str != "-" and check_out_str != "-":
                    ci_dt = datetime.datetime.combine(selected_date, datetime.datetime.strptime(check_in_str, "%H:%M").time())
                    co_dt = datetime.datetime.combine(selected_date, datetime.datetime.strptime(check_out_str, "%H:%M").time())
                    actual_hours = round((co_dt - ci_dt).seconds / 3600.0, 1)

        elif is_vp:
            is_sunday = selected_date.weekday() == 6
            if is_sunday:
                notes.append("Nghỉ Chủ Nhật")
            else:
                std_in = datetime.datetime.combine(selected_date, datetime.time(8, 0))
                std_out = datetime.datetime.combine(selected_date, datetime.time(17, 0))
                day_punches = emp_punches[emp_punches['ThoiGian'].dt.date == selected_date]
                
                if day_punches.empty:
                    notes.append("Vắng")
                else:
                    ci = day_punches.iloc[0]['ThoiGian']
                    co = day_punches.iloc[-1]['ThoiGian'] if len(day_punches) > 1 else None
                    
                    check_in_str = ci.strftime("%H:%M")
                    if ci > std_in + pd.Timedelta(minutes=5):
                        notes.append("Đi trễ")
                    
                    if co:
                        check_out_str = co.strftime("%H:%M")
                        actual_hours = round((co - ci).seconds / 3600.0, 1)
                        if actual_hours < 8.0:
                            notes.append("Về sớm")
                    else:
                        notes.append("BR")
                    
                    if not notes:
                        notes.append("Đúng giờ")

        elif is_worker and df_shift is not None:
            shift_row = df_shift[(df_shift['MaNV'] == manv) & (df_shift['Ngay'] == selected_date)]
            shift_type = shift_row['Ca'].values[0] if not shift_row.empty else "Nghỉ"
            
            if shift_type in ["Nghỉ", "Chủ Nhật"]:
                notes.append("Ngày nghỉ")
            elif shift_type == "Lễ":
                notes.append("Lễ")
            elif shift_type == "N": 
                std_in = datetime.datetime.combine(selected_date, datetime.time(7, 0))
                std_out = datetime.datetime.combine(selected_date, datetime.time(19, 0))
                day_p = emp_punches[emp_punches['ThoiGian'].dt.date == selected_date]
                if day_p.empty:
                    notes.append("Vắng")
                else:
                    ci = day_p.iloc[0]['ThoiGian']
                    co = day_p.iloc[-1]['ThoiGian'] if len(day_p) > 1 else None
                    check_in_str = ci.strftime("%H:%M")
                    if ci > std_in + pd.Timedelta(minutes=10):
                        notes.append("Đi trễ")
                    if co:
                        check_out_str = co.strftime("%H:%M")
                        actual_hours = round((co - ci).seconds / 3600.0, 1)
                        if co < std_out - pd.Timedelta(minutes=10):
                            notes.append("Về sớm")
                    else:
                        notes.append("BR")
                    if not notes:
                        notes.append("Đúng giờ")
            elif shift_type == "Đ": 
                std_in = datetime.datetime.combine(selected_date, datetime.time(19, 0))
                next_date = selected_date + pd.Timedelta(days=1)
                std_out = datetime.datetime.combine(next_date, datetime.time(7, 0))
                
                night_p = emp_punches[(emp_punches['ThoiGian'] >= std_in - pd.Timedelta(hours=2)) & (emp_punches['ThoiGian'] <= std_out + pd.Timedelta(hours=2))]
                if night_p.empty:
                    notes.append("Vắng")
                else:
                    ci = night_p.iloc[0]['ThoiGian']
                    co = night_p.iloc[-1]['ThoiGian'] if len(night_p) > 1 else None
                    check_in_str = ci.strftime("%H:%M")
                    check_out_str = co.strftime("%H:%M") if co else "-"
                    if co:
                        actual_hours = round((co - ci).seconds / 3600.0, 1)
                    if not notes:
                        notes.append("Đúng giờ ca đêm")
        else:
            day_p = emp_punches[emp_punches['ThoiGian'].dt.date == selected_date]
            if not day_p.empty:
                check_in_str = day_p.iloc[0]['ThoiGian'].strftime("%H:%M")
                if len(day_p) > 1:
                    check_out_str = day_p.iloc[-1]['ThoiGian'].strftime("%H:%M")
                    actual_hours = round((day_p.iloc[-1]['ThoiGian'] - day_p.iloc[0]['ThoiGian']).seconds / 3600.0, 1)
                notes.append("Bình thường")
            else:
                notes.append("Vắng")

        results.append({
            t['col_manv']: manv,
            t['col_hoten']: name,
            t['col_loai']: emp_type,
            t['col_vao']: check_in_str,
            t['col_ra']: check_out_str,
            t['col_giolam']: actual_hours,
            t['col_ghichu']: ", ".join(notes) if notes else "Đúng giờ"
        })

    df_result = pd.DataFrame(results)

    # Statistics Overview
    st.header(t['stats_overview'])
    col1, col2, col3, col4, col5 = st.columns(5)
    total_emp_count = len(df_result)
    present_count = len(df_result[~df_result[t['col_ghichu']].str.contains("Vắng|Nghỉ", na=False)])
    absent_count = len(df_result[df_result[t['col_ghichu']].str.contains("Vắng", na=False)])
    late_count = len(df_result[df_result[t['col_ghichu']].str.contains("Đi trễ", na=False)])
    early_count = len(df_result[df_result[t['col_ghichu']].str.contains("Về sớm", na=False)])

    col1.metric(t['total_emp'], total_emp_count)
    col2.metric(t['present'], present_count)
    col3.metric(t['absent'], absent_count)
    col4.metric(t['late'], late_count)
    col5.metric(t['early'], early_count)

    st.markdown("---")

    # Charts
    c1, c2 = st.columns(2)
    with c1:
        st.subheader(t['chart_status'])
        status_counts = df_result[t['col_ghichu']].value_counts().reset_index()
        status_counts.columns = ['Trạng thái', 'Số lượng']
        fig1 = px.pie(status_counts, names='Trạng thái', values='Số lượng', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig1, use_container_width=True)

    with c2:
        st.subheader(t['chart_hours'])
        fig2 = px.histogram(df_result, x=t['col_giolam'], nbins=10, color=t['col_loai'], marginal="box", color_discrete_sequence=px.colors.qualitative.Set2)
        st.plotly_chart(fig2, use_container_width=True)

    # Table
    st.subheader(t['table_title'])
    st.dataframe(df_result, use_container_width=True)

    # Download buttons
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='ChamCong')
        excel_data = output.getvalue()
        st.download_button(
            label=f"📥 {t['download_excel']}",
            data=excel_data,
            file_name=f"cham_cong_{selected_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col_dl2:
        html_report = f"""
        <html>
        <head><meta charset="utf-8"><style>
            body {{ font-family: DejaVu Sans, Arial, sans-serif; margin: 20px; color: #333; }}
            h1 {{ color: #2c3e50; text-align: center; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 12px; }}
            th {{ background-color: #f2f2f2; color: #333; }}
        </style></head>
        <body>
            <h1>Báo Cáo Chấm Công Ngày {selected_date}</h1>
            <p>Tổng NV: {total_emp_count} | Có mặt: {present_count} | Vắng: {absent_count} | Trễ: {late_count} | Sớm: {early_count}</p>
            {df_result.to_html(index=False)}
        </body>
        </html>
        """
        st.download_button(
            label=f"📥 {t['download_pdf']}",
            data=html_report,
            file_name=f"bao_cao_cham_cong_{selected_date}.html",
            mime="text/html"
        )
else:
    st.info(t['no_data'])
