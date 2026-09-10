import streamlit as st
import pandas as pd
import datetime
import json
import io
from google import genai
from google.genai import types
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. CẤU HÌNH & KHỞI TẠO GEMINI CLIENT
# ==========================================
api_key = st.secrets.get("GEMINI_API_KEY", "")
if api_key:
    client = genai.Client(api_key=api_key)
else:
    client = None

# ==========================================
# 2. ĐỊNH NGHĨA NGÔN NGỮ (SONG NGỮ TRUNG - VIỆT)
# ==========================================
LANG = {
    'VI': {
        'title': 'HỆ THỐNG QUẢN LÝ VÀ THỐNG KÊ NHÂN VIÊN ĐI LÀM',
        'upload_section': 'TẢI DỮ LIỆU ĐẦU VÀO',
        'btn_finger': '1. Tải file Vân Tay (.xlsx)',
        'btn_schedule': '2. Tải Lịch Xếp Ca (.xlsx)',
        'btn_office': '3. Tải Danh Sách VP (.xlsx)',
        'btn_worker': '4. Tải Danh Sách Công Nhân (.xlsx)',
        'filter_section': 'BỘ LỌC THỜI GIAN',
        'select_date': 'Chọn ngày kiểm tra',
        'stats_section': 'THỐNG KÊ TRỰC QUAN',
        'report_section': 'BÁO CÁO CHI TIẾT',
        'download_excel': 'Tải file Excel báo cáo',
        'download_pdf': 'Tải file PDF giao diện',
        'total_emp': 'Tổng số NV',
        'present': 'Đi làm',
        'absent': 'Vắng',
        'late_early': 'Trễ/Về sớm',
        'chart_title': 'Biểu đồ tình trạng đi làm theo Nhóm',
        'no_data': 'Vui lòng tải đủ file và cấu hình thông tin.',
    },
    'ZH': {
        'title': '员工出勤考勤管理与统计系统',
        'upload_section': '上传输入 table 数据',
        'btn_finger': '1. 上传指纹打卡文件 (.xlsx)',
        'btn_schedule': '2. 上传排班表文件 (.xlsx)',
        'btn_office': '3. 上传办公室人员名单 (.xlsx)',
        'btn_worker': '4. 上传工人名单 (.xlsx)',
        'filter_section': '时间筛选器',
        'select_date': '选择 non-working / 工作检查日期',
        'stats_section': '直观数据统计',
        'report_section': '详细报表',
        'download_excel': '下载 Excel 报表',
        'download_pdf': '下载 PDF 报表',
        'total_emp': '总人数',
        'present': '出勤',
        'absent': '缺勤',
        'late_early': '迟到/早退',
        'chart_title': '各组别出勤状态分析图',
        'no_data': '请上传完整文件并配置文件信息。',
    }
}

# ==========================================
# 3. LỚP XỬ LÝ THÔNG MINH BẰNG HỆ THỐNG PHÂN TÍCH (ẨN DANH)
# ==========================================
def smart_parse_columns(df, expected_cols, file_description):
    """
    Sử dụng mô hình phân tích tự động và đồng nhất tiêu đề cột nếu Python quét không khớp.
    """
    if client is None:
        return {col: col for col in expected_cols if col in df.columns}
        
    sample_data = df.head(3).to_string()
    prompt = f"""
    Bạn là một chuyên gia xử lý dữ liệu. Hệ thống cần tìm các cột tương ứng với danh sách mục tiêu: {expected_cols}.
    Dưới đây là các cột thực tế và một ít dữ liệu mẫu của file '{file_description}':
    Các cột thực tế: {list(df.columns)}
    Dữ liệu mẫu:
    {sample_data}
    
    Hãy trả về một JSON duy nhất map từ tên cột mục tiêu sang tên cột thực tế chính xác nhất. 
    Ví dụ: {{"Mã NV": "Mã nhân viên thực tế", "Giờ vào": "Giờ vào thực tế"}}
    Chỉ trả ra JSON, không viết thêm lời thoại nào khác.
    """
    try:
        response = client.models.generate_content(
           model='gemini-2.5-flash',
           contents=prompt,
           config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        mapping = json.loads(response.text.strip())
        return mapping
    except Exception:
        return {col: col for col in expected_cols if col in df.columns}

# ==========================================
# 4. HÀM XỬ LÝ LOGIC TÍNH TOÁN THEO QUY TẮC
# ==========================================
def process_attendance(df_finger, df_schedule, df_office, df_worker, target_date):
    finger_map = {'Mã NV': 'Mã NV', 'Ngày': 'Ngày', 'Giờ vào': 'Giờ vào', 'Giờ ra': 'Giờ ra', 'Tổng giờ': 'Tổng giờ'}
    
    df_finger_today = df_finger.copy()
    
    try:
       df_finger_today['Ngày'] = pd.to_datetime(df_finger_today['Ngày']).dt.date
    except:
       pass
        
    df_finger_today = df_finger_today[df_finger_today['Ngày'] == target_date]
    
    results = []
    
    office_ids = df_office['Mã NV'].astype(str).tolist() if df_office is not None else []
    worker_ids = df_worker['Mã NV'].astype(str).tolist() if df_worker is not None else []
    
    maintenance_ids = ['673', 'A068']
    qc_ids = ['749', '949']
    cleaner_ids = ['575']
    
    all_employees = []
    if df_office is not None:
        for idx, row in df_office.iterrows():
            all_employees.append({'Mã NV': str(row['Mã NV']), 'Tên': row.get('Họ và tên', 'N/A'), 'Nhóm': 'VP'})
    if df_worker is not None:
        for idx, row in df_worker.iterrows():
            emp_id = str(row['Mã NV'])
            group = 'Công nhân'
            if emp_id in maintenance_ids: group = 'Bảo trì'
            elif emp_id in qc_ids: group = 'QC'
            elif emp_id in cleaner_ids: group = 'Tạp vụ'
            all_employees.append({'Mã NV': emp_id, 'Tên': row.get('Họ và tên', 'N/A'), 'Nhóm': group})

    for emp in all_employees:
        emp_id = emp['Mã NV']
        name = emp['Tên']
        group = emp['Nhóm']
        
        emp_finger = df_finger_today[df_finger_today['Mã NV'].astype(str) == emp_id]
        
        time_in = None
        time_out = None
        actual_hours = 0
        note = ""
        
        if not emp_finger.empty:
            row_f = emp_finger.iloc[0]
            time_in = row_f.get('Giờ vào')
            time_out = row_f.get('Giờ ra')
            try:
               actual_hours = float(row_f.get('Tổng giờ', 0))
            except:
               actual_hours = 0
                
        is_sunday = target_date.weekday() == 6 
        
        if group == 'VP':
            if is_sunday:
                note = "Nghỉ CN / 周日休息"
            elif emp_finger.empty:
                note = "Vắng / 缺勤"
            else:
                if pd.isna(time_in): note += "BV "
                if pd.isna(time_out): note += "BR "
                
                if actual_hours < 8:
                    note += f"Về sớm ({actual_hours}h) / 早退"
                elif actual_hours >= 8:
                    note = "Bình thường / 正常"
                    
        elif group == 'Tạp vụ' and emp_id == '575':
            if emp_finger.empty:
                note = "Vắng / 缺勤"
            else:
                if actual_hours < 12: note = f"Về sớm ({actual_hours}h) / 早退"
                else: note = "Bình thường / 正常"
                
        elif group == 'QC' and emp_id in qc_ids:
            if emp_finger.empty:
                note = "Vắng / 缺勤"
            else:
                if actual_hours < 12: note = f"Về sớm ({actual_hours}h) / 早退"
                else: note = "Bình thường / 正常"
                
        else:
            shift = "Nghỉ"
            if df_schedule is not None:
               sched_emp = df_schedule[df_schedule['Mã NV'].astype(str) == emp_id]
               if not sched_emp.empty:
                   day_col = str(target_date.day)
                   if day_col in df_schedule.columns:
                       shift = str(sched_emp.iloc[0][day_col]).strip()

            if shift in ['Nghỉ', 'nan', 'Nghỉ tuần']:
                note = "Nghỉ theo lịch / 排班休息"
            elif emp_finger.empty:
                note = "Vắng / 缺勤"
            else:
                if shift == 'Đ':
                    if actual_hours < 12: note = f"Làm không đúng lịch ({actual_hours}h) / 未按班"
                    else: note = "Ca Đêm OK / 夜班正常"
                elif shift == 'N':
                    if actual_hours < 12: note = f"Làm không đúng lịch ({actual_hours}h) / 未按班"
                    else: note = "Ca Ngày OK / 白班正常"
                else:
                    note = f"Lịch ghi: {shift}"

        results.append({
            'Mã NV': emp_id,
            'Họ và tên': name,
            'Nhóm': group,
            'Giờ vào': time_in if not pd.isna(time_in) else "Thiếu/无",
            'Giờ ra': time_out if not pd.isna(time_out) else "Thiếu/无",
            'Giờ làm thực tế': actual_hours,
            'Ghi chú': note
        })
        
    return pd.DataFrame(results)

# ==========================================
# 5. GIAO DIỆN STREAMLIT
# ==========================================
st.set_page_config(layout="wide", page_title="HR Dashboard")

lang_choice = st.sidebar.selectbox("Ngôn ngữ / 语言", ["Tiếng Việt", "中文"])
lang = LANG['VI'] if lang_choice == "Tiếng Việt" else LANG['ZH']

st.title(f"📊 {lang['title']}")

st.header(f"📁 {lang['upload_section']}")
col1, col2, col3, col4 = st.columns(4)
with col1:
    file_finger = st.file_uploader(lang['btn_finger'], type=['xlsx'])
with col2:
    file_schedule = st.file_uploader(lang['btn_schedule'], type=['xlsx'])
with col3:
    file_office = st.file_uploader(lang['btn_office'], type=['xlsx'])
with col4:
    file_worker = st.file_uploader(lang['btn_worker'], type=['xlsx'])

st.header(f"📅 {lang['filter_section']}")
target_date = st.date_input(lang['select_date'], datetime.date(2026, 9, 10))

if file_finger and file_office:
    df_finger = pd.read_excel(file_finger)
    df_office = pd.read_excel(file_office)
    df_schedule = pd.read_excel(file_schedule) if file_schedule else None
    df_worker = pd.read_excel(file_worker) if file_worker else None

    df_result = process_attendance(df_finger, df_schedule, df_office, df_worker, target_date)

    st.header(f"📈 {lang['stats_section']}")
    total_emp = len(df_result)
    vắng_count = len(df_result[df_result['Ghi chú'].str.contains("Vắng|缺勤")])
    bth_count = total_emp - vắng_count
    
    m1, m2, m3 = st.columns(3)
    m1.metric(lang['total_emp'], total_emp)
    m2.metric(lang['present'], bth_count)
    m3.metric(lang['absent'], vắng_count)

    fig = px.bar(df_result, x='Nhóm', color='Ghi chú',
                 title=lang['chart_title'],
                 barmode='stack', text_auto=True,
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    st.plotly_chart(fig, use_container_width=True)

    st.header(f"📋 {lang['report_section']}")
    st.dataframe(df_result, use_container_width=True)

    st.subheader("📥 Xuất dữ liệu / Export")
    ex_col, pdf_col = st.columns(2)

    output_excel = io.BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        df_result[['Mã NV', 'Họ và tên', 'Giờ vào', 'Giờ ra', 'Giờ làm thực tế', 'Ghi chú']].to_excel(writer, index=False, sheet_name='Thống kê')
    excel_data = output_excel.getvalue()

    with ex_col:
        st.download_button(
            label=lang['download_excel'],
            data=excel_data,
            file_name=f"Thong_ke_Nhan_vien_{target_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    output_pdf = io.BytesIO()
    doc = SimpleDocTemplate(output_pdf, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"{lang['title']}", styles['Title']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Ngày kiểm tra / 日期: {target_date}", styles['Normal']))
    story.append(Paragraph(f"Tổng số nhân viên: {total_emp} | Đi làm: {bth_count} | Vắng: {vắng_count}", styles['Normal']))
    story.append(Spacer(1, 20))

    pdf_data = [list(df_result.columns)] + df_result.values.tolist()
    t = Table(pdf_data)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.grey),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BOTTOMPADDING', (0,0), (-1,0), 8),
        ('GRID', (0,0), (-1,-1), 1, colors.black)
    ]))
    story.append(t)
    doc.build(story)
    pdf_data_bytes = output_pdf.getvalue()

    with pdf_col:
        st.download_button(
            label=lang['download_pdf'],
            data=pdf_data_bytes,
            file_name=f"Dashboard_Report_{target_date}.pdf",
            mime="application/pdf"
        )
else:
    st.info(lang['no_data'])
