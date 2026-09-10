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
# 1. CẤU HÌNH & KHỞI TẠO HỆ THỐNG
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
        'chart_title': 'Biểu đồ tình trạng đi làm theo Nhóm và Trạng Thái',
        'no_data': 'Vui lòng tải đủ file dữ liệu ở thanh bên để hệ thống xử lý.',
        'processing': 'Hệ thống đang quét và đồng nhất cấu trúc dữ liệu...',
        'success_proc': 'Đồng nhất cấu trúc dữ liệu thành công!',
    },
    'ZH': {
        'title': '员工出勤考勤管理与统计系统',
        'upload_section': '上传输入数据',
        'btn_finger': '1. 上传指纹打卡文件 (.xlsx)',
        'btn_schedule': '2. 上传排班表文件 (.xlsx)',
        'btn_office': '3. 上传办公室人员名单 (.xlsx)',
        'btn_worker': '4. 上传工人名单 (.xlsx)',
        'filter_section': '时间筛选器',
        'select_date': '选择工作检查日期',
        'stats_section': '直观数据统计',
        'report_section': '详细报表',
        'download_excel': '下载 Excel 报表',
        'download_pdf': '下载 PDF 报表',
        'total_emp': '总人数',
        'present': '出勤',
        'absent': '缺勤',
        'chart_title': '各组别及考勤状态图表分析',
        'no_data': '请在侧边栏上传完整的文件以便系统处理。',
        'processing': '系统正在扫描并对齐数据结构...',
        'success_proc': '数据结构对齐成功！',
    }
}

# ==========================================
# 3. HÀM XỬ LÝ TIÊU ĐỀ THÔNG MINH
# ==========================================
def smart_parse_columns(df, expected_cols, file_description):
    if client is None:
        return {col: col for col in expected_cols if col in df.columns}
        
    sample_data = df.head(2).to_string()
    prompt = f"""
    Bạn là một chuyên gia xử lý dữ liệu nhân sự. Hệ thống cần tìm các cột tương ứng với danh sách mục tiêu: {expected_cols}.
    Dưới đây là các cột thực tế và dữ liệu mẫu của file '{file_description}':
    Cột thực tế: {list(df.columns)}
    Dữ liệu mẫu:
    {sample_data}
    
    Hãy trả về một JSON duy nhất map từ tên cột mục tiêu sang tên cột thực tế chính xác nhất. 
    Chỉ trả ra JSON chuẩn, không bọc markdown hay chú thích thêm.
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
    df_f = df_finger.copy()
    
    def get_c(df, keys):
        for c in df.columns:
            if any(k in str(c).lower() for k in keys):
                return c
        return None

    c_date = get_c(df_f, ['ngày', 'date', 'ngay'])
    c_id = get_c(df_f, ['mã nv', 'ma nv', 'manv', 'code', 'id', 'mã'])
    c_in = get_c(df_f, ['giờ vào', 'gio vao', 'vào', 'vao', 'in'])
    c_out = get_c(df_f, ['giờ ra', 'gio ra', 'ra', 'out'])
    c_total = get_c(df_f, ['tổng giờ', 'tong gio', 'total', 'giờ làm', 'gio lam'])

    if c_date:
        try:
            df_f['Parsed_Date'] = pd.to_datetime(df_f[c_date], errors='coerce').dt.date
            df_f = df_f[df_f['Parsed_Date'] == target_date]
        except:
            pass

    results = []
    
    off_id_c = get_c(df_office, ['mã nv', 'ma nv', 'manv', 'code', 'id', 'mã']) if df_office is not None else None
    off_name_c = get_c(df_office, ['tên', 'ten', 'name', 'họ và tên']) if df_office is not None else None
    
    wr_id_c = get_c(df_worker, ['mã nv', 'ma nv', 'manv', 'code', 'id', 'mã']) if df_worker is not None else None
    wr_name_c = get_c(df_worker, ['tên', 'ten', 'name', 'họ và tên']) if df_worker is not None else None

    maintenance_ids = ['673', 'A068']
    qc_ids = ['749', '949']
    cleaner_ids = ['575']
    
    all_employees = []
    if df_office is not None and off_id_c:
        for _, r in df_office.iterrows():
            all_employees.append({'Mã NV': str(r[off_id_c]), 'Tên': r.get(off_name_c, 'N/A') if off_name_c else 'N/A', 'Nhóm': 'Văn phòng'})
            
    if df_worker is not None and wr_id_c:
        for _, r in df_worker.iterrows():
            eid = str(r[wr_id_c])
            grp = 'Công nhân'
            if eid in maintenance_ids: grp = 'Bảo trì'
            elif eid in qc_ids: grp = 'QC'
            elif eid in cleaner_ids: grp = 'Tạp vụ'
            all_employees.append({'Mã NV': eid, 'Tên': r.get(wr_name_c, 'N/A') if wr_name_c else 'N/A', 'Nhóm': grp})

    for emp in all_employees:
        eid = emp['Mã NV']
        name = emp['Tên']
        grp = emp['Nhóm']
        
        emp_f = df_f[df_f[c_id].astype(str) == eid] if c_id and not df_f.empty else pd.DataFrame()
        
        t_in, t_out, actual_h, note = None, None, 0.0, ""
        
        if not emp_f.empty:
            rf = emp_f.iloc[0]
            t_in = rf.get(c_in) if c_in else None
            t_out = rf.get(c_out) if c_out else None
            try:
                actual_h = float(rf.get(c_total, 0.0)) if c_total else 0.0
            except:
                actual_h = 0.0
                
        is_sun = target_date.weekday() == 6
        
        # --- ÁP DỤNG QUY TẮC CHO TỪNG NHÓM ---
        if grp == 'Văn phòng':
            if is_sun:
                note = "Nghỉ CN / 周日休息"
            elif emp_f.empty:
                note = "Vắng / 缺勤"
            else:
                if pd.isna(t_in): note += "BV "
                if pd.isna(t_out): note += "BR "
                if actual_h < 8.0:
                    note = f"về sớm ({actual_h}h) / 早退"
                else:
                    note = "Đúng giờ / 准时"
                    
        elif grp == 'Tạp vụ' and eid == '575':
            if emp_f.empty:
                note = "Vắng / 缺勤"
            else:
                if actual_h < 12.0:
                    note = f"về sớm ({actual_h}h) / 早退"
                else:
                    note = "Đúng giờ / 准时"
                
        elif grp == 'QC' and eid in qc_ids:
            if emp_f.empty:
                note = "Vắng / 缺勤"
            else:
                if actual_h < 12.0:
                    note = f"về sớm ({actual_h}h) / 早退"
                else:
                    note = "Đúng giờ / 准时"
                
        else:
            shift = "Nghỉ"
            sched_id_c = get_c(df_schedule, ['mã nv', 'ma nv', 'manv', 'code', 'id', 'mã']) if df_schedule is not None else None
            if df_schedule is not None and sched_id_c:
                sched_row = df_schedule[df_schedule[sched_id_c].astype(str) == eid]
                if not sched_row.empty:
                    d_col = str(target_date.day)
                    if d_col in df_schedule.columns:
                        shift = str(sched_row.iloc[0][d_col]).strip()

            if shift in ['Nghỉ', 'nan', 'Nghỉ tuần', 'None', '']:
                note = "Nghỉ / 休息"
            elif emp_f.empty:
                note = "Vắng / 缺勤"
            else:
                if shift == 'Đ':
                    if actual_h < 12.0:
                        note = f"làm không đúng lịch ({actual_h}h) / 未按班"
                    else:
                        note = "Ca đêm OK / 夜班正常"
                elif shift == 'N':
                    if actual_h < 12.0:
                        note = f"làm không đúng lịch ({actual_h}h) / 未按班"
                    else:
                        note = "Ca ngày OK / 白班正常"
                else:
                    note = f"Lịch: {shift}"

        results.append({
            'mã NV': eid,
            'họ và tên': name,
            'nhóm': grp,
            'giờ vào': t_in if (t_in is not None and not pd.isna(t_in)) else "Thiếu/无",
            'giờ ra': t_out if (t_out is not None and not pd.isna(t_out)) else "Thiếu/无",
            'giờ làm thực tế': actual_h,
            'ghi chú': note
        })
        
    return pd.DataFrame(results)

# ==========================================
# 5. GIAO DIỆN STREAMLIT
# ==========================================
st.set_page_config(layout="wide", page_title="HR Attendance Dashboard")

lang_choice = st.sidebar.selectbox("🌐 Ngôn ngữ / 语言", ["Tiếng Việt", "中文"])
lang = LANG['VI'] if lang_choice == "Tiếng Việt" else LANG['ZH']

st.title(f"📊 {lang['title']}")

st.header(f"📁 {lang['upload_section']}")
c1, c2, c3, c4 = st.columns(4)
with c1:
    file_finger = st.file_uploader(lang['btn_finger'], type=['xlsx', 'xls'])
with c2:
    file_schedule = st.file_uploader(lang['btn_schedule'], type=['xlsx', 'xls'])
with c3:
    file_office = st.file_uploader(lang['btn_office'], type=['xlsx', 'xls'])
with c4:
    file_worker = st.file_uploader(lang['btn_worker'], type=['xlsx', 'xls'])

st.header(f"📅 {lang['filter_section']}")

col_d, col_m, col_y = st.columns(3)
with col_d:
    sel_day = st.selectbox("Ngày / 日", list(range(1, 32)), index=9)
with col_m:
    sel_month = st.selectbox("Tháng / 月", list(range(1, 13)), index=8)
with col_y:
    sel_year = st.selectbox("Năm / 年", [2025, 2026, 2027], index=1)

try:
    target_date = datetime.date(sel_year, sel_month, sel_day)
except ValueError:
    target_date = datetime.date(2026, 9, 10)

if file_finger and file_office:
    with st.spinner(lang['processing']):
        df_finger = pd.read_excel(file_finger)
        df_office = pd.read_excel(file_office)
        df_schedule = pd.read_excel(file_schedule) if file_schedule else None
        df_worker = pd.read_excel(file_worker) if file_worker else None

        if client:
            smart_parse_columns(df_finger, ["mã nv", "ngày", "giờ vào", "giờ ra", "tổng giờ"], "File vân tay")
            
        df_result = process_attendance(df_finger, df_schedule, df_office, df_worker, target_date)
    st.toast(lang['success_proc'])

    # HIỂN THỊ LUÔN KHU VỰC THỐNG KÊ VÀ BIỂU ĐỒ TRỰC QUAN CHUYÊN NGHIỆP
    st.header(f"📈 {lang['stats_section']}")
    total_emp = len(df_result)
    absent_count = len(df_result[df_result['ghi chú'].str.contains("Vắng|缺勤")]) if not df_result.empty else 0
    present_count = total_emp - absent_count
    
    m1, m2, m3 = st.columns(3)
    m1.metric(lang['total_emp'], total_emp)
    m2.metric(lang['present'], present_count)
    m3.metric(lang['absent'], absent_count)

    if not df_result.empty:
        # Biểu đồ cột phân tích trực quan chuyên nghiệp phân theo Nhóm và Trạng thái Ghi chú
        fig = px.bar(
            df_result, 
            x='nhóm', 
            color='ghi chú',
            title=lang['chart_title'],
            barmode='stack', 
            text_auto=True,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig.update_layout(xaxis_title="Nhóm Nhân Sự / 人员组别", yaxis_title="Số Lượng / 数量")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Không có dữ liệu nhân sự phù hợp để hiển thị biểu đồ phân tích cho ngày này.")

    st.header(f"📋 {lang['report_section']} ({target_date})")
    
    # Hiển thị bảng lược bỏ cột nhóm phụ trong bảng xuất chính nhưng giữ nội dung chuẩn xác
    display_df = df_result[['mã NV', 'họ và tên', 'giờ vào', 'giờ ra', 'giờ làm thực tế', 'ghi chú']] if not df_result.empty else df_result
    st.dataframe(display_df, use_container_width=True)

    if not df_result.empty:
        st.subheader("📥 Tải Xuống Báo Cáo / 下载报告")
        ex_col, pdf_col = st.columns(2)

        # Xuất Excel đúng chuẩn yêu cầu
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='ThongKe_Attendance')
        excel_bytes = output_excel.getvalue()

        with ex_col:
            st.download_button(
                label=lang['download_excel'],
                data=excel_bytes,
                file_name=f"ThongKe_NhanVien_{target_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Xuất PDF đồng nhất giao diện
        output_pdf = io.BytesIO()
        doc = SimpleDocTemplate(output_pdf, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        story.append(Paragraph(f"<b>{lang['title']}</b>", styles['Title']))
        story.append(Spacer(1, 10))
        story.append(Paragraph(f"Ngày kiểm tra / 日期: {target_date}", styles['Normal']))
        story.append(Paragraph(f"Tổng NV: {total_emp} | Đi làm: {present_count} | Vắng: {absent_count}", styles['Normal']))
        story.append(Spacer(1, 15))

        pdf_data = [list(display_df.columns)] + display_df.values.tolist()
        t = Table(pdf_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1f4e78")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0,0), (-1,0), 6),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey)
        ]))
        story.append(t)
        doc.build(story)
        pdf_bytes = output_pdf.getvalue()

        with pdf_col:
            st.download_button(
                label=lang['download_pdf'],
                data=pdf_bytes,
                file_name=f"Dashboard_Report_{target_date}.pdf",
                mime="application/pdf"
            )
else:
    st.info(lang['no_data'])
