import streamlit as st
import pandas as pd
import numpy as np
import datetime
from google import genai
import io
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 0. CẤU HÌNH GIAO DIỆN & NGÔN NGỮ (SONG NGỮ KHÔNG CHỨA TỪ "AI")
# ==========================================
st.set_page_config(page_title="Hệ Thống Thống Kê Chấm Công", layout="wide")

if "lang" not in st.session_state:
    st.session_state.lang = "vi"

# Bộ từ điển song ngữ Trung - Việt (Không chứa từ "AI")
LANG_DICT = {
    "vi": {
        "title": "📊 HỆ THỐNG PHÂN TÍCH CHẤM CÔNG THÔNG MINH BẤT ĐỊNH VĂN BẢN",
        "upload_section": "📁 Tải Lên Dữ Liệu Hệ Thống",
        "btn_lang": "切换至中文 (Chuyển sang tiếng Trung)",
        "file_finger": "1. File dữ liệu bấm vân tay máy chấm công",
        "file_schedule": "2. File lịch xếp ca làm việc (Công nhân)",
        "file_vp": "3. Danh sách nhân viên Văn phòng (VP)",
        "file_cn": "4. Danh sách Công nhân nhà máy",
        "filter_section": "📆 Bộ Lọc Thời Gian Kiểm Tra",
        "select_date": "Chọn ngày cần kết xuất báo cáo:",
        "analyze_btn": "🚀 Bắt Đầu Phân Tích & Đối Chiếu Dữ Liệu",
        "dashboard_stat": "📈 Bảng Thống Kê Phân Tích Chuyên Nghiệp",
        "abnormal_focus": "🚨 Trọng Tâm Các Trường Hợp Bất Thường (Cần Lưu Ý)",
        "total_staff": "Tổng nhân sự dự kiến",
        "total_present": "Đi làm đúng / đủ",
        "total_abnormal": "Số ca bất thường",
        "total_absent": "Vắng mặt",
        "export_section": "📥 Xuất Báo Cáo Thống Kê",
        "download_excel": "Tải file kết quả Excel (.xlsx)",
        "download_pdf": "Tải file giao diện PDF",
        "col_eid": "Mã NV", "col_name": "Họ và tên", "col_in": "Giờ vào", 
        "col_out": "Giờ ra", "col_hours": "Giờ làm thực tế", "col_note": "Ghi chú"
    },
    "zh": {
        "title": "📊 智能文本泛化考勤分析与对账系统",
        "upload_section": "📁 系统数据上传",
        "btn_lang": "Chuyển sang tiếng Việt (切换至越南语)",
        "file_finger": "1. 考勤机指纹打卡原始数据文件",
        "file_schedule": "2. 排班表文件 (车间工人)",
        "file_vp": "3. 办公室人员花名册 (VP)",
        "file_cn": "4. 车间工人花名册",
        "filter_section": "📆 检查时间筛选",
        "select_date": "选择需要生成报告的日期:",
        "analyze_btn": "🚀 开始智能数据对账与分析",
        "dashboard_stat": "📈 专业考勤数据图表分析",
        "abnormal_focus": "🚨 异常考勤重点关注摘要 (精简)",
        "total_staff": "预计总出勤人数",
        "total_present": "正常出勤人数",
        "total_abnormal": "异常打卡人数",
        "total_absent": "旷工/缺勤人数",
        "export_section": "📥 导出统计报告",
        "download_excel": "下载 Excel 统计表 (.xlsx)",
        "download_pdf": "下载当前 PDF 报表页面",
        "col_eid": "工号", "col_name": "姓名", "col_in": "签到时间", 
        "col_out": "签退時間", "col_hours": "实际工时", "col_note": "考勤备注"
    }
}

l = LANG_DICT[st.session_state.lang]

# Nút chuyển đổi ngôn ngữ nhanh trên thanh Sidebar
if st.sidebar.button(l["btn_lang"]):
    st.session_state.lang = "zh" if st.session_state.lang == "vi" else "vi"
    st.rerun()

st.title(l["title"])

# ==========================================
# 1. KẾT NỐI HỆ THỐNG TRÍ TUỆ ĐIỆN TỬ (GEMINI CLIENT)
# ==========================================
if "GEMINI_API_KEY" in st.secrets:
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Chưa cấu hình khóa bảo mật GEMINI_API_KEY trên Streamlit Cloud!")
    st.stop()

# Hàm trung gian gửi danh sách cột qua hệ thống Trí tuệ điện tử để phân tích ngữ nghĩa tự động
def dynamic_column_mapping(columns_list, target_concept):
    prompt = f"""
    Bạn là một chuyên gia chuẩn hóa dữ liệu. Tôi có danh sách các tiêu đề cột quét từ file Excel như sau: {columns_list}
    Hãy tìm ra tên cột khớp nhất với khái niệm: '{target_concept}'.
    Yêu cầu bắt buộc: Chỉ trả về duy nhất tên cột chính xác tuyệt đối lấy từ danh sách trên, không kèm giải thích, không dấu câu. Nếu hoàn toàn không thấy khái niệm tương đương, trả về 'NONE'.
    """
    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text.strip()
    except:
        return "NONE"

# ==========================================
# 2. KHÔNG GIAN TẢI FILE DỮ LIỆU ĐẦU VÀO
# ==========================================
st.header(l["upload_section"])
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

with col1:
    f_finger = st.file_uploader(l["file_finger"], type=["xlsx", "xls"])
with col2:
    f_schedule = st.file_uploader(l["file_schedule"], type=["xlsx", "xls"])
with col3:
    f_vp = st.file_uploader(l["file_vp"], type=["xlsx", "xls"])
with col4:
    f_cn = st.file_uploader(l["file_cn"], type=["xlsx", "xls"])

# ==========================================
# 3. BỘ LỌC THỜI GIAN THEO YÊU CẦU NGƯỜI DÙNG
# ==========================================
st.header(l["filter_section"])
target_date = st.date_input(l["select_date"], datetime.date(2026, 9, 6))

# Tải cấu trúc tĩnh để tránh lỗi xử lý trống
final_report = pd.DataFrame()

if f_finger and f_schedule and f_vp and f_cn:
    if st.button(l["analyze_btn"], type="primary"):
        with st.spinner("Hệ thống kiểm toán dữ liệu đang quét tự động cấu trúc văn bản..."):
            
            # Đọc toàn bộ dữ liệu thô
            df_finger = pd.read_excel(f_finger)
            df_sched = pd.read_excel(f_schedule)
            df_master_vp = pd.read_excel(f_vp)
            df_master_cn = pd.read_excel(f_cn)
            
            # --- Hệ thống tự dịch và tìm ánh xạ cột bằng mô hình Trí Tuệ Điện Tử ---
            id_col_finger = dynamic_column_mapping(list(df_finger.columns), "mã số nhân viên hoặc số thẻ")
            time_col_finger = dynamic_column_mapping(list(df_finger.columns), "thời gian bấm giờ vào ra hoặc ngày giờ quét vân tay")
            
            # KIỂM TRA AN TOÀN: Đảm bảo cột tồn tại, nếu không cho phép chọn thủ công để tránh KeyError
            if time_col_finger not in df_finger.columns or time_col_finger == "NONE":
                st.warning("⚠️ Không thể tự động nhận diện cột thời gian. Vui lòng chọn cột chứa Thời gian/Ngày giờ quét vân tay:")
                time_col_finger = st.selectbox("Chọn cột thời gian (File vân tay):", df_finger.columns, key="sel_time_finger")

            if id_col_finger not in df_finger.columns or id_col_finger == "NONE":
                st.warning("⚠️ Không thể tự động nhận diện cột mã nhân viên. Vui lòng chọn cột chứa Mã nhân viên:")
                id_col_finger = st.selectbox("Chọn cột Mã nhân viên (File vân tay):", df_finger.columns, key="sel_id_finger")
            
            id_col_vp = dynamic_column_mapping(list(df_master_vp.columns), "mã nhân viên hoặc số thẻ văn phòng")
            name_col_vp = dynamic_column_mapping(list(df_master_vp.columns), "họ và tên nhân viên văn phòng")
            
            id_col_cn = dynamic_column_mapping(list(df_master_cn.columns), "mã công nhân hoặc số thẻ nhà máy")
            name_col_cn = dynamic_column_mapping(list(df_master_cn.columns), "họ và tên công nhân")
            
            # Đồng bộ dữ liệu ngày tháng của file vân tay về chuẩn datetime
            df_finger['Standard_DateTime'] = pd.to_datetime(df_finger[time_col_finger], errors='coerce')
            df_finger['Standard_Date'] = df_finger['Standard_DateTime'].dt.date
            
            # Lọc dữ liệu vân tay đúng ngày người dùng chọn trên dashboard
            df_finger_today = df_finger[df_finger['Standard_Date'] == target_date]
            
            rows_output = []
            
            # ==========================================
            # LOGIC XỬ LÝ KHỐI 1: NHÂN VIÊN VĂN PHÒNG (VP)
            # ==========================================
            for _, row in df_master_vp.iterrows():
                eid = str(row[id_col_vp]).strip()
                name = str(row[name_col_vp]).strip()
                
                # Lấy dữ liệu quẹt thẻ của nhân viên này trong ngày
                emp_logs = df_finger_today[df_finger_today[id_col_finger].astype(str).str.strip() == eid].sort_values(by='Standard_DateTime')
                
                # Quy tắc riêng biệt cho mã đặc thù
                if eid == "575":
                    std_in, std_out, expected_hours = "07:00:00", "15:00:00", 12.0
                elif eid in ["749", "949"]:
                    std_in, std_out, expected_hours = "07:00:00", "19:00:00", 12.0
                else:
                    std_in, std_out, expected_hours = "08:00:00", "17:00:00", 8.0
                
                # Mặc định Chủ Nhật của Khối VP là ngày nghỉ tuần
                if target_date.weekday() == 6:
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name,
                        l["col_in"]: "-", l["col_out"]: "-",
                        l["col_hours"]: 0, l["col_note"]: "Nghỉ tuần" if st.session_state.lang=="vi" else "周休"
                    })
                    continue
                
                if emp_logs.empty:
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name,
                        l["col_in"]: "-", l["col_out"]: "-",
                        l["col_hours"]: 0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                    })
                else:
                    t_in = emp_logs['Standard_DateTime'].min()
                    t_out = emp_logs['Standard_DateTime'].max() if len(emp_logs) > 1 else pd.NaT
                    
                    in_str = t_in.strftime("%H:%M") if not pd.isna(t_in) else "-"
                    out_str = t_out.strftime("%H:%M") if not pd.isna(t_out) else "-"
                    
                    notes = []
                    actual_hours = 0.0
                    if not pd.isna(t_in) and not pd.isna(t_out):
                        actual_hours = round((t_out - t_in).total_seconds() / 3600, 1)
                        
                        limit_in = datetime.datetime.combine(target_date, datetime.time.fromisoformat(std_in))
                        limit_out = datetime.datetime.combine(target_date, datetime.time.fromisoformat(std_out))
                        
                        if t_in > limit_in + datetime.timedelta(minutes=5):
                            notes.append("Đi trễ" if st.session_state.lang=="vi" else "迟到")
                        if t_out < limit_out:
                            notes.append("Về sớm" if st.session_state.lang=="vi" else "早退")
                        if actual_hours < expected_hours and "Về sớm" not in notes:
                            notes.append(f"Về sớm ({actual_hours}h)")
                    else:
                        if pd.isna(t_in): notes.append("BV")
                        if pd.isna(t_out): notes.append("BR")
                        
                    note_str = ", ".join(notes) if notes else ("Đủ công" if st.session_state.lang=="vi" else "全勤")
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name,
                        l["col_in"]: in_str, l["col_out"]: out_str,
                        l["col_hours"]: actual_hours, l["col_note"]: note_str
                    })

            # ==========================================
            # LOGIC XỬ LÝ KHỐI 2: CÔNG NHÂN (THEO LỊCH XẾP CA)
            # ==========================================
            id_col_sched = dynamic_column_mapping(list(df_sched.columns), "mã nhân viên hoặc mã số công nhân trong bảng lịch ca")
            day_str_target = str(target_date.day) 
            
            sched_day_col = [c for c in df_sched.columns if str(c).strip() == day_str_target]
            if sched_day_col:
                target_day_col_name = sched_day_col[0]
                
                for _, row in df_master_cn.iterrows():
                    eid = str(row[id_col_cn]).strip()
                    name = str(row[name_col_cn]).strip()
                    
                    sched_row = df_sched[df_sched[id_col_sched].astype(str).str.strip() == eid]
                    shift_type = "OFF"
                    
                    if not sched_row.empty:
                        shift_type = str(sched_row[target_day_col_name].values[0]).strip()
                        
                    if shift_type in ["OFF", "nan", "-"]:
                        rows_output.append({
                           l["col_eid"]: eid, l["col_name"]: name,
                           l["col_in"]: "-", l["col_out"]: "-",
                           l["col_hours"]: 0, l["col_note"]: "Nghỉ ca" if st.session_state.lang=="vi" else "轮休"
                        })
                        continue
                        
                    emp_logs = df_finger_today[df_finger_today[id_col_finger].astype(str).str.strip() == eid].sort_values(by='Standard_DateTime')
                    
                    if shift_type == "N": 
                        if emp_logs.empty:
                            rows_output.append({
                                __import__('builtins').dict(l)["col_eid"]: eid, l["col_name"]: name,
                                l["col_in"]: "-", l["col_out"]: "-",
                                l["col_hours"]: 0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                            })
                        else:
                            t_in = emp_logs['Standard_DateTime'].min()
                            t_out = emp_logs['Standard_DateTime'].max() if len(emp_logs) > 1 else pd.NaT
                            actual_hours = round((t_out - t_in).total_seconds() / 3600, 1) if not pd.isna(t_out) else 0
                            
                            notes = []
                            if t_in.time() > datetime.time(7, 10): 
                                notes.append("Đi trễ" if st.session_state.lang=="vi" else "迟到")
                            if pd.isna(t_out) or t_out.time() < datetime.time(19, 0):
                                notes.append("Làm không đúng lịch" if st.session_state.lang=="vi" else "未按排班出勤")
                                
                            note_str = ", ".join(notes) if notes else "N"
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name,
                                l["col_in"]: t_in.strftime("%H:%M"),
                                l["col_out"]: t_out.strftime("%H:%M") if not pd.isna(t_out) else "-",
                                l["col_hours"]: actual_hours, l["col_note"]: note_str
                            })
                            
                    elif shift_type == "Đ": 
                        t_in = emp_logs[emp_logs['Standard_DateTime'].dt.time >= datetime.time(18, 0)]['Standard_DateTime'].min()
                        
                        next_day = target_date + datetime.timedelta(days=1)
                        df_finger_next = df_finger[df_finger['Standard_Date'] == next_day]
                        emp_logs_next = df_finger_next[df_finger_next[id_col_finger].astype(str).str.strip() == eid].sort_values(by='Standard_DateTime')
                        
                        t_out = emp_logs_next[emp_logs_next['Standard_DateTime'].dt.time <= datetime.time(8, 0)]['Standard_DateTime'].max()
                        notes = []
                        
                        if pd.isna(t_in) and pd.isna(t_out):
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name,
                                l["col_in"]: "-", l["col_out"]: "-",
                                l["col_hours"]: 0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                            })
                        else:
                            actual_hours = round((t_out - t_in).total_seconds() / 3600, 1) if not pd.isna(t_in) and not pd.isna(t_out) else 0
                            
                            if not pd.isna(t_in) and t_in.time() > datetime.time(19, 10):
                                notes.append("Đi trễ" if st.session_state.lang=="vi" else "迟到")
                            if pd.isna(t_out) or t_out.time() < datetime.time(7, 0):
                                notes.append("Làm không đúng lịch" if st.session_state.lang=="vi" else "未按排班出勤")
                                
                            note_str = ", ".join(notes) if notes else "Đ"
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name,
                                l["col_in"]: t_in.strftime("%H:%M") if not pd.isna(t_in) else "-",
                                l["col_out"]: t_out.strftime("%H:%M") if not pd.isna(t_out) else "-",
                                l["col_hours"]: actual_hours, l["col_note"]: note_str
                            })
                            
            final_report = pd.DataFrame(rows_output)
            st.session_state.final_report = final_report

# ==========================================
# 4. HIỂN THỊ KẾT QUẢ ĐỐI CHIẾU DỮ LIỆU
# ==========================================
if "final_report" in st.session_state and not st.session_state.final_report.empty:
    df_res = st.session_state.final_report
    
    # 4.1 Bảng Chỉ Số Đo Lường (Metrics)
    st.header(l["dashboard_stat"])
    
    m_total = len(df_res)
    m_absent = len(df_res[df_res[l["col_note"]].str.contains("Vắng|旷工", na=False)])
    m_abnormal = len(df_res[df_res[l["col_note"]].str.contains("Đi trễ|Về sớm|lịch|迟到|早退|未按", na=False)])
    m_present = m_total - m_absent - m_abnormal
    
    mx1, mx2, mx3, mx4 = st.columns(4)
    mx1.metric(l["total_staff"], m_total)
    mx2.metric(l["total_present"], m_present)
    mx3.metric(l["total_abnormal"], m_abnormal, delta_color="inverse")
    mx4.metric(l["total_absent"], m_absent, delta_color="inverse")
    
    # 4.2 Đồ Thị Trực Quan Hóa Chuyên Nghiệp
    g1, g2 = st.columns([1, 1])
    with g1:
        fig_pie = px.pie(
            names=[l["total_present"], l["total_abnormal"], l["total_absent"]],
            values=[m_present, m_abnormal, m_absent],
            color_discrete_sequence=px.colors.qualitative.Pastel,
            title="Tỷ lệ trạng thái chấm công / 出勤状态分布图"
        )
        st.plotly_chart(fig_pie, use_container_width=True)
        
    with g2:
        fig_bar = px.bar(
            df_res, x=l["col_name"],
            y=l["col_hours"], # Đã sửa từ col_name sang col_hours (giá trị số thực tế)
            color=l["col_note"],
            title="Phân bố chi tiết theo từng cá nhân / 个人考勤具体分布"
        )
        st.plotly_chart(fig_bar, use_container_width=True)
        
    # 4.3 Trọng Tâm Các Trường Hợp Bất Thường
    st.header(l["abnormal_focus"])
    df_bad = df_res[df_res[l["col_note"]].str.contains("Vắng|Đi trễ|Về sớm|lịch|旷工|迟到|早退|未按", na=False)]
    
    if not df_bad.empty:
        summary_prompt = f"""
        Dựa trên bảng danh sách nhân sự lỗi chấm công sau đây, hãy viết một báo cáo tóm tắt trọng tâm siêu ngắn gọn (dưới 5 dòng), chỉ rõ các mã nhân viên vi phạm nghiêm trọng (Vắng, Đi trễ hoặc sai ca). Không viết dài dòng.
        Danh sách dữ liệu lỗi:
        {df_bad[[l["col_eid"], l["col_name"], l["col_note"]]].to_string(index=False)}
        Language response requirement: Trả về kết quả song ngữ Trung-Việt tương ứng. Tuyệt đối không dùng chữ 'AI' trong câu trả lời.
        """
        summary_response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=summary_prompt,
        )
        st.warning(summary_response.text)
    else:
        st.success("Không có trường hợp bất thường nào trong ngày! / 今日无异常考勤记录。")
        
    # 4.4 Bảng Dữ Liệu Chi Tiết Giao Diện Người Dùng
    st.dataframe(df_res, use_container_width=True)

# ==========================================
# 5. XUẤT FILE BÁO CÁO (EXCEL / PDF TRỰC TIẾP)
# ==========================================
st.header(l["export_section"])

buffer_excel = io.BytesIO()
with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
    if not final_report.empty:
        final_report.to_excel(writer, index=False, sheet_name='ThongKe_Attendance')

c_down1, c_down2 = st.columns(2)
with c_down1:
    st.download_button(
        label=l["download_excel"],
        data=buffer_excel.getvalue(),
        file_name=f"Attendance_Report_{target_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with c_down2:
    st.button(
        l["download_pdf"], 
        on_click=st.js_sandbox if hasattr(st, "js_sandbox") else None, 
        help="Nhấn Ctrl+P hoặc Cmd+P để lưu file PDF khớp giao diện Dashboard"
    )
