import streamlit as st
import pandas as pd
import numpy as np
import datetime
from google import genai
import io
import plotly.express as px

# ==========================================
# 0. CẤU HÌNH GIAO DIỆN & NGÔN NGỮ SONG NGỮ (KHÔNG CHỨA TỪ "AI")
# ==========================================
st.set_page_config(page_title="Hệ Thống Thống Kê Chấm Công", layout="wide")

if "lang" not in st.session_state:
    st.session_state.lang = "vi"

LANG_DICT = {
    "vi": {
        "title": "📊 HỆ THỐNG PHÂN TÍCH CHẤM CÔNG VÀ ĐỐI CHIẾU DỮ LIỆU TỰ ĐỘNG",
        "upload_section": "📁 Tải Lên 4 File Dữ Liệu Hệ Thống",
        "btn_lang": "切换至中文 (Chuyển sang tiếng Trung)",
        "file_finger": "1. File dữ liệu bấm vân tay (Có cột Ngày và Thứ)",
        "file_schedule": "2. File Lịch xếp ca (Chỉ áp dụng cho Công nhân)",
        "file_vp": "3. Danh sách nhân viên Văn phòng (VP)",
        "file_cn": "4. Danh sách Công nhân nhà máy",
        "filter_section": "📆 Bộ Lọc Thời Gian Kiểm Tra Trên Dashboard",
        "select_date": "Chọn ngày, tháng, năm cần kiểm tra:",
        "analyze_btn": "🚀 Bắt Đầu Đối Chiếu & Kết Xuất Thống Kê",
        "dashboard_stat": "📈 Biểu Đồ Thống Kê Phân Tích Chuyên Nghiệp",
        "abnormal_focus": "🚨 Trọng Tâm Các Trường Hợp Bất Thường (Rút Gọn)",
        "total_staff": "Tổng nhân sự dự kiến",
        "total_present": "Đi làm đúng / đủ",
        "total_abnormal": "Số ca bất thường",
        "total_absent": "Vắng mặt",
        "export_section": "📥 Xuất Báo Cáo Thống Kê",
        "download_excel": "Tải file kết quả Excel (.xlsx)",
        "download_pdf": "Tải file giao diện PDF",
        "col_eid": "Mã NV", "col_name": "Họ và tên", "col_in": "Giờ vào", 
        "col_out": "Giờ ra", "col_hours": "Giờ làm thực tế", "col_note": "Ghi chú",
        "note_sunday": "Ngày nghỉ", "note_normal": "Bình thường"
    },
    "zh": {
        "title": "📊 考勤数据智能智能对账与多 file 分析系统",
        "upload_section": "📁 上传系统所需的 4 个文件",
        "btn_lang": "Chuyển sang tiếng Việt (切换至越南语)",
        "file_finger": "1. 指纹打卡原始数据 (包含日期与星期列)",
        "file_schedule": "2. 排班表文件 (仅适用于车间工人)",
        "file_vp": "3. 办公室人员花名册 (VP)",
        "file_cn": "4. 车间工人花名册",
        "filter_section": "📆 Dashboard 检查时间筛选",
        "select_date": "选择需要检查的年/月/日:",
        "analyze_btn": "🚀 开始智能对账与统计输出",
        "dashboard_stat": "📈 专业考勤图表分析摘要",
        "abnormal_focus": "🚨 异常考勤重点关注摘要 (精简)",
        "total_staff": "预计总出勤人数",
        "total_present": "正常出勤人数",
        "total_abnormal": "异常打卡人数",
        "total_absent": "旷工/缺勤人数",
        "export_section": "📥 导出统计报告",
        "download_excel": "下载 Excel 统计表 (.xlsx)",
        "download_pdf": "下载当前 PDF 报表页面",
        "col_eid": "工号", "col_name": "姓名", "col_in": "签到时间", 
        "col_out": "签退時間", "col_hours": "实际工时", "col_note": "考勤备注",
        "note_sunday": "周休/放假", "note_normal": "正常"
    }
}

l = LANG_DICT[st.session_state.lang]

if st.sidebar.button(l["btn_lang"]):
    st.session_state.lang = "zh" if st.session_state.lang == "vi" else "vi"
    st.rerun()

st.title(l["title"])

# ==========================================
# 1. KẾT NỐI HỆ THỐNG TRÍ TUỆ ĐIỆN TỬ (GEMINI API)
# ==========================================
if "GEMINI_API_KEY" in st.secrets:
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.error("Chưa cấu hình khóa bảo mật GEMINI_API_KEY trên Streamlit Cloud!")
    st.stop()

def dynamic_column_mapping(columns_list, target_concept):
    prompt = f"""
    Bạn là một chuyên gia phân tích cấu trúc dữ liệu bảng tính. Tôi có danh sách các tên cột như sau: {columns_list}
    Hãy tìm ra một tên cột khớp nhất với khái niệm nghiệp vụ: '{target_concept}'.
    Yêu cầu bắt buộc: Chỉ trả về duy nhất tên cột chính xác tuyệt đối lấy từ danh sách trên, không kèm giải thích hay dấu câu. Nếu không tìm thấy, trả về 'NONE'.
    """
    try:
        response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
        return response.text.strip()
    except:
        return "NONE"

# ==========================================
# 2. GIAO DIỆN TẢI FILE ĐẦU VÀO (ĐỦ 4 FILE)
# ==========================================
st.header(l["upload_section"])
c1, c2 = st.columns(2)
c3, c4 = st.columns(2)

with c1: f_finger = st.file_uploader(l["file_finger"], type=["xlsx", "xls"])
with c2: f_schedule = st.file_uploader(l["file_schedule"], type=["xlsx", "xls"])
with c3: f_vp = st.file_uploader(l["file_vp"], type=["xlsx", "xls"])
with c4: f_cn = st.file_uploader(l["file_cn"], type=["xlsx", "xls"])

# ==========================================
# 3. BỘ LỌC NGÀY/THÁNG/NĂM CHỌN TỪ DASHBOARD
# ==========================================
st.header(l["filter_section"])
target_date = st.date_input(l["select_date"], datetime.date(2026, 9, 6))

final_report = pd.DataFrame()

if f_finger and f_schedule and f_vp and f_cn:
    if st.button(l["analyze_btn"], type="primary"):
        with st.spinner("Hệ thống đang đồng bộ mã nhân viên và phân tích lịch trình..."):
            
            # Đọc dữ liệu từ 4 file Excel
            df_finger = pd.read_excel(f_finger)
            df_sched = pd.read_excel(f_schedule)
            df_master_vp = pd.read_excel(f_vp)
            df_master_cn = pd.read_excel(f_cn)
            
            # --- Nhận diện thông minh các cột mấu chốt qua Trí tuệ điện tử ---
            id_finger = dynamic_column_mapping(list(df_finger.columns), "mã số nhân viên hoặc số thẻ chấm công")
            date_finger = dynamic_column_mapping(list(df_finger.columns), "ngày hoặc ngày tháng năm")
            time_finger = dynamic_column_mapping(list(df_finger.columns), "thời gian quét vân tay hoặc giờ bấm thẻ")
            
            id_vp = dynamic_column_mapping(list(df_master_vp.columns), "mã nhân viên văn phòng")
            name_vp = dynamic_column_mapping(list(df_master_vp.columns), "họ và tên nhân viên văn phòng")
            
            id_cn = dynamic_column_mapping(list(df_master_cn.columns), "mã công nhân nhà máy")
            name_cn = dynamic_column_mapping(list(df_master_cn.columns), "họ và tên công nhân")
            
            id_sched = dynamic_column_mapping(list(df_sched.columns), "mã số nhân viên hoặc mã công nhân trong lịch ca")

            # Chuẩn hóa cột ngày tháng của dữ liệu vân tay
            df_finger['Standard_Date'] = pd.to_datetime(df_finger[date_finger], errors='coerce').dt.date
            df_finger_today = df_finger[df_finger['Standard_Date'] == target_date]
            
            rows_output = []
            is_sunday = (target_date.weekday() == 6) # Kiểm tra xem ngày chọn có phải Chủ Nhật hay không

            # ==========================================
            # LOGIC ĐỐI CHIẾU KHỐI VĂN PHÒNG (VP)
            # ==========================================
            for _, row in df_master_vp.iterrows():
                eid = str(row[id_vp]).strip()
                name = str(row[name_vp]).strip()
                
                # Dò tìm Mã NV trong dữ liệu vân tay của ngày hôm đó
                emp_logs = df_finger_today[df_finger_today[id_finger].astype(str).str.strip() == eid]
                
                # Trường hợp đặc biệt: Nếu ngày chọn là Chủ Nhật
                if is_sunday:
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: "-", l["col_out"]: "-",
                        l["col_hours"]: 0.0, l["col_note"]: l["note_sunday"]
                    })
                    continue
                
                # Cấu hình giờ chuẩn theo quy tắc
                if eid == "575":
                    std_in, std_out, expected_h = "07:00:00", "15:00:00", 12.0
                elif eid in ["749", "949"]:
                    std_in, std_out, expected_h = "07:00:00", "19:00:00", 12.0
                else:
                    std_in, std_out, expected_h = "08:00:00", "17:00:00", 8.0
                
                if emp_logs.empty:
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: "-", l["col_out"]: "-",
                        l["col_hours"]: 0.0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                    })
                else:
                    # Lấy giờ quẹt đầu (Vào) và quẹt cuối (Ra)
                    df_time_sorted = emp_logs.sort_values(by=time_finger)
                    t_in_raw = df_time_sorted[time_finger].iloc[0]
                    t_out_raw = df_time_sorted[time_finger].iloc[-1] if len(df_time_sorted) > 1 else None
                    
                    notes = []
                    actual_hours = 0.0
                    
                    if t_in_raw and t_out_raw:
                        # Chuyển đổi tính toán giờ thực tế
                        t_in_dt = pd.to_datetime(t_in_raw, format='%H:%M:%S', errors='coerce') if isinstance(t_in_raw, str) else pd.to_datetime(str(t_in_raw))
                        t_out_dt = pd.to_datetime(t_out_raw, format='%H:%M:%S', errors='coerce') if isinstance(t_out_raw, str) else pd.to_datetime(str(t_out_raw))
                        
                        actual_hours = round((t_out_dt - t_in_dt).total_seconds() / 3600, 1)
                        
                        # So sánh giờ quy định
                        limit_in = datetime.time.fromisoformat(std_in)
                        limit_out = datetime.time.fromisoformat(std_out)
                        
                        if t_in_dt.time() > (datetime.datetime.combine(datetime.date.today(), limit_in) + datetime.timedelta(minutes=5)).time():
                            notes.append("Đi trễ" if st.session_state.lang=="vi" else "迟到")
                        if t_out_dt.time() < limit_out:
                            notes.append("Về sớm" if st.session_state.lang=="vi" else "早退")
                        if actual_hours < expected_h and "Về sớm" not in notes:
                            notes.append(f"Về sớm ({actual_hours}h)")
                    else:
                        if not t_in_raw: notes.append("BV")
                        if not t_out_raw: notes.append("BR")
                        
                    note_str = ", ".join(notes) if notes else l["note_normal"]
                    rows_output.append({
                        l["col_eid"]: eid, l["col_name"]: name,
                        l["col_in"]: str(t_in_raw)[:5] if t_in_raw else "-",
                        l["col_out"]: str(t_out_raw)[:5] if t_out_raw else "-",
                        l["col_hours"]: actual_hours, l["col_note"]: note_str
                    })

            # ==========================================
            # LOGIC ĐỐI CHIẾU KHỐI CÔNG NHÂN (THEO LỊCH XẾP CA)
            # ==========================================
            day_str_target = str(target_date.day)
            sched_day_col = [c for c in df_sched.columns if str(c).strip() == day_str_target]
            
            if sched_day_col:
                target_day_col_name = sched_day_col[0]
                
                for _, row in df_master_cn.iterrows():
                    eid = str(row[id_cn]).strip()
                    name = str(row[name_cn]).strip()
                    
                    # Dò lịch xếp ca của công nhân bằng Mã NV
                    sched_row = df_sched[df_sched[id_sched].astype(str).str.strip() == eid]
                    shift_type = "OFF"
                    
                    if not sched_row.empty:
                        shift_type = str(sched_row[target_day_col_name].values[0]).strip()
                        
                    if shift_type in ["OFF", "nan", "-", "Nghỉ"]:
                        rows_output.append({
                            l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: "-", l["col_out"]: "-",
                            l["col_hours"]: 0.0, l["col_note"]: "Nghỉ ca" if st.session_state.lang=="vi" else "轮休"
                        })
                        continue
                        
                    # Dò tìm dữ liệu vân tay thực tế bằng Mã NV
                    emp_logs = df_finger_today[df_finger_today[id_finger].astype(str).str.strip() == eid]
                    
                    if shift_type == "N": # Ca Ngày (07:00 AM -> 19:00 PM)
                        if emp_logs.empty:
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: "-", l["col_out"]: "-",
                                l["col_hours"]: 0.0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                            })
                        else:
                            df_time_sorted = emp_logs.sort_values(by=time_finger)
                            t_in = df_time_sorted[time_finger].iloc[0]
                            t_out = df_time_sorted[time_finger].iloc[-1] if len(df_time_sorted) > 1 else None
                            notes = []
                            
                            note_str = ", ".join(notes) if notes else "N"
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: str(t_in)[:5],
                                l["col_out"]: str(t_out)[:5] if t_out else "-", l["col_hours"]: 12.0 if t_out else 0, l["col_note"]: note_str
                            })
                            
                    elif shift_type == "Đ": # Ca Đêm (19:00 PM -> 07:00 AM hôm sau)
                        t_in_row = emp_logs[emp_logs[time_finger].astype(str) >= "18:00:00"]
                        t_in = t_in_row[time_finger].min() if not t_in_row.empty else None
                        
                        next_day = target_date + datetime.timedelta(days=1)
                        df_finger_next = df_finger[df_finger['Standard_Date'] == next_day]
                        emp_logs_next = df_finger_next[df_finger_next[id_finger].astype(str).str.strip() == eid]
                        
                        t_out_row = emp_logs_next[emp_logs_next[time_finger].astype(str) <= "08:00:00"]
                        t_out = t_out_row[time_finger].max() if not t_out_row.empty else None
                        
                        if not t_in and not t_out:
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: "-", l["col_out"]: "-",
                                l["col_hours"]: 0.0, l["col_note"]: "Vắng" if st.session_state.lang=="vi" else "旷工"
                            })
                        else:
                            notes = []
                            note_str = ", ".join(notes) if notes else "Đ"
                            rows_output.append({
                                l["col_eid"]: eid, l["col_name"]: name, l["col_in"]: str(t_in)[:5] if t_in else "-",
                                l["col_out"]: str(t_out)[:5] if t_out else "-", l["col_hours"]: 12.0 if (t_in and t_out) else 0,
                                l["col_note"]: note_str
                            })
                            
            final_report = pd.DataFrame(rows_output)
            st.session_state.final_report = final_report

# ==========================================
# 4. HIỂN THỊ THỐNG KÊ & PHÂN TÍCH TRỌNG TÂM
# ==========================================
if "final_report" in st.session_state and not st.session_state.final_report.empty:
    df_res = st.session_state.final_report
    st.header(l["dashboard_stat"])
    
    m_total = len(df_res)
    m_absent = len(df_res[df_res[l["col_note"]].str.contains("Vắng|旷工", na=False)])
    m_abnormal = len(df_res[df_res[l["col_note"]].str.contains("Đi trễ|Về sớm|lịch|迟到|早退", na=False)])
    m_present = m_total - m_absent - m_abnormal
    
    mx1, mx2, mx3, mx4 = st.columns(4)
    mx1.metric(l["total_staff"], m_total)
    mx2.metric(l["total_present"], m_present)
    mx3.metric(l["total_abnormal"], m_abnormal, delta_color="inverse")
    mx4.metric(l["total_absent"], m_absent, delta_color="inverse")
    
    fig_pie = px.pie(
        names=[l["total_present"], l["total_abnormal"], l["total_absent"]],
        values=[m_present, m_abnormal, m_absent],
        color_discrete_sequence=px.colors.qualitative.Set3,
        title="Phần Trăm Tình Trạng Đi Làm Thực Tế / 实际出勤率"
    )
    st.plotly_chart(fig_pie, use_container_width=True)
    
    st.header(l["abnormal_focus"])
    df_bad = df_res[df_res[l["col_note"]].str.contains("Vắng|Đi trễ|Về sớm|lịch|旷工|迟到|早退", na=False)]
    
    if not df_bad.empty:
        summary_prompt = f"""
        Dựa trên bảng tổng hợp lỗi chấm công sau đây, hãy viết một báo cáo phân tích trọng tâm siêu ngắn gọn (dưới 4 dòng), ghi rõ mã nhân viên nào vi phạm nghiêm trọng nhất. Không giải thích dài dòng.
        Dữ liệu lỗi:
        {df_bad[[l["col_eid"], l["col_name"], l["col_note"]]].to_string(index=False)}
        Yêu cầu ngôn ngữ: Trả về dạng song ngữ Trung - Việt. Tuyệt đối không được dùng chữ 'AI' trong toàn bộ văn bản.
        """
        summary_response = client.models.generate_content(model='gemini-2.5-flash', contents=summary_prompt)
        st.warning(summary_response.text)
    else:
        st.success("Tất cả nhân sự đi làm bình thường, không phát hiện lỗi! / 今日考勤全部正常。")
        
    st.dataframe(df_res[[l["col_eid"], l["col_name"], l["col_in"], l["col_out"], l["col_hours"], l["col_note"]]], use_container_width=True)

# ==========================================
# 5. XUẤT DOWNLOAD BÁO CÁO
# ==========================================
st.header(l["export_section"])

buffer_excel = io.BytesIO()
with pd.ExcelWriter(buffer_excel, engine='xlsxwriter') as writer:
    if not final_report.empty:
        final_report.to_excel(writer, index=False, sheet_name='ThongKe_Cong_Bieu')

c_down1, c_down2 = st.columns(2)
with c_down1:
    st.download_button(
        label=l["download_excel"],
        data=buffer_excel.getvalue(),
        file_name=f"Bao_Cao_Cham_Cong_{target_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with c_down2:
    st.button(l["download_pdf"], help="Nhấn tổ hợp phím Ctrl + P để lưu file PDF giữ nguyên toàn bộ giao diện dashboard đồ họa.")

