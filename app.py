import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import plotly.express as px
from google import genai
from google.genai import types
from fpdf import FPDF
import io

# ==========================================
# 1. CẤU HÌNH TRANG & CÀI ĐẶT CHUNG
# ==========================================
st.set_page_config(page_title="系统 - 报表", layout="wide", initial_sidebar_state="expanded")

# Thiết lập Từ điển Song ngữ Trung - Việt (Yêu cầu 11: Không chứa từ "AI")
LANG = {
    "vi": {
       "title": "HỆ THỐNG QUẢN LÝ VÀ THỐNG KÊ NHÂN VIÊN ĐI LÀM",
       "upload_section": "TẢI LÊN DỮ LIỆU ĐẦU VÀO",
       "btn_fingerprint": "1. Tải file bấm vân tay (Excel)",
       "btn_schedule": "2. Tải lịch xếp ca (Excel)",
       "btn_office_staff": "3. Tải danh sách nhân viên VP (Excel)",
       "btn_factory_staff": "4. Tải danh sách công nhân (Excel)",
       "filter_section": "BỘ LỌC THỜI GIAN THỐNG KÊ",
       "select_date": "Chọn Ngày",
       "select_month": "Chọn Tháng",
       "select_year": "Chọn Năm",
       "process_status": "Trạng thái xử lý hệ thống",
       "processing": "Hệ thống phân tích đang đọc cấu trúc file và xử lý...",
       "success": "Xử lý dữ liệu thành công!",
       "error_api": "Không thể kết nối với Hệ thống phân tích. Đang chuyển sang chế độ tự động quét tiêu đề bằng thuật toán Python...",
       "error_file": "Vui lòng tải đầy đủ các file dữ liệu để hệ thống thống kê.",
       "dashboard_title": "BẢNG THỐNG KÊ CHI TIẾT NGÀY",
       "chart_title": "BIỂU ĐỒ PHÂN TÍCH TRỰC QUAN TRẠNG THÁI NHÂN SỰ",
       "col_id": "Mã NV",
       "col_name": "Họ và tên",
       "col_in": "Giờ vào",
       "col_out": "Giờ ra",
       "col_hours": "Giờ làm thực tế",
       "col_note": "Ghi chú",
       "download_excel": "Xuất file Excel thống kê",
       "download_pdf": "Xuất file PDF giao diện Dashboard",
       "stat_total": "Tổng số nhân sự dự kiến",
       "stat_present": "Đi làm",
       "stat_absent": "Vắng",
       "stat_late": "Đi trễ / Về sớm / Sai lịch"
    },
    "zh": {
       "title": "员工出勤管理与统计系统",
       "upload_section": "输入数据上传",
       "btn_fingerprint": "1. 上传指纹打卡文件 (Excel)",
       "btn_schedule": "2. 上传排班表 (Excel)",
       "btn_office_staff": "3. 上传办公室人员名单 (Excel)",
       "btn_factory_staff": "4. 上传工人名单 (Excel)",
       "filter_section": "统计时间筛选",
       "select_date": "选择日期",
       "select_month": "选择月份",
       "select_year": "选择年份",
       "process_status": "系统处理状态",
       "processing": "分析系统正在读取文件结构并进行处理...",
       "success": "数据处理成功！",
       "error_api": "无法连接到分析系统。 正在切换到Python算法自动扫描表头...",
       "error_file": "请上传完整的各项数据文件以进行统计。",
       "dashboard_title": "当日详细出勤统计表",
       "chart_title": "人员出勤状态直观分析图表",
       "col_id": "工号",
       "col_name": "姓名",
       "col_in": "签到时间",
       "col_out": "签退时间",
       "col_hours": "实际工时",
       "col_note": "备注",
       "download_excel": "导出 Excel 统计报表",
       "download_pdf": "导出 PDF 看板界面",
       "stat_total": "预计总人数",
       "stat_present": "出勤",
       "stat_absent": "缺勤",
       "stat_late": "迟到 / 早退 / 异常"
    }
}

# Chọn ngôn ngữ hiển thị mặc định
if 'lang_idx' not in st.session_state:
    st.session_state.lang_idx = "vi"

col_l1, col_l2 = st.sidebar.columns(2)
if col_l1.button("Tiếng Việt"):
    st.session_state.lang_idx = "vi"
if col_l2.button("简体中文"): 
    st.session_state.lang_idx = "zh"

L = LANG[st.session_state.lang_idx]

st.title(f"📊 {L['title']}")
st.write("---")

# Kích hoạt kết nối Gemini API thông qua Streamlit Secrets hoặc biến môi trường
GEMINI_KEY = st.secrets.get("GEMINI_API_KEY", None)

# ==========================================
# 2. HÀM TRỢ GIÚP DÙNG GEMINI AI ĐỒNG NHẤT CỘT
# ==========================================
def ask_gemini_for_headers(columns_list, file_type_desc):
    """
    Sử dụng Gemini AI để phân tích danh sách tiêu đề thực tế từ file excel 
    và ánh xạ về các từ khóa chuẩn hóa (Nguyên lý vận hành).
    """
    if not GEMINI_KEY:
        return None
    try:
        client = genai.Client(api_key=GEMINI_KEY)
        prompt = f"""
        Bạn là một kỹ sư xử lý dữ liệu hệ thống nhân sự. Bạn nhận được danh sách các cột thực tế từ một file Excel {file_type_desc}.
        Các cột thực tế nhận được là: {str(columns_list)}
        
        Nhiệm vụ của bạn là phân tích cấu trúc và tìm ra cột nào tương ứng với các trường thông tin chuẩn sau:
        1. 'ma_nv' (Mã nhân viên / 工号)
        2. 'ho_ten' (Họ và tên / 姓名)
        3. 'ngay' (Ngày tháng / 日期)
        4. 'gio_vao' (Giờ vào / 签到时间 / 上班时间)
        5. 'gio_ra' (Giờ ra / 签退时间 / 下班时间)
        6. 'tong_gio' (Tổng giờ / 实际工时)
        7. 'thu' (Thứ trong tuần / 星期)
        
        Hãy chỉ trả về kết quả dưới dạng cấu trúc JSON sạch, định dạng đúng như sau (ví dụ nếu cột thực tế là 'Mã Số NV' thì gán 'ma_nv': 'Mã Số NV'). Không thêm bất kỳ văn bản giải thích nào khác ngoài JSON:
        {{
           "ma_nv": "tên cột thực tế",
           "ho_ten": "tên cột thực tế hoặc null nếu không có",
           "ngay": "tên cột thực tế hoặc null nếu không có",
           "gio_vao": "tên cột thực tế hoặc null nếu không có",
           "gio_ra": "tên cột thực tế hoặc null nếu không có",
           "tong_gio": "tên cột thực tế hoặc null nếu không có",
           "thu": "tên cột thực tế hoặc null nếu không có"
        }}
        """
        response = client.models.generate_content(
           model='gemini-2.5-flash',
           contents=prompt,
           config=types.GenerateContentConfig(response_mime_type="application/json")
        )
        import json
        mapping = json.loads(response.text.strip())
        return mapping
    except Exception as e:
        return None

def fallback_header_match(columns_list):
    """
    Hàm Python tự động quét tiêu đề dự phòng nếu Gemini không phản hồi (Bước 2 của Nguyên lý vận hành).
    """
    mapping = {"ma_nv": None, "ho_ten": None, "ngay": None, "gio_vao": None, "gio_ra": None, "tong_gio": None, "thu": None}
    for c in columns_list:
        c_lower = str(c).lower()
        if 'mã' in c_lower or 'ma' in c_lower or 'id' in c_lower or '工号' in c_lower:
           mapping['ma_nv'] = c
        elif 'tên' in c_lower or 'ten' in c_lower or 'ho' in c_lower or '姓名' in c_lower:
           mapping['ho_ten'] = c
        elif 'ngày' in c_lower or 'ngay' in c_lower or '日期' in c_lower:
           mapping['ngay'] = c
        elif 'vào' in c_lower or 'vao' in c_lower or 'in' in c_lower or '签到' in c_lower or '上' in c_lower:
           mapping['gio_vao'] = c
        elif 'ra' in c_lower or 'out' in c_lower or '签退' in c_lower or '下' in c_lower:
           mapping['gio_ra'] = c
        elif 'tổng' in c_lower or 'tong' in c_lower or '工时' in c_lower or 'duration' in c_lower:
           mapping['tong_gio'] = c
        elif 'thứ' in c_lower or 'thu' in c_lower or '星期' in c_lower or '周' in c_lower:
           mapping['thu'] = c
    return mapping

# ==========================================
# 3. GIAO DIỆN TẢI FILE & BỘ LỌC (Yêu cầu 1,2,3,4,5)
# ==========================================
st.sidebar.header(f"📥 {L['upload_section']}")
file_fingerprint = st.sidebar.file_uploader(L['btn_fingerprint'], type=["xlsx", "xls"])
file_schedule = st.sidebar.file_uploader(L['btn_schedule'], type=["xlsx", "xls"])
file_office = st.sidebar.file_uploader(L['btn_office_staff'], type=["xlsx", "xls"])
file_factory = st.sidebar.file_uploader(L['btn_factory_staff'], type=["xlsx", "xls"])

st.sidebar.header(f"📅 {L['filter_section']}")
today = datetime.now()
selected_year = st.sidebar.selectbox(L['select_year'], list(range(2020, 2031)), index=list(range(2020, 2031)).index(today.year))
selected_month = st.sidebar.selectbox(L['select_month'], list(range(1, 13)), index=today.month - 1)
selected_date = st.sidebar.selectbox(L['select_date'], list(range(1, 32)), index=today.day - 1)

target_date_str = f"{selected_year}-{selected_month:02d}-{selected_date:02d}"
target_datetime = datetime(selected_year, selected_month, selected_date)

# ==========================================
# 4. THỰC THI LOGIC XỬ LÝ CHÍNH
# ==========================================
if file_fingerprint and file_schedule and file_office and file_factory:
    with st.spinner(L['processing']):
        try:
            # Đọc thô các file excel đầu vào
            df_finger_raw = pd.read_excel(file_fingerprint)
            df_sched_raw = pd.read_excel(file_schedule)
            df_off_raw = pd.read_excel(file_office)
            df_fac_raw = pd.read_excel(file_factory)
            
            # --- Thực hiện Nguyên Lý Vận Hành: Nhận diện cột qua Gemini AI ---
            map_finger = ask_gemini_for_headers(list(df_finger_raw.columns), "chứa dữ liệu quẹt vân tay hàng ngày")
            if not map_finger:
               st.sidebar.warning(L['error_api'])
               map_finger = fallback_header_match(list(df_finger_raw.columns))
               map_off = fallback_header_match(list(df_off_raw.columns))
               map_fac = fallback_header_match(list(df_fac_raw.columns))
            else:
               map_off = ask_gemini_for_headers(list(df_off_raw.columns), "danh sách nhân viên văn phòng") or fallback_header_match(list(df_off_raw.columns))
               map_fac = ask_gemini_for_headers(list(df_fac_raw.columns), "danh sách công nhân nhà máy") or fallback_header_match(list(df_fac_raw.columns))
            
            # --- Chuẩn hóa danh sách Nhân viên Văn phòng & Công nhân ---
            # Nhóm 1: Công nhân
            df_fac_raw = df_fac_raw.rename(columns={map_fac['ma_nv']: 'Mã NV', map_fac['ho_ten']: 'Họ và tên'})
            df_fac_raw['Mã NV'] = df_fac_raw['Mã NV'].astype(str).str.strip()
            workers_list = df_fac_raw[['Mã NV', 'Họ và tên']].dropna(subset=['Mã NV']).drop_duplicates()
            workers_list['Nhóm'] = 'Công nhân'
            
            # Nhóm 2: Văn phòng (Loại bỏ các mã đặc biệt thuộc nhóm 3, 4, 5 nếu có)
            df_off_raw = df_off_raw.rename(columns={map_off['ma_nv']: 'Mã NV', map_off['ho_ten']: 'Họ và tên'})
            df_off_raw['Mã NV'] = df_off_raw['Mã NV'].astype(str).str.strip()
            office_list = df_off_raw[['Mã NV', 'Họ và tên']].dropna(subset=['Mã NV']).drop_duplicates()
            
            special_ids = ['673', 'A068', '749', '949', '575']
            office_list = office_list[~office_list['Mã NV'].isin(special_ids)]
            office_list['Nhóm'] = 'Văn phòng'
            
            # Hợp nhất cây danh mục nhân sự tổng thể dựa trên Mã NV
            all_staff = pd.concat([office_list, workers_list], ignore_index=True)
            
            # Thêm các nhóm đặc biệt theo định nghĩa của đề bài vào cây danh mục chính nếu tồn tại trong dữ liệu thô
            all_staff.loc[all_staff['Mã NV'].isin(['673', 'A068']), 'Nhóm'] = 'Bảo trì'
            all_staff.loc[all_staff['Mã NV'] == '749', 'Nhóm'] = 'QC 749'
            all_staff.loc[all_staff['Mã NV'] == '949', 'Nhóm'] = 'QC 949'
            all_staff.loc[all_staff['Mã NV'] == '575', 'Nhóm'] = 'Tạp vụ'
            
            # Đảm bảo các nhân viên đặc biệt luôn nằm trong danh sách kiểm tra dù không có trong file danh mục gốc
            for sp_id, grp, name in [('673', 'Bảo trì', 'Nhân viên Bảo Trì 673'), ('A068', 'Bảo trì', 'Nhân viên Bảo Trì A068'),
                                     ('749', 'QC 749', 'Nhân viên QC 749'), ('949', 'QC 949', 'Nhân viên QC 949'),
                                     ('575', 'Tạp vụ', 'Nhân viên Tạp Vụ 575')]:
                if sp_id not in all_staff['Mã NV'].values:
                    all_staff = pd.concat([all_staff, pd.DataFrame([{'Mã NV': sp_id, 'Họ và tên': name, 'Nhóm': grp}])], ignore_index=True)
            
            # --- Chuẩn hóa dữ liệu File Bấm Vân Tay ---
            df_finger = df_finger_raw.rename(columns=map_finger)
            df_finger['ma_nv'] = df_finger['ma_nv'].astype(str).str.strip()
            
            # Ép kiểu dữ liệu ngày tháng chuẩn hóa để truy vấn đúng ngày chọn trên Dashboard
            df_finger['ngay_parsed'] = pd.to_datetime(df_finger['ngay'], errors='coerce')
            df_day_data = df_finger[df_finger['ngay_parsed'].dt.strftime('%Y-%m-%d') == target_date_str]
            
            # --- MẢNG LƯU TRỮ KẾT QUẢ THỐNG KÊ CUỐI CÙNG ---
            final_rows = []
            
            # --- VÒNG LẶP DUYỆT QUA TỪNG NHÂN VIÊN ĐỂ ĐỐI CHIẾU QUY TẮC ---
            for idx, row in all_staff.iterrows():
                emp_id = row['Mã NV']
                emp_name = row['Họ và tên']
                emp_group = row['Nhóm']
                
                # Trích xuất dòng quẹt thẻ tương ứng của nhân viên trong ngày đang kiểm tra
                emp_log = df_day_data[df_day_data['ma_nv'] == emp_id]
                
                # Khởi tạo các giá trị trống mặc định theo cấu trúc yêu cầu 13
                g_vao, g_ra, g_lam = None, None, 0.0
                ghi_chu = ""
                
                # Đọc thông tin giờ từ file vân tay nếu có quẹt thẻ
                if not emp_log.empty:
                    log_row = emp_log.iloc[0]
                    g_vao = log_row.get('gio_vao', None)
                    g_ra = log_row.get('gio_ra', None)
                    try:
                        g_lam = float(log_row.get('tong_gio', 0.0))
                    except:
                        g_lam = 0.0
                    
                    # Chuyển đổi định dạng giờ hiển thị sang String sạch
                    g_vao_str = str(g_vao) if pd.notna(g_vao) else "BR"
                    g_ra_str = str(g_ra) if pd.notna(g_ra) else "BV"
                else:
                    g_vao_str = "Vắng"
                    g_ra_str = "Vắng"
                    g_lam = 0.0
                
                # ----------------------------------------------------
                # ÁP DỤNG QUY TẮC LOGIC PHÂN LOẠI CHO TỪNG NHÓM (Yêu cầu 8, 13 & Quy tắc)
                # ----------------------------------------------------
                is_sunday = (target_datetime.weekday() == 6)
                
                # Xử lý Quy tắc chung về thiếu giờ: Thiếu Vào = BV, Thiếu Ra = BR. Cả 2 thiếu = Vắng.
                if g_vao_str == "Vắng" and g_ra_str == "Vắng":
                    status_vắng = True
                else:
                    status_vắng = False
                
                # 1. Nhóm Văn Phòng
                if emp_group == 'Văn phòng':
                    if is_sunday:
                        ghi_chu = "Chủ nhật (Nghỉ)" if not status_vắng else "Nghỉ tuần"
                    elif status_vắng:
                        ghi_chu = "Vắng"
                    else:
                        # Kiểm tra đi trễ / về sớm
                        notes_list = []
                        if pd.notna(g_vao):
                            try:
                                vao_time = pd.to_datetime(str(g_vao)).time()
                                if vao_time > time(8, 0): notes_list.append("Đi trễ")
                            except: pass
                        else: notes_list.append("BV")
                        
                        if pd.notna(g_ra):
                            try:
                                ra_time = pd.to_datetime(str(g_ra)).time()
                                if ra_time < time(17, 0): notes_list.append("Về sớm")
                            except: pass
                        else: notes_list.append("BR")
                        
                        if g_lam < 8.0:
                            notes_list.append(f"Về sớm ({g_lam}h)")
                            
                        ghi_chu = ", ".join(list(set(notes_list))) if notes_list else "Đủ tiêu chuẩn"
                
                # 2. Nhóm 5: Tạp vụ (Mã 575: 7:00 AM - 15:00 PM cùng ngày)
                elif emp_group == 'Tạp vụ':
                    if status_vắng:
                        ghi_chu = "Vắng"
                    else:
                        notes_list = []
                        if g_lam < 12.0: notes_list.append(f"Làm không đủ lịch ({g_lam}h)")
                        ghi_chu = ", ".join(notes_list) if notes_list else "Đủ tiêu chuẩn"
                
                # 3. Nhóm 4: QC (Mã 749 & 949: 7:00 AM - 19:00 PM)
                elif emp_group in ['QC 749', 'QC 949']:
                    if status_vắng:
                        ghi_chu = "Vắng"
                    else:
                        notes_list = []
                        if g_lam < 12.0: notes_list.append(f"Làm không đủ lịch ({g_lam}h)")
                        ghi_chu = ", ".join(notes_list) if notes_list else "Đủ tiêu chuẩn"
                
                # 4. Nhóm 3: Bảo trì (Mã 673 & A068 - Theo dõi dựa trên thực tế vân tay)
                elif emp_group == 'Bảo trì':
                    if status_vắng:
                        ghi_chu = "Vắng"
                    else:
                        ghi_chu = "Đủ tiêu chuẩn" if g_lam >= 8 else f"Làm không đủ ca ({g_lam}h)"
                
                # 5. Nhóm 1: Công nhân (Xử lý dựa trên Lịch xếp ca ca Ngày 'N' / ca Đêm 'Đ')
                elif emp_group == 'Công nhân':
                    # Định vị dòng xếp ca của công nhân
                    emp_sched = df_sched_raw[df_sched_raw.iloc[:, 0].astype(str).str.strip() == emp_id]
                    shift_type = ""
                    if not emp_sched.empty:
                        # Giả định lịch xếp ca có tiêu đề cột chứa thông tin Ngày hoặc dùng vị trí số cột tương ứng ngày chọn
                        try:
                            # Quét tìm cột tương ứng với ngày trong tháng
                            col_target = [c for c in df_sched_raw.columns if str(selected_date) == str(c) or f"{selected_date:02d}" in str(c)]
                            if col_target:
                                shift_type = str(emp_sched.iloc[0][col_target[0]]).strip()
                        except:
                            shift_type = ""
                            
                    if shift_type in ['N', 'Đ']:
                        if status_vắng:
                            ghi_chu = "Vắng"
                        else:
                            if g_lam < 12.0:
                                ghi_chu = f"Làm không đúng lịch ({g_lam}h)"
                            else:
                                ghi_chu = f"Làm ca {shift_type}"
                    else:
                        # Các ô màu đỏ/nâu hoặc trống ký hiệu là ngày nghỉ hàng tuần theo lịch xếp ca
                        if not status_vắng:
                            ghi_chu = "Làm ngày nghỉ"
                        else:
                            ghi_chu = "Nghỉ tuần"
                
                # Chèn hàng dữ liệu hoàn chỉnh
                final_rows.append({
                    "Mã NV": emp_id,
                    "Họ và tên": emp_name,
                    "Giờ vào": g_vao_str,
                    "Giờ ra": g_ra_str,
                    "Giờ làm thực tế": g_lam,
                    "Ghi chú": ghi_chu if ghi_chu != "" else "Vắng"
                })
                
            df_result = pd.DataFrame(final_rows)
            
            # ==========================================
            # 5. HIỂN THỊ KẾT QUẢ KANBAN CARD & THỐNG KÊ (Yêu cầu 10)
            # ==========================================
            st.success(L['success'])
            
            # Chỉ số thẻ KPI Tổng quan
            total_staff_count = len(df_result)
            absent_count = len(df_result[df_result['Ghi chú'] == 'Vắng'])
            present_count = total_staff_count - absent_count
            late_or_early = len(df_result[df_result['Ghi chú'].str.contains('trễ|sớm|không đúng|Khuyết', case=False, na=False)])
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(L['stat_total'], total_staff_count)
            m2.metric(L['stat_present'], present_count)
            m3.metric(L['stat_absent'], absent_count, delta_color="inverse")
            m4.metric(L['stat_late'], late_or_early)
            
            # --- Biểu đồ phân tích trực quan chuyên nghiệp ---
            st.subheader(f"📈 {L['chart_title']}")
            fig_data = df_result['Ghi chú'].value_counts().reset_index()
            fig_data.columns = ['Trạng thái / 状态', 'Số lượng / 数量']
            fig = px.pie(fig_data, values='Số lượng / 数量', names='Trạng thái / 状态', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            fig.update_layout(margin=dict(l=20, r=20, t=30, b=20))
            st.plotly_chart(fig, use_container_width=True)
            
            # --- Bảng hiển thị giao diện chính ---
            st.subheader(f"📋 {L['dashboard_title']}: {target_date_str}")
            
            # Đổi tiêu đề cột theo ngôn ngữ đã chọn
            df_display = df_result.rename(columns={
                "Mã NV": L['col_id'],
                "Họ và tên": L['col_name'],
                "Giờ vào": L['col_in'],
                "Giờ ra": L['col_out'],
                "Giờ làm thực tế": L['col_hours'],
                "Ghi chú": L['col_note']
            })
            st.dataframe(df_display, use_container_width=True)
            
            # ==========================================
            # 6. XUẤT FILE ĐẦU RA EXCEL & PDF (Yêu cầu 9)
            # ==========================================
            st.write("---")
            col_d1, col_d2 = st.columns(2)
            
            # Xuất dữ liệu Excel chuẩn hóa chỉ gồm các cột theo quy định (Yêu cầu 13)
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Thống kê chấm công')
            excel_data = output_excel.getvalue()
            
            col_d1.download_button(
                label=f"📥 {L['download_excel']}",
                data=excel_data,
                file_name=f"Thong_Ke_Cham_Cong_{target_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Xuất file PDF khớp định dạng thiết kế của Dashboard bảng dữ liệu
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=f"{L['title']}", ln=1, align="C")
            pdf.cell(200, 10, txt=f"Date: {target_date_str}", ln=2, align="C")
            pdf.ln(10)
            
            # Thêm tiêu đề bảng dữ liệu vào PDF
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(25, 8, L['col_id'], 1)
            pdf.cell(45, 8, L['col_name'], 1)
            pdf.cell(25, 8, L['col_in'], 1)
            pdf.cell(25, 8, L['col_out'], 1)
            pdf.cell(25, 8, L['col_hours'], 1)
            pdf.cell(45, 8, L['col_note'], 1)
            pdf.ln()
            
            # Đổ dữ liệu hàng vào bảng PDF
            pdf.set_font("Arial", size=9)
            for idx, r in df_result.iterrows():
                pdf.cell(25, 7, str(r["Mã NV"]), 1)
                pdf.cell(45, 7, str(r["Họ và tên"])[:20], 1)
                pdf.cell(25, 7, str(r["Giờ vào"]), 1)
                pdf.cell(25, 7, str(r["Giờ ra"]), 1)
                pdf.cell(25, 7, str(r["Giờ làm thực tế"]), 1)
                pdf.cell(45, 7, str(r["Ghi chú"])[:22], 1)
                pdf.ln()
                
            pdf_output = pdf.output()
            col_d2.download_button(
                label=f"📥 {L['download_pdf']}",
                data=bytes(pdf_output),
                file_name=f"Dashboard_Report_{target_date_str}.pdf",
                mime="application/pdf"
            )
            
        except Exception as ex:
            st.error(f"Đã xảy ra lỗi khi tính toán dữ liệu: {ex}")
else:
    st.info(f"💡 {L['error_file']}")
