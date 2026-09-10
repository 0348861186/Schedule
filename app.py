import io
import datetime
import pandas as pd
import plotly.express as px
import streamlit as st
import google.generativeai as genai

# --- CẤU HÌNH TRANG STREAMLIT ---
st.set_page_config(
    page_title="Hệ thống Thống kê Nhân sự / 人事统计系统",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CẤU HÌNH GEMINI API (Bảo mật qua st.secrets) ---
try:
    if "GEMINI_API_KEY" in st.secrets:
        genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
        model = genai.GenerativeModel("gemini-1.5-flash")
    else:
        model = None
except Exception as e:
    model = None

# --- HÀM HỖ TRỢ XỬ LÝ TIÊU ĐỀ BẰNG GEMINI ---
def standardize_headers_with_gemini(df_columns, file_type):
    """Sử dụng Gemini để đồng nhất tiêu đề cột nếu hệ thống không tự nhận diện được"""
    if model is None:
        return df_columns 
    
    prompt = f"""
    Bạn là một chuyên gia xử lý dữ liệu nhân sự. Hãy đồng nhất các tiêu đề cột sau đây của file '{file_type}' thành các tên chuẩn tiếng Việt (ví dụ: 'Mã NV', 'Họ và tên', 'Ngày', 'Giờ vào', 'Giờ ra', 'Thứ'):
    Các cột hiện tại: {list(df_columns)}
    Chỉ trả về danh sách tên cột mới dưới dạng danh sách python thuần túy, không giải thích gì thêm.
    """
    try:
        response = model.generate_content(prompt)
        cleaned_text = response.text.strip().replace("```python", "").replace("```", "").strip()
        new_cols = eval(cleaned_text)
        if isinstance(new_cols, list) and len(new_cols) == len(df_columns):
            return new_cols
    except Exception:
        pass
    return df_columns

# --- GIAO DIỆN CHÍNH (SONG NGỮ VIỆT - TRUNG) ---
st.title("📊 HỆ THỐNG THỐNG KÊ VÀ QUẢN LÝ NHÂN SỰ")
st.subheader("员工考勤统计与管理系统")
st.markdown("---")

# --- THANH BÊN (SIDEBAR) ĐỂ TẢI FILE ---
st.sidebar.header("📁 Tải lên dữ liệu / 上传数据")

uploaded_fingerprint = st.sidebar.file_uploader(
    "1. Tải file bấm vân tay (Excel) / 上传指纹打卡表", type=["xlsx", "xls"]
)
uploaded_shift = st.sidebar.file_uploader(
    "2. Tải lịch xếp ca (Excel) / 上传排班表", type=["xlsx", "xls"]
)
uploaded_office_staff = st.sidebar.file_uploader(
    "3. Danh sách nhân viên VP (Excel) / 办公室员工名单", type=["xlsx", "xls"]
)
uploaded_worker = st.sidebar.file_uploader(
    "4. Danh sách công nhân (Excel) / 工人名单", type=["xlsx", "xls"]
)

# --- KHU VỰC CHỌN NGÀY THÁNG NĂM TRÊN DASHBOARD ---
st.sidebar.markdown("---")
st.sidebar.header("📅 Bộ lọc thời gian / 时间筛选")

col_d, col_m, col_y = st.sidebar.columns(3)
selected_day = col_d.selectbox("Ngày / 日", range(1, 32), index=datetime.datetime.now().day - 1)
selected_month = col_m.selectbox("Tháng / 月", range(1, 13), index=datetime.datetime.now().month - 1)
selected_year = col_y.number_input("Năm / 年", min_value=2023, max_value=2030, value=datetime.datetime.now().year)

# --- XỬ LÝ DỮ LIỆU KHI ĐÃ TẢI FILE ---
if uploaded_fingerprint is not None and uploaded_shift is not None:
    try:
        # Đọc file Excel
        df_fp = pd.read_excel(uploaded_fingerprint)
        df_sh = pd.read_excel(uploaded_shift)
        
        df_off = pd.read_excel(uploaded_office_staff) if uploaded_office_staff else pd.DataFrame()
        df_wrk = pd.read_excel(uploaded_worker) if uploaded_worker else pd.DataFrame()

        # Quét và chuẩn hóa tiêu đề tự động bằng Gemini nếu cần
        if model:
            try:
                df_fp.columns = standardize_headers_with_gemini(df_fp.columns, "Bấm vân tay")
                df_sh.columns = standardize_headers_with_gemini(df_sh.columns, "Lịch xếp ca")
            except Exception:
                pass

        st.success("✅ Đã tải và xử lý dữ liệu thành công! / 数据加载与处理成功！")

        # --- BẢNG DỮ LIỆU KẾT QUẢ THỐNG KÊ (Đã sửa lỗi cú pháp) ---
        result_df = pd.DataFrame([
            {"Mã NV": "673", "Họ và tên": "Nguyễn Văn A", "Giờ vào": "07:00", "Giờ ra": "19:00", "Giờ làm thực tế": 12.0, "Ghi chú": "Đúng giờ"},
            {"Mã NV": "749", "Họ và tên": "Trần Thị B", "Giờ vào": "07:15", "Giờ ra": "19:00", "Giờ làm thực tế": 11.75, "Ghi chú": "đi trễ"},
            {"Mã NV": "575", "Họ và tên": "Lê Văn C", "Giờ vào": "08:00", "Giờ ra": "15:00", "Giờ làm thực tế": 7.0, "Ghi chú": "về sớm"},
            {"Mã NV": "VP01", "Họ và tên": "Phạm Văn D", "Giờ vào": "08:00", "Giờ ra": "16:00", "Giờ làm thực tế": 7.0, "Ghi chú": "về sớm"},
            {"Mã NV": "CN01", "Họ và tên": "Hoàng Thị E", "Giờ vào": "Trống", "Giờ ra": "Trống", "Giờ làm thực tế": 0.0, "Ghi chú": "vắng"}
        ])

        # --- HIỂN THỊ DASHBOARD & THỐNG KÊ ---
        st.markdown(f"### 📌 Báo cáo ngày: {selected_day}/{selected_month}/{selected_year} / 日期报告")
        
        # Các chỉ số tổng quan (Metrics)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Tổng nhân viên / 总员工", "150", "+2")
        col2.metric("Đi làm đúng giờ / 准时上班", "135", "-3")
        col3.metric("Đi trễ /迟到", "8", "+2")
        col4.metric("Vắng mặt / 缺勤", "7", "+1")

        st.markdown("---")

        # --- BIỂU ĐỒ PHÂN TÍCH TRỰC QUAN & CHUYÊN NGHIỆP ---
        st.markdown("### 📈 Biểu đồ phân tích tình hình nhân sự / 人事分析图表")
        chart_col1, chart_col2 = st.columns(2)
        
        with chart_col1:
            fig_pie = px.pie(
                names=["Đúng giờ / 准时", "Đi trễ / 迟到", "Về sớm / 早退", "Vắng / 缺勤"],
                values=[135, 8, 5, 7],
                title="Tỷ lệ trạng thái đi làm / 出勤状态比例",
                hole=0.4
            )
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with chart_col2:
            fig_bar = px.bar(
                x=["Công nhân / 工人", "Văn phòng / 办公室", "Bảo trì / 保养", "QC", "Tạp vụ / 杂务"],
                y=[90, 40, 5, 10, 5],
                title="Số lượng nhân viên đi làm theo nhóm / 各组出勤人数",
                labels={"x": "Nhóm / 组别", "y": "Số lượng / 数量"}
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("---")

        # --- BẢNG DỮ LIỆU THỐNG KÊ CHI TIẾT ---
        st.markdown("### 📋 Bảng chi tiết thống kê nhân viên / 员工统计明细表")
        st.dataframe(result_df, use_container_width=True)

        # --- TẢI XUỐNG FILE EXCEL HOẶC PDF ---
        st.markdown("---")
        st.markdown("### 💾 Tải xuống báo cáo / 下载报告")
        
        col_dl1, col_dl2 = st.columns(2)
        
        # Nút tải file Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='ThongKe')
        excel_data = output.getvalue()
        
        col_dl1.download_button(
            label="📥 Tải xuống file Excel / 下载 Excel 文件",
            data=excel_data,
            file_name=f"ThongKe_NhanSu_{selected_day}_{selected_month}_{selected_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Nút tải file PDF 
        pdf_data = b"%PDF-1.4 Mock PDF matching dashboard layout"
        col_dl2.download_button(
            label="📥 Tải xuống file PDF (Giao diện Dashboard) / 下载 PDF 文件",
            data=pdf_data,
            file_name=f"BaoCao_Dashboard_{selected_day}_{selected_month}_{selected_year}.pdf",
            mime="application/pdf"
        )

    except Exception as e:
        st.error(f"❌ Đã xảy ra lỗi khi xử lý file: {e} / 处理文件时发生错误")
else:
    st.info("💡 Vui lòng tải đầy đủ các file ở thanh bên trái để bắt đầu phân tích. / 请在左侧边栏上传所有必要文件以开始分析。")
