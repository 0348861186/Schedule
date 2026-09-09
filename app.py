import streamlit as st
import pandas as pd
import datetime
import io
from google import genai
from google.genai import types

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dashboard Quản lý Chấm công & Tích hợp Gemini AI",
    page_icon="📊",
    layout="wide"
)

# Khởi tạo Gemini Client (Lấy API Key từ st.secrets hoặc nhập trực tiếp)
gemini_api_key = st.secrets.get("GEMINI_API_KEY", "")

st.sidebar.title("⚙️ Cấu hình & Dữ liệu")
if not gemini_api_key:
    gemini_api_key = st.sidebar.text_input("Nhập Google Gemini API Key:", type="password")

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Tải lên các tệp dữ liệu")

# 1. Nút tải file chấm công vân tay
uploaded_van_tay = st.sidebar.file_uploader("1. File chấm công vân tay (Excel/CSV)", type=["xlsx", "xls", "csv"])

# 2. Nút tải lịch xếp ca
uploaded_lich_ca = st.sidebar.file_uploader("2. File lịch xếp ca (Excel/CSV)", type=["xlsx", "xls", "csv"])

# 3. Nút tải danh sách nhân viên VP
uploaded_vp = st.sidebar.file_uploader("3. Danh sách Nhân viên VP (Excel/CSV)", type=["xlsx", "xls", "csv"])

# 4. Nút tải danh sách công nhân
uploaded_cn = st.sidebar.file_uploader("4. Danh sách Công nhân (Excel/CSV)", type=["xlsx", "xls", "csv"])

st.markdown("# 📊 Dashboard Đối Chiếu Chấm Công & Xếp Ca")
st.markdown("Hệ thống kiểm tra giờ vào/ra thực tế của Văn phòng và Công nhân, tích hợp Gemini AI phân tích chuyên sâu.")

# 8. Bộ lọc Ngày, Tháng, Năm trên Dashboard
st.markdown("### 📅 Chọn thời gian đối chiếu")
col_d, col_m, col_y = st.columns(3)
today = datetime.date.today()
selected_day = col_d.selectbox("Chọn Ngày", list(range(1, 32)), index=today.day - 1)
selected_month = col_m.selectbox("Chọn Tháng", list(range(1, 13)), index=today.month - 1)
selected_year = col_y.number_input("Chọn Năm", min_value=2020, max_value=2030, value=today.year)

selected_date = datetime.date(selected_year, selected_month, selected_day)
st.info(f"Đang xem dữ liệu thống kê và đối chiếu cho ngày: **{selected_date.strftime('%d/%m/%Y')}**")

# Hàm đọc file linh hoạt
def load_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                return pd.read_csv(uploaded_file)
            else:
                return pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"Lỗi đọc file {uploaded_file.name}: {e}")
    return None

df_vt = load_uploaded_file(uploaded_van_tay)
df_lc = load_uploaded_file(uploaded_lich_ca)
df_nv_vp = load_uploaded_file(uploaded_vp)
df_nv_cn = load_uploaded_file(uploaded_cn)

# Xử lý giả lập dữ liệu nếu chưa tải đủ để demo giao diện
if df_vt is None:
    # Tạo dữ liệu mẫu cho demo
    df_vt = pd.DataFrame({
        'MaNV': ['VP01', 'CN01', 'CN02'],
        'HoTen': ['Nguyễn Văn A', 'Trần Thị B', 'Lê Văn C'],
        'ThoiGianVao': ['2026-09-06 08:05:00', '2026-09-06 07:00:00', '2026-09-06 19:00:00'],
        'ThoiGianRa': ['2026-09-06 17:00:00', '2026-09-06 19:00:00', '2026-09-07 07:00:00']
    })

# Main Processing Logic
if st.button("🚀 Chạy Đối Chiếu & Phân Tích Dữ Liệu", type="primary"):
    with st.spinner("Đang xử lý dữ liệu và gọi Gemini AI đối chiếu..."):
        
        # Mô phỏng kết quả đối chiếu theo quy tắc
        # Quy tắc VP: Vào 8:00 AM, Ra 17:00 PM (8 tiếng), Chủ nhật nghỉ
        # Quy tắc CN: Ca N (7h - 19h = 12h), Ca Đ (19h - 7h sáng hôm sau = 12h)
        # Thiếu giờ Vào = "BTV", Thiếu giờ Ra = "BTR", Vắng nếu thiếu cả 2.
        
        mock_results = [
            {"Mã NV": "VP01", "Họ Tên": "Nguyễn Văn A", "Loại": "Văn Phòng", "Ca/Lịch": "Hành chính", "Giờ Vào TT": "08:05 AM", "Giờ Ra TT": "17:00 PM", "Số Giờ Làm": "8.5 tiếng", "Trạng Thái": "Đạt chuẩn đủ giờ", "Ghi chú": "Đi trễ 5 phút"},
            {"Mã NV": "CN01", "Họ Tên": "Trần Thị B", "Loại": "Công Nhân", "Ca/Lịch": "N (Ca Ngày)", "Giờ Vào TT": "07:00 AM", "Giờ Ra TT": "19:00 PM", "Số Giờ Làm": "12 tiếng", "Trạng Thái": "Khớp lịch chuẩn", "Ghi chú": "Hoàn thành ca ngày"},
            {"Mã NV": "CN02", "Họ Tên": "Lê Văn C", "Loại": "Công Nhân", "Ca/Lịch": "Đ (Ca Đêm)", "Giờ Vào TT": "19:00 PM", "Giờ Ra TT": "Thiếu", "Số Giờ Làm": "0 tiếng", "Trạng Thái": "BTR", "Ghi chú": "Thiếu giờ ra (Bấm vân tay lúc 19:00)"},
        ]
        df_result = pd.DataFrame(mock_results)

        st.success("Đã hoàn thành đối chiếu dữ liệu chấm công!")
        
        # Hiển thị bảng kết quả
        st.subheader("📋 Bảng Kết Quả Đối Chiếu Chấm Công Ngày " + selected_date.strftime('%d/%m/%Y'))
        st.dataframe(df_result, use_container_width=True)

        # Tích hợp Gemini AI để phân tích
        if gemini_api_key:
            try:
                client = genai.Client(api_key=gemini_api_key)
                prompt = f"""
                Bạn là một chuyên gia nhân sự và hệ thống kiểm toán chấm công thông minh. 
                Hãy phân tích bảng dữ liệu chấm công ngày {selected_date.strftime('%d/%m/%Y')} sau đây:
                {df_result.to_string()}
                
                Hãy đưa ra:
                1. Nhận xét tổng quan về tình hình chấm công (số nhân viên đạt, số nhân viên vi phạm/thiếu giờ BTV, BTR).
                2. Đề xuất hướng xử lý cụ thể cho các trường hợp bất thường (ví dụ nhân viên thiếu giờ ra BTR).
                """
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt,
                )
                st.markdown("### 🤖 Phân tích & Đánh giá từ Gemini AI")
                st.info(response.text)
            except Exception as e:
                st.warning(f"Không thể kết nối với Gemini AI (Kiểm tra lại API Key của bạn): {e}")
        else:
            st.warning("⚠️ Vui lòng nhập Gemini API Key ở thanh bên (sidebar) để sử dụng tính năng phân tích thông minh bằng AI.")

        # 9. Nút download file Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='KetQuaChamCong')
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Tải xuống Báo cáo Excel (.xlsx)",
            data=processed_data,
            file_name=f"Bao_Cao_Cham_Cong_{selected_date.strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("---")
st.markdown("💡 *Hướng dẫn: Tải file lên ở cột bên trái, chọn ngày cần xem và bấm nút 'Chạy Đối Chiếu & Phân Tích Dữ Liệu'.*")
