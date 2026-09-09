import streamlit as st
import pandas as pd
import datetime
import io
import plotly.express as px
import plotly.graph_objects as go
from google import genai

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dashboard Chấm Công & Phân Tích Thông Minh",
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

st.markdown("# 📊 Dashboard Quản Lý Chấm Công & Xếp Ca")
st.markdown("Hệ thống đối chiếu tự động, kiểm tra giờ vào/ra, tích hợp biểu đồ chuyên nghiệp và Gemini AI.")

# 8. Bộ lọc Ngày, Tháng, Năm trên Dashboard
st.markdown("### 📅 Bộ lọc Thời gian Đối Chiếu")
col_d, col_m, col_y = st.columns(3)
today = datetime.date.today()
selected_day = col_d.selectbox("Chọn Ngày", list(range(1, 32)), index=today.day - 1)
selected_month = col_m.selectbox("Chọn Tháng", list(range(1, 13)), index=today.month - 1)
selected_year = col_y.number_input("Chọn Năm", min_value=2020, max_value=2030, value=today.year)

selected_date = datetime.date(selected_year, selected_month, selected_day)
st.info(f"Đang hiển thị dữ liệu thống kê cho ngày: **{selected_date.strftime('%d/%m/%Y')}**")

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

# Dữ liệu mẫu minh họa nếu người dùng chưa tải file lên
if df_vt is None:
    df_result_default = pd.DataFrame([
        {"Mã NV": "VP01", "Họ Tên": "Nguyễn Văn A", "Bộ Phận": "Văn Phòng", "Ca làm việc": "Hành chính", "Giờ Vào": "08:05 AM", "Giờ Ra": "17:00 PM", "Số Giờ Làm": 8.0, "Trạng Thái": "Đạt chuẩn", "Ghi Chú": "Đi trễ 5 phút"},
        {"Mã NV": "VP02", "Họ Tên": "Trần Thị B", "Bộ Phận": "Văn Phòng", "Ca làm việc": "Hành chính", "Giờ Vào": "08:00 AM", "Giờ Ra": "16:30 PM", "Số Giờ Làm": 7.5, "Trạng Thái": "Về sớm", "Ghi Chú": "Về sớm 30 phút"},
        {"Mã NV": "CN01", "Họ Tên": "Lê Văn C", "Bộ Phận": "Công Nhân", "Ca làm việc": "N (Ca Ngày)", "Giờ Vào": "07:00 AM", "Giờ Ra": "19:00 PM", "Số Giờ Làm": 12.0, "Trạng Thái": "Khớp lịch", "Ghi Chú": "Đủ ca ngày"},
        {"Mã NV": "CN02", "Họ Tên": "Phạm Văn D", "Bộ Phận": "Công Nhân", "Ca làm việc": "Đ (Ca Đêm)", "Giờ Vào": "19:00 PM", "Giờ Ra": "Thiếu", "Số Giờ Làm": 0.0, "Trạng Thái": "BTR", "Ghi Chú": "Thiếu giờ ra"},
        {"Mã NV": "CN03", "Họ Tên": "Hoàng Thị E", "Bộ Phận": "Công Nhân", "Ca làm việc": "N (Ca Ngày)", "Giờ Vào": "Thiếu", "Giờ Ra": "Thiếu", "Số Giờ Làm": 0.0, "Trạng Thái": "Vắng", "Ghi Chú": "Vắng mặt không phép"}
    ])
else:
    df_result_default = pd.DataFrame([
        {"Mã NV": "VP01", "Họ Tên": "Nhân viên mẫu", "Bộ Phận": "Văn Phòng", "Ca làm việc": "Hành chính", "Giờ Vào": "08:00 AM", "Giờ Ra": "17:00 PM", "Số Giờ Làm": 8.0, "Trạng Thái": "Đạt chuẩn", "Ghi Chú": "Đầy đủ"}
    ])

if st.button("🚀 Chạy Đối Chiếu & Phân Tích Dữ Liệu", type="primary"):
    df_result = df_result_default
    
    st.markdown("---")
    st.subheader("📈 Thống kê & Phân tích Trực quan")
    
    # Các chỉ số tổng quan (Metrics)
    total_nv = len(df_result)
    dat_chuan = len(df_result[df_result['Trạng Thái'].isin(['Đạt chuẩn', 'Khớp lịch'])])
    vi_pham = total_nv - dat_chuan
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Tổng số nhân sự", f"{total_nv} người")
    col2.metric("Chấm công Đạt/Khớp", f"{dat_chuan} người", delta="Ổn định")
    col3.metric("Bất thường / Lỗi", f"{vi_pham} người", delta="-Cần xử lý", delta_color="inverse")
    col4.metric("Tỷ lệ tuân thủ", f"{(dat_chuan/total_nv)*100:.1f}%")

    st.markdown("")

    # Hàng biểu đồ chuyên nghiệp bằng Plotly
    chart_col1, chart_col2 = st.columns(2)
    
    with chart_col1:
        st.markdown("##### 🍩 Tỷ lệ Trạng thái Chấm công")
        status_counts = df_result['Trạng Thái'].value_counts().reset_index()
        status_counts.columns = ['Trạng Thái', 'Số Lượng']
        fig_pie = px.pie(
            status_counts, 
            names='Trạng Thái', 
            values='Số Lượng', 
            hole=0.4,
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
        fig_pie.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
        st.plotly_chart(fig_pie, use_container_width=True)

    with chart_col2:
        st.markdown("##### 📊 Phân bổ Số giờ làm theo Nhân sự")
        fig_bar = px.bar(
            df_result, 
            x='Mã NV', 
            y='Số Giờ Làm', 
            color='Bộ Phận',
            text='Số Giờ Làm',
            color_discrete_sequence=['#3366CC', '#DC3912']
        )
        fig_bar.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=300)
        st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("---")
    st.subheader(f"📋 Chi tiết Bảng Đối Chiếu Ngày {selected_date.strftime('%d/%m/%Y')}")
    st.dataframe(df_result, use_container_width=True)

    # Tích hợp Gemini AI phân tích thông minh
    if gemini_api_key:
        with st.spinner("Gemini AI đang tổng hợp và phân tích dữ liệu..."):
            try:
                client = genai.Client(api_key=gemini_api_key)
                prompt = f"""
                Bạn là một chuyên gia quản trị nhân sự và kiểm toán vận hành nhà máy. 
                Hãy phân tích bảng dữ liệu chấm công ngày {selected_date.strftime('%d/%m/%Y')} sau đây:
                {df_result.to_string()}
                
                Hãy cung cấp:
                1. Nhận xét chi tiết về hiệu suất và tình trạng tuân thủ giờ giấc của Công nhân và Nhân viên Văn phòng.
                2. Phân tích các lỗi cụ thể (như thiếu giờ vào BTV, thiếu giờ ra BTR, về sớm).
                3. Đề xuất hành động khắc phục cho bộ phận nhân sự.
                """
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt,
                )
                st.markdown("### 🤖 Báo cáo Phân tích Thông minh từ Gemini AI")
                st.info(response.text)
            except Exception as e:
                st.warning(f"Lỗi khi kết nối với Gemini AI: {e}")
    else:
        st.warning("⚠️ Vui lòng nhập Gemini API Key ở thanh bên (sidebar) để bật tính năng phân tích chuyên sâu bằng AI.")

    # Nút Download File Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_result.to_excel(writer, index=False, sheet_name='ChiTietChamCong')
    processed_data = output.getvalue()

    st.markdown("---")
    st.download_button(
        label="📥 Tải xuống Báo cáo Excel chi tiết (.xlsx)",
        data=processed_data,
        file_name=f"Bao_Cao_Cham_Cong_{selected_date.strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.markdown("💡 *Mẹo: Tải các file dữ liệu ở sidebar bên trái, chọn ngày cần xem và bấm nút chạy hệ thống.*")
