import io
import os
import time
from datetime import datetime, date
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google import genai
from google.genai import types

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Dashboard Thống Kê Nhân Sự & Chấm Công",
    page_icon="📊",
    layout="wide"
)

# Sidebar - Cấu hình API Key Gemini
st.sidebar.header("⚙️ Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập Google Gemini API Key", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

st.title("📊 Dashboard Thống Kê Nhân Viên Đi Làm & Chấm Công")
st.markdown("Hệ thống tích hợp **Gemini AI** để xử lý lịch ca, đối chiếu vân tay, phân tích văn phòng và công nhân.")

# --- PHẦN 1: TẢI CÁC FILE DỮ LIỆU LÊN DASHBOARD ---
st.subheader("📁 1. Tải lên các tệp dữ liệu")
col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader("Tải file Excel bấm vân tay", type=["xlsx", "xls", "csv"])
    uploaded_schedule = st.file_uploader("Tải file lịch xếp ca (dành cho CN)", type=["xlsx", "xls", "csv"])

with col2:
    uploaded_staff_vp = st.file_uploader("Tải file danh sách nhân viên VP (Mã NV)", type=["xlsx", "xls", "csv"])
    uploaded_staff_cn = st.file_uploader("Tải file danh sách công nhân (Mã CN)", type=["xlsx", "xls", "csv"])

# --- PHẦN 2: Ô CHỌN NGÀY, THÁNG, NĂM ĐỐI CHIẾU ---
st.subheader("📅 2. Chọn ngày tháng năm đối chiếu")
col_d, col_m, col_y = st.columns(3)
with col_d:
    selected_day = st.selectbox("Chọn ngày", list(range(1, 32)), index=5) # Mặc định ngày 6
with col_m:
    selected_month = st.selectbox("Chọn tháng", list(range(1, 13)), index=8) # Mặc định tháng 9
with col_y:
    selected_year = st.selectbox("Chọn năm", [2025, 2026, 2027], index=1) # Mặc định 2026

target_date_str = f"{selected_year}-{selected_month:02d}-{selected_day:02d}"
st.info(f"Đang chọn ngày đối chiếu: **{selected_day}/{selected_month}/{selected_year}**")


# Hàm hỗ trợ đọc file linh hoạt
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

df_fp = load_uploaded_file(uploaded_fingerprint)
df_sc = load_uploaded_file(uploaded_schedule)
df_vp = load_uploaded_file(uploaded_staff_vp)
df_cn = load_uploaded_file(uploaded_staff_cn)


# --- PHẦN 3: XỬ LÝ VÀ GỌI GEMINI AI (CÓ RETRY CHỐNG LỖI 503) ---
st.subheader("🤖 3. Phân tích dữ liệu bằng Gemini AI")

if st.button("🚀 Chạy phân tích chấm công với Gemini AI", type="primary"):
    if not api_key:
        st.warning("Vui lòng nhập Google Gemini API Key ở thanh bên trái (Sidebar) để tiếp tục.")
    elif df_fp is None or df_vp is None or df_cn is None:
        st.warning("Vui lòng tải lên ít nhất file Bấm vân tay, Danh sách Nhân viên VP và Danh sách Công nhân.")
    else:
        with st.spinner("Gemini AI đang xử lý, đối chiếu dữ liệu ca làm việc và vân tay..."):
            try:
                client = genai.Client()
                
                # Chuyển toàn bộ dữ liệu thực tế thành chuỗi để đưa vào Prompt cho AI
                fp_data = df_fp.to_string() if df_fp is not None else "Không có"
                vp_data = df_vp.to_string() if df_vp is not None else "Không có"
                cn_data = df_cn.to_string() if df_cn is not None else "Không có"
                sc_data = df_sc.to_string() if df_sc is not None else "Không có"

                prompt = f"""
                Bạn là một AI chuyên phân tích nhân sự và chấm công tự động. 
                Hãy thực hiện phân tích dữ liệu cho ngày: {target_date_str}.
                
                Dữ liệu đầu vào:
                1. File vân tay:
                {fp_data}
                
                2. Danh sách Nhân viên Văn Phòng (VP) (Tổng số: {len(df_vp) if df_vp is not None else 0} người):
                {vp_data}
                
                3. Danh sách Công Nhân (CN) (Tổng số: {len(df_cn) if df_cn is not None else 0} người):
                {cn_data}
                
                4. Lịch xếp ca Công Nhân:
                {sc_data}
                
                QUY TẮC ÁP DỤNG:
                - VP: Giờ vào chuẩn 08:00 AM, Giờ ra chuẩn 17:00 PM (8 tiếng/ngày). Chủ nhật là ngày nghỉ. Ngày {target_date_str} nếu là thứ 2 đến thứ 7 mà không thấy bấm giờ vào/ra -> Ghi chú "vắng". Nếu thiếu giờ vào -> "BTV", thiếu giờ ra -> "BTR". Đủ 8 tiếng hoặc làm bù. Không đủ 8 tiếng tính là về sớm và ghi chú số giờ thực tế.
                - Công Nhân (CN): 
                  + Ca Ngày (N): Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 19:00 PM (12 tiếng).
                  + Ca Đêm (Đ): Giờ vào chuẩn 19:00 PM, Giờ ra chuẩn 7:00 AM hôm sau (12 tiếng).
                  + Ngày nghỉ hàng tuần được tô màu cam hoặc quy định trong lịch ca.
                  + Đối chiếu mã nhân viên CN ngày {target_date_str} với lịch ca xem làm ca gì (Đ hay N), sau đó check file vân tay xem có khớp giờ vào/ra không. Thiếu giờ vào: "BTV", thiếu giờ ra: "BTR". Cả hai thiếu -> "Vắng".
                
                Hãy trả về kết quả dưới dạng cấu trúc bảng Markdown rõ ràng gồm các cột: 
                Mã NV/CN | Họ Tên | Loại (VP/CN) | Ca làm việc | Giờ vào thực tế | Giờ ra thực tế | Số giờ làm | Trạng thái (Đi làm / Vắng / Về sớm / BTV / BTR) | Ghi chú chi tiết.
                """

                # Cơ chế thử lại (Retry) khi gặp lỗi quá tải 503
                max_retries = 3
                response = None
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-2.5-flash', # Dùng model ổn định
                            contents=prompt,
                        )
                        break
                    except Exception as err:
                        if "503" in str(err) and attempt < max_retries - 1:
                            time.sleep(3)
                            continue
                        else:
                            raise err

                st.session_state['ai_analysis_result'] = response.text
                st.success("✅ Phân tích hoàn tất!")
                
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi kết nối với Gemini AI: {e}")

# Hiển thị kết quả phân tích nếu có sẵn trong session
if 'ai_analysis_result' in st.session_state:
    st.markdown("### 📋 Kết quả đối chiếu chi tiết từ Gemini AI")
    st.markdown(st.session_state['ai_analysis_result'])


# --- PHẦN 4: THỐNG KÊ & PHÂN TÍCH BẰNG BIỂU ĐỒ (ĐỌC SỐ LƯỢNG THỰC TẾ) ---
st.subheader("📈 4. Thống kê & Phân tích bằng Biểu đồ")

total_vp = len(df_vp) if df_vp is not None else 0
total_cn = len(df_cn) if df_cn is not None else 0
total_nhan_su = total_vp + total_cn

# Hiển thị các chỉ số tổng quan chuẩn xác theo file tải lên
col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Tổng số Nhân viên VP", f"{total_vp} người")
col_m2.metric("Tổng số Công nhân (CN)", f"{total_cn} người")
col_m3.metric("Tổng nhân sự hệ thống", f"{total_nhan_su} người")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("#### Tỷ lệ phân bố nhân sự")
    dept_data = pd.DataFrame({
        'Bộ phận': ['Văn Phòng', 'Công Nhân'],
        'Số lượng': [total_vp, total_cn]
    })
    fig_pie = px.pie(dept_data, values='Số lượng', names='Bộ phận', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
    st.plotly_chart(fig_pie, use_container_width=True)

with col_chart2:
    st.markdown("#### Thống kê quy mô nhân sự theo bộ phận")
    fig_bar = px.bar(dept_data, x='Bộ phận', y='Số lượng', color='Bộ phận', text_auto=True)
    st.plotly_chart(fig_bar, use_container_width=True)


# --- PHẦN 5: XUẤT BÁO CÁO (EXCEL / PDF) ---
st.subheader("💾 5. Xuất báo cáo (Excel / PDF)")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if 'ai_analysis_result' in st.session_state:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_summary_export = pd.DataFrame({
                "Ngày": [target_date_str]*2,
                "Phân loại": ["Văn Phòng", "Công Nhân"],
                "Tổng số": [total_vp, total_cn],
            })
            df_summary_export.to_excel(writer, index=False, sheet_name='ThongKeChamCong')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Tải xuống file Excel báo cáo",
            data=processed_data,
            file_name=f"Bao_cao_cham_cong_{target_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Hãy chạy phân tích AI để kích hoạt nút tải Excel chi tiết.")

with col_dl2:
    if st.button("📥 Tải xuống file PDF báo cáo giao diện"):
        try:
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=f"BAO CAO CHAM CONG - NGAY {target_date_str}", ln=1, align="C")
            pdf.ln(10)
            pdf.cell(200, 10, txt=f"- Tong so nhan vien VP: {total_vp}", ln=1)
            pdf.cell(200, 10, txt=f"- Tong so cong nhan: {total_cn}", ln=1)
            pdf.cell(200, 10, txt=f"- Tong nhan su: {total_nhan_su}", ln=1)
            
            pdf_output = pdf.output(dest='S').encode('latin1')
            st.download_button(
                label="📄 Xác nhận tải file PDF",
                data=pdf_output,
                file_name=f"Dashboard_Report_{target_date_str}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.warning(f"Cần cài đặt thư viện fpdf để xuất file PDF (`pip install fpdf`): {e}")
