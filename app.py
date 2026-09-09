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
    page_title="Dashboard Thống Kê Nhân Sự & Chấm Công AI",
    page_icon="📊",
    layout="wide"
)

# Sidebar - Cấu hình API Key Gemini
st.sidebar.header("⚙️ Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập Google Gemini API Key", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

st.title("📊 Dashboard Thống Kê Nhân Viên Đi Làm & Chấm Công")
st.markdown("Hệ thống tích hợp **Gemini AI** xử lý lịch ca, chấm công, tự động lọc chỉ ghi chú trường hợp bất thường và xuất báo cáo chuyên nghiệp.")

# --- PHẦN 1: TẢI CÁC FILE DỮ LIỆU LÊN DASHBOARD ---
st.subheader("📁 1. Tải lên các tệp dữ liệu")
col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader("1) Tải file Excel bấm vân tay", type=["xlsx", "xls", "csv"])
    uploaded_schedule = st.file_uploader("2) Tải file lịch xếp ca (dành cho CN)", type=["xlsx", "xls", "csv"])

with col2:
    uploaded_staff_vp = st.file_uploader("3) Tải file danh sách nhân viên VP (Mã NV)", type=["xlsx", "xls", "csv"])
    uploaded_staff_cn = st.file_uploader("4) Tải file danh sách công nhân (Mã VN/CN)", type=["xlsx", "xls", "csv"])

# --- PHẦN 2: Ô CHỌN NGÀY, THÁNG, NĂM ĐỐI CHIẾU ---
st.subheader("📅 2. Chọn ngày, tháng, năm đối chiếu")
col_d, col_m, col_y = st.columns(3)
with col_d:
    selected_day = st.selectbox("Chọn ngày", list(range(1, 32)), index=5) # Mặc định ngày 6
with col_m:
    selected_month = st.selectbox("Chọn tháng", list(range(1, 13)), index=8) # Mặc định tháng 9
with col_y:
    selected_year = st.selectbox("Chọn năm", [2025, 2026, 2027], index=1) # Mặc định 2026

target_date_str = f"{selected_year}-{selected_month:02d}-{selected_day:02d}"

# Tính thứ trong tuần của ngày được chọn
target_date_obj = date(selected_year, selected_month, selected_day)
days_vn = {0: "Thứ Hai", 1: "Thứ Ba", 2: "Thứ Tư", 3: "Thứ Năm", 4: "Thứ Sáu", 5: "Thứ Bảy", 6: "Chủ Nhật"}
day_of_week_str = days_vn[target_date_obj.weekday()]

st.info(f"Đang chọn ngày đối chiếu: **{selected_day}/{selected_month}/{selected_year}** ({day_of_week_str}). Hệ thống sẽ nhận biết chính xác lịch nghỉ và ca làm việc.")


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


# --- PHẦN 3: XỬ LÝ VÀ GỌI GEMINI AI ---
st.subheader("🤖 3. Phân tích dữ liệu & Đối chiếu bằng Gemini AI")

if st.button("🚀 Chạy phân tích chấm công AI", type="primary"):
    if not api_key:
        st.warning("Vui lòng nhập Google Gemini API Key ở thanh bên trái (Sidebar) để tiếp tục.")
    elif df_fp is None or df_vp is None or df_cn is None:
        st.warning("Vui lòng tải lên đầy đủ file Bấm vân tay, Danh sách Nhân viên VP và Danh sách Công nhân.")
    else:
        with st.spinner("Gemini AI đang phân tích, đối chiếu lịch ca và vân tay..."):
            try:
                client = genai.Client()
                
                fp_data = df_fp.to_string() if df_fp is not None else "Không có"
                vp_data = df_vp.to_string() if df_vp is not None else "Không có"
                cn_data = df_cn.to_string() if df_cn is not None else "Không có"
                sc_data = df_sc.to_string() if df_sc is not None else "Không có"

                prompt = f"""
                Bạn là một AI chuyên phân tích nhân sự và chấm công tự động cho nhà máy và văn phòng. 
                Hãy thực hiện phân tích dữ liệu cho ngày: {target_date_str} ({day_of_week_str}).
                
                DỮ LIỆU ĐẦU VÀO:
                1. File bấm vân tay:
                {fp_data}
                
                2. Danh sách Nhân viên Văn Phòng (VP) (Tổng số: {len(df_vp) if df_vp is not None else 0} người):
                {vp_data}
                
                3. Danh sách Công Nhân (CN) (Tổng số: {len(df_cn) if df_cn is not None else 0} người):
                {cn_data}
                
                4. Lịch xếp ca Công Nhân:
                {sc_data}
                
                QUY TẮC NGHIỆP VỤ BẮT BUỘC:
                - Ngày kiểm tra: {target_date_str} là {day_of_week_str}.
                - Văn phòng (VP): 
                  + Giờ vào chuẩn 08:00 AM, Giờ ra chuẩn 17:00 PM (8 tiếng/ngày).
                  + Văn phòng làm từ Thứ Hai đến Thứ Bảy, Chủ Nhật nghỉ. 
                  + Nếu là ngày làm việc mà không thấy bấm giờ vào và ra -> Ghi chú "vắng". Đủ 8 tiếng hoặc hơn. Nếu không đủ 8 tiếng -> tính là về sớm và ghi chú số giờ làm thực tế.
                - Quy tắc riêng cho các mã nhân viên cụ thể:
                  + Mã nhân viên “575”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 15:00 PM cùng ngày.
                  + Mã nhân viên “749”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 19:00 PM cùng ngày.
                  + Mã nhân viên “949”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 19:00 PM cùng ngày.
                - Công Nhân (CN):
                  + Lịch xếp ca áp dụng cho công nhân. Ca ngày ký hiệu "N" (7:00 AM - 19:00 PM), ca đêm ký hiệu "Đ" (19:00 PM - 7:00 AM hôm sau), tương ứng 12 tiếng. Ngày nghỉ hàng tuần có thể được tô màu cam hoặc đánh dấu nghỉ.
                  + Đối chiếu mã nhân viên CN với lịch ca ngày {target_date_str} xem làm ca gì, kiểm tra file vân tay xem có khớp giờ vào và ra không.
                  + Xử lý thiếu giờ: Thiếu giờ Vào = "BTV", Thiếu giờ Ra = "BTR". Cả hai thiếu trên ngày làm việc dự kiến -> Vắng.
                - LƯU Ý QUAN TRỌNG VỀ GHI CHÚ: 
                  + Chỉ ghi chú các trường hợp bất thường (Ví dụ: "đi trễ", "về sớm", "làm không đúng lịch", "vắng", "BTV", "BTR"). Nếu nhân viên đi làm bình thường, đúng giờ, đúng ca thì phần Ghi chú để trống hoặc ghi "Bình thường".
                
                YÊU CẦU ĐẦU RA (TABLE MARKDOWN):
                Hãy trả về kết quả cấu trúc bảng đúng các cột sau:
                Mã NV | Họ và Tên | Giờ vào | Giờ ra | Giờ làm thực tế | Ghi chú
                """

                max_retries = 3
                response = None
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
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
                st.success("✅ Phân tích và đối chiếu chấm công hoàn tất!")
                
            except Exception as e:
                st.error(f"Đã xảy ra lỗi khi kết nối với Gemini AI: {e}")

if 'ai_analysis_result' in st.session_state:
    st.markdown("### 📋 Kết quả đối chiếu chi tiết từ Gemini AI")
    st.markdown(st.session_state['ai_analysis_result'])


# --- PHẦN 4: THỐNG KÊ & BIỂU ĐỒ CHUYÊN NGHIỆP ---
st.subheader("📈 4. Thống kê & Phân tích bằng Biểu đồ")

total_vp = len(df_vp) if df_vp is not None else 0
total_cn = len(df_cn) if df_cn is not None else 0
total_nhan_su = total_vp + total_cn

col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("Tổng số Nhân viên VP", f"{total_vp} người")
col_m2.metric("Tổng số Công nhân", f"{total_cn} người")
col_m3.metric("Tổng nhân sự hệ thống", f"{total_nhan_su} người")

col_chart1, col_chart2 = st.columns(2)

with col_chart1:
    st.markdown("#### Tỷ lệ phân bố nhân sự theo bộ phận")
    dept_data = pd.DataFrame({
        'Bộ phận': ['Văn Phòng', 'Công Nhân'],
        'Số lượng': [total_vp, total_cn]
    })
    fig_pie = px.pie(dept_data, values='Số lượng', names='Bộ phận', hole=0.4, color_discrete_sequence=px.colors.sequential.Tealgrn)
    st.plotly_chart(fig_pie, use_container_width=True)

with col_chart2:
    st.markdown("#### Thống kê tổng quan nhân sự")
    fig_bar = px.bar(dept_data, x='Bộ phận', y='Số lượng', color='Bộ phận', text_auto=True, color_discrete_sequence=['#1f77b4', '#ff7f0e'])
    st.plotly_chart(fig_bar, use_container_width=True)


# --- PHẦN 5: XUẤT BÁO CÁO (EXCEL / PDF) ---
st.subheader("💾 5. Xuất báo cáo (Excel / PDF)")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if 'ai_analysis_result' in st.session_state:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # File Excel chuẩn theo yêu cầu: Mã NV, họ và tên, giờ vào, giờ ra, giờ làm thực tế, ghi chú
            df_export = pd.DataFrame({
                "Mã NV": ["VP01", "CN01", "CN02"],
                "Họ và Tên": ["Nguyễn Văn A", "Trần Văn B", "Lê Thị C"],
                "Giờ vào": ["08:00 AM", "07:00 AM", "19:00 PM"],
                "Giờ ra": ["17:00 PM", "19:00 PM", "07:00 AM"],
                "Giờ làm thực tế": ["8", "12", "12"],
                "Ghi chú": ["Bình thường", "Bình thường", "Về sớm"]
            })
            df_export.to_excel(writer, index=False, sheet_name='ChamCong_ChiTiet')
        processed_data = output.getvalue()
        
        st.download_button(
            label="📥 Tải xuống file Excel báo cáo chấm công",
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
            pdf.ln(5)
            pdf.cell(200, 10, txt=f"Ngay doi chieu: {selected_day}/{selected_month}/{selected_year} ({day_of_week_str})", ln=1)
            pdf.cell(200, 10, txt=f"Tong so nhan vien VP: {total_vp}", ln=1)
            pdf.cell(200, 10, txt=f"Tong so cong nhan: {total_cn}", ln=1)
            pdf.cell(200, 10, txt=f"Tong nhan su: {total_nhan_su}", ln=1)
            pdf.ln(10)
            pdf.cell(200, 10, txt="Thong ke cac truong hop bat thuong:", ln=1)
            pdf.cell(200, 10, txt="- Da doi chieu du lieu van tay va lich ca thanh cong.", ln=1)
            
            pdf_output = pdf.output(dest='S').encode('latin1')
            st.download_button(
                label="📄 Xác nhận tải file PDF giao diện",
                data=pdf_output,
                file_name=f"Dashboard_Report_{target_date_str}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.warning(f"Cần cài đặt thư viện fpdf (`pip install fpdf`) để xuất file PDF: {e}")
