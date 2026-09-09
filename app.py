import io
import os
import time
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px
from google import genai

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="考勤统计仪表盘 / Dashboard Thống Kê Chấm Công",
    page_icon="📊",
    layout="wide"
)

# Cấu hình API Key Gemini ở Sidebar (Không có chữ AI trên giao diện)
st.sidebar.header("⚙️ 系统配置 / Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập API Key", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

st.title("📊 员工考勤与统计仪表盘")
st.markdown("### Dashboard Thống Kê & Phân Tích Chấm Công Nhân Sự")

# --- PHẦN 1: TẢI CÁC FILE DỮ LIỆU LÊN DASHBOARD ---
st.subheader("📁 1. 数据文件上传 / Tải lên tệp dữ liệu")
col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader("1) 上传指纹打卡 Excel 文件 / Tải file Excel bấm vân tay", type=["xlsx", "xls", "csv"])
    uploaded_schedule = st.file_uploader("2) 上传排班表文件 (仅限工人) / Tải file lịch xếp ca (cho công nhân)", type=["xlsx", "xls", "csv"])

with col2:
    uploaded_staff_vp = st.file_uploader("3) 上传办公室员工名单 (NV) / Tải file danh sách nhân viên VP", type=["xlsx", "xls", "csv"])
    uploaded_staff_cn = st.file_uploader("4) 上传工人名单 (CN) / Tải file danh sách công nhân", type=["xlsx", "xls", "csv"])

# Hàm hỗ trợ đọc file linh hoạt
def load_uploaded_file(uploaded_file):
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                return pd.read_csv(uploaded_file)
            else:
                return pd.read_excel(uploaded_file)
        except Exception as e:
            st.error(f"文件读取错误 / Lỗi đọc file: {e}")
    return None

df_fp = load_uploaded_file(uploaded_fingerprint)
df_sc = load_uploaded_file(uploaded_schedule)
df_vp = load_uploaded_file(uploaded_staff_vp)
df_cn = load_uploaded_file(uploaded_staff_cn)


# --- PHẦN 2: CHỌN NGÀY TỪ CỘT 'Ngày' TRONG FILE VÂN TAY ---
st.subheader("📅 2. 选择考勤核对日期 / Chọn ngày kiểm tra từ file vân tay")

selected_date_str = ""
df_fp_filtered = None

if df_fp is not None:
    # Chuẩn hóa tên cột để tránh lỗi khoảng trắng thừa
    df_fp.columns = df_fp.columns.str.strip()
    
    if 'Ngày' in df_fp.columns:
        df_fp['Ngày_Clean'] = pd.to_datetime(df_fp['Ngày'], errors='coerce').dt.date
        available_dates = sorted(df_fp['Ngày_Clean'].dropna().unique())
        
        if available_dates:
            selected_date = st.selectbox("选择文件中的日期 / Chọn ngày có trong file vân tay", available_dates)
            selected_date_str = str(selected_date)
            
            # Lọc dữ liệu vân tay đúng theo ngày được chọn
            df_fp_filtered = df_fp[df_fp['Ngày_Clean'] == selected_date]
            st.success(f"✅ 已选择日期 / Đã chọn ngày: **{selected_date_str}** (Số bản ghi trong ngày: {len(df_fp_filtered)} dòng)")
        else:
            st.warning("⚠️ Không nhận diện được định dạng ngày tháng trong cột 'Ngày'.")
    else:
        st.error("❌ Không tìm thấy cột 'Ngày' trong file vân tay.")
else:
    st.info("ℹ️ Vui lòng tải file Excel bấm vân tay lên ở bước 1.")


# --- PHẦN 3: XỬ LÝ VÀ PHÂN TÍCH ---
st.subheader("🤖 3. 智能数据核对与分析 / Phân tích & Đối chiếu thông minh")

if st.button("🚀 开始考勤核对分析 / Chạy phân tích chấm công", type="primary"):
    if not api_key:
        st.warning("请输入 API Key 以继续 / Vui lòng nhập API Key ở thanh bên.")
    elif df_fp_filtered is None or df_vp is None or df_cn is None:
        st.warning("请完整上传文件并选择有效日期 / Vui lòng tải đủ file và chọn ngày hợp lệ.")
    else:
        with st.spinner("系统正在智能核对排班与考勤数据，请稍候..."):
            try:
                client = genai.Client()
                
                fp_data = df_fp_filtered.to_string()
                vp_data = df_vp.to_string() if df_vp is not None else "Không có"
                cn_data = df_cn.to_string() if df_cn is not None else "Không có"
                sc_data = df_sc.to_string() if df_sc is not None else "Không có"

                prompt = f"""
                Bạn là một hệ thống phân tích nhân sự và chấm công tự động thông minh. 
                Hãy thực hiện đối chiếu và phân tích dữ liệu CHO ĐÚNG NGÀY: {selected_date_str}.
                
                DỮ LIỆU ĐẦU VÀO:
                1. File bấm vân tay (các cột: Mã Nhân Viên, Tên nhân viên, Phòng ban, Ngày, Thứ, Giờ vào, Giờ ra):
                {fp_data}
                
                2. Danh sách Nhân viên Văn Phòng (VP):
                {vp_data}
                
                3. Danh sách Công Nhân (CN):
                {cn_data}
                
                4. Lịch xếp ca Công Nhân:
                {sc_data}
                
                QUY TẮC NGHIỆP VỤ BẮT BUỘC:
                - Ngày kiểm tra: {selected_date_str}.
                - Văn phòng (VP): Giờ vào chuẩn 08:00 AM, Giờ ra chuẩn 17:00 PM (8 tiếng/ngày). Chủ nhật nghỉ. Nếu đủ giờ nhưng về sớm -> ghi chú "về sớm" kèm giờ làm thực tế. Nếu ngày làm việc không bấm thẻ -> "vắng".
                - Mã nhân viên đặc biệt:
                  + Mã “575”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 15:00 PM.
                  + Mã “749”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 19:00 PM.
                  + Mã “949”: Giờ vào chuẩn 7:00 AM, Giờ ra chuẩn 19:00 PM.
                - Công Nhân (CN):
                  + Căn cứ lịch xếp ca ngày {selected_date_str} ("N" = Ca ngày 7:00-19:00, "Đ" = Ca đêm 19:00-7:00 hôm sau, hoặc ngày lễ ghi chú "Lễ").
                  + Kiểm tra file vân tay xem khớp hay không. Thiếu giờ vào = "BV", thiếu giờ ra = "BR". Thiếu cả hai = "Vắng".
                - TRỌNG TÂM: Chỉ tập trung đưa ra các trường hợp bất thường (Ví dụ: "đi trễ", "về sớm", "làm không đúng lịch", "vắng", "BV", "BR", "Lễ"). Không dài dòng.
                
                YÊU CẦU ĐẦU RA:
                Trả về một bảng Markdown gồm đúng các cột sau:
                Mã NV | Họ và Tên | Giờ vào | Giờ ra | Giờ làm thực tế | Ghi chú
                """

                max_retries = 3
                response = None
                for attempt in range(max_retries):
                    try:
                        response = client.models.generate_content(
                            model='gemini-1.5-flash',
                            contents=prompt,
                        )
                        break
                    except Exception as err:
                        if "503" in str(err) and attempt < max_retries - 1:
                            time.sleep(3)
                            continue
                        else:
                            raise err

                st.session_state['analysis_result'] = response.text
                st.success("✅ 考勤核对完成！ / Hoàn tất đối chiếu chấm công!")
                
            except Exception as e:
                st.error(f"处理出错 / Lỗi xử lý: {e}")

if 'analysis_result' in st.session_state:
    st.markdown(f"### 📋 异常考勤与核对结果 ({selected_date_str})")
    st.markdown(st.session_state['analysis_result'])


# --- PHẦN 4: THỐNG KÊ BIỂU ĐỒ THEO NGÀY ĐÃ CHỌN ---
st.subheader("📈 4. 考勤数据统计图表 / Biểu đồ thống kê theo ngày")

total_vp = len(df_vp) if df_vp is not None else 0
total_cn = len(df_cn) if df_cn is not None else 0
active_count = len(df_fp_filtered['Mã Nhân Viên'].unique()) if (df_fp_filtered is not None and 'Mã Nhân Viên' in df_fp_filtered.columns) else 0

col_m1, col_m2, col_m3 = st.columns(3)
col_m1.metric("办公室员工总数 / Tổng NV Văn Phòng", f"{total_vp} 人")
col_m2.metric("工人总数 / Tổng Công Nhân", f"{total_cn} 人")
col_m3.metric(f"当日打卡人数 ({selected_date_str})", f"{active_count} 人")

if total_vp > 0 or total_cn > 0:
    chart_data = pd.DataFrame({
        '状态 / Trạng thái': ['Văn phòng tổng', 'Công nhân tổng', 'Đã bấm vân tay trong ngày'],
        '人数 / Số lượng': [total_vp, total_cn, active_count]
    })
    fig = px.bar(chart_data, x='状态 / Trạng thái', y='人数 / Số lượng', color='状态 / Trạng thái', text_auto=True)
    st.plotly_chart(fig, use_container_width=True)


# --- PHẦN 5: XUẤT BÁO CÁO EXCEL / PDF ---
st.subheader("💾 5. 导出报 cáo / Xuất báo cáo")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if 'analysis_result' in st.session_state:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export = pd.DataFrame({
                "Mã NV": ["VP01", "CN01"],
                "Họ và Tên": ["Nguyễn Văn A", "Trần Văn B"],
                "Giờ vào": ["08:00 AM", "07:00 AM"],
                "Giờ ra": ["17:00 PM", "19:00 PM"],
                "Giờ làm thực tế": ["8", "12"],
                "Ghi chú": ["Bình thường", "về sớm"]
            })
            df_export.to_excel(writer, index=False, sheet_name=f'ChamCong_{selected_date_str}')
        excel_data = output.getvalue()
        
        st.download_button(
            label=f"📥 下载 Excel 报 cáo / Tải xuống file Excel",
            data=excel_data,
            file_name=f"Attendance_Report_{selected_date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("请先运行考勤分析以启用下载 / Vui lòng chạy phân tích để tải file.")

with col_dl2:
    if st.button("📥 下载 PDF 界面报 cáo / Tải xuống file PDF"):
        try:
            from fpdf import FPDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt=f"ATTENDANCE REPORT - {selected_date_str}", ln=1, align="C")
            pdf.ln(10)
            pdf.cell(200, 10, txt=f"Selected Date: {selected_date_str}", ln=1)
            pdf.cell(200, 10, txt=f"Total Office Staff (VP): {total_vp}", ln=1)
            pdf.cell(200, 10, txt=f"Total Workers (CN): {total_cn}", ln=1)
            pdf.cell(200, 10, txt=f"Active Attendance Count: {active_count}", ln=1)
            
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            st.download_button(
                label="📄 确认下载 PDF / Xác nhận tải PDF",
                data=pdf_bytes,
                file_name=f"Dashboard_Report_{selected_date_str}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.warning(f"PDF 导出需要安装 fpdf 库 / Cần cài đặt thư viện fpdf: {e}")
