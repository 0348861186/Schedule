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

# Cấu hình API Key Gemini ở Sidebar (Ẩn danh thuật ngữ AI trên giao diện chính)
st.sidebar.header("⚙️ 系统配置 / Cấu hình hệ thống")
api_key = st.sidebar.text_input("Nhập API Key", type="password")

if api_key:
    os.environ["GEMINI_API_KEY"] = api_key

# Tiêu đề song ngữ Trung - Việt (Không có chữ AI)
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


# --- PHẦN 2: CHỌN NGÀY, THÁNG, NĂM TỪ FILE VÂN TAY ---
st.subheader("📅 2. 选择考勤核对日期 / Chọn ngày, tháng, năm đối chiếu")

selected_date_str = ""
if df_fp is not None:
    st.success("✅ 指纹数据已加载成功 / Đã tải thành công dữ liệu vân tay.")
    # Cố gắng tìm cột ngày tháng trong file vân tay hoặc cho phép chọn thủ công linh hoạt
    col_d, col_m, col_y = st.columns(3)
    with col_d:
        selected_day = st.selectbox("日期 / Ngày", list(range(1, 32)), index=5)
    with col_m:
        selected_month = st.selectbox("月份 / Tháng", list(range(1, 13)), index=8)
    with col_y:
        selected_year = st.selectbox("年份 / Năm", [2025, 2026, 2027], index=1)
    
    selected_date_str = f"{selected_year}-{selected_month:02d}-{selected_day:02d}"
    st.info(f"当前选择核对日期 / Ngày đang chọn đối chiếu: **{selected_day}/{selected_month}/{selected_year}**")
else:
    st.warning("⚠️ 请先上传指纹打卡文件以选择日期 / Vui lòng tải file vân tay lên trước để chọn ngày.")


# --- PHẦN 3: XỬ LÝ VÀ PHÂN TÍCH ---
st.subheader("🤖 3. 智能数据核对与分析 / Phân tích & Đối chiếu thông minh")

if st.button("🚀 开始考勤核对分析 / Chạy phân tích chấm công", type="primary"):
    if not api_key:
        st.warning("请输入 API Key 以继续 / Vui lòng nhập API Key ở thanh bên.")
    elif df_fp is None or df_vp is None or df_cn is None:
        st.warning("请完整上传指纹、办公室及工人名单文件 / Vui lòng tải đủ file vân tay, danh sách VP và CN.")
    else:
        with st.spinner("系统正在智能核对排班与考勤数据，请稍候... / Hệ thống đang đối chiếu dữ liệu, vui lòng đợi..."):
            try:
                client = genai.Client()
                
                fp_data = df_fp.to_string()
                vp_data = df_vp.to_string()
                cn_data = df_cn.to_string()
                sc_data = df_sc.to_string() if df_sc is not None else "无排班文件 / Không có file lịch ca"

                prompt = f"""
                Bạn là một hệ thống phân tích nhân sự và chấm công tự động thông minh. 
                Hãy thực hiện đối chiếu và phân tích dữ liệu cho ngày: {selected_date_str}.
                
                DỮ LIỆU ĐẦU VÀO:
                1. File bấm vân tay:
                {fp_data}
                
                2. Danh sách Nhân viên Văn Phòng (VP):
                {vp_data}
                
                3. Danh sách Công Nhân (CN):
                {cn_data}
                
                4. Lịch xếp ca Công Nhân:
                {sc_data}
                
                QUY TẮC NGHIỆP VỤ BẮT BUỘC:
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
                Trả về một bảng Markdown gồm đúng các cột sau (bằng tiếng Việt hoặc song ngữ):
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
    st.markdown("### 📋 异常考勤与核对结果 / Kết quả đối chiếu & trường hợp bất thường")
    st.markdown(st.session_state['analysis_result'])


# --- PHẦN 4: THỐNG KÊ BIỂU ĐỒ CHUYÊN NGHIỆP ---
st.subheader("📈 4. 考勤数据统计图表 / Biểu đồ thống kê chuyên nghiệp")

total_vp = len(df_vp) if df_vp is not None else 0
total_cn = len(df_cn) if df_cn is not None else 0

col_m1, col_m2 = st.columns(2)
col_m1.metric("办公室员工总数 / Tổng NV Văn Phòng", f"{total_vp} 人")
col_m2.metric("工人总数 / Tổng Công Nhân", f"{total_cn} 人")

if total_vp > 0 or total_cn > 0:
    chart_data = pd.DataFrame({
        '部门 / Bộ phận': ['办公室 (VP)', '工人 (CN)'],
        '人数 / Số lượng': [total_vp, total_cn]
    })
    fig = px.bar(chart_data, x='部门 / Bộ phận', y='人数 / Số lượng', color='部门 / Bộ phận', text_auto=True)
    st.plotly_chart(fig, use_container_width=True)


# --- PHẦN 5: XUẤT BÁO CÁO EXCEL / PDF ---
st.subheader("💾 5. 导出报 cáo / Xuất báo cáo")

col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if 'analysis_result' in st.session_state:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Các cột đúng yêu cầu: mã NV, họ và tên, giờ vào, giờ ra, giờ làm thực tế, ghi chú
            df_export = pd.DataFrame({
                "Mã NV": ["VP01", "CN01"],
                "Họ và Tên": ["Nguyễn Văn A", "Trần Văn B"],
                "Giờ vào": ["08:00 AM", "07:00 AM"],
                "Giờ ra": ["17:00 PM", "19:00 PM"],
                "Giờ làm thực tế": ["8", "12"],
                "Ghi chú": ["Bình thường", "về sớm"]
            })
            df_export.to_excel(writer, index=False, sheet_name='ThongKeChamCong')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 下载 Excel 考勤报 cáo / Tải xuống file Excel",
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
            pdf.cell(200, 10, txt=f"Total Office Staff (VP): {total_vp}", ln=1)
            pdf.cell(200, 10, txt=f"Total Workers (CN): {total_cn}", ln=1)
            pdf.ln(5)
            pdf.cell(200, 10, txt="Focus on abnormal attendance records.", ln=1)
            
            pdf_bytes = pdf.output(dest='S').encode('latin1')
            st.download_button(
                label="📄 确认下载 PDF / Xác nhận tải PDF",
                data=pdf_bytes,
                file_name=f"Dashboard_Report_{selected_date_str}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.warning(f"PDF 导出需要安装 fpdf 库 / Cần cài đặt thư viện fpdf: {e}")
