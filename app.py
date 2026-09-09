import io
import os
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types

st.set_page_config(
    page_title="Hệ thống Chấm công & Xử lý Nhân sự AI",
    layout="wide",
)

st.title("📊 Hệ thống Chấm công Tự động tích hợp Gemini AI")
st.markdown(
    "Tải lên các file dữ liệu theo yêu cầu để hệ thống tổng hợp và phân tích."
)

# Khởi tạo Gemini Client (Lấy API key từ st.secrets hoặc biến môi trường)
api_key = st.secrets.get("GEMINI_API_KEY") or os.getenv("GEMINI_API_KEY")

if not api_key:
    st.warning("⚠️ Vui lòng cấu hình `GEMINI_API_KEY` trong Streamlit Secrets.")
    st.stop()

client = genai.Client(api_key=api_key)

# --- KHU VỰC TẢI FILE (SIDEBAR HOẶC MAIN) ---
st.subheader("1. Tải lên dữ liệu đầu vào")
col1, col2 = st.columns(2)

with col1:
    uploaded_fingerprint = st.file_uploader(
        "Tải file Excel chấm công (bấm vân tay)", type=["xlsx", "xls", "csv"]
    )
    uploaded_shift = st.file_uploader(
        "Tải file lịch xếp ca (dành cho công nhân)",
        type=["xlsx", "xls", "csv"],
    )

with col2:
    uploaded_vp = st.file_uploader(
        "Tải file danh sách Nhân viên Văn phòng (có Mã NV)",
        type=["xlsx", "xls", "csv"],
    )
    uploaded_cn = st.file_uploader(
        "Tải file danh sách Công nhân (có Mã VN)", type=["xlsx", "xls", "csv"]
    )

# --- XỬ LÝ KHI BẤM NÚT ---
if st.button("🚀 Bắt đầu Xử lý và Phân tích với Gemini AI", type="primary"):
  if not (
      uploaded_fingerprint
      and uploaded_shift
      and uploaded_vp
      and uploaded_cn
  ):
    st.error("Vui lòng tải lên đầy đủ cả 4 file dữ liệu!")
  else:
    with st.spinner(
        "Đang đọc dữ liệu và gửi yêu cầu phân tích cho Gemini AI..."
    ):
      try:
        # Đọc dữ liệu cơ bản bằng Pandas để kiểm tra và chuyển sang dạng văn bản/JSON cho AI
        df_fp = pd.read_excel(uploaded_fingerprint)
        df_shift = pd.read_excel(uploaded_shift)
        df_vp = pd.read_excel(uploaded_vp)
        df_cn = pd.read_excel(uploaded_cn)

        # Chuyển đổi dữ liệu mẫu thành chuỗi để đưa vào prompt cho Gemini
        # (Đối với file lớn, bạn nên lọc dữ liệu hoặc chỉ gửi các mẫu/summary cần thiết)
        prompt = f"""
                Bạn là một chuyên gia nhân sự và hệ thống xử lý dữ liệu chấm công. 
                Hãy xử lý dữ liệu dựa trên các quy tắc nghiêm ngặt sau:

                QUY TẮC NGHIỆP VỤ:
                1. Danh sách nhân sự gồm Nhân viên Văn phòng (VP) và Công nhân (CN).
                2. Giờ làm việc chuẩn:
                   - Nhân viên VP: Giờ vào chuẩn 08:00 AM, Giờ ra chuẩn 17:00 PM. Định mức đủ 8 tiếng/ngày. Nếu không đủ là về sớm, ghi chú giờ làm thực tế.
                   - Công nhân (CN): Giờ vào chuẩn 07:00 AM, Giờ ra chuẩn 19:00 PM (áp dụng lịch xếp ca).
                3. Lịch xếp ca chỉ áp dụng cho Công nhân:
                   - Ngày nghỉ hàng tuần được tô màu cam (hoặc đánh dấu nghỉ).
                   - Ngày đi làm ban ngày ghi là "N" (tương ứng 12 tiếng).
                   - Ca đêm ghi là "Đ" (tương ứng 12 tiếng).
                   - Không làm đủ ca thì ghi số giờ thực tế làm trong ngày đó.
                4. Xử lý thiếu giờ:
                   - Chỉ dò thấy giờ vào hoặc giờ ra: Ghi chú "Thiếu giờ Vào = BTV" hoặc "Thiếu giờ Ra = BTR".
                5. Yêu cầu đầu ra:
                   Tạo một bảng thống kê chi tiết cho từng nhân viên bao gồm các cột: 
                   Mã nhân viên, Họ tên, Loại nhân sự (VP/CN), Số giờ làm thực tế, Đi trễ, Về sớm, Vắng, và Tổng số lần đi trễ trên/tuần.

                DƯỚI ĐÂY LÀ DỮ LIỆU ĐẦU VÀO (được chuyển đổi dạng bảng):
                --- DANH SÁCH VĂN PHÒNG ---
                {df_vp.head(20).to_string()}

                --- DANH SÁCH CÔNG NHÂN ---
                {df_cn.head(20).to_string()}

                --- LỊCH XẾP CA (MẪU) ---
                {df_shift.head(20).to_string()}

                --- DỮ LIỆU CHẤM CÔNG (VÂN TAY - MẪU) ---
                {df_fp.head(20).to_string()}

                Hãy phân tích toàn bộ logic trên, thực hiện tính toán và trả về kết quả cuối cùng dưới dạng bảng dữ liệu Markdown rõ ràng hoặc cấu trúc bảng có thể chuyển đổi thành file Excel.
                """

        # Gọi Gemini Model (Sử dụng model gemini-2.5-flash hoặc gemini-2.5-pro tùy nhu cầu)
        response = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt,
        )

        st.success("✨ Đã xử lý xong dữ liệu!")
        st.markdown("### 📋 Kết quả thống kê từ Gemini AI:")
        st.markdown(response.text)

        # Tính năng hỗ trợ tải file kết quả (Giả lập chuyển đổi markdown thành file excel để tải về)
        # (Thực tế bạn có thể yêu cầu Gemini trả về JSON để convert sang DataFrame và xuất file .xlsx)

      except Exception as e:
        st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
