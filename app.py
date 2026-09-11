import streamlit as st
import pandas as pd
import google.generativeai as genai

# ==========================================
# CẤU HÌNH GEMINI AI
# ==========================================
# Khuyên dùng: Nên cấu hình key trong mục "Settings" -> "Secrets" của Streamlit Cloud với biến GEMINI_KEY
if "GEMINI_KEY" in st.secrets:
    GOOGLE_API_KEY = st.secrets["GEMINI_KEY"]
else:
    GOOGLE_API_KEY = "THAY_KEY_GEMINI_CỦA_BẠN_VÀO_ĐÂY" # Hoặc dán trực tiếp key vào đây nếu chạy local

genai.configure(api_key=GOOGLE_API_KEY)
model = genai.GenerativeModel('gemini-1.5-flash')

# Cấu hình giao diện Streamlit
st.set_page_config(page_title="Hệ thống chấm công AI", layout="wide")
st.title("📊 Hệ Thống Quản Lý & Phân Tích Chấm Công Thông Minh (AI)")

# ==========================================
# SIDEBAR - TẢI VÀ LOAD CÁC FILE EXCEL
# ==========================================
st.sidebar.header("📁 Upload dữ liệu đầu vào")

file_van_tay = st.sidebar.file_uploader("1. Tải file Bấm Vân Tay (.xlsx)", type=["xlsx"])
file_xep_ca = st.sidebar.file_uploader("2. Tải file Lịch Xếp Ca (.xlsx)", type=["xlsx"])
file_vp = st.sidebar.file_uploader("3. Tải Danh sách Nhân viên VP (.xlsx)", type=["xlsx"])
file_cn = st.sidebar.file_uploader("4. Tải Danh sách Công nhân (.xlsx)", type=["xlsx"])

# Hàm tiện ích giúp chuẩn hóa tên cột (Xóa khoảng trắng ở đầu/cuối, viết thường)
def chuan_hoa_columns(df):
    if df is not None:
        df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

# Kiểm tra xem đã tải đủ file bắt buộc chưa
if file_van_tay and file_vp and file_cn:
    # Đọc dữ liệu thô từ các file
    df_van_tay = chuan_hoa_columns(pd.read_excel(file_van_tay))
    df_vp = chuan_hoa_columns(pd.read_excel(file_vp))
    df_cn = chuan_hoa_columns(pd.read_excel(file_cn))
    df_xep_ca = chuan_hoa_columns(pd.read_excel(file_xep_ca)) if file_xep_ca else None
    
    # Tự động tìm cột 'ngày' sau khi đã chuẩn hóa viết thường
    ten_cot_ngay = 'ngày' if 'ngày' in df_van_tay.columns else df_van_tay.columns[4] # Nếu không thấy chữ 'ngày', mặc định lấy cột thứ 5 (vị trí index 4) theo ảnh mẫu
    
    # Ép kiểu dữ liệu Cột Ngày sang chuỗi để bộ lọc hoạt động chính xác
    df_van_tay[ten_cot_ngay] = df_van_tay[ten_cot_ngay].astype(str).str.strip()

    # ==========================================
    # BỘ LỌC THỜI GIAN TRÊN DASHBOARD
    # ==========================================
    st.subheader("🔍 Chọn thời gian cần kiểm tra")
    
    # Lấy danh sách các ngày duy nhất có trong file vân tay để user chọn
    danh_sach_ngay = sorted(df_van_tay[ten_cot_ngay].unique())
    ngay_chon = st.selectbox("Chọn Ngày/Tháng/Năm cần thống kê:", danh_sach_ngay)
    
    if st.button("🚀 Bắt đầu phân tích dữ liệu bằng AI"):
        with st.spinner("AI đang đối chiếu dữ liệu vân tay và lịch trình làm việc..."):
            
            # 1. Lọc dữ liệu vân tay của ngày được chọn
            df_van_tay_loc = df_van_tay[df_van_tay[ten_cot_ngay] == ngay_chon]
            
            # Chuyển dataframe thành dạng Text/Markdown để gửi cho Gemini
            Text_Van_Tay = df_van_tay_loc.to_markdown(index=False)
            Text_VP = df_vp.to_markdown(index=False)
            Text_CN = df_cn.to_markdown(index=False)
            Text_Xep_Ca = df_xep_ca.to_markdown(index=False) if df_xep_ca is not None else "Không có lịch xếp ca công nhân"

            # 2. Xây dựng Prompt gửi cho Gemini AI
            prompt = f"""
            Bạn là một chuyên gia nhân sự và phân tích dữ liệu chấm công cấp cao. 
            Nhiệm vụ của bạn là đối chiếu dữ liệu từ các bảng sau để xác định trạng thái đi làm của nhân viên vào ngày {ngay_chon}.

            DANH SÁCH DỮ LIỆU ĐẦU VÀO (Tên các cột đã được chuyển về chữ viết thường):
            ---
            1. BẢNG BẤM VÂN TAY (Ngày {ngay_chon}):
            {Text_Van_Tay}
            
            2. BẢNG DANH SÁCH NHÂN VIÊN VĂN PHÒNG (VP):
            {Text_VP}
            
            3. BẢNG DANH SÁCH CÔNG NHÂN (CN):
            {Text_CN}
            
            4. BẢNG LỊCH XẾP CA (Dành riêng cho Công nhân):
            {Text_Xep_Ca}
            ---

            QUY TẮC ĐỐI CHIẾU VÀ ĐÁNH GIÁ (BẮT BUỘC TUÂN THỦ):

            Nhóm 1: CÔNG NHÂN (Mã NV nằm trong danh sách Công nhân)
            - Hãy xem Mã NV đó được xếp ca nào trong bảng "LỊCH XẾP CA" vào ngày {ngay_chon}.
            - Ví dụ: Nếu ca là "Đ" (Ca đêm): Giờ vào chuẩn là 19:00 PM cùng ngày và Giờ ra chuẩn là 7:00 AM ngày hôm sau. 
            - Lưu ý đặc biệt ca đêm: Kiểm tra giờ ra 7:00 AM trong dữ liệu vân tay của ngày kế tiếp hoặc bản ghi kế cận. Nếu vào sau 19:00 là Đi trễ, ra trước 7:00 là Về sớm. Nếu thiếu giờ vào hoặc ra là "Quên bấm thẻ". Nếu không có dữ liệu là "Vắng mặt".
            - Đối với các ký hiệu ca ngày khác, hãy tự suy luận khung giờ chuẩn dựa theo mẫu thông thường của nhà máy hoặc lịch ca.

            Nhóm 2: NHÂN VIÊN VĂN PHÒNG (Mã NV nằm trong danh sách Nhân viên VP)
            - Giờ làm việc chuẩn mặc định: Vào 8:00 AM - Ra 17:00 PM cùng ngày.
            - Đối chiếu trực tiếp Mã NV sang bảng Vân tay ngày {ngay_chon} để kết luận: Đi trễ (Vào > 8:00), Về sớm (Ra < 17:00), Quên bấm thẻ (thiếu 1 trong 2 giờ), hoặc Vắng mặt (Không có dữ liệu).

            Nhóm 3: NHÓM ĐẶC BIỆT CỐ ĐỊNH (Mã NV không thuộc 2 danh sách trên):
            - Mã "575": Vào chuẩn 7:00 AM, Ra chuẩn 15:00 PM cùng ngày.
            - Mã "749": Vào chuẩn 7:00 AM, Ra chuẩn 19:00 PM cùng ngày.
            - Mã "949": Vào chuẩn 7:00 AM, Ra chuẩn 19:00 PM cùng ngày.
            - Mã "673": Vào chuẩn 7:00 AM, Ra chuẩn 15:00 PM cùng ngày.
            - Mã "A068": Vào chuẩn 10:00 AM, Ra chuẩn 18:00 PM cùng ngày.
            Dò các mã này trực tiếp trong Bảng Vân Tay dựa theo giờ chuẩn trên.

            YÊU CẦU ĐẦU RA:
            Hãy trả về kết quả dưới dạng một bảng tổng hợp báo cáo chi tiết gồm các cột:

            | Mã Nhân Viên | Tên Nhân Viên | Bộ phận/Nhóm | Ca làm việc | Giờ vào thực tế | Giờ ra thực tế | Trạng thái (Đúng giờ / Đi trễ / Về sớm / Quên bấm thẻ / Vắng) | Ghi chú chi tiết lý do |
            
            Sau đó, liệt kê danh sách tóm tắt nhanh:
            - Danh sách đi trễ:
            - Danh sách về sớm:
            - Danh sách quên bấm thẻ:
            - Danh sách vắng không lý do:
            """

            # 3. Gọi Gemini AI xử lý
            response = model.generate_content(prompt)
            
            # 4. Hiển thị kết quả lên màn hình Dashboard
            st.success("🎉 Đã phân tích xong dữ liệu bằng AI!")
            st.markdown("### 📋 Kết Quả Phân Tích Chi Tiết Từ Gemini AI")
            st.write(response.text)

else:
    st.info("💡 Vui lòng upload đầy đủ tối thiểu các file: Bấm vân tay, Danh sách Nhân viên VP, Danh sách Công nhân ở thanh menu bên trái để bắt đầu.")
