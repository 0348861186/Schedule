import streamlit as st
import pandas as pd
from datetime import datetime, time
import google.generativeai as genai
import json
import io
 
# ==========================================
# CẤU HÌNH CONFIG & TOÀN CỤC
# ==========================================
if "GEMINI_API_KEY" in st.secrets:
   genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
   genai.configure(api_key="YOUR_GEMINI_API_KEY_HERE") # Thay API Key của bạn nếu chạy Local
 
st.set_page_config(page_title="AI Attendance Dashboard", layout="wide")
st.title("🤖 DASHBOARD QUẢN LÝ GIỜ CÔNG THÔNG MINH (PYTHON + GEMINI AI)")
 
# ==========================================
# 1. THANH BÊN (SIDEBAR) - TẢI FILE EXCEL
# ==========================================
st.sidebar.header("📁 Tải Lên Dữ Liệu Excel")
file_vantay = st.sidebar.file_uploader("1. File Bấm Vân Tay", type=["xlsx", "xls"])
file_lichca = st.sidebar.file_uploader("2. File Lịch Xếp Ca (Công nhân)", type=["xlsx", "xls"])
file_vp = st.sidebar.file_uploader("3. Danh Sách Nhân Viên VP", type=["xlsx", "xls"])
file_cn = st.sidebar.file_uploader("4. Danh Sách Công Nhân", type=["xlsx", "xls"])
 
# Cấu hình khung giờ cố định cho nhóm Đặc biệt & Văn phòng
NHOM_DAC_BIET = {
    "575": {"vao": time(7, 0), "ra": time(15, 0)},
    "749": {"vao": time(7, 0), "ra": time(19, 0)},
    "949": {"vao": time(7, 0), "ra": time(19, 0)},
    "673": {"vao": time(7, 0), "ra": time(15, 0)},
    "A068": {"vao": time(10, 0), "ra": time(18, 0)},
}
GIO_VP = {"vao": time(8, 0), "ra": time(17, 0)}
 
# Khung giờ làm việc theo ca của Công nhân
GIO_CA_CN = {
    "Đ": {"vao": time(19, 0), "ra": time(7, 0), "qua_dem": True},
    "H": {"vao": time(7, 0), "ra": time(16, 0), "qua_dem": False},
}
 
# ==========================================
# 2. XỬ LÝ DỮ LIỆU LOGIC KHI ĐỦ FILE
# ==========================================
if file_vantay and file_vp and file_cn:
    # Đọc dữ liệu thô ban đầu
    df_vantay = pd.read_excel(file_vantay)
    df_vp = pd.read_excel(file_vp)
    df_cn = pd.read_excel(file_cn)
    df_lichca = pd.read_excel(file_lichca) if file_lichca else None
    
    # 🌟 CHUẨN HÓA TÊN CỘT THÔNG MINH (Tránh hoàn toàn lỗi KeyError)
    for df in [df_vantay, df_vp, df_cn] + ([df_lichca] if df_lichca is not None else []):
        df.columns = df.columns.astype(str).str.strip()
        
        # Tìm cột tương đương Mã Nhân Viên
        ma_nv_col = [c for c in df.columns if c.lower() in ['mã nv', 'manv', 'ma nv', 'mã nhân viên', 'ma nhan vien', 'id', 'mã số']]
        if ma_nv_col:
           df.rename(columns={ma_nv_col[0]: 'Mã NV'}, inplace=True)
           df['Mã NV'] = df['Mã NV'].astype(str).str.strip()
            
    # Đồng bộ hóa cột Ngày, Giờ trong file Vân tay
    df_vantay.columns = df_vantay.columns.astype(str).str.strip()
    ngay_col = [c for c in df_vantay.columns if c.lower() in ['ngày', 'ngay', 'date']]
    if ngay_col:
        df_vantay.rename(columns={ngay_col[0]: 'Ngày'}, inplace=True)
        
    gio_vao_col = [c for c in df_vantay.columns if c.lower() in ['giờ vào', 'gio vao', 'vào', 'time in', 'giờ checkin', 'checkin']]
    if gio_vao_col:
        df_vantay.rename(columns={gio_vao_col[0]: 'Giờ Vào'}, inplace=True)
        
    gio_ra_col = [c for c in df_vantay.columns if c.lower() in ['giờ ra', 'gio ra', 'ra', 'time out', 'giờ checkout', 'checkout']]
    if gio_ra_col:
        df_vantay.rename(columns={gio_ra_col[0]: 'Giờ Ra'}, inplace=True)
 
    # Đồng bộ hóa dữ liệu file Lịch ca công nhân
    if df_lichca is not None:
        ngay_ca_col = [c for c in df_lichca.columns if c.lower() in ['ngày', 'ngay', 'date']]
        if ngay_ca_col: 
            df_lichca.rename(columns={ngay_ca_col[0]: 'Ngày'}, inplace=True)
        df_lichca['Ngày'] = pd.to_datetime(df_lichca['Ngày']).dt.date
        
        ca_col = [c for c in df_lichca.columns if c.lower() in ['ca', 'ca làm việc', 'shift']]
        if ca_col:
            df_lichca.rename(columns={ca_col[0]: 'Ca'}, inplace=True)
 
    # Chuyển đổi định dạng Ngày của Vân Tay
    df_vantay['Ngày'] = pd.to_datetime(df_vantay['Ngày']).dt.date
    
    # 🎯 BỘ CHỌN THỜI GIAN TRÊN DASHBOARD
    cac_ngay_co_san = sorted(df_vantay['Ngày'].unique())
    st.subheader("🎯 Chọn Thời Gian Kiểm Tra")
    ngay_chon = st.selectbox("Chọn ngày cần kết xuất báo cáo từ file vân tay:", cac_ngay_co_san)
    
    # Lọc dữ liệu vân tay đúng theo ngày được chỉ định
    df_vantay_ngay = df_vantay[df_vantay['Ngày'] == ngay_chon]
    
    # Gom danh sách phân loại nhân sự tổng thể
    ds_ma_vp = df_vp['Mã NV'].tolist()
    ds_ma_cn = df_cn['Mã NV'].tolist()
    ds_dac_biet = list(NHOM_DAC_BIET.keys())
    tat_ca_ma = set(ds_ma_vp + ds_ma_cn + ds_dac_biet)
    
    ket_qua_python = []
    dong_cho_ai_xu_ly = [] # Hàng đợi lưu các dòng bị lỗi hoặc ca đêm phức tạp chuyển giao sang AI
 
    # --- BƯỚC 1: LỚP PYTHON XỬ LÝ LÕI CỨNG ---
    for ma in tat_ca_ma:
        loai_nv = "Chưa phân loại"
        gio_vao_chuan, gio_ra_chuan = None, None
        ca_lam_viec = "Hành chính"
        
        # Phân loại nhóm và gán khung giờ đích
        if ma in ds_dac_biet:
            loai_nv = "Nhóm Đặc Biệt"
            gio_vao_chuan = NHOM_DAC_BIET[ma]["vao"]
            gio_ra_chuan = NHOM_DAC_BIET[ma]["ra"]
        elif ma in ds_ma_vp:
            loai_nv = "Văn Phòng"
            gio_vao_chuan = GIO_VP["vao"]
            gio_ra_chuan = GIO_VP["ra"]
        elif ma in ds_ma_cn:
            loai_nv = "Công Nhân"
            if df_lichca is not None:
               lich_hom_nay = df_lichca[(df_lichca['Mã NV'] == ma) & (df_lichca['Ngày'] == ngay_chon)]
               if not lich_hom_nay.empty:
                   ca_lam_viec = str(lich_hom_nay.iloc[0]['Ca']).strip()
                   if ca_lam_viec in GIO_CA_CN:
                       gio_vao_chuan = GIO_CA_CN[ca_lam_viec]["vao"]
                       gio_ra_chuan = GIO_CA_CN[ca_lam_viec]["ra"]
                       
                       # Nếu là ca đêm có tính chất gãy ngày, đưa thẳng vào hàng đợi để AI kiểm tra chéo vân tay
                       if GIO_CA_CN[ca_lam_viec].get("qua_dem"):
                           gio_vao_chuan, gio_ra_chuan = "AI_NEEDED", "AI_NEEDED"
                   else:
                       gio_vao_chuan, gio_ra_chuan = "AI_NEEDED", "AI_NEEDED" # Gặp ký hiệu ca lạ
               else:
                   loai_nv = "Công Nhân (Nghỉ/Không lịch ca)"
            else:
               gio_vao_chuan, gio_ra_chuan = "AI_NEEDED", "AI_NEEDED"
 
        # Đẩy thẳng sang hàng đợi AI nếu thuộc ca phức tạp hoặc Python không tra cứu được giờ chuẩn
        if gio_vao_chuan == "AI_NEEDED" or gio_vao_chuan is None:
           dong_van_tay_loi = df_vantay_ngay[df_vantay_ngay['Mã NV'] == ma]
           dong_cho_ai_xu_ly.append({
               "Mã NV": ma, "Bộ phận": loai_nv, "Ca": ca_lam_viec,
               "Dữ liệu vân tay ngày này": dong_van_tay_loi.to_dict('records')
           })
           continue
 
        # Tìm kiếm dòng tương ứng trong file Vân Tay
        dong_van_tay = df_vantay_ngay[df_vantay_ngay['Mã NV'] == ma]
        
        if dong_van_tay.empty:
            ket_qua_python.append({
                "Mã NV": ma, "Bộ phận": loai_nv, "Ca": ca_lam_viec,
                "Giờ Vào Chuẩn": gio_vao_chuan.strftime("%H:%M:%S"), "Giờ Ra Chuẩn": gio_ra_chuan.strftime("%H:%M:%S"),
                "Giờ Vào Thực Tế": "N/A", "Giờ Ra Thực Tế": "N/A",
                "Trạng Thái": "Vắng mặt (Không bấm thẻ)", "Xử lý bởi": "Python"
            })
        else:
            try:
                v_vao = dong_van_tay.iloc[0]['Giờ Vào']
                v_ra = dong_van_tay.iloc[0]['Giờ Ra']
                
                # Ép kiểu dữ liệu thời gian thô từ file Excel
                g_vao = datetime.strptime(str(v_vao).strip(), "%H:%M:%S").time() if pd.notna(v_vao) else None
                g_ra = datetime.strptime(str(v_ra).strip(), "%H:%M:%S").time() if pd.notna(v_ra) else None
                
                if not g_vao or not g_ra:
                   raise ValueError("Thiếu dữ liệu check-in/out")
                
                # Logic phân tích Đi trễ / Về sớm
                ly_do = []
                if g_vao > gio_vao_chuan: ly_do.append("Đi trễ")
                if g_ra < gio_ra_chuan: ly_do.append("Về sớm")
                
                trang_thai = " + ".join(ly_do) if ly_do else "Đúng giờ"
                
                ket_qua_python.append({
                   "Mã NV": ma, "Bộ phận": loai_nv, "Ca": ca_lam_viec,
                   "Giờ Vào Chuẩn": gio_vao_chuan.strftime("%H:%M:%S"), "Giờ Ra Chuẩn": gio_ra_chuan.strftime("%H:%M:%S"),
                   "Giờ Vào Thực Tế": g_vao.strftime("%H:%M:%S"), "Giờ Ra Thực Tế": g_ra.strftime("%H:%M:%S"),
                   "Trạng Thái": trang_thai, "Xử lý bởi": "Python"
                })
            except Exception:
                # Gặp bất cứ lỗi định dạng nào trong ô, chuyển tiếp sang cho Gemini giải cứu
                dong_cho_ai_xu_ly.append({
                    "Mã NV": ma, "Bộ phận": loai_nv, "Ca": ca_lam_viec,
                    "Giờ Vào Chuẩn": gio_vao_chuan.strftime("%H:%M:%S") if isinstance(gio_vao_chuan, time) else str(gio_vao_chuan),
                    "Giờ Ra Chuẩn": gio_ra_chuan.strftime("%H:%M:%S") if isinstance(gio_ra_chuan, time) else str(gio_ra_chuan),
                    "Dữ liệu thô lỗi": dong_van_tay.to_dict('records')
                })
 
    df_sach_python = pd.DataFrame(ket_qua_python)
 
    # --- BƯỚC 2: PHÂN CẤP SỬ LÝ CỨU HỘ BỞI GEMINI AI ---
    df_cuoi_cung = df_sach_python
    if dong_cho_ai_xu_ly:
        st.warning(f"⚡ Phát hiện {len(dong_cho_ai_xu_ly)} trường hợp dữ liệu ca gãy hoặc ca đêm phức tạp. Đang chuyển giao dữ liệu qua Gemini AI phân tích...")
        
        with st.spinner("Gemini AI đang suy luận thông minh từ dữ liệu thô..."):
            prompt = f"""
            Bạn là một chuyên gia AI quản lý chấm công. Dưới đây là danh sách nhân sự có lỗi dữ liệu hoặc có ca đêm phức tạp nhảy ngày mà mã Python không tự giải quyết được:
            {json.dumps(dong_cho_ai_xu_ly, default=str, ensure_ascii=False)}
            Nhiệm vụ của bạn:

            1. Đọc và phân tích thông tin của từng nhân viên.

            2. Trích xuất giờ vào thực tế và giờ ra thực tế từ "Dữ liệu vân tay ngày này" hoặc "Dữ liệu thô lỗi".

            3. Nếu là ca Đêm ("Ca": "Đ"), giờ vào chuẩn là 19:00:00 và giờ ra chuẩn là 07:00:00 sáng hôm sau. Đối chiếu xem nhân viên có bấm thẻ khớp hay không.

            4. Trả về kết quả dưới dạng một MẢNG JSON DUY NHẤT. Mỗi phần tử trong mảng có cấu trúc chuẩn như sau:

            - "Mã NV": (giữ nguyên)

            - "Bộ phận": (giữ nguyên)

            - "Ca": (giữ nguyên)

            - "Giờ Vào Chuẩn": (định dạng HH:MM:SS)

            - "Giờ Ra Chuẩn": (định dạng HH:MM:SS)

            - "Giờ Vào Thực Tế": (định dạng HH:MM:SS hoặc "N/A" nếu không có)

            - "Giờ Ra Thực Tế": (định dạng HH:MM:SS hoặc "N/A" nếu không có)

            - "Trạng Thái": (Điền chính xác: "Đi trễ", "Về sớm", "Đi trễ + Về sớm", "Đúng giờ", hoặc "Vắng mặt (Không bấm thẻ)")

            - "Xử lý bởi": "Gemini AI"
            LƯU Ý: Không thêm bất kỳ dòng văn bản giải thích nào ngoài đoạn mã JSON để Python có thể đọc trực tiếp.

            """
 
            try:
                model = genai.GenerativeModel("gemini-1.5-flash")
                response = model.generate_content(prompt)
                clean_text = response.text.replace("```json", "").replace("```", "").strip()
 
                ket_qua_ai = json.loads(clean_text)
                df_ai = pd.DataFrame(ket_qua_ai)
                
                # Trộn báo cáo đã qua xử lý của cả hai lớp Python và AI
                df_cuoi_cung = pd.concat([df_sach_python, df_ai], ignore_index=True)
 
            except Exception as e:
                st.error(f"⚠️ Gemini AI lỗi cấu trúc trả về. Sử dụng tạm kết quả từ Python. Chi tiết lỗi: {e}")
 
# ==========================================
# 3. HIỂN THỊ KẾT QUẢ & NÚT XUẤT FILE EXCEL
# ==========================================
if file_vantay and file_vp and file_cn:
    st.subheader(f"📋 Bảng Thống Kê Giờ Công Tổng Hợp - Ngày {ngay_chon.strftime('%d/%m/%Y')}")
    st.dataframe(df_cuoi_cung, use_container_width=True)
 
    # Lọc riêng những người vi phạm
    st.subheader("⚠️ Danh Sách Nhân Viên Vi Phạm (Đi Trễ / Về Sớm / Vắng Mặt)")
    df_vi_pham = df_cuoi_cung[df_cuoi_cung['Trạng Thái'] != "Đúng giờ"]
    st.dataframe(df_vi_pham, use_container_width=True)
 
    # 📤 TÍNH NĂNG XUẤT FULL BÁO CÁO RA FILE EXCEL
    st.subheader("📥 Xuất Báo Cáo")
    
    # Tạo luồng dữ liệu nhị phân lưu trữ file Excel tạm thời
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_cuoi_cung.to_excel(writer, index=False, sheet_name="Tổng Hợp Giờ Công")
        df_vi_pham.to_excel(writer, index=False, sheet_name="Danh Sách Vi Phạm")
    buffer.seek(0)
    
    # Nút bấm tải dữ liệu xuống thiết bị máy tính
    st.download_button(
        label="📥 Tải xuống file Excel báo cáo tổng hợp",
        data=buffer,
        file_name=f"Bao_cao_gio_cong_{ngay_chon.strftime('%d_%m_%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👋 Vui lòng tải đầy đủ cả 4 file Excel cần thiết ở thanh bên trái (Sidebar) để kích hoạt hệ thống tự động đồng bộ.")
