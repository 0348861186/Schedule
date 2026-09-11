import streamlit as st
import pandas as pd
from datetime import datetime, time
import google.generativeai as genai

# Cấu hình Gemini AI
# Thay bằng API Key của bạn nếu muốn dùng tính năng AI phân tích báo cáo sâu
genai.configure(api_key="YOUR_GEMINI_API_KEY")

st.set_page_config(page_title="Dashboard Quản Lý Giờ Công", layout="wide")
st.title("📊 HỆ THỐNG QUẢN LÝ GIỜ CÔNG TỰ ĐỘNG (PYTHON & GEMINI AI)")

# ==========================================
# 1. THANH BÊN (SIDEBAR) - NÚT TẢI FILE EXCEL
# ==========================================
st.sidebar.header("📁 Tải Lên Dữ Liệu Excel")
file_vantay = st.sidebar.file_uploader("1. File Bấm Vân Tay", type=["xlsx", "xls"])
file_lichca = st.sidebar.file_uploader("2. File Lịch Xếp Ca (Công nhân)", type=["xlsx", "xls"])
file_vp = st.sidebar.file_uploader("3. Danh Sách Nhân Viên VP", type=["xlsx", "xls"])
file_cn = st.sidebar.file_uploader("4. Danh Sách Công Nhân", type=["xlsx", "xls"])

# Định nghĩa giờ chuẩn cho nhóm đặc biệt
NHOM_DAC_BIET = {
    "575": {"vao": time(7, 0), "ra": time(15, 0), "qua_dem": False},
    "749": {"vao": time(7, 0), "ra": time(19, 0), "qua_dem": False},
    "949": {"vao": time(7, 0), "ra": time(19, 0), "qua_dem": False},
    "673": {"vao": time(7, 0), "ra": time(15, 0), "qua_dem": False},
    "A068": {"vao": time(10, 0), "ra": time(18, 0), "qua_dem": False},
}

# Giờ chuẩn mặc định cho Văn Phòng
GIO_VP = {"vao": time(8, 0), "ra": time(17, 0)} 

# Giờ chuẩn theo Ca của Công Nhân
GIO_CA_CN = {
    "Đ": {"vao": time(19, 0), "ra": time(7, 0), "qua_dem": True}, # Ca đêm
    "H": {"vao": time(7, 0), "ra": time(16, 0), "qua_dem": False}, # Ca hành chính ví dụ
}

# ==========================================
# 2. XỬ LÝ DỮ LIỆU CHÍNH
# ==========================================
if file_vantay and file_vp and file_cn:
    # Đọc dữ liệu từ excel
    df_vantay = pd.read_excel(file_vantay)
    df_vp = pd.read_excel(file_vp)
    df_cn = pd.read_excel(file_cn)
    df_lichca = pd.read_excel(file_lichca) if file_lichca else None
    
    # Ép kiểu dữ liệu Mã NV về chuỗi để đồng bộ
    df_vantay['Mã NV'] = df_vantay['Mã NV'].astype(str).str.strip()
    df_vp['Mã NV'] = df_vp['Mã NV'].astype(str).str.strip()
    df_cn['Mã NV'] = df_cn['Mã NV'].astype(str).str.strip()
    
    if df_lichca is not None:
        df_lichca['Mã NV'] = df_lichca['Mã NV'].astype(str).str.strip()
        df_lichca['Ngày'] = pd.to_datetime(df_lichca['Ngày']).dt.date

    # Chuẩn hóa cột ngày trong file vân tay
    df_vantay['Ngày'] = pd.to_datetime(df_vantay['Ngày']).dt.date
    
    # Lấy danh sách các ngày có trong file để hiển thị lên bộ chọn
    cac_ngay_co_san = sorted(df_vantay['Ngày'].unique())
    
    # 🎯 BỘ CHỌN NGÀY THÁNG NĂM TRÊN DASHBOARD
    st.subheader("🎯 Chọn Thời Gian Kiểm Tra")
    ngay_chon = st.selectbox("Chọn ngày cần kết xuất báo cáo:", cac_ngay_co_san)
    
    # Lọc dữ liệu vân tay theo ngày đã chọn
    df_vantay_ngay = df_vantay[df_vantay['Ngày'] == ngay_chon]
    
    # Lấy danh sách tất cả nhân viên cần kiểm tra trong ngày
    ds_ma_vp = df_vp['Mã NV'].tolist()
    ds_ma_cn = df_cn['Mã NV'].tolist()
    ds_dac_biet = list(NHOM_DAC_BIET.keys())
    
    ket_qua = []
    
    # --- LOGIC TOÀN BỘ NHÂN SỰ ---
    # Tổng hợp tất cả mã nhân viên từ các danh sách
    tat_ca_ma = set(ds_ma_vp + ds_ma_cn + ds_dac_biet)
    
    for ma in tat_ca_ma:
        loai_nv = "Chưa phân loại"
        gio_vao_chuan, gio_ra_chuan = None, None
        qua_dem = False
        
        # 1. Xác định nhóm và khung giờ chuẩn
        if ma in ds_dac_biet:
            loai_nv = "Nhóm Đặc Biệt"
            gio_vao_chuan = NHOM_DAC_BIET[ma]["vao"]
            gio_ra_chuan = NHOM_DAC_BIET[ma]["ra"]
            qua_dem = NHOM_DAC_BIET[ma]["qua_dem"]
            
        elif ma in ds_ma_vp:
            loai_nv = "Văn Phòng"
            gio_vao_chuan = GIO_VP["vao"]
            gio_ra_chuan = GIO_VP["ra"]
            
        elif ma in ds_ma_cn:
            loai_nv = "Công Nhân"
            if df_lichca is not None:
                # Tra lịch ca của công nhân ngày hôm đó
                lich_hom_nay = df_lichca[(df_lichca['Mã NV'] == ma) & (df_lichca['Ngày'] == ngay_chon)]
                if not lich_hom_nay.empty:
                    ca = lich_hom_nay.iloc[0]['Ca']
                    if ca in GIO_CA_CN:
                        gio_vao_chuan = GIO_CA_CN[ca]["vao"]
                        gio_ra_chuan = GIO_CA_CN[ca]["ra"]
                        qua_dem = GIO_CA_CN[ca]["qua_dem"]
                    else:
                        loai_nv = "Công Nhân (Nghỉ/Không xếp ca)"
                else:
                    loai_nv = "Công Nhân (Không có lịch ca)"
        
        # Nếu không xác định được giờ chuẩn (ví dụ công nhân không đi ca ngày đó)
        if not gio_vao_chuan:
            continue
            
        # 2. Dò dữ liệu bấm vân tay
        dong_van_tay = df_vantay_ngay[df_vantay_ngay['Mã NV'] == ma]
        
        if file_vantay and file_vp and file_cn:
    # Đọc dữ liệu từ excel
    df_vantay = pd.read_excel(file_vantay)
    df_vp = pd.read_excel(file_vp)
    df_cn = pd.read_excel(file_cn)
    df_lichca = pd.read_excel(file_lichca) if file_lichca else None
    
    # 🌟 ĐOẠN SỬA ĐỔI: TỰ ĐỘNG CHUẨN HÓA TÊN CỘT (Xóa khoảng trắng, viết thường để dò)
    for df in [df_vantay, df_vp, df_cn] + ([df_lichca] if df_lichca is not None else []):
        # Đổi tên cột về dạng viết thường và xóa khoảng trắng hai đầu để dễ dò
        df.columns = df.columns.astype(str).str.strip()
        
        # Tự động tìm và đổi tên cột Mã Nhân Viên về chuẩn 'Mã NV'
        ma_nv_col = [c for c in df.columns if c.lower() in ['mã nv', 'manv', 'ma nv', 'mã nhân viên', 'ma nhan vien', 'id', 'mã số']]
        if ma_nv_col:
            df.rename(columns={ma_nv_col[0]: 'Mã NV'}, inplace=True)
            df['Mã NV'] = df['Mã NV'].astype(str).str.strip() # Ép kiểu chuỗi
            
    # Tự động chuẩn hóa cột Ngày, Giờ cho file vân tay
    df_vantay.columns = df_vantay.columns.astype(str).str.strip()
    ngay_col = [c for c in df_vantay.columns if c.lower() in ['ngày', 'ngay', 'date']]
    if ngay_col: df.rename(columns={ngay_col[0]: 'Ngày'}, inplace=True)
        
    gio_vao_col = [c for c in df_vantay.columns if c.lower() in ['giờ vào', 'gio vao', 'vào', 'time in', 'giờ checkin']]
    if gio_vao_col: df_vantay.rename(columns={gio_vao_col[0]: 'Giờ Vào'}, inplace=True)
        
    gio_ra_col = [c for c in df_vantay.columns if c.lower() in ['giờ ra', 'gio ra', 'ra', 'time out', 'giờ checkout']]
    if gio_ra_col: df_vantay.rename(columns={gio_ra_col[0]: 'Giờ Ra'}, inplace=True)

    # Chuẩn hóa ngày cho file lịch ca
    if df_lichca is not None:
        ngay_ca_col = [c for c in df_lichca.columns if c.lower() in ['ngày', 'ngay', 'date']]
        if ngay_ca_col: df_lichca.rename(columns={ngay_ca_col[0]: 'Ngày'}, inplace=True)
        df_lichca['Ngày'] = pd.to_datetime(df_lichca['Ngày']).dt.date
        
        ca_col = [c for c in df_lichca.columns if c.lower() in ['ca', 'ca làm việc', 'shift']]
        if ca_col: df_lichca.rename(columns={ca_col[0]: 'Ca'}, inplace=True)

    # Tiếp tục xử lý ngày của file vân tay
    df_vantay['Ngày'] = pd.to_datetime(df_vantay['Ngày']).dt.date
        # Thêm vào bảng kết quả
        ket_qua.append({
            "Mã NV": ma,
            "Bộ phận": loai_nv,
            "Giờ Vào Chuẩn": gio_vao_chuan,
            "Giờ Ra Chuẩn": gio_ra_chuan,
            "Giờ Vào Thực Tế": g_vao,
            "Giờ Ra Thực Tế": g_ra,
            "Trạng Thái": trang_thai
        })
        
    # Hiện thị kết quả lên Dashboard dưới dạng bảng
    df_ket_qua = pd.DataFrame(ket_qua)
    
    st.subheader(f"📋 Kết quả kiểm tra ngày {ngay_chon.strftime('%d/%m/%Y')}")
    st.dataframe(df_ket_qua, use_container_width=True)
    
    # Bộ lọc nhanh trạng thái vi phạm
    st.subheader("⚠️ Danh sách nhân viên đi trễ / về sớm / vắng")
    df_vi_pham = df_ket_qua[df_ket_qua['Trạng Thái'] != "Đúng giờ"]
    st.dataframe(df_vi_pham, use_container_width=True)

    # ==========================================
    # 3. KẾT HỢP GEMINI AI ĐỂ PHÂN TÍCH BÁO CÁO
    # ==========================================
    st.subheader("🤖 Gemini AI Phân Tích & Đánh Giá Giao Ca")
    if st.button("Yêu cầu Gemini AI phân tích dữ liệu ngày này"):
        with st.spinner("Gemini đang đọc bảng dữ liệu dữ liệu..."):
            # Chuyển data vi phạm thành text gửi cho AI
            du_lieu_text = df_vi_pham.to_string()
            
            prompt = f"""
            Bạn là một chuyên gia nhân sự chuyên nghiệp. Hãy đọc dữ liệu thống kê sai lệch giờ công (đi trễ, về sớm, vắng) của ngày {ngay_chon} dưới đây:
            {du_lieu_text}
            
            Nhiệm vụ:
            1. Tóm tắt nhanh số lượng đi trễ, về sớm và vắng theo từng bộ phận (Công nhân, Văn phòng, Nhóm đặc biệt).
            2. Chỉ ra những mã nhân viên vi phạm nghiêm trọng nhất (vừa trễ vừa về sớm, hoặc vắng không lý do).
            3. Đưa ra 3 đề xuất ngắn gọn bằng tiếng Việt để quản lý nhắc nhở hoặc cải thiện tình hình kỷ luật lao động dựa trên dữ liệu trên.
            """
            
            try:
                model = genai.GenerativeModel("gemini-1.5-flash")
                response = model.generate_content(prompt)
                st.markdown(response.text)
            except Exception as e:
                st.error("Không thể kết nối Gemini AI. Vui lòng kiểm tra lại API Key trong code.")
else:
    st.info("👋 Chào mừng bạn! Vui lòng tải đầy đủ các file Excel ở thanh bên trái (Sidebar) để Dashboard bắt đầu xử lý dữ liệu.")
