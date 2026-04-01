import streamlit as st
import pandas as pd
import io
import time
import os
import glob
import gc
import base64
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Lệnh set_page_config phải luôn nằm trên cùng
st.set_page_config(page_title="Phần mềm lập PL01 Chuyên nghiệp", layout="wide")

# ==========================================
# TÙY BIẾN GIAO DIỆN WEB (CHỈ CSS, KHÔNG SỬA LOGIC)
# ==========================================
st.markdown("""
<style>
    /* Import font hiện đại */
    @import url('https://fonts.googleapis.com/css2?family=Inter:opsz,wght@14..32,300;14..32,400;14..32,500;14..32,600;14..32,700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Nền tổng thể gradient nhẹ */
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #eef2f5 100%);
    }

    /* Sidebar làm đẹp */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        border-right: 1px solid rgba(0,0,0,0.05);
        box-shadow: 4px 0 12px rgba(0,0,0,0.02);
    }
    section[data-testid="stSidebar"] .css-1d391kg {
        background: transparent;
    }

    /* Tiêu đề sidebar */
    .sidebar-title {
        font-size: 1.1rem;
        font-weight: 600;
        color: #1e3a8a;
        text-align: center;
        margin-bottom: 0.5rem;
    }

    /* Các expander trong sidebar */
    .streamlit-expanderHeader {
        font-weight: 500;
        color: #2c3e66;
        background-color: rgba(30, 58, 138, 0.05);
        border-radius: 8px;
        transition: 0.2s;
    }
    .streamlit-expanderHeader:hover {
        background-color: rgba(30, 58, 138, 0.1);
    }

    /* Nút bấm chính */
    .stButton > button {
        border-radius: 30px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        border: none !important;
        background: linear-gradient(95deg, #1e3a8a 0%, #2563eb 100%) !important;
        color: white !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 14px rgba(37,99,235,0.25);
        background: linear-gradient(95deg, #1e40af 0%, #3b82f6 100%) !important;
    }

    /* Nút download */
    .stDownloadButton > button {
        background: linear-gradient(95deg, #10b981 0%, #34d399 100%) !important;
        border-radius: 30px !important;
        font-weight: 600 !important;
    }
    .stDownloadButton > button:hover {
        background: linear-gradient(95deg, #059669 0%, #10b981 100%) !important;
    }

    /* Các khung thông báo */
    .stAlert {
        border-radius: 12px;
        border-left: 5px solid;
    }
    .stAlert[data-baseweb="notification"] {
        background-color: #fef9e3;
        border-left-color: #f59e0b;
    }
    .stSuccess {
        background-color: #e6f9ed;
        border-left-color: #10b981;
    }
    .stError {
        background-color: #fee2e2;
        border-left-color: #ef4444;
    }

    /* Data Editor (bảng chỉnh sửa) */
    [data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
        border: 1px solid #e2e8f0;
        background: white;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    [data-testid="stDataFrame"] table {
        font-size: 13px;
    }
    [data-testid="stDataFrame"] th {
        background-color: #f1f5f9;
        font-weight: 600;
        color: #1e293b;
    }

    /* Text input, number input */
    .stTextInput > div > div > input, .stNumberInput > div > div > input {
        border-radius: 12px;
        border: 1px solid #cbd5e1;
        transition: 0.2s;
    }
    .stTextInput > div > div > input:focus, .stNumberInput > div > div > input:focus {
        border-color: #3b82f6;
        box-shadow: 0 0 0 2px rgba(59,130,246,0.2);
    }

    /* Expander tiêu đề chính */
    .streamlit-expanderHeader {
        font-size: 1rem;
        font-weight: 500;
        background-color: rgba(30, 58, 138, 0.03);
    }

    /* Ẩn chữ "Press Enter..." - giữ nguyên CSS cũ */
    [data-testid="InputInstructions"] {
        display: none !important;
    }

    /* Header chính */
    h1, h2, h3 {
        font-weight: 600;
        letter-spacing: -0.01em;
    }
    h1 {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.2rem;
    }
    .st-emotion-cache-1v0mbdj {
        margin-top: -0.5rem;
    }

    /* Card cho các phần */
    .card {
        background: white;
        border-radius: 20px;
        padding: 1rem;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
        border: 1px solid #eef2ff;
    }

    /* Thanh progress */
    .stProgress > div > div {
        background-color: #3b82f6 !important;
        border-radius: 20px;
    }

    /* Dropdown và selectbox */
    .stSelectbox > div > div {
        border-radius: 12px;
    }

    /* Căn chỉnh cột trong sidebar */
    .css-1kyxreq {
        gap: 0.5rem;
    }

    /* Responsive: giảm padding trên mobile */
    @media (max-width: 768px) {
        .stApp {
            padding: 1rem;
        }
        section[data-testid="stSidebar"] {
            padding: 0.5rem;
        }
    }

    /* Nút toggle expander mượt */
    .streamlit-expanderHeader svg {
        transition: transform 0.2s;
    }
    .streamlit-expanderHeader:hover svg {
        transform: scale(1.1);
    }

    /* Đẹp cho bảng trong expander cảnh báo */
    .stExpander .stDataFrame {
        border: none;
        box-shadow: none;
    }

    /* Footer nhẹ */
    .footer {
        text-align: center;
        font-size: 0.75rem;
        color: #6c757d;
        margin-top: 2rem;
        border-top: 1px solid #e9ecef;
        padding-top: 1rem;
    }

    /* ===== HIỆU ỨNG CHO GÓC NHỎ TÁC GIẢ ===== */
    .donate-box {
        position: relative;
        overflow: hidden;
        transition: all 0.3s ease;
        animation: borderGlow 2s infinite;
    }

    @keyframes borderGlow {
        0% { box-shadow: 0 0 0 0 rgba(0, 86, 179, 0.2); }
        50% { box-shadow: 0 0 0 8px rgba(0, 86, 179, 0.1); }
        100% { box-shadow: 0 0 0 0 rgba(0, 86, 179, 0.2); }
    }

    /* Hiệu ứng chữ chạy (marquee) cho dòng cảm ơn */
    .marquee-text {
        display: inline-block;
        white-space: nowrap;
        animation: marquee 15s linear infinite;
        font-weight: 500;
    }

    @keyframes marquee {
        0% { transform: translateX(100%); }
        100% { transform: translateX(-100%); }
    }

    /* Hiệu ứng nhấp nháy cho số tài khoản */
    .blink-number {
        animation: blink 1s step-end infinite;
    }

    @keyframes blink {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.5; }
    }

    /* ===== HIỆU ỨNG CHO LOGO VÀ TÊN CÔNG TY ===== */
    /* Hiệu ứng chữ chạy cho tên công ty */
    .company-name {
        text-align: center;
        font-size: 1.1em;
        font-weight: 600;
        color: #1E3A8A;
        white-space: nowrap;
        overflow: hidden;
        animation: slideIn 1s ease-out, glow 2s ease-in-out infinite;
    }

    @keyframes slideIn {
        from { transform: translateX(-30px); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
    }

    @keyframes glow {
        0% { text-shadow: 0 0 0px #1E3A8A; }
        50% { text-shadow: 0 0 8px #3b82f6; }
        100% { text-shadow: 0 0 0px #1E3A8A; }
    }

    /* Hiệu ứng nhấp nháy cho dòng trạm */
    .station-name {
        text-align: center;
        color: gray;
        font-size: 0.9em;
        margin-top: 5px;
        animation: blinkSoft 1.5s step-start infinite;
    }

    @keyframes blinkSoft {
        0%, 100% { opacity: 1; color: #6c757d; }
        50% { opacity: 0.6; color: #1e3a8a; }
    }

    /* Hiệu ứng xoay nhẹ cho logo */
    .logo-container {
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 10px;
        transition: transform 0.3s ease;
    }
    .logo-container:hover {
        transform: scale(1.02);
    }
    .logo-img {
        width: 140px;
        height: auto;
        animation: gentlePulse 2s infinite;
    }

    @keyframes gentlePulse {
        0% { filter: drop-shadow(0 0 0px rgba(30,58,138,0.2)); transform: scale(1); }
        50% { filter: drop-shadow(0 0 6px rgba(30,58,138,0.4)); transform: scale(1.02); }
        100% { filter: drop-shadow(0 0 0px rgba(30,58,138,0.2)); transform: scale(1); }
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# LỚP BẢO MẬT: KHÓA MẬT KHẨU (MÃ PIN)
# ==========================================
def check_password():
    """Trả về True nếu người dùng nhập đúng mật khẩu."""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    # Giao diện nhập mật khẩu
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("🔐 ĐĂNG NHẬP HỆ THỐNG")
        st.info("Phần mềm thuộc bản quyền tác giả. Vui lòng đăng nhập để tiếp tục.")
        
        # SỬ DỤNG FORM ĐỂ CHO PHÉP ẤN ENTER
        with st.form("login_form"):
            password = st.text_input("Nhập Mật Khẩu :", type="password")
            
            # Trả lại nút Đăng nhập to, đẹp, tràn viền
            submitted = st.form_submit_button("🚀 Đăng nhập", type="primary", use_container_width=True)
            
            if submitted:
                if password == "429751": 
                    st.session_state["password_correct"] = True
                    st.rerun()
                else:
                    st.error("""
                    ❌ **Mật khẩu không chính xác.**
                    
                    📞 Vui lòng liên hệ tác giả - **Trần Thọ**: 098.7575.691
                    """)
    return False

if not check_password():
    st.stop()

# ==========================================
# 0. KHU VỰC CHÈN ẢNH CHỦ QUYỀN (SIDEBAR) - CÓ HIỆU ỨNG
# ==========================================
script_dir = os.path.dirname(os.path.abspath(__file__))
image_files = glob.glob(os.path.join(script_dir, "anh_cua_toi*"))

st.sidebar.markdown("""
<div class="company-name">
    Công ty TNHH MTV Khai Thác<br>Công Trình Thủy Lợi Kon Tum
</div>
""", unsafe_allow_html=True)

if image_files:
    with open(image_files[0], "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode()
    
    st.sidebar.markdown(
        f"""
        <div class="logo-container">
            <img src="data:image/png;base64,{encoded_string}" class="logo-img">
        </div>
        """,
        unsafe_allow_html=True
    )
    st.sidebar.markdown('<div class="station-name">✨ TRẠM QLTN KHU VỰC 1</div>', unsafe_allow_html=True)
else:
    st.sidebar.info("💡 Mẹo: Hãy copy 1 tấm ảnh, đổi tên thành `anh_cua_toi` và ném chung vào thư mục code nhé!")
st.sidebar.markdown("---")

# ==========================================
# CHÈN NÚT DONATE CÀ PHÊ (CÓ HIỆU ỨNG)
# ==========================================
st.sidebar.markdown("### ☕&nbsp;&nbsp;Góc nhỏ của Tác giả", unsafe_allow_html=True)

st.sidebar.markdown("""
<div class="donate-box" style="text-align: justify; background-color: #e6f3fd; padding: 15px; border-radius: 5px; color: #0056b3; font-size: 14.5px;">
<span class="marquee-text">Một ly cà phê từ bạn là sự ghi nhận tuyệt vời nhất cho những nỗ lực tự động hóa công việc này. Xin chân thành cảm ơn! ❤️</span>
<br><br>
🏦 <b>Ngân hàng:</b> Vietcom Bank<br>
💳 <b>STK:</b> <span class="blink-number" style="font-size: 15px; color: #d9534f;"><b>0761002363642</b></span><br>
👤 <b>Chủ TK:</b> Trần Văn Thọ
</div>
""", unsafe_allow_html=True)
st.sidebar.markdown("---")

# ==========================================
# 1. CẤU TRÚC 30 CỘT CHUẨN Y HỆT PL01
# ==========================================
COLS = [str(i) for i in range(1, 31)]

def to_float(val):
    try:
        if pd.isna(val) or str(val).strip() == "" or str(val) == "<NA>": return 0.0
        return float(val)
    except: return 0.0

def clean_zero(val):
    return val if val > 0 else ""

def clean_text(val):
    if pd.isna(val) or str(val) == "<NA>": return ""
    s = str(val).strip()
    if s.lower() in ['nan', 'none', 'nat', 'null', '<na>', '']: return ""
    if s.endswith('.0'): s = s[:-2]
    return s

# ==========================================
# 2. XUẤT EXCEL BẢNG PL01 
# ==========================================
def export_pl01_excel(df_raw, cfg):
    df_raw = df_raw.fillna("")
    wb = Workbook()
    ws = wb.active
    ws.title = "PL01"

    font_title = Font(name='Times New Roman', size=12, bold=True)
    font_bold = Font(name='Times New Roman', size=11, bold=True)
    font_italic = Font(name='Times New Roman', size=11, italic=True)
    font_normal = Font(name='Times New Roman', size=11)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.append(["PHỤ LỤC 01. BẢNG KÊ ĐỐI TƯỢNG VÀ DIỆN TÍCH, BIỆN PHÁP TƯỚI, TIÊU ĐƯỢC HỖ TRỢ TIỀN SỬ DỤNG SẢN PHẨM, DỊCH VỤ CÔNG ÍCH THỦY LỢI GIAI ĐOẠN 2026-2030"] + [""]*29)
    ws.append(["TỔ CHỨC THỦY LỢI CƠ SỞ (HTXDVNN,......): (chỉ áp dụng đối với HTX Đoàn kết và các Công ty cà phê), XÃ/PHƯỜNG: ............................................................."] + [""]*29)
    
    ws.merge_cells('A1:AD1'); ws.merge_cells('A2:AD2')
    ws['A1'].font = font_title; ws['A1'].alignment = align_center
    ws['A2'].font = Font(name='Times New Roman', size=11, bold=True, italic=True); ws['A2'].alignment = align_center

    ws.append([
        "TT", "Hộ gia đình, cá nhân", "Bản đồ địa chính", "", "Diện tích thửa (m2)", 
        "Tên công trình cấp nước\n(Tuyến kênh, HCN, đập dâng,...)", "TỔNG DIỆN TÍCH (M2)", 
        "Tổng diện tích lúa", "DIỆN TÍCH TRỒNG LÚA (M2)"] + [""]*5 + 
        ["Tổng DT rau, màu, cây CNDN", "DIỆN TÍCH TRỒNG CÂY CÔNG NGHIỆP DÀI NGÀY(M2)"] + [""]*5 + 
        ["Tổng DT rau, màu, cây CNNN", "DIỆN TÍCH TRỒNG RAU, MÀU, CÂY CÔNG NGHIỆP NGẮN NGÀY (M2)"] + [""]*5 + 
        ["DIỆN TÍCH NUÔI TRỒNG THUỶ SẢN (M2)", "Ký xác nhận của đại diện hộ gia đình, cá nhân"]
    )
    ws.append([
        "", "", "Số tờ bản đồ", "Số thửa", "", "", "", "", 
        "Tưới tiêu bằng trọng lực", "", "", "Tưới tiêu bằng động lực", "", "", 
        "", "Tưới tiêu bằng trọng lực", "", "", "Tưới bằng động lực", "", "", 
        "", "Tưới tiêu bằng trọng lực", "", "", "Tưới bằng động lực", "", "", 
        "", ""
    ])
    ws.append([
        "", "", "", "", "", "", "", "", 
        "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", 
        "", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", 
        "", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", 
        "", ""
    ])
    ws.append([
        "1", "2", "3", "4", "5", "6", "7=\n(8+15+22+29)", "8=\n(9+..+14)", 
        "9", "10", "11", "12", "13", "14", 
        "15=\n(16+..+21)", "16", "17", "18", "19", "20", "21", 
        "22=\n(23+..+28)", "23", "24", "25", "26", "27", "28", 
        "29", "30"
    ])

    merges = [
        'A3:A5', 'B3:B5', 'C3:D3', 'C4:C5', 'D4:D5', 'E3:E5', 'F3:F5', 'G3:G5', 'H3:H5',
        'I3:N3', 'I4:K4', 'L4:N4', 
        'O3:O5', 'P3:U3', 'P4:R4', 'S4:U4', 
        'V3:V5', 'W3:AB3', 'W4:Y4', 'Z4:AB4', 
        'AC3:AC5', 'AD3:AD5'
    ]
    for m in merges: ws.merge_cells(m)

    for r in range(3, 7):
        for c in range(1, 31):
            cell = ws.cell(row=r, column=c)
            cell.alignment = align_center
            cell.border = thin_border
            if r == 6: cell.font = font_italic 
            else: cell.font = font_bold

    vertical_cols = [8, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19, 20, 22, 23, 24, 25, 26, 27, 28]

    data_rows = []
    current_excel_row = 10
    tt = 1
    
    for ho, group in df_raw.groupby("2", sort=False, dropna=False):
        if pd.isna(ho) or str(ho).strip() == "" or str(ho) == "1": continue
        
        has_dx = any(v >= 1 for v in cfg.values())
        has_mua = any(v == 2 for v in cfg.values())
        seasons = []
        if has_dx: seasons.append(("- Vụ Đông Xuân", 1))
        if has_mua: seasons.append(("- Vụ Mùa", 2))
        
        ho_parcels_exist = False
        ho_season_indices = []
        temp_ho_items = [] 
        ho_row_idx = current_excel_row
        
        for season_name, season_level in seasons:
            season_parcels = []
            for _, row in group.iterrows():
                l9 = to_float(row.get("9")) if cfg.get("9", 0) >= season_level else 0.0
                l10 = to_float(row.get("10")) if cfg.get("10", 0) >= season_level else 0.0
                l11 = to_float(row.get("11")) if cfg.get("11", 0) >= season_level else 0.0
                l12 = to_float(row.get("12")) if cfg.get("12", 0) >= season_level else 0.0
                l13 = to_float(row.get("13")) if cfg.get("13", 0) >= season_level else 0.0
                l14 = to_float(row.get("14")) if cfg.get("14", 0) >= season_level else 0.0
                
                c16 = to_float(row.get("16")) if cfg.get("16", 0) >= season_level else 0.0
                c17 = to_float(row.get("17")) if cfg.get("17", 0) >= season_level else 0.0
                c18 = to_float(row.get("18")) if cfg.get("18", 0) >= season_level else 0.0
                c19 = to_float(row.get("19")) if cfg.get("19", 0) >= season_level else 0.0
                c20 = to_float(row.get("20")) if cfg.get("20", 0) >= season_level else 0.0
                c21 = to_float(row.get("21")) if cfg.get("21", 0) >= season_level else 0.0
                
                m23 = to_float(row.get("23")) if cfg.get("23", 0) >= season_level else 0.0
                m24 = to_float(row.get("24")) if cfg.get("24", 0) >= season_level else 0.0
                m25 = to_float(row.get("25")) if cfg.get("25", 0) >= season_level else 0.0
                m26 = to_float(row.get("26")) if cfg.get("26", 0) >= season_level else 0.0
                m27 = to_float(row.get("27")) if cfg.get("27", 0) >= season_level else 0.0
                m28 = to_float(row.get("28")) if cfg.get("28", 0) >= season_level else 0.0
                
                ca29 = to_float(row.get("29")) if cfg.get("29", 0) >= season_level else 0.0
                
                sum_total = sum([l9,l10,l11,l12,l13,l14, c16,c17,c18,c19,c20,c21, m23,m24,m25,m26,m27,m28, ca29])
                if sum_total == 0: continue

                r_data = [""] * 30
                r_data[2] = clean_text(row.get("3"))
                r_data[3] = clean_text(row.get("4"))
                r_data[4] = to_float(row.get("5"))
                r_data[5] = clean_text(row.get("6"))
                
                r_data[8] = l9; r_data[9] = l10; r_data[10] = l11
                r_data[11] = l12; r_data[12] = l13; r_data[13] = l14
                
                r_data[15] = c16; r_data[16] = c17; r_data[17] = c18
                r_data[18] = c19; r_data[19] = c20; r_data[20] = c21
                
                r_data[22] = m23; r_data[23] = m24; r_data[24] = m25
                r_data[25] = m26; r_data[26] = m27; r_data[27] = m28
                
                r_data[28] = ca29   
                r_data[29] = clean_text(row.get("30"))
                
                season_parcels.append(r_data)
            
            if len(season_parcels) > 0:
                ho_parcels_exist = True
                season_row_idx = ho_row_idx + 1 + sum([len(s['parcels']) + 1 for s in temp_ho_items])
                parcel_start = season_row_idx + 1
                parcel_end = season_row_idx + len(season_parcels)
                
                season_row_data = [""] * 30
                season_row_data[1] = season_name
                for i in vertical_cols:
                    col_letter = get_column_letter(i + 1)
                    if parcel_start == parcel_end: season_row_data[i] = f"={col_letter}{parcel_start}"
                    else: season_row_data[i] = f"=SUM({col_letter}{parcel_start}:{col_letter}{parcel_end})"
                
                temp_ho_items.append({
                    "season_row_data": season_row_data,
                    "season_row_idx": season_row_idx,
                    "parcels": season_parcels
                })
                ho_season_indices.append(season_row_idx)
                
        if ho_parcels_exist:
            ho_row_data = [""] * 30
            ho_row_data[0] = tt
            ho_row_data[1] = ho
            for i in vertical_cols:
                col_letter = get_column_letter(i + 1)
                if len(ho_season_indices) == 1: ho_row_data[i] = f"={col_letter}{ho_season_indices[0]}"
                else: ho_row_data[i] = f"=SUM({','.join([f'{col_letter}{idx}' for idx in ho_season_indices])})"
            
            data_rows.append({"type": "ho", "data": ho_row_data})
            current_excel_row += 1
            for s_item in temp_ho_items:
                data_rows.append({"type": "season", "data": s_item["season_row_data"]})
                current_excel_row += 1
                for p_data in s_item["parcels"]:
                    data_rows.append({"type": "parcel", "data": p_data})
                    current_excel_row += 1
            tt += 1

    max_row = current_excel_row - 1
    
    row_tong = ["1", "Tổng cộng", "", "", "", "THOKEEN PRO"] + [""] * 24
    row_dx = ["a", "Vụ Đông Xuân", "", "", "", ""] + [""] * 24
    row_mua = ["b", "Vụ Mùa", "", "", "", ""] + [""] * 24

    row_tong[6] = "=H7+O7+V7+AC7"; row_tong[7] = "=SUM(I7:N7)"; row_tong[14] = "=SUM(P7:U7)"; row_tong[21] = "=SUM(W7:AB7)"
    row_dx[6] = "=H8+O8+V8+AC8"; row_dx[7] = "=SUM(I8:N8)"; row_dx[14] = "=SUM(P8:U8)"; row_dx[21] = "=SUM(W8:AB8)"
    row_mua[6] = "=H9+O9+V9+AC9"; row_mua[7] = "=SUM(I9:N9)"; row_mua[14] = "=SUM(P9:U9)"; row_mua[21] = "=SUM(W9:AB9)"

    if max_row >= 10:
        for i in vertical_cols:
            col_letter = get_column_letter(i + 1)
            row_dx[i] = f'=SUMIF($B$10:$B${max_row}, "- Vụ Đông Xuân", {col_letter}$10:{col_letter}${max_row})'
            row_mua[i] = f'=SUMIF($B$10:$B${max_row}, "- Vụ Mùa", {col_letter}$10:{col_letter}${max_row})'
            row_tong[i] = f"={col_letter}8+{col_letter}9"

    ws.append(row_tong)
    ws.append(row_dx)
    ws.append(row_mua)

    for c_idx in range(1, 31):
        for r_idx in range(7, 10): 
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = Font(name='Times New Roman', size=11, bold=True, color="FF0000")
            cell.border = thin_border
            cell.alignment = align_center if c_idx != 2 else align_left
            if c_idx >= 5 and c_idx <= 29: 
                cell.number_format = '#,##0.00;-#,##0.00;""'

    start_row = 10
    for item in data_rows:
        r_data = item["data"]
        r_data[6] = f"=H{start_row}+O{start_row}+V{start_row}+AC{start_row}"
        r_data[7] = f"=SUM(I{start_row}:N{start_row})"
        r_data[14] = f"=SUM(P{start_row}:U{start_row})"
        r_data[21] = f"=SUM(W{start_row}:AB{start_row})"
        
        ws.append([clean_zero(v) if isinstance(v, float) else v for v in r_data])
        
        for col_idx, cell in enumerate(ws[start_row], start=1):
            cell.border = thin_border
            cell.font = font_normal
            cell.alignment = align_center
            if col_idx == 2: cell.alignment = align_left
            if item["type"] == "ho" and col_idx in [1, 2]: cell.font = font_bold
            
            if col_idx >= 5 and col_idx <= 29:
                cell.number_format = '#,##0.00;-#,##0.00;""'
                
        start_row += 1

    ws.column_dimensions['A'].width = 5; ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 8; ws.column_dimensions['D'].width = 8   
    for i in range(5, 31): ws.column_dimensions[get_column_letter(i)].width = 10

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ==========================================
# 3. THUẬT TOÁN ĐỊNH DẠNG DATA GỐC CHUẨN 30 CỘT
# ==========================================
def export_formatted_data_goc(df):
    df_clean = df.fillna("")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data_Goc"

    font_bold = Font(name='Times New Roman', size=11, bold=True)
    font_normal = Font(name='Times New Roman', size=11)
    font_italic = Font(name='Times New Roman', size=11, italic=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    ws.append(["PHỤ LỤC 01. BẢNG KÊ ĐỐI TƯỢNG VÀ DIỆN TÍCH... (DATA NỘI BỘ ĐỂ NHẬP LIỆU)"] + [""]*29)
    ws.append(["TỔ CHỨC THỦY LỢI CƠ SỞ..."] + [""]*29)
    ws.append([
        "TT", "Hộ gia đình, cá nhân", "Bản đồ địa chính", "", "Diện tích thửa (m2)", "Tên công trình cấp nước", "TỔNG DIỆN TÍCH (M2)", "Tổng diện tích lúa", "DIỆN TÍCH TRỒNG LÚA (M2)"] + [""]*5 + ["Tổng DT rau, màu, cây CNDN", "DIỆN TÍCH TRỒNG CÂY CÔNG NGHIỆP DÀI NGÀY(M2)"] + [""]*5 + ["Tổng DT rau, màu, cây CNNN", "DIỆN TÍCH TRỒNG RAU, MÀU, CÂY CÔNG NGHIỆP NGẮN NGÀY (M2)"] + [""]*5 + ["DIỆN TÍCH NUÔI TRỒNG THUỶ SẢN (M2)", "Ký xác nhận"]
    )
    ws.append(["", "", "Số tờ bản đồ", "Số thửa", "", "", "", "", "Tưới tiêu bằng trọng lực", "", "", "Tưới tiêu bằng động lực", "", "", "", "Tưới tiêu bằng trọng lực", "", "", "Tưới bằng động lực", "", "", "", "Tưới tiêu bằng trọng lực", "", "", "Tưới bằng động lực", "", "", "", ""])
    ws.append(["", "", "", "", "", "", "", "", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "Chủ động", "CĐ 1 phần", "Tạo nguồn", "", ""])
    ws.append(COLS)

    merges = ['A3:A5', 'B3:B5', 'C3:D3', 'C4:C5', 'D4:D5', 'E3:E5', 'F3:F5', 'G3:G5', 'H3:H5', 'I3:N3', 'I4:K4', 'L4:N4', 'O3:O5', 'P3:U3', 'P4:R4', 'S4:U4', 'V3:V5', 'W3:AB3', 'W4:Y4', 'Z4:AB4', 'AC3:AC5', 'AD3:AD5']
    for m in merges: ws.merge_cells(m)

    for r in range(3, 7):
        for c in range(1, 31):
            cell = ws.cell(row=r, column=c)
            cell.alignment = align_center
            cell.border = thin_border
            cell.font = font_bold if r < 6 else font_italic
            if r == 6: cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    for r_idx, row in enumerate(df_clean.values, start=7):
        ws.append(list(row))
        for col_idx, cell in enumerate(ws[r_idx], start=1):
            cell.font = font_normal
            cell.border = thin_border
            cell.alignment = align_left if col_idx == 2 else align_center
            if col_idx >= 5 and col_idx <= 29:
                cell.number_format = '#,##0.00;-#,##0.00;""'

    ws.column_dimensions['B'].width = 25
    for i in range(5, 31): ws.column_dimensions[get_column_letter(i)].width = 11

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ==========================================
# GIAO DIỆN APP PHẦN 1: TẠO BÁO CÁO PL01
# ==========================================
st.header("1. 🚀 Xây dựng và Xuất báo cáo PL01")
uploaded_file = st.file_uploader("📥 Tải lên file Data Excel hoặc file Báo cáo PL01", type=["xlsx", "xls"])

if uploaded_file is not None:
    if st.session_state.get('last_file_id') != uploaded_file.file_id:
        st.session_state['last_file_id'] = uploaded_file.file_id
        
        for key in ['raw_data', 'pl01_data', 'goc_data', 'cfg_hash']:
            if key in st.session_state:
                del st.session_state[key]
        gc.collect()
        
        progress_text = "⏳ Đang phân tích và đồng bộ hóa cấu trúc file..."
        my_bar = st.progress(0, text=progress_text)
        for percent in range(100):
            time.sleep(0.01)
            my_bar.progress(percent + 1, text=f"{progress_text} {percent + 1}%")
        my_bar.empty()
        
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        header_idx = -1
        for i, row in df_raw.iterrows():
            vals = [str(x).strip().replace('.0', '') for x in row.values[:5]]
            if '1' in vals and '2' in vals and '3' in vals:
                header_idx = i
                break
                
        if header_idx != -1:
            data_part = df_raw.iloc[header_idx+1:].reset_index(drop=True)
            if data_part.shape[1] < 30:
                for i in range(data_part.shape[1], 30): data_part[i] = ""
            data_part = data_part.iloc[:, :30]
            data_part.columns = COLS
            
            extracted_rows = []
            current_ho = ""
            
            for _, row in data_part.iterrows():
                c2 = str(row['2']).strip()
                c3 = str(row['3']).strip()
                c4 = str(row['4']).strip()
                
                if c2 in ['Tổng cộng', 'Vụ Đông Xuân', 'Vụ Mùa'] or c2.startswith('Tổng cộng'):
                    continue
                    
                if c2 != "" and c3 in ["", "nan", "None", "<NA>"] and c4 in ["", "nan", "None", "<NA>"] and not c2.startswith("- Vụ"):
                    current_ho = c2
                    continue
                    
                if (c3 not in ["", "nan", "None", "<NA>"]) or (c4 not in ["", "nan", "None", "<NA>"]):
                    actual_ho = c2 if (c2 not in ["", "nan", "None", "<NA>"] and not c2.startswith("- Vụ")) else current_ho
                    
                    new_row = row.copy()
                    new_row['2'] = actual_ho
                    new_row['7'] = 0; new_row['8'] = 0; new_row['15'] = 0; new_row['22'] = 0
                    extracted_rows.append(new_row)
                    
            df_final = pd.DataFrame(extracted_rows, columns=COLS)
            df_final = df_final.dropna(subset=['2'])
            df_final = df_final[df_final['2'].astype(str).str.strip() != ""]
            for col in ['3', '4']: df_final[col] = df_final[col].apply(clean_text)
            
            st.session_state.raw_data = df_final
            st.session_state.pop('pl01_data', None)
            st.session_state.pop('goc_data', None)
            st.toast("✅ Đã nhận diện và bóc tách dữ liệu thành công!", icon="🚀")
        else:
            st.error("❌ File không đúng định dạng (Không tìm thấy hàng số 1->30). Vui lòng kiểm tra lại file của bạn.")

if 'raw_data' in st.session_state:
    st.subheader("🔎 Radar Quét Lỗi Dữ Liệu")
    
    df_check = st.session_state.raw_data.copy()
    df_check['3'] = df_check['3'].astype(str).str.strip()
    df_check['4'] = df_check['4'].astype(str).str.strip()
    df_valid = df_check[(df_check['3'] != '') & (df_check['4'] != '') & (df_check['3'] != 'nan') & (df_check['4'] != 'nan') & (df_check['3'] != '<NA>') & (df_check['4'] != '<NA>')]
    
    dup_mask = df_valid.duplicated(subset=['3', '4'], keep=False)
    if dup_mask.any():
        dup_df = df_valid[dup_mask]
        dup_summary = dup_df.groupby(['3', '4']).size().reset_index(name='count')
        with st.expander("⚠️ PHÁT HIỆN TRÙNG LẶP THỬA ĐẤT (Bấm để xem chi tiết)", expanded=True):
            st.error("Hệ thống phát hiện các thửa đất sau đang bị nhập nhiều lần:")
            for _, r in dup_summary.iterrows():
                st.write(f"👉 **Tờ bản đồ {r['3']}, thửa {r['4']}**: xuất hiện **{r['count']}** lần.")
    else:
        st.success("✅ Dữ liệu sạch: Không phát hiện trùng lặp thửa đất!")

    invalid_area_rows = []
    crop_cols = ['9','10','11','12','13','14', '16','17','18','19','20','21', '23','24','25','26','27','28', '29']
    
    for idx, row in df_valid.iterrows():
        dt_thua = to_float(row['5'])
        dt_hotro = sum([to_float(row[c]) for c in crop_cols])
        
        if round(dt_hotro, 2) > round(dt_thua, 2):
            invalid_area_rows.append({
                'to_bd': row['3'],
                'thua': row['4'],
                'dt_thua': dt_thua,
                'dt_hotro': dt_hotro
            })

    if invalid_area_rows:
        with st.expander("🚨 CẢNH BÁO LỖI LOGIC: DIỆN TÍCH VƯỢT QUÁ HẠN MỨC", expanded=True):
            st.error("Các thửa đất sau có Tổng diện tích tưới tiêu lớn hơn Diện tích thực tế của thửa đất:")
            for r in invalid_area_rows:
                st.write(f"👉 **Tờ {r['to_bd']}, thửa {r['thua']}**: DT Thửa = **{r['dt_thua']:,.2f}** | Tổng tưới = **{r['dt_hotro']:,.2f}**")
    else:
        st.success("✅ Diện tích hợp lệ: Không phát hiện thửa đất nào vượt quá diện tích quy định!")

    # --- SIDEBAR ---
    st.sidebar.markdown("### ⚙️ CÀI ĐẶT MÙA VỤ CHI TIẾT")
    st.sidebar.caption("Chỉ định phân bổ diện tích tự động cho từng cột.")
    
    def ui_select(key, label=""):
        return st.selectbox(label, [0, 1, 2], format_func=lambda x: "❌ Bỏ qua" if x==0 else ("🌱 1 Vụ (ĐX)" if x==1 else "🌾 Cả 2 Vụ"), key=key, label_visibility="collapsed")

    cfg = {}
    
    st.sidebar.markdown("#### 🌾 1. DIỆN TÍCH TRỒNG LÚA")
    with st.sidebar.expander("💧 Tưới tiêu bằng trọng lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 9):"); cfg["9"] = ui_select("l9")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 10):"); cfg["10"] = ui_select("l10")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 11):"); cfg["11"] = ui_select("l11")
    with st.sidebar.expander("🚰 Tưới tiêu bằng động lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 12):"); cfg["12"] = ui_select("l12")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 13):"); cfg["13"] = ui_select("l13")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 14):"); cfg["14"] = ui_select("l14")

    st.sidebar.markdown("#### 🌳 2. CÂY CÔNG NGHIỆP DÀI NGÀY")
    with st.sidebar.expander("💧 Tưới tiêu bằng trọng lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 16):"); cfg["16"] = ui_select("c16")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 17):"); cfg["17"] = ui_select("c17")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 18):"); cfg["18"] = ui_select("c18")
    with st.sidebar.expander("🚰 Tưới tiêu bằng động lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 19):"); cfg["19"] = ui_select("c19")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 20):"); cfg["20"] = ui_select("c20")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 21):"); cfg["21"] = ui_select("c21")

    st.sidebar.markdown("#### 🥬 3. RAU, MÀU, CÂY CNNN")
    with st.sidebar.expander("💧 Tưới tiêu bằng trọng lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 23):"); cfg["23"] = ui_select("m23")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 24):"); cfg["24"] = ui_select("m24")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 25):"); cfg["25"] = ui_select("m25")
    with st.sidebar.expander("🚰 Tưới tiêu bằng động lực", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động (Cột 26):"); cfg["26"] = ui_select("m26")
        c1, c2 = st.columns([5, 4]); c1.write("Chủ động 1 phần (Cột 27):"); cfg["27"] = ui_select("m27")
        c1, c2 = st.columns([5, 4]); c1.write("Tạo nguồn (Cột 28):"); cfg["28"] = ui_select("m28")

    st.sidebar.markdown("#### 🐟 4. THỦY SẢN")
    with st.sidebar.expander("🌊 Diện tích ao cá", expanded=False):
        c1, c2 = st.columns([5, 4]); c1.write("Ao cá (Cột 29):"); cfg["29"] = ui_select("ca29")

    # --- BỘ LỌC THÔNG MINH CHO BẢNG DỮ LIỆU ---
    st.subheader("Bảng tính Data Nội bộ (Tìm kiếm & Chỉnh sửa)")
    
    col1, col2, col3 = st.columns(3)
    search_chu_ho = col1.text_input("🔍 Tìm theo Tên Chủ Hộ:")
    search_to = col2.text_input("🗺️ Tìm theo Số Tờ BĐ:")
    search_thua = col3.text_input("🟩 Tìm theo Số Thửa:")
    
    df_display = st.session_state.raw_data.copy()
    
    if search_chu_ho:
        df_display = df_display[df_display['2'].astype(str).str.contains(search_chu_ho, case=False, na=False)]
    if search_to:
        df_display = df_display[df_display['3'].astype(str) == search_to]
    if search_thua:
        df_display = df_display[df_display['4'].astype(str) == search_thua]
        
    is_filtered = bool(search_chu_ho or search_to or search_thua)
    
    if is_filtered:
        st.info(f"💡 Đang hiển thị {len(df_display)} kết quả lọc. (Lưu ý: Bạn có thể thêm/xóa hàng trực tiếp trên bảng lọc này, những thay đổi sẽ áp dụng cho dữ liệu gốc)")

    col_config = {col: st.column_config.NumberColumn(format="%.2f") for col in COLS[4:29]}
    col_config["2"] = st.column_config.TextColumn(width="large")
    
    # Cho phép dynamic để thêm/xóa hàng trong cả hai chế độ
    edited_df = st.data_editor(
        df_display, 
        column_config=col_config, 
        num_rows="dynamic",  # luôn dynamic để có thể xóa hàng khi lọc
        use_container_width=True, 
        height=500
    )
    # Nút xóa nhanh khi đang lọc
    if is_filtered:
        if st.button("🗑️ Xóa các dòng đang hiển thị (vĩnh viễn)", type="secondary"):
            with st.spinner("Đang xóa..."):
                # Lấy danh sách index các dòng đang hiển thị
                indices_to_drop = df_display.index.tolist()
                # Xóa khỏi raw_data
                st.session_state.raw_data.drop(indices_to_drop, inplace=True)
                st.success(f"✅ Đã xóa {len(indices_to_drop)} dòng.")
                # Reset lại bộ lọc để tránh hiển thị dòng đã xóa
                st.rerun()
    # ===== XỬ LÝ LƯU: HỖ TRỢ XÓA HÀNG KHI ĐANG LỌC =====
    if st.button("💾 Lưu thay đổi bảng tính"):
        with st.spinner("Đang cập nhật thay đổi..."):
            time.sleep(0.5)
            # Ép kiểu số
            for col in COLS[4:29]:
                edited_df[col] = pd.to_numeric(edited_df[col], errors='coerce').fillna(0)
            
            if is_filtered:
                # Tìm các hàng đã bị xóa trong kết quả lọc
                # df_display là DataFrame hiển thị (trước chỉnh sửa), edited_df là sau chỉnh sửa (có thể thiếu hàng)
                # Xóa các hàng trong raw_data mà index nằm trong df_display nhưng không nằm trong edited_df
                deleted_indexes = set(df_display.index) - set(edited_df.index)
                if deleted_indexes:
                    st.session_state.raw_data = st.session_state.raw_data.drop(index=deleted_indexes)
                
                # Cập nhật các hàng còn lại (có thể đã sửa)
                # Chỉ cập nhật những hàng có index chung giữa edited_df và raw_data
                common_indexes = edited_df.index.intersection(st.session_state.raw_data.index)
                for idx in common_indexes:
                    for col in edited_df.columns:
                        st.session_state.raw_data.loc[idx, col] = edited_df.loc[idx, col]
                
                # Nếu có hàng mới được thêm vào (index chưa có trong raw_data)
                new_indexes = edited_df.index.difference(st.session_state.raw_data.index)
                if not new_indexes.empty:
                    new_rows = edited_df.loc[new_indexes]
                    st.session_state.raw_data = pd.concat([st.session_state.raw_data, new_rows], ignore_index=False)
                    # Reset index để tránh trùng lặp
                    st.session_state.raw_data.reset_index(drop=True, inplace=True)
            else:
                # Không lọc: gán lại toàn bộ DataFrame (đã ép kiểu)
                st.session_state.raw_data = edited_df.copy()
            
        st.success("✅ Đã lưu! Cấu trúc Data gốc được bảo toàn an toàn tuyệt đối.")
        st.rerun()

    st.markdown("---")
    st.subheader("Xuất Biểu mẫu Báo cáo")
    
    if st.button("🚀 TỔNG HỢP VÀ TẠO BÁO CÁO (EXCEL)", type="primary"):
        progress_text = "⏳ Đang tính toán Ma trận và Gắn CÔNG THỨC EXCEL..."
        my_bar = st.progress(0, text=progress_text)
        for percent in range(1, 40):
            time.sleep(0.01)
            my_bar.progress(percent, text=f"{progress_text} {percent}%")
            
        my_bar.progress(40, text="⏳ HỆ THỐNG ĐANG CHẠY...")
        pl01_data = export_pl01_excel(st.session_state.raw_data, cfg)
        
        my_bar.progress(70, text="⏳ VUI LÒNG CHỜ TRONG GIÂY LÁT... ")
        goc_data = export_formatted_data_goc(st.session_state.raw_data)
        
        for percent in range(70, 101):
            time.sleep(0.01)
            my_bar.progress(percent, text=f"🎉 ĐÃ HOÀN THÀNH XONG! {percent}%")
        time.sleep(0.5)
        my_bar.empty()
        
        st.session_state['pl01_data'] = pl01_data
        st.session_state['goc_data'] = goc_data
        st.session_state['cfg_hash'] = str(cfg)

    if 'pl01_data' in st.session_state:
        if st.session_state.get('cfg_hash') == str(cfg):
            col1, col2 = st.columns(2)
            with col1:
                st.download_button(label="📥 TẢI XUỐNG PL01 CHUẨN", data=st.session_state['pl01_data'], file_name="BieuMau_PL01.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            with col2:
                st.download_button(label="🔄 Tải file Data Nội bộ ", data=st.session_state['goc_data'], file_name="Data_Goc.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("⚠️ Cấu hình mùa vụ đã thay đổi. Vui lòng bấm 'TỔNG HỢP VÀ TẠO BÁO CÁO' lại để cập nhật.")

# ==========================================
# GIAO DIỆN APP PHẦN 2: KIỂM TRA TRÙNG LẶP TỪ PL01
# ==========================================
st.markdown("---")
st.header("2. 🕵️ Kiểm tra Trùng lặp Tờ/Thửa từ PL01")
st.info("Hệ thống sẽ quét toàn bộ file PL01 để phát hiện các thửa đất bị nhập trùng (Mỗi thửa chỉ được xuất hiện tối đa 1 lần trong cùng 1 Vụ).")

check_file = st.file_uploader("📥 Tải lên file PL01 cần kiểm tra (.xlsx)", type=["xlsx"], key="check_file")

if check_file is not None:
    try:
        import openpyxl
        import pandas as pd
        
        wb = openpyxl.load_workbook(check_file, data_only=True)
        ws = wb.active
        df_check = pd.DataFrame(ws.values)
        
        start_row_idx = -1
        for i, row in df_check.iterrows():
            if str(row[1]).strip() == "Tổng cộng":
                start_row_idx = i + 3 
                break
                
        if start_row_idx == -1:
            st.error("❌ Không tìm thấy cấu trúc chuẩn. Vui lòng đảm bảo đây là file PL01.")
        else:
            data = df_check.iloc[start_row_idx:].copy()
            current_season = "Không xác định"
            parcels = []
            
            for idx, row in data.iterrows():
                excel_row_num = idx + 1 
                col2_name = str(row[1]).strip()
                col3_to = str(row[2]).strip()
                col4_thua = str(row[3]).strip()
                
                if col2_name.startswith("- Vụ"):
                    current_season = col2_name
                    
                elif (col3_to not in ["None", "nan", ""]) and (col4_thua not in ["None", "nan", ""]):
                    parcels.append({
                        'Vụ': current_season,
                        'Tờ': col3_to,
                        'Thửa': col4_thua,
                        'Dòng Excel': excel_row_num
                    })

            df_parcels = pd.DataFrame(parcels)
            
            if df_parcels.empty:
                st.warning("⚠️ Không tìm thấy dữ liệu thửa đất nào.")
            else:
                parcel_counts = df_parcels.groupby(['Vụ', 'Tờ', 'Thửa']).agg(
                    Số_lần_xuất_hiện=('Tờ', 'size'),
                    Vị_trí_dòng_Excel=('Dòng Excel', lambda x: ', '.join(map(str, x))) 
                ).reset_index()
                
                parcel_counts.rename(columns={'Số_lần_xuất_hiện': 'Số lần xuất hiện', 'Vị_trí_dòng_Excel': 'Nằm tại các hàng (Excel)'}, inplace=True)

                duplicates = parcel_counts[parcel_counts['Số lần xuất hiện'] > 1].copy()
                
                if not duplicates.empty:
                    st.error(f"🚨 PHÁT HIỆN THỬA ĐẤT BỊ TRÙNG LẶP TRONG CÙNG MỘT VỤ!")
                    duplicates = duplicates.reset_index(drop=True)
                    duplicates.insert(0, 'STT', range(1, len(duplicates) + 1))
                    
                    duplicates = duplicates[['STT', 'Vụ', 'Tờ', 'Thửa', 'Số lần xuất hiện', 'Nằm tại các hàng (Excel)']]
                    
                    st.dataframe(duplicates, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Tuyệt vời! Không phát hiện thửa đất nào bị trùng lặp trong các vụ.")
                    st.balloons()
                    
    except Exception as e:
        st.error(f"❌ Lỗi đọc file: {e}")
