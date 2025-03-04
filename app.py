import streamlit as st
import pandas as pd
import openpyxl
from THONG_KE_VNINDEX_VN30 import process_excel  # Import function từ file function.py

st.title("📊 Xem & Xử Lý File Excel Trực Tiếp")

# Upload file Excel
uploaded_file = st.file_uploader("📂 Chọn file Excel (.xlsm, .xlsx)", type=["xlsm", "xlsx"])

if uploaded_file:
    st.success("✅ File đã tải lên!")
    
    # Đọc file và hiển thị dữ liệu
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    st.write("### 🔍 Dữ liệu trong file:")
    st.dataframe(df)

    # Chạy function.py khi nhấn nút
    if st.button("🚀 Xử lý dữ liệu"):
        output_file = process_excel(uploaded_file)
        with open(output_file, "rb") as f:
            st.download_button("📥 Tải file kết quả", f, file_name="output.xlsx")
