import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from cover_letter_processor import CoverLetterProcessor

st.title("Google Sheets Data Viewer của Thành Nhật")
# sidebar là thanh bên trái
sheet_id = st.sidebar.text_input("Enter Google Sheet ID")
sheet_name = st.sidebar.text_input("Enter Sheet Name")

if sheet_id != "" and sheet_name != "":
    try:
        # load credentials từ json key
        creds = Credentials.from_service_account_file('google_service.json', scopes =['https://www.googleapis.com/auth/spreadsheets'])
        # kết nối tới google sheet
        client = gspread.authorize(creds)
        # mở file google sheet và lấy dữ liệu từ sheet
        sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        all_values = sheet.get_all_values()
        # lấy ra những hàng đầu tiên => tên cột
        headers = all_values[0]
        # lấy ra những hàng còn lại => dữ liệu
        data = all_values[1:] if len(all_values) > 1 else []
        df = pd.DataFrame(data, columns=headers)
        st.success("Data loaded successfully!")
    except FileNotFoundError:
        st.error("Credential file not found.")
    except gspread.SpreadsheetNotFound:
        st.error('Spreadsheet not found. Please check the Sheet ID.')
    except gspread.exceptions.APIError as e:
        st.error(f"API Error: {e}")
    
    # Form nhập liệu thủ công
    with st.form(key='user_form'):
        ho_ten = st.text_input("Họ và tên (Viết in hoa)")
        gioi_tinh = st.selectbox("Giới tính", ["Nam", "Nữ"])
        ngay_sinh = st.text_input("Ngày sinh (ví dụ: 07 tháng 05 năm 2025)")
        noi_sinh = st.text_input("Nơi sinh")
        nguyen_quan = st.text_input("Nguyên quán")
        ho_khau = st.text_input("Hộ khẩu")
        cho_o_hien_nay = st.text_input("Chỗ ở hiện nay")
        dien_thoai = st.text_input("Điện thoại")
        dan_toc = st.text_input("Dân tộc", value="Kinh")
        cccd_cmnd = st.text_input("CCCD/CMND")
        ngay_cap = st.text_input("Ngày cấp (Ví dụ: 07/05/2025)")
        noi_cap = st.text_input("Nơi cấp", value="Cục Cảnh sát")
        trinh_do_van_hoa = st.selectbox("Trình độ văn hóa", ["Học sinh tiểu học", "Trung học cơ sở", "Trung học phổ thông", "Đại học", "Sinh viên"])
        so_truong = st.text_input("Sở trường")

        # Nút submit cho form thủ công
        submit_button = st.form_submit_button(label='Gửi thông tin')

        if submit_button:
            user_data = {
                "Họ và tên": ho_ten,
                "Giới tính": gioi_tinh,
                "Ngày sinh": ngay_sinh,
                "Nơi sinh": noi_sinh,
                "Nguyên quán": nguyen_quan,
                "Hộ khẩu": ho_khau,
                "Chỗ ở hiện nay": cho_o_hien_nay,
                "Điện thoại": dien_thoai,
                "Dân tộc": dan_toc,
                "CCCD/CMND": cccd_cmnd,
                "Ngày cấp": ngay_cap,
                "Nơi cấp": noi_cap,
                "Trình độ văn hóa": trinh_do_van_hoa,
                "Sở trường": so_truong
            }
            new_row = pd.DataFrame([user_data])
            df = pd.concat([df, new_row], ignore_index=True)
            df['Điện thoại'] = df['Điện thoại'].astype(str)
            df['CCCD/CMND'] = df['CCCD/CMND'].astype(str)
            sheet.clear()
            sheet.update([df.columns.values.tolist()] + df.values.tolist())
            st.success("Data submitted successfully!")
    #upload file .docx
    uploaded_files = st.file_uploader('Tải lên file .docx', type=['docx'], accept_multiple_files=True)
    processor = CoverLetterProcessor()
    
    if uploaded_files:
        if st.button("Trích xuất và lưu thông tin"):
            extracted_data = processor.process_documents(uploaded_files)
            if extracted_data:
                df_new = pd.DataFrame(extracted_data, columns=processor.headers)
                df = pd.concat([df, df_new], ignore_index=True)
                df['Điện thoại'] = df['Điện thoại'].astype(str)
                df['CCCD/CMND'] = df['CCCD/CMND'].astype(str)
                sheet.clear()
                sheet.update([df.columns.values.tolist()] + df.values.tolist())
                st.success("Data extracted and saved successfully!")
                uploaded_files.clear
            else:
                st.warning("No valid data extracted from the uploaded files.")