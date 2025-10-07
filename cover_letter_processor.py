from docx import Document
import re

# Lớp xử lý file .docx
class CoverLetterProcessor:
    def __init__(self):
        self.patterns = {
            "Họ và tên": r"Họ và tên: (.*?)\s+Nam/Nữ",
            "Giới tính": r"Nam/Nữ: (.*?)$",
            "Ngày sinh": r"Sinh ngày: (.*?)\s+Nơi sinh",
            "Nơi sinh": r"Nơi sinh: (.*?)$",
            "Nguyên quán": r"Nguyên quán: (.*?)$",
            "Hộ khẩu": r"Nơi đăng ký hộ khẩu thường trú: (.*?)$",
            "Chỗ ở hiện nay": r"Chỗ ở hiện nay: (.*?)$",
            "Điện thoại": r"Điện thoại liên hệ: (.*?)$",
            "Dân tộc": r"Dân tộc: (.*?)\s+Tôn giáo",
            "CCCD/CMND": r"Số CCCD/CMND: (.*?)\s+Cấp ngày",
            "Ngày cấp": r"Cấp ngày: (.*?)\s+Nơi cấp",
            "Nơi cấp": r"Nơi cấp: (.*?)$",
            "Trình độ văn hóa": r"Trình độ văn hóa: (.*?)$",
            "Sở trường": r"Sở trường: (.*?)$"
        }
        self.headers = [
            "Họ và tên", "Giới tính", "Ngày sinh", "Nơi sinh", "Nguyên quán",
            "Hộ khẩu", "Chỗ ở hiện nay", "Điện thoại", "Dân tộc", "CCCD/CMND",
            "Ngày cấp", "Nơi cấp", "Trình độ văn hóa", "Sở trường"
        ]

    def read_docx(self, file_content):
        """Read text from a .docx file content."""
        try:
            doc = Document(file_content)
            return "\n".join(para.text for para in doc.paragraphs)
        except Exception as e:
            print(f"Error reading file: {e}")
            return None

    def extract_info(self, text):
        """Extract information from text using regex patterns."""
        info = {}
        for key, pattern in self.patterns.items():
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                info[key] = match.group(1).strip()
        return info

    def process_documents(self, uploaded_files):
        """Process uploaded .docx files and return data."""
        data_list = []
        for uploaded_file in uploaded_files:
            document_text = self.read_docx(uploaded_file)
            if document_text is None:
                continue
            data = self.extract_info(document_text)
            data_list.append([data.get(header, '') for header in self.headers])
        return data_list