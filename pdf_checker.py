import os
from PyPDF2 import PdfReader
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog

class PDFChecker:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def parse_pdf_date(self, date_str):
        """Chuyển đổi chuỗi ngày tháng từ metadata của PDF thành datetime object"""
        if not date_str:
            return None
        try:
            # Xử lý các định dạng ngày tháng phổ biến trong metadata PDF
            if isinstance(date_str, str):
                # Loại bỏ D: nếu có
                if date_str.startswith('D:'):
                    date_str = date_str[2:]
                
                # Xử lý múi giờ nếu có
                if "+" in date_str:
                    date_str = date_str.split("+")[0]
                elif "Z" in date_str:
                    date_str = date_str.replace("Z", "")

                # Thử các định dạng phổ biến
                formats = [
                    '%Y%m%d%H%M%S',
                    '%Y%m%d%H%M',
                    '%Y%m%d'
                ]
                
                for fmt in formats:
                    try:
                        return datetime.strptime(date_str[:len(fmt)], fmt)
                    except ValueError:
                        continue
            
            return None
        except Exception as e:
            print(f"Lỗi khi xử lý ngày tháng: {str(e)}")
            return None

    def check_pdf_metadata(self):
        """Kiểm tra metadata của các file PDF trong thư mục được chọn"""
        print("Vui lòng chọn thư mục chứa các file PDF...")
        folder_path = filedialog.askdirectory(title="Chọn thư mục chứa file PDF")
        if not folder_path:
            print("Không có thư mục nào được chọn.")
            return

        results = []
        print("\nĐang phân tích các file PDF...")

        for filename in os.listdir(folder_path):
            if filename.lower().endswith('.pdf'):
                file_path = os.path.join(folder_path, filename)
                try:
                    reader = PdfReader(file_path)
                    metadata = reader.metadata

                    # Lấy thông tin ngày tháng từ metadata
                    creation_date = self.parse_pdf_date(metadata.get('/CreationDate', ''))
                    mod_date = self.parse_pdf_date(metadata.get('/ModDate', ''))

                    # Kiểm tra xem file có bị chỉnh sửa không
                    is_modified = False
                    if creation_date and mod_date:
                        is_modified = mod_date > creation_date
                        time_diff = mod_date - creation_date if is_modified else None
                    else:
                        time_diff = None

                    # Thêm kết quả
                    results.append({
                        'Tên file': filename,
                        'Ngày tạo': creation_date.strftime('%Y-%m-%d %H:%M:%S') if creation_date else 'Không có',
                        'Ngày chỉnh sửa': mod_date.strftime('%Y-%m-%d %H:%M:%S') if mod_date else 'Không có',
                        'Đã chỉnh sửa': 'Có' if is_modified else 'Không',
                        'Thời gian chỉnh sửa': str(time_diff) if time_diff else 'N/A',
                        'Sample': 'Sample' if is_modified else ''
                    })

                    # In thông tin chi tiết
                    print(f"\nPhân tích file: {filename}")
                    print(f"Ngày tạo: {creation_date.strftime('%Y-%m-%d %H:%M:%S') if creation_date else 'Không có'}")
                    print(f"Ngày chỉnh sửa: {mod_date.strftime('%Y-%m-%d %H:%M:%S') if mod_date else 'Không có'}")
                    if is_modified:
                        print(f"File đã được chỉnh sửa sau khi tạo {time_diff}")
                    else:
                        print("File chưa bị chỉnh sửa sau khi tạo")

                except Exception as e:
                    print(f"\nLỗi khi đọc file {filename}: {str(e)}")
                    results.append({
                        'Tên file': filename,
                        'Ngày tạo': 'Lỗi',
                        'Ngày chỉnh sửa': 'Lỗi',
                        'Đã chỉnh sửa': 'Lỗi',
                        'Thời gian chỉnh sửa': 'Lỗi',
                        'Sample': 'Lỗi'
                    })

        # Tạo DataFrame và lưu kết quả
        df = pd.DataFrame(results)
        excel_path = os.path.join(folder_path, 'pdf_metadata_analysis.xlsx')
        df.to_excel(excel_path, index=False)
        print(f"\nĐã lưu kết quả phân tích vào file: {excel_path}")

if __name__ == "__main__":
    checker = PDFChecker()
    checker.check_pdf_metadata() 