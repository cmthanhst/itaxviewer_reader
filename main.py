import sys
import csv
import os
import xml.etree.ElementTree as ET
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QFileDialog, QTableWidget, QTableWidgetItem, QHeaderView
)
from PySide6.QtCore import Qt

class XMLApp(QMainWindow):
    """
    Một ứng dụng PySide6 để đọc dữ liệu từ nhiều tệp XML của Tổng cục Thuế,
    hiển thị nó trong bảng và lưu nó vào một tệp CSV.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Trình đọc XML (Tờ khai thuế) và xuất CSV")
        self.setGeometry(100, 100, 600, 700)

        # Danh sách các trường cần trích xuất từ XML
        self.fields_to_extract = [
            "21", "22", "23", "24", "23a", "24a", "25", "26", "27", "28", "29",
            "30", "31", "32", "33", "32a", "34", "35", "36", "37", "38", "39a",
            "40a", "40b", "40", "41", "42", "43"
        ]
        
        # Dữ liệu được trích xuất sẽ được lưu trữ ở đây
        self.extracted_data = []

        # Thiết lập giao diện người dùng
        self._setup_ui()

    def _setup_ui(self):
        """Thiết lập các thành phần giao diện người dùng."""
        layout = QVBoxLayout()

        # Nút mở tệp(nhiều tệp)
        self.open_button = QPushButton("Mở (các) tệp XML")
        self.open_button.clicked.connect(self.open_xml_file)
        layout.addWidget(self.open_button)

        # Bảng hiển thị dữ liệu
        self.table = QTableWidget()
        # Cấu hình bảng để hiển thị dữ liệu theo hàng
        self.table.setColumnCount(len(self.fields_to_extract) + 1)
        self.table.setHorizontalHeaderLabels(["Tên tệp"] + self.fields_to_extract)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        layout.addWidget(self.table)

        # Nút lưu tệp
        self.save_button = QPushButton("Lưu ra CSV")
        self.save_button.clicked.connect(self.save_to_csv)
        self.save_button.setEnabled(False)
        layout.addWidget(self.save_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        
        self.statusBar().showMessage("Sẵn sàng. Vui lòng mở một tệp XML.")

    def open_xml_file(self):
        """Mở hộp thoại tệp để chọn và xử lý nhiều tệp XML."""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Chọn (các) tệp XML", "", "XML Files (*.xml)"
        )
        if file_paths:
            self.extracted_data.clear()
            self.table.setRowCount(0) # Xóa dữ liệu cũ trên bảng

            for file_path in file_paths:
                self.parse_and_add_row(file_path)

            self.populate_table()
            self.save_button.setEnabled(True)
            self.statusBar().showMessage(f"Đã tải thành công {len(file_paths)} tệp.")

    def parse_and_add_row(self, file_path):
        """
        Phân tích một tệp XML, trích xuất dữ liệu và chuẩn bị một hàng
        để thêm vào bảng.
        """
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Tự động phát hiện namespace từ thẻ root
            namespace = ''
            if '}' in root.tag:
                namespace = root.tag.split('}')[0][1:]
            
            ns_map = {'ns': namespace}

            file_data = {}
            for field in self.fields_to_extract:
                tag_name = f"ct{field}"
                element = root.find(f".//ns:{tag_name}", ns_map)
                value = element.text if element is not None and element.text else "0"
                file_data[field] = value
            
            # Tạo một hàng dữ liệu hoàn chỉnh
            row_data = [os.path.basename(file_path)] # Thêm tên tệp vào cột đầu tiên
            for field in self.fields_to_extract:
                row_data.append(file_data.get(field, "0"))
            
            self.extracted_data.append(row_data)

        except ET.ParseError:
            self.statusBar().showMessage(f"Lỗi: Tệp '{os.path.basename(file_path)}' không hợp lệ.")
        except Exception as e:
            self.statusBar().showMessage(f"Lỗi khi xử lý tệp '{os.path.basename(file_path)}': {e}")

    def populate_table(self):
        """Điền dữ liệu đã trích xuất vào QTableWidget."""
        self.table.setRowCount(len(self.extracted_data))
        for row_idx, row_data in enumerate(self.extracted_data):
            for col_idx, cell_data in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(cell_data)))
        self.table.resizeColumnsToContents()

    def save_to_csv(self):
        """Lưu dữ liệu từ bảng vào một tệp CSV."""
        if not self.extracted_data:
            self.statusBar().showMessage("Không có dữ liệu để lưu.")
            return

        # Lấy tên tệp mặc định từ tệp XML đầu tiên
        default_filename = self.extracted_data[0][0].replace('.xml', '.csv') if self.extracted_data else "output.csv"
        file_path, _ = QFileDialog.getSaveFileName(self, "Lưu tệp CSV", default_filename, "CSV Files (*.csv)")

        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.writer(csvfile)
                    # Ghi tiêu đề cột
                    headers = ["TenTep"] + self.fields_to_extract
                    writer.writerow(headers)
                    # Ghi tất cả các hàng dữ liệu
                    writer.writerows(self.extracted_data)
                self.statusBar().showMessage(f"Đã lưu thành công vào: {file_path}")
            except Exception as e:
                self.statusBar().showMessage(f"Lỗi khi lưu tệp: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = XMLApp()
    window.show()
    sys.exit(app.exec())
