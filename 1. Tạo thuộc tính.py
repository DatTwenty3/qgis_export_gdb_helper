from qgis.core import QgsProject, QgsField, QgsMapLayerType
from qgis.PyQt.QtCore import QVariant, Qt
from qgis.PyQt.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                                 QLineEdit, QListWidget, QListWidgetItem, 
                                 QPushButton, QDialogButtonBox, QScrollArea, QWidget, QMessageBox, QCheckBox)

# --- XỬ LÝ PHIÊN BẢN PYQT ---
try:
    CHK_CHECKED = Qt.CheckState.Checked
    CHK_UNCHECKED = Qt.CheckState.Unchecked
    FLAG_CHECKABLE = Qt.ItemFlag.ItemIsUserCheckable
    USER_ROLE = Qt.ItemDataRole.UserRole
    BTN_OK = QDialogButtonBox.StandardButton.Ok
    BTN_CANCEL = QDialogButtonBox.StandardButton.Cancel
except AttributeError:
    CHK_CHECKED = Qt.Checked
    CHK_UNCHECKED = Qt.Unchecked
    FLAG_CHECKABLE = Qt.ItemIsUserCheckable
    USER_ROLE = Qt.UserRole
    BTN_OK = QDialogButtonBox.Ok
    BTN_CANCEL = QDialogButtonBox.Cancel

class LayerSelectionDialog(QDialog):
    def __init__(self, vector_layers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cập nhật Thuộc Tính Quy Hoạch")
        self.resize(500, 650)
        
        layout = QVBoxLayout(self)

        # Tạo vùng cuộn (Scroll Area) để chứa nhiều ô nhập liệu không bị tràn màn hình
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)

        # 1. Khu vực nhập liệu (Đã bổ sung đầy đủ)
        scroll_layout.addWidget(QLabel("Mã thông tin quy hoạch (maThongTinQH):"))
        self.input_ma_tt = QLineEdit(self)
        scroll_layout.addWidget(self.input_ma_tt)

        scroll_layout.addWidget(QLabel("Mã hồ sơ quy hoạch (maHoSoQH) - VD: 84QHC1000001:"))
        self.input_ma_hs = QLineEdit("84QHC1000001", self)
        scroll_layout.addWidget(self.input_ma_hs)

        scroll_layout.addWidget(QLabel("Mã đối tượng (maDoiTuong) - do người dùng nhập:"))
        self.input_ma_dt = QLineEdit(self)
        scroll_layout.addWidget(self.input_ma_dt)
        
        scroll_layout.addWidget(QLabel("Tên đối tượng (tenDoiTuong) - Bỏ trống nếu nhập sau:"))
        self.input_ten_dt = QLineEdit(self)
        scroll_layout.addWidget(self.input_ten_dt)
        
        scroll_layout.addWidget(QLabel("Phân loại (phanLoai) - Bỏ trống nếu nhập sau:"))
        self.input_phan_loai = QLineEdit(self)
        scroll_layout.addWidget(self.input_phan_loai)
        
        scroll_layout.addWidget(QLabel("Ghi chú (ghiChu) - Bỏ trống nếu nhập sau:"))
        self.input_ghi_chu = QLineEdit(self)
        scroll_layout.addWidget(self.input_ghi_chu)

        # Tùy chọn xử lý thuộc tính cũ
        self.chk_delete_old_fields = QCheckBox("Xóa toàn bộ thuộc tính cũ của layer trước khi thêm thuộc tính mới", self)
        self.chk_delete_old_fields.setChecked(False)
        scroll_layout.addWidget(self.chk_delete_old_fields)

        # 2. Khu vực danh sách Layer
        scroll_layout.addWidget(QLabel("Chọn các layer cần thêm trường dữ liệu:"))
        self.list_widget = QListWidget(self)
        self.layer_map = {}
        
        for layer in vector_layers:
            item = QListWidgetItem(layer.name())
            item.setFlags(item.flags() | FLAG_CHECKABLE)
            item.setCheckState(CHK_CHECKED)
            item.setData(USER_ROLE, layer.id()) 
            self.list_widget.addItem(item)
            self.layer_map[layer.id()] = layer
            
        scroll_layout.addWidget(self.list_widget)

        # 3. Các nút chức năng chọn nhanh
        btn_layout = QHBoxLayout()
        btn_select_all = QPushButton("Chọn tất cả")
        btn_deselect_all = QPushButton("Bỏ chọn tất cả")
        
        btn_select_all.clicked.connect(self.select_all)
        btn_deselect_all.clicked.connect(self.deselect_all)
        
        btn_layout.addWidget(btn_select_all)
        btn_layout.addWidget(btn_deselect_all)
        scroll_layout.addLayout(btn_layout)

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        # 4. Nút OK / Cancel 
        self.button_box = QDialogButtonBox(BTN_OK | BTN_CANCEL)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)

    def select_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(CHK_CHECKED)

    def deselect_all(self):
        for i in range(self.list_widget.count()):
            self.list_widget.item(i).setCheckState(CHK_UNCHECKED)

    def get_selected_layers(self):
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == CHK_CHECKED:
                layer_id = item.data(USER_ROLE)
                selected.append(self.layer_map[layer_id])
        return selected

# Hàm hỗ trợ lấy Index cột (Khắc phục lỗi Shapefile cắt tên thành 10 ký tự)
def get_field_idx(layer, field_name):
    idx = layer.fields().indexOf(field_name)
    if idx == -1:
        # Thử tìm tên cột bị cắt bớt 10 ký tự (Đặc thù của file .shp)
        idx = layer.fields().indexOf(field_name[:10])
    return idx

def add_fields_and_data():
    all_layers = QgsProject.instance().mapLayers().values()
    vector_layers = [lyr for lyr in all_layers if lyr.type() == QgsMapLayerType.VectorLayer]
    
    if not vector_layers:
        print("Không tìm thấy layer Vector nào trong dự án.")
        return

    dialog = LayerSelectionDialog(vector_layers)
    
    if hasattr(dialog, 'exec'):
        is_accepted = dialog.exec()
    else:
        is_accepted = dialog.exec_()

    if is_accepted:
        # Lấy dữ liệu từ giao diện
        ma_tt_qh = dialog.input_ma_tt.text().strip()
        ma_hs_qh = dialog.input_ma_hs.text().strip()
        ma_dt_goc = dialog.input_ma_dt.text().strip()
        ten_dt = dialog.input_ten_dt.text().strip()
        phan_loai = dialog.input_phan_loai.text().strip()
        ghi_chu = dialog.input_ghi_chu.text().strip()
        delete_old_fields = dialog.chk_delete_old_fields.isChecked()
        
        selected_layers = dialog.get_selected_layers()
        if not selected_layers:
            print("Bạn chưa chọn layer nào để xử lý.")
            return
            
        print(f"Bắt đầu xử lý {len(selected_layers)} layer đã chọn...\n")

        # Cấu trúc trường theo chuẩn (Độ dài được khai báo đúng quy định)
        fields_to_add = [
            QgsField("maThongTinQH", QVariant.String, len=15),
            QgsField("maHoSoQH", QVariant.String, len=15),
            QgsField("maDoiTuong", QVariant.String, len=100),
            QgsField("tenDoiTuong", QVariant.String, len=100),
            QgsField("phanLoai", QVariant.String, len=250),
            QgsField("ghiChu", QVariant.String, len=250)
        ]

        for layer in selected_layers:
            print(f"Đang cập nhật: {layer.name()}")
            layer.startEditing()
            pr = layer.dataProvider()

            # 1. Tùy chọn xóa toàn bộ trường cũ
            if delete_old_fields:
                old_field_count = len(layer.fields())
                if old_field_count > 0:
                    pr.deleteAttributes(list(range(old_field_count)))
                    layer.updateFields()
                    print(f"  + Đã xóa {old_field_count} thuộc tính cũ")

            # 2. Thêm trường nếu chưa có
            existing_fields = layer.fields().names()
            new_fields = [f for f in fields_to_add if f.name() not in existing_fields and f.name()[:10] not in existing_fields]
            if new_fields:
                pr.addAttributes(new_fields)
                layer.updateFields()

            # 3. Lấy Index chính xác của cột (Kể cả khi bị cắt 10 ký tự)
            idx_maThongTinQH = get_field_idx(layer, "maThongTinQH")
            idx_maHoSoQH = get_field_idx(layer, "maHoSoQH")
            idx_maDoiTuong = get_field_idx(layer, "maDoiTuong")
            idx_tenDoiTuong = get_field_idx(layer, "tenDoiTuong")
            idx_phanLoai = get_field_idx(layer, "phanLoai")
            idx_ghiChu = get_field_idx(layer, "ghiChu")

            # 4. Cập nhật dữ liệu
            update_dict = {}
            for feat in layer.getFeatures():
                obj_id = feat.id()
                ma_doi_tuong = ma_dt_goc
                
                # Chỉ cập nhật các cột tìm thấy
                attr_map = {}
                if idx_maThongTinQH != -1: attr_map[idx_maThongTinQH] = ma_tt_qh
                if idx_maHoSoQH != -1: attr_map[idx_maHoSoQH] = ma_hs_qh
                if idx_maDoiTuong != -1: attr_map[idx_maDoiTuong] = ma_doi_tuong
                if idx_tenDoiTuong != -1 and ten_dt: attr_map[idx_tenDoiTuong] = ten_dt
                if idx_phanLoai != -1 and phan_loai: attr_map[idx_phanLoai] = phan_loai
                if idx_ghiChu != -1 and ghi_chu: attr_map[idx_ghiChu] = ghi_chu
                
                update_dict[obj_id] = attr_map

            pr.changeAttributeValues(update_dict)
            layer.commitChanges()
            print(f"  + Hoàn tất: {layer.name()}")

        print("\n--- ĐÃ CHẠY XONG TOÀN BỘ YÊU CẦU ---")
        QMessageBox.information(
            None,
            "Hoàn tất",
            "Quá trình đã hoàn tất.\nPhần mềm được phát triển bởi LEDAT.\nCảm ơn đã sử dụng."
        )
    else:
        print("Đã hủy thao tác.")

# Gọi tool
add_fields_and_data()