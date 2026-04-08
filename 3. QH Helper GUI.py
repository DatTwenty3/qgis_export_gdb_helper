import os
from qgis.core import (
    QgsField,
    QgsFeatureRequest,
    QgsLayerTreeGroup,
    QgsLayerTreeLayer,
    QgsMapLayerType,
    QgsProject,
    QgsVectorFileWriter,
    QgsVectorLayer,
)
from qgis.gui import QgsProjectionSelectionDialog
from qgis.PyQt.QtCore import QVariant, Qt
from qgis.PyQt.QtGui import QTextCursor
from qgis.PyQt.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# --- FIX TRUNCATION TEN LAYER DXF (OGR "Layer" HAY BI CAT 10 KY TU) ---
def _extract_dxf_layer_names(file_path: str):
    """
    Trả về list tên layer đầy đủ trong bảng LAYER của DXF (ASCII).
    """
    try:
        with open(file_path, "rb") as f:
            head = f.read(32)
        if b"AutoCAD Binary DXF" in head:
            return {}

        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            lines = [ln.rstrip("\r\n") for ln in f]

        in_layer_table = False
        pending_layer_entity = False
        full_names = []
        i = 0
        n = len(lines)
        while i + 1 < n:
            code = lines[i].strip()
            value = lines[i + 1].strip()

            if code == "0" and value == "TABLE":
                if i + 3 < n and lines[i + 2].strip() == "2" and lines[i + 3].strip().upper() == "LAYER":
                    in_layer_table = True
                    i += 4
                    continue

            if in_layer_table and code == "0" and value == "ENDTAB":
                in_layer_table = False
                pending_layer_entity = False
                i += 2
                continue

            if in_layer_table and code == "0" and value == "LAYER":
                pending_layer_entity = True
                i += 2
                continue

            if in_layer_table and pending_layer_entity and code == "2":
                if value:
                    full_names.append(value)
                pending_layer_entity = False
                i += 2
                continue

            i += 2

        return full_names
    except Exception:
        return []


def _restore_layer_name(cad_layer_str: str, dxf_full_layer_names):
    if not dxf_full_layer_names:
        return cad_layer_str

    if cad_layer_str in dxf_full_layer_names:
        return cad_layer_str

    matches = [n for n in dxf_full_layer_names if n.startswith(cad_layer_str)]
    if len(matches) == 1:
        return matches[0]

    key10 = cad_layer_str[:10]
    matches10 = [n for n in dxf_full_layer_names if n.startswith(key10)]
    if len(matches10) == 1:
        return matches10[0]

    return cad_layer_str

# --- TUONG THICH ENUM GIUA CAC PHIEN BAN PYQT/QGIS ---
try:
    ORIENTATION_VERTICAL = Qt.Orientation.Vertical
except AttributeError:
    ORIENTATION_VERTICAL = Qt.Vertical
try:
    ORIENTATION_HORIZONTAL = Qt.Orientation.Horizontal
except AttributeError:
    ORIENTATION_HORIZONTAL = Qt.Horizontal

try:
    CHK_CHECKED = Qt.CheckState.Checked
    CHK_UNCHECKED = Qt.CheckState.Unchecked
except AttributeError:
    CHK_CHECKED = Qt.Checked
    CHK_UNCHECKED = Qt.Unchecked

try:
    FLAG_CHECKABLE = Qt.ItemFlag.ItemIsUserCheckable
except AttributeError:
    FLAG_CHECKABLE = Qt.ItemIsUserCheckable

try:
    USER_ROLE = Qt.ItemDataRole.UserRole
except AttributeError:
    USER_ROLE = Qt.UserRole

try:
    TEXT_CURSOR_END = QTextCursor.MoveOperation.End
except AttributeError:
    TEXT_CURSOR_END = QTextCursor.End


class QHHelperWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("QH Export GDB Helper - LEDAT")
        self.resize(1120, 760)
        self.setStyleSheet(self._build_style())
        self._build_ui()
        self.refresh_vector_layers()

    def _build_style(self):
        return """
        QMainWindow { background-color: #f3f6fb; }
        QLabel#title { font-size: 20px; font-weight: 700; color: #1f2937; }
        QLabel#subtitle { color: #4b5563; }
        QGroupBox {
            background: #ffffff;
            border: 1px solid #dbe3ef;
            border-radius: 10px;
            margin-top: 10px;
            font-weight: 600;
            color: #1f2937;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px;
        }
        QLineEdit, QListWidget, QTextEdit {
            border: 1px solid #c9d4e5;
            border-radius: 8px;
            padding: 6px;
            background: #ffffff;
        }
        QPushButton {
            background-color: #2563eb;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 8px 12px;
            font-weight: 600;
        }
        QPushButton:hover { background-color: #1d4ed8; }
        QPushButton:pressed { background-color: #1e40af; }
        QPushButton#ghost {
            background-color: #e8eefb;
            color: #1e40af;
        }
        QTabWidget::pane {
            border: 1px solid #dbe3ef;
            border-radius: 10px;
            background: #ffffff;
            top: -1px;
        }
        QTabBar::tab {
            background: #e8eefb;
            color: #1e3a8a;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            min-width: 120px;
            padding: 8px 12px;
            margin-right: 4px;
        }
        QTabBar::tab:selected {
            background: #ffffff;
            color: #111827;
            border: 1px solid #dbe3ef;
            border-bottom: 1px solid #ffffff;
        }
        """

    def _build_ui(self):
        container = QWidget()
        self.setCentralWidget(container)

        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(14, 14, 14, 14)
        root_layout.setSpacing(10)

        header = QFrame()
        header_layout = QVBoxLayout(header)
        lbl_title = QLabel("QH Export GDB Helper")
        lbl_title.setObjectName("title")
        lbl_sub = QLabel("Tích hợp đầy đủ: Nhập bản vẽ CAD -> Tạo thuộc tính -> Xuất File Geodatabase")
        lbl_sub.setObjectName("subtitle")
        header_layout.addWidget(lbl_title)
        header_layout.addWidget(lbl_sub)
        root_layout.addWidget(header)

        splitter = QSplitter(ORIENTATION_HORIZONTAL)
        root_layout.addWidget(splitter, 1)

        left_panel = self._build_left_panel()
        splitter.addWidget(left_panel)

        log_box = QGroupBox("Nhật ký xử lý")
        log_layout = QVBoxLayout(log_box)
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)
        splitter.addWidget(log_box)
        splitter.setSizes([760, 360])

    def _build_left_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        # Bước 1
        box = QGroupBox("Bước 1 - Nhập CAD và tách lớp theo hình học")
        box_layout = QVBoxLayout(box)

        row = QHBoxLayout()
        self.input_cad_path = QLineEdit()
        self.input_cad_path.setPlaceholderText("Chọn file CAD (khuyến nghị DXF)")
        btn_browse = QPushButton("Chọn file")
        btn_browse.clicked.connect(self.choose_cad_file)
        row.addWidget(self.input_cad_path, 1)
        row.addWidget(btn_browse)
        box_layout.addLayout(row)

        btn_run = QPushButton("Chạy nhập bản vẽ")
        btn_run.clicked.connect(self.import_and_split_cad)
        box_layout.addWidget(btn_run)
        layout.addWidget(box)

        # Bước 2
        form_box = QGroupBox("Bước 2 - Thông tin thuộc tính quy hoạch")
        form_layout = QVBoxLayout(form_box)

        self.input_ma_tt = QLineEdit()
        self.input_ma_hs = QLineEdit("84QHC1000001")
        self.input_ma_dt = QLineEdit()
        self.input_ten_dt = QLineEdit()
        self.input_phan_loai = QLineEdit()
        self.input_ghi_chu = QLineEdit()

        self._add_labeled_input(form_layout, "Mã thông tin quy hoạch (maThongTinQH):", self.input_ma_tt)
        self._add_labeled_input(form_layout, "Mã hồ sơ quy hoạch (maHoSoQH):", self.input_ma_hs)
        self._add_labeled_input(form_layout, "Mã đối tượng (maDoiTuong):", self.input_ma_dt)
        self._add_labeled_input(form_layout, "Tên đối tượng (tenDoiTuong):", self.input_ten_dt)
        self._add_labeled_input(form_layout, "Phân loại (phanLoai):", self.input_phan_loai)
        self._add_labeled_input(form_layout, "Ghi chú (ghiChu):", self.input_ghi_chu)

        self.chk_delete_old = QCheckBox("Xóa toàn bộ thuộc tính cũ trước khi thêm thuộc tính chuẩn")
        form_layout.addWidget(self.chk_delete_old)
        layout.addWidget(form_box)

        layer_box = QGroupBox("Bước 2.1 - Chọn layer cần cập nhật")
        layer_layout = QVBoxLayout(layer_box)
        self.list_layers = QListWidget()
        layer_layout.addWidget(self.list_layers)

        btn_row = QHBoxLayout()
        btn_select_all = QPushButton("Chọn tất cả")
        btn_select_all.setObjectName("ghost")
        btn_unselect_all = QPushButton("Bỏ chọn tất cả")
        btn_unselect_all.setObjectName("ghost")
        btn_refresh = QPushButton("Làm mới danh sách")
        btn_refresh.setObjectName("ghost")
        btn_select_all.clicked.connect(self.select_all_layers)
        btn_unselect_all.clicked.connect(self.unselect_all_layers)
        btn_refresh.clicked.connect(self.refresh_vector_layers)
        btn_row.addWidget(btn_select_all)
        btn_row.addWidget(btn_unselect_all)
        btn_row.addWidget(btn_refresh)
        layer_layout.addLayout(btn_row)

        btn_run = QPushButton("Chạy tạo thuộc tính")
        btn_run.clicked.connect(self.add_fields_and_data)

        layout.addWidget(layer_box, 1)
        layout.addWidget(btn_run)

        # Bước 3
        box = QGroupBox("Bước 3 - Xuất toàn bộ layer vector theo nhóm ra FileGDB")
        box_layout = QVBoxLayout(box)
        tip = QLabel("Layer sẽ được xuất theo cấu trúc group trong Layer Tree. Bỏ qua nhóm tên 'base map'.")
        tip.setWordWrap(True)
        box_layout.addWidget(tip)

        btn_run = QPushButton("Chạy xuất GDB")
        btn_run.clicked.connect(self.export_to_gdb)
        box_layout.addWidget(btn_run)

        layout.addWidget(box)
        return panel

    def _add_labeled_input(self, layout, text, widget):
        layout.addWidget(QLabel(text))
        layout.addWidget(widget)

    def log(self, message):
        text = str(message)
        lower_text = text.lower()

        color = "#1f2937"  # mặc định
        if "lỗi" in lower_text or "không thành công" in lower_text or "thất bại" in lower_text:
            color = "#b91c1c"
        elif "hoàn tất" in lower_text or text.strip().startswith("  +"):
            color = "#047857"
        elif "đang" in lower_text or "bắt đầu" in lower_text:
            color = "#1d4ed8"
        elif "hủy" in lower_text:
            color = "#b45309"

        safe_text = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        self.log_edit.append(f'<span style="color:{color};">{safe_text}</span>')
        self.log_edit.moveCursor(TEXT_CURSOR_END)
        QApplication.processEvents()

    def show_done_message(self):
        QMessageBox.information(
            self,
            "Hoàn tất",
            "Quá trình đã hoàn tất.\nPhần mềm được phát triển bởi LEDAT.\nCảm ơn đã sử dụng.",
        )

    def choose_cad_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file CAD (Nên dùng DXF)",
            "",
            "DXF Files (*.dxf);;CAD Files (*.dxf *.dwg)",
        )
        if file_path:
            self.input_cad_path.setText(file_path)

    def refresh_vector_layers(self):
        self.list_layers.clear()
        all_layers = QgsProject.instance().mapLayers().values()
        vector_layers = [lyr for lyr in all_layers if lyr.type() == QgsMapLayerType.VectorLayer]
        for layer in vector_layers:
            item = QListWidgetItem(layer.name())
            item.setFlags(item.flags() | FLAG_CHECKABLE)
            item.setCheckState(CHK_CHECKED)
            item.setData(USER_ROLE, layer.id())
            self.list_layers.addItem(item)
        self.log(f"Đã tải danh sách layer: {len(vector_layers)} layer vector.")

    def select_all_layers(self):
        for i in range(self.list_layers.count()):
            self.list_layers.item(i).setCheckState(CHK_CHECKED)

    def unselect_all_layers(self):
        for i in range(self.list_layers.count()):
            self.list_layers.item(i).setCheckState(CHK_UNCHECKED)

    def _selected_vector_layers(self):
        id_map = QgsProject.instance().mapLayers()
        selected = []
        for i in range(self.list_layers.count()):
            item = self.list_layers.item(i)
            if item.checkState() == CHK_CHECKED:
                layer_id = item.data(USER_ROLE)
                layer = id_map.get(layer_id)
                if layer and layer.type() == QgsMapLayerType.VectorLayer:
                    selected.append(layer)
        return selected

    def _get_field_idx(self, layer, field_name):
        idx = layer.fields().indexOf(field_name)
        if idx == -1:
            idx = layer.fields().indexOf(field_name[:10])
        return idx

    # --- CHUC NANG 1: NHAP BAN VE ---
    def import_and_split_cad(self):
        file_path = self.input_cad_path.text().strip()
        if not file_path:
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Chọn file CAD (Nên dùng DXF)",
                "",
                "DXF Files (*.dxf);;CAD Files (*.dxf *.dwg)",
            )
            if not file_path:
                self.log("Đã hủy chọn file CAD.")
                return
            self.input_cad_path.setText(file_path)

        crs_dialog = QgsProjectionSelectionDialog()
        crs_dialog.setWindowTitle("Chọn hệ tọa độ cho bản vẽ CAD")
        if not crs_dialog.exec():
            self.log("Đã hủy chọn hệ tọa độ.")
            return
        selected_crs = crs_dialog.crs()

        file_name = os.path.basename(file_path)
        file_base_name = os.path.splitext(file_name)[0]
        self.log(f"Đang xử lý file: {file_name} với hệ tọa độ {selected_crs.authid()}...")

        dxf_full_layer_names = _extract_dxf_layer_names(file_path) if file_path.lower().endswith(".dxf") else []

        root = QgsProject.instance().layerTreeRoot()
        cad_group = root.addGroup(f"CAD_{file_base_name}")

        geometry_mapping = {
            "Point": ("", "Điểm (Point)"),
            "LineString": ("", "Đường (Line)"),
            "Polygon": ("", "Vùng (Polygon)"),
        }

        layer_count = 0
        subgroups = {}

        for geom_type, (suffix, group_name) in geometry_mapping.items():
            uri = f"{file_path}|layername=entities|geometrytype={geom_type}"
            temp_layer = QgsVectorLayer(uri, "temp", "ogr")

            if not temp_layer.isValid() or temp_layer.featureCount() == 0:
                continue

            idx = temp_layer.fields().indexOf("Layer")
            if idx == -1:
                continue

            unique_cad_layers = temp_layer.uniqueValues(idx)

            for cad_layer in unique_cad_layers:
                if not cad_layer:
                    continue

                cad_layer_str = str(cad_layer)
                restored_layer = _restore_layer_name(cad_layer_str, dxf_full_layer_names)

                clean_name = restored_layer.strip().replace(" ", "_").replace("-", "_")
                layer_name = clean_name
                new_layer = QgsVectorLayer(uri, layer_name, "ogr")
                safe_cad_layer = cad_layer_str.replace("'", "''")
                new_layer.setSubsetString(f"\"Layer\" = '{safe_cad_layer}'")

                if new_layer.isValid() and new_layer.featureCount() > 0:
                    editable_layer = new_layer.materialize(QgsFeatureRequest())
                    editable_layer.setName(layer_name)
                    editable_layer.setCrs(selected_crs)
                    QgsProject.instance().addMapLayer(editable_layer, False)

                    if group_name not in subgroups:
                        subgroups[group_name] = cad_group.addGroup(group_name)

                    subgroups[group_name].addLayer(editable_layer)
                    layer_count += 1
                    self.log(f"  + Đã tách và mở khóa: {layer_name} -> Nhóm: {group_name}")

        if layer_count > 0:
            self.log(f"HOÀN TẤT! Đã import và phân loại {layer_count} layer.")
            self.refresh_vector_layers()
            self.show_done_message()
        else:
            self.log("KHÔNG THÀNH CÔNG! Không tìm thấy dữ liệu hợp lệ (thử Save As DWG -> DXF).")

    # --- CHUC NANG 2: TAO THUOC TINH ---
    def add_fields_and_data(self):
        selected_layers = self._selected_vector_layers()
        if not selected_layers:
            self.log("Bạn chưa chọn layer nào để xử lý.")
            return

        ma_tt_qh = self.input_ma_tt.text().strip()
        ma_hs_qh = self.input_ma_hs.text().strip()
        ma_dt_goc = self.input_ma_dt.text().strip()
        ten_dt = self.input_ten_dt.text().strip()
        phan_loai = self.input_phan_loai.text().strip()
        ghi_chu = self.input_ghi_chu.text().strip()
        delete_old_fields = self.chk_delete_old.isChecked()

        self.log(f"Bắt đầu xử lý {len(selected_layers)} layer đã chọn...")

        fields_to_add = [
            QgsField("maThongTinQH", QVariant.String, len=15),
            QgsField("maHoSoQH", QVariant.String, len=15),
            QgsField("maDoiTuong", QVariant.String, len=100),
            QgsField("tenDoiTuong", QVariant.String, len=100),
            QgsField("phanLoai", QVariant.String, len=250),
            QgsField("ghiChu", QVariant.String, len=250),
        ]

        for layer in selected_layers:
            self.log(f"Đang cập nhật: {layer.name()}")
            layer.startEditing()
            pr = layer.dataProvider()

            if delete_old_fields:
                old_field_count = len(layer.fields())
                if old_field_count > 0:
                    pr.deleteAttributes(list(range(old_field_count)))
                    layer.updateFields()
                    self.log(f"  + Đã xóa {old_field_count} thuộc tính cũ")

            existing_fields = layer.fields().names()
            new_fields = [
                f
                for f in fields_to_add
                if f.name() not in existing_fields and f.name()[:10] not in existing_fields
            ]
            if new_fields:
                pr.addAttributes(new_fields)
                layer.updateFields()

            idx_maThongTinQH = self._get_field_idx(layer, "maThongTinQH")
            idx_maHoSoQH = self._get_field_idx(layer, "maHoSoQH")
            idx_maDoiTuong = self._get_field_idx(layer, "maDoiTuong")
            idx_tenDoiTuong = self._get_field_idx(layer, "tenDoiTuong")
            idx_phanLoai = self._get_field_idx(layer, "phanLoai")
            idx_ghiChu = self._get_field_idx(layer, "ghiChu")

            update_dict = {}
            for feat in layer.getFeatures():
                attr_map = {}
                if idx_maThongTinQH != -1:
                    attr_map[idx_maThongTinQH] = ma_tt_qh
                if idx_maHoSoQH != -1:
                    attr_map[idx_maHoSoQH] = ma_hs_qh
                if idx_maDoiTuong != -1:
                    attr_map[idx_maDoiTuong] = ma_dt_goc
                if idx_tenDoiTuong != -1 and ten_dt:
                    attr_map[idx_tenDoiTuong] = ten_dt
                if idx_phanLoai != -1 and phan_loai:
                    attr_map[idx_phanLoai] = phan_loai
                if idx_ghiChu != -1 and ghi_chu:
                    attr_map[idx_ghiChu] = ghi_chu
                update_dict[feat.id()] = attr_map

            pr.changeAttributeValues(update_dict)
            layer.commitChanges()
            self.log(f"  + Hoàn tất: {layer.name()}")

        self.log("--- ĐÃ CHẠY XONG TOÀN BỘ YÊU CẦU ---")
        self.show_done_message()

    # --- CHUC NANG 3: TAO GDB ---
    def export_to_gdb(self):
        gdb_path, _ = QFileDialog.getSaveFileName(
            self,
            "Chọn nơi lưu và đặt tên file Geodatabase",
            "",
            "Esri File Geodatabase (*.gdb)",
        )

        if not gdb_path:
            self.log("Đã hủy thao tác lưu file.")
            return

        if not gdb_path.endswith(".gdb"):
            gdb_path += ".gdb"

        self.log(f"Bắt đầu xuất dữ liệu ra: {gdb_path}")

        gdb_exists = os.path.exists(gdb_path)
        exported_layers = 0

        def process_group(group):
            nonlocal gdb_exists, exported_layers
            group_name = group.name()
            if group_name == "base map":
                return

            self.log(f"Đang quét nhóm: {group_name}")
            for child in group.children():
                if isinstance(child, QgsLayerTreeGroup):
                    process_group(child)
                elif isinstance(child, QgsLayerTreeLayer):
                    layer = child.layer()
                    if layer and layer.type() == QgsMapLayerType.VectorLayer:
                        safe_layer_name = layer.name().replace(" ", "_").replace("-", "_")
                        options = QgsVectorFileWriter.SaveVectorOptions()
                        options.driverName = "OpenFileGDB"
                        options.layerName = safe_layer_name
                        if not gdb_exists:
                            options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteFile
                        else:
                            options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteLayer
                        options.layerOptions = [f"FEATURE_DATASET={group_name}"]

                        error = QgsVectorFileWriter.writeAsVectorFormatV3(
                            layer,
                            gdb_path,
                            QgsProject.instance().transformContext(),
                            options,
                        )
                        if error[0] == QgsVectorFileWriter.NoError:
                            self.log(f"  + Đã xuất: {layer.name()} -> Nhóm GDB: {group_name}")
                            gdb_exists = True
                            exported_layers += 1
                        else:
                            self.log(f"  - Lỗi khi xuất {layer.name()}: {error}")

        root = QgsProject.instance().layerTreeRoot()
        for child in root.children():
            if isinstance(child, QgsLayerTreeGroup):
                process_group(child)

        if exported_layers > 0:
            self.log(f"--- HOÀN TẤT --- Đã xuất {exported_layers} layer.")
            self.show_done_message()
        else:
            self.log("--- HOÀN TẤT --- Không có layer nào được xuất.")


def show_qh_helper_gui():
    global QH_HELPER_WINDOW
    try:
        QH_HELPER_WINDOW.close()
    except Exception:
        pass

    QH_HELPER_WINDOW = QHHelperWindow()
    QH_HELPER_WINDOW.show()
    QH_HELPER_WINDOW.raise_()
    QH_HELPER_WINDOW.activateWindow()


show_qh_helper_gui()
