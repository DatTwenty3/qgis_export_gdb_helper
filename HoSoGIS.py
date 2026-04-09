import os
import tempfile
import zipfile
from collections import defaultdict

from qgis.core import (
    QgsField,
    QgsFeatureRequest,
    QgsLayerTreeGroup,
    QgsLayerTreeLayer,
    QgsMapLayerType,
    QgsProject,
    QgsVectorFileWriter,
    QgsVectorLayer,
    QgsWkbTypes,
)
from qgis.gui import QgsProjectionSelectionDialog
from qgis.PyQt.QtCore import QDate, QDateTime, QVariant, Qt
from qgis.PyQt.QtGui import QTextCursor
from qgis.PyQt.QtWidgets import (
    QApplication,
    QCheckBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QInputDialog,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

# --- EXCEL: xuất/nhập thuộc tính (nhúng trong file, không cần thêm file .py khác; cần openpyxl) ---
try:
    from qgis.core import NULL as QGIS_NULL
except ImportError:
    QGIS_NULL = object()

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    Workbook = None
    load_workbook = None

_FID_HEADER = "qgis_fid"

_QUICK_LAYER_NAME_OPTIONS = [
    "BanKinhBoViaBanKinhTimDuong_P",
    "BoViaDaiPhanCach_L",
    "CaoDoCongTNM_P",
    "CaoDoCongThoatTNT_P",
    "CaoDoNen_P",
    "CayXanh_P",
    "ChiGioiDuongDo_L",
    "ChiGioiXayDung_L",
    "ChucNangCongTrinh_P",
    "ChucNangSuDungDat_A",
    "CongTrinhCBKT_A",
    "CongTrinhCBKT_L",
    "CongTrinhCBKT_P",
    "CongTrinhCapDien_A",
    "CongTrinhCapDien_P",
    "CongTrinhChieuSang_P",
    "CongTrinhGiaoThong_A",
    "CongTrinhGiaoThong_L",
    "CongTrinhGiaoThong_P",
    "CongTrinhNangLuong_A",
    "CongTrinhNangLuong_P",
    "CongTrinhNgam_A",
    "CongTrinhNgam_L",
    "CongTrinhNgam_P",
    "CongTrinhTNTvaVSMT_A",
    "CongTrinhTNTvaVSMT_L",
    "CongTrinhTNTvaVSMT_P",
    "CongTrinhThongTin_A",
    "CongTrinhThongTin_P",
    "CongTrinh_A",
    "CongTrinh_L",
    "CongtrinhCapNuocPCCC_A",
    "CongtrinhCapNuocPCCC_P",
    "DanhGiaMoiTruong_A",
    "DanhGiaMoiTruong_L",
    "DanhGiaMoiTruong_P",
    "DiemDauNoi_P",
    "DiemNhanChinh_P",
    "DiemQuanTrac_P",
    "DiemToaDoTimDuongChuyenHuongTimDuong_P",
    "DongMucThietKe_L",
    "DuAnLienQuan_A",
    "GiaiPhapBaoVeMoiTruong_A",
    "GiaiPhapBaoVeMoiTruong_L",
    "GiaiPhapBaoVeMoiTruong_P",
    "HanhLangAnToan_L",
    "HuongDi_L",
    "HuongThoatNuocMua_L",
    "HuongThoatNuocThai_L",
    "KhongGianKTCQ_A",
    "KhongGianKTCQ_L",
    "KhuVucPhoiCanh_A",
    "MangLuoiCapNuoc_L",
    "MangLuoiCapThongTin_L",
    "MangLuoiChieuSang_L",
    "MangLuoiGiaoThongDuongBo_A",
    "MangLuoiGiaoThongDuongBo_L",
    "MangLuoiGiaoThongDuongKhong_L",
    "MangLuoiGiaoThongDuongSat_L",
    "MangLuoiGiaoThongDuongThuy_L",
    "MangLuoiNangLuong_L",
    "MangLuoiPhanPhoiDien_L",
    "MangLuoiThoatNuocMua_L",
    "MangLuoiThoatNuocThai_L",
    "MangLuoiTuyenBus_L",
    "MatCatNgang_L",
    "MatNuoc_A",
    "MocGioiQuyHoach_A",
    "MocGioiQuyHoach_L",
    "MocGioiQuyHoach_P",
    "NutTinhToanTNT_P",
    "PhanKhuQuyHoach_A",
    "PhanLuuThoatNuocMua_L",
    "PhanLuuThoatNuocThai_L",
    "PhanOQuyHoach_A",
    "PhanVungCapDien_A",
    "PhanVungCapNuoc_A",
    "PhanVungDanhGia_A",
    "PhanVungLuuVuc_A",
    "PhanVungPhucVu_A",
    "PhanVungSDDkhac_A",
    "PhanVungSanNen_A",
    "RanhGioiHanhChinh_L",
    "RanhGioiQuyHoach_A",
    "TenDonViHanhChinh_P",
    "ThongTinSanNen_P",
    "TuyenTKDT_L",
]


def _openpyxl_available():
    return Workbook is not None and load_workbook is not None


def sanitize_filename(value):
    text = str(value).strip()
    if not text:
        return "layer"
    for ch in '<>:"/\\|?*':
        text = text.replace(ch, "_")
    text = text.replace(" ", "_").replace("-", "_")
    return text.strip("._") or "layer"


def _excel_scalar(val):
    if val is None or val == QGIS_NULL:
        return None
    if isinstance(val, QVariant) and val.isNull():
        return None
    if isinstance(val, QDateTime):
        if val.isValid():
            return val.toPyDateTime()
        return None
    if isinstance(val, QDate):
        if val.isValid():
            return val.toPyDate()
        return None
    return val


def _write_layer_workbook(layer):
    wb = Workbook()
    ws = wb.active
    ws.title = "thuoc_tinh"
    fields = layer.fields()
    headers = [_FID_HEADER] + [fields.at(i).name() for i in range(fields.count())]
    ws.append(headers)
    for feat in layer.getFeatures():
        row = [feat.id()]
        for i in range(fields.count()):
            row.append(_excel_scalar(feat.attribute(i)))
        ws.append(row)
    return wb


def _save_workbook(wb, path):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)


def export_layers_attributes_excel(layers, output_path):
    if not _openpyxl_available():
        return "", "Thiếu thư viện openpyxl. Cài trong Python của QGIS: pip install openpyxl"
    if not layers:
        return "", "Không có lớp nào được chọn."
    n = len(layers)
    if n == 1:
        layer = layers[0]
        path = output_path.strip().rstrip("/\\")
        if not path:
            path = os.path.join(".", sanitize_filename(layer.name()) + ".xlsx")
        elif os.path.isdir(path):
            path = os.path.join(path, sanitize_filename(layer.name()) + ".xlsx")
        else:
            base_dir, fname = os.path.split(path)
            stem, ext = os.path.splitext(fname)
            if not stem:
                path = os.path.join(base_dir or ".", sanitize_filename(layer.name()) + ".xlsx")
        if not path.lower().endswith(".xlsx"):
            path = path + ".xlsx" if not path.endswith((".xlsx", ".zip")) else path
        if not path.lower().endswith(".xlsx"):
            path = os.path.splitext(path)[0] + ".xlsx"
        wb = _write_layer_workbook(layer)
        _save_workbook(wb, path)
        return path, None
    zip_path = output_path.strip().rstrip("/\\")
    if not zip_path:
        zip_path = os.path.join(".", f"{sanitize_filename(layers[0].name())}_{n}lop.zip")
    elif os.path.isdir(zip_path):
        zip_path = os.path.join(zip_path, f"{sanitize_filename(layers[0].name())}_{n}lop.zip")
    if not zip_path.lower().endswith(".zip"):
        zip_path = os.path.splitext(zip_path)[0] + ".zip"
    used_names = {}
    with tempfile.TemporaryDirectory(prefix="hsg_excel_") as tmp:
        files_to_zip = []
        for layer in layers:
            base = sanitize_filename(layer.name())
            k = used_names.get(base, 0)
            used_names[base] = k + 1
            fname = f"{base}.xlsx" if k == 0 else f"{base}_{k + 1}.xlsx"
            fpath = os.path.join(tmp, fname)
            wb = _write_layer_workbook(layer)
            _save_workbook(wb, fpath)
            files_to_zip.append((fpath, fname))
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for fpath, arcname in files_to_zip:
                zf.write(fpath, arcname)
    return zip_path, None


def _header_map(ws):
    row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if not row:
        return {}
    return {str(c).strip(): i for i, c in enumerate(row) if c is not None and str(c).strip()}


def _apply_sheet_to_layer(layer, ws):
    hmap = _header_map(ws)
    if _FID_HEADER not in hmap:
        return -1
    fid_col = hmap[_FID_HEADER]
    fields = layer.fields()
    field_names = [fields.at(i).name() for i in range(fields.count())]
    col_to_field_idx = {}
    for name in field_names:
        for hdr, cidx in hmap.items():
            if hdr == _FID_HEADER:
                continue
            if hdr == name or hdr == name[:10]:
                fi = fields.indexOf(name)
                if fi >= 0:
                    col_to_field_idx[cidx] = fi
                break
    layer.startEditing()
    pr = layer.dataProvider()
    updated = 0
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    for row in rows:
        if row is None or fid_col >= len(row):
            continue
        raw_fid = row[fid_col]
        if raw_fid is None or str(raw_fid).strip() == "":
            continue
        try:
            fid = int(raw_fid)
        except (TypeError, ValueError):
            continue
        if not layer.getFeature(fid).isValid():
            continue
        attrs = {}
        for cidx, fidx in col_to_field_idx.items():
            if cidx >= len(row):
                continue
            val = row[cidx]
            attrs[fidx] = val
        if attrs:
            pr.changeAttributeValues({fid: attrs})
            updated += 1
    if not layer.commitChanges():
        layer.rollBack()
        return -2
    return updated


def import_layers_attributes_excel(layers, input_path, log_print=print):
    if not _openpyxl_available():
        return "Thiếu thư viện openpyxl. Cài trong Python của QGIS: pip install openpyxl"
    if not layers:
        return "Không có lớp nào được chọn."
    if not input_path or not os.path.isfile(input_path):
        return "File không tồn tại."
    ext = os.path.splitext(input_path)[1].lower()
    if ext == ".zip":
        groups = defaultdict(list)
        for lyr in layers:
            groups[sanitize_filename(lyr.name())].append(lyr)
        with tempfile.TemporaryDirectory(prefix="hsg_excel_in_") as tmp:
            with zipfile.ZipFile(input_path, "r") as zf:
                zf.extractall(tmp)
            stem_to_path = {}
            for root, _, files in os.walk(tmp):
                for fn in files:
                    if not fn.lower().endswith(".xlsx"):
                        continue
                    stem = os.path.splitext(fn)[0]
                    stem_to_path[stem] = os.path.join(root, fn)
            matched = 0
            for base, lyr_list in groups.items():
                for i, layer in enumerate(lyr_list):
                    stem_key = base if i == 0 else f"{base}_{i + 1}"
                    fpath = stem_to_path.get(stem_key)
                    if not fpath:
                        log_print(
                            f"  − Trong ZIP không có file tương ứng lớp «{layer.name()}» (mong đợi: {stem_key}.xlsx)."
                        )
                        continue
                    wb = load_workbook(fpath, read_only=True, data_only=True)
                    try:
                        ws = wb.active
                        n = _apply_sheet_to_layer(layer, ws)
                        if n == -2:
                            log_print(
                                f"  − Không lưu được thay đổi cho lớp {layer.name()} (từ {os.path.basename(fpath)})."
                            )
                        elif n < 0:
                            log_print(f"  − File {os.path.basename(fpath)} thiếu cột bắt buộc '{_FID_HEADER}'.")
                        else:
                            matched += 1
                            log_print(
                                f"  + Đã cập nhật {n} đối tượng từ {os.path.basename(fpath)} → {layer.name()}"
                            )
                    finally:
                        wb.close()
            if matched == 0:
                return "Không khớp được file .xlsx nào trong ZIP với các lớp đã chọn (tên file = tên lớp đã làm sạch)."
        return None
    if ext != ".xlsx":
        return "Chỉ hỗ trợ file .xlsx hoặc .zip."
    wb = load_workbook(input_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        if len(layers) == 1:
            n = _apply_sheet_to_layer(layers[0], ws)
            if n == -2:
                return f"Không lưu được thay đổi cho lớp {layers[0].name()}."
            if n < 0:
                return f"File thiếu cột bắt buộc '{_FID_HEADER}'."
            log_print(f"  + Đã cập nhật {n} đối tượng cho lớp {layers[0].name()}.")
            return None
        stem = sanitize_filename(os.path.splitext(os.path.basename(input_path))[0])
        target = None
        for lyr in layers:
            if sanitize_filename(lyr.name()) == stem or lyr.name() == stem:
                target = lyr
                break
        if target is None:
            return (
                "Nhiều lớp được chọn: hãy dùng file ZIP từ bước xuất, hoặc chọn đúng một lớp, "
                "hoặc đặt tên file .xlsx trùng tên lớp (đã làm sạch)."
            )
        n = _apply_sheet_to_layer(target, ws)
        if n == -2:
            return f"Không lưu được thay đổi cho lớp {target.name()}."
        if n < 0:
            return f"File thiếu cột bắt buộc '{_FID_HEADER}'."
        log_print(f"  + Đã cập nhật {n} đối tượng cho lớp {target.name()}.")
        return None
    finally:
        wb.close()


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

try:
    ALIGN_VCENTER = Qt.AlignmentFlag.AlignVCenter
except AttributeError:
    ALIGN_VCENTER = Qt.AlignVCenter

# Thương hiệu / bản quyền (hiển thị trên giao diện)
_DEV_BRAND = "LEDAT"
_COPYRIGHT_YEAR = "2026"


class HoSoGISWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("HoSoGIS")
        self.resize(1120, 760)
        self.setStyleSheet(self._build_style())
        self._build_ui()
        self.refresh_vector_layers()

    def _build_style(self):
        return """
        QMainWindow { background-color: #f1f5f9; }
        QWidget {
            color: #334155;
            font-size: 12px;
        }
        QFrame#headerCard {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
        }
        QLabel#title {
            font-size: 26px;
            font-weight: 800;
            color: #0f172a;
            letter-spacing: 0.02em;
        }
        QLabel#titleKicker {
            color: #2563eb;
            font-size: 10px;
            font-weight: 700;
            letter-spacing: 0.12em;
            text-transform: uppercase;
        }
        QLabel#subtitle { color: #64748b; font-size: 12px; }
        QFrame#footerBar {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            min-height: 40px;
        }
        QLabel#footerBrand {
            color: #2563eb;
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 0.08em;
        }
        QLabel#footerCopyright {
            color: #64748b;
            font-size: 11px;
            font-weight: 500;
        }
        QLabel#footerDot {
            color: #94a3b8;
            font-size: 11px;
            padding: 0 6px;
        }
        QGroupBox {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            margin-top: 10px;
            font-weight: 600;
            color: #0f172a;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px;
            color: #0f172a;
        }
        QLineEdit, QListWidget, QTextEdit {
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 6px 8px;
            background: #ffffff;
            selection-background-color: #bfdbfe;
            selection-color: #0f172a;
        }
        QLineEdit:focus, QListWidget:focus, QTextEdit:focus {
            border: 1px solid #2563eb;
            background: #ffffff;
        }
        QLineEdit::placeholder {
            color: #94a3b8;
        }
        QPushButton {
            background-color: #2563eb;
            color: white;
            border: none;
            border-radius: 8px;
            padding: 7px 12px;
            font-weight: 600;
        }
        QPushButton:hover { background-color: #1d4ed8; }
        QPushButton:pressed { background-color: #1e40af; }
        QPushButton:disabled {
            background-color: #cbd5e1;
            color: #f8fafc;
        }
        QPushButton#ghost {
            background-color: #ffffff;
            color: #334155;
            border: 1px solid #cbd5e1;
        }
        QPushButton#ghost:hover {
            background-color: #f8fafc;
            border: 1px solid #94a3b8;
        }
        QTabWidget#moduleTabs::pane {
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            background: #ffffff;
            margin-top: 8px;
            top: 0px;
        }
        QTabWidget#moduleTabs QTabBar {
            qproperty-drawBase: 0;
        }
        QTabWidget#moduleTabs QTabBar::tab {
            background: #e2e8f0;
            color: #475569;
            border: 1px solid transparent;
            border-radius: 8px;
            min-width: 122px;
            padding: 8px 14px;
            margin-right: 6px;
        }
        QTabWidget#moduleTabs QTabBar::tab:hover {
            background: #dbe4ef;
            color: #1e293b;
        }
        QTabWidget#moduleTabs QTabBar::tab:selected {
            background: #ffffff;
            color: #0f172a;
            border: 1px solid #cbd5e1;
            font-weight: 600;
        }
        QTabWidget#moduleTabs QTabBar::tab:!selected {
            margin-top: 2px;
        }
        QLabel#sectionHint {
            color: #64748b;
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 6px 8px;
        }
        QCheckBox {
            spacing: 7px;
            color: #475569;
        }
        QCheckBox::indicator {
            width: 16px;
            height: 16px;
        }
        QSplitter::handle {
            background: #e2e8f0;
            width: 4px;
            border-radius: 2px;
            margin: 6px 0;
        }
        QTextEdit#logView {
            background: #ffffff;
        }
        QMessageBox {
            background-color: #ffffff;
        }
        QMessageBox QLabel {
            color: #334155;
            font-size: 12px;
        }
        QMessageBox QPushButton {
            min-width: 88px;
            padding: 7px 12px;
            border-radius: 8px;
        }
        """

    def _build_ui(self):
        container = QWidget()
        self.setCentralWidget(container)

        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(8)

        header = QFrame()
        header.setObjectName("headerCard")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(14, 12, 14, 12)
        header_layout.setSpacing(1)
        lbl_kicker = QLabel("PLANNING DATA WORKSPACE")
        lbl_kicker.setObjectName("titleKicker")
        lbl_title = QLabel('HoSo<span style="color:#2563eb;">GIS</span>')
        lbl_title.setObjectName("title")
        lbl_sub = QLabel("Quản lý dữ liệu quy hoạch: nhập CAD, cập nhật thuộc tính, xuất GDB/GPKG")
        lbl_sub.setObjectName("subtitle")
        header_layout.addWidget(lbl_kicker)
        header_layout.addWidget(lbl_title)
        header_layout.addWidget(lbl_sub)
        root_layout.addWidget(header)

        left_panel = self._build_left_panel()
        root_layout.addWidget(left_panel, 1)

        footer = QFrame()
        footer.setObjectName("footerBar")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(16, 10, 16, 10)
        footer_layout.setSpacing(0)
        lbl_footer_brand = QLabel(_DEV_BRAND)
        lbl_footer_brand.setObjectName("footerBrand")
        lbl_dot = QLabel("·")
        lbl_dot.setObjectName("footerDot")
        lbl_footer_copy = QLabel(
            f"Bản quyền phần mềm HoSoGIS © {_COPYRIGHT_YEAR} {_DEV_BRAND}. "
            "Không sao chép, phân phối hoặc sửa đổi mà không được phép."
        )
        lbl_footer_copy.setObjectName("footerCopyright")
        lbl_footer_copy.setWordWrap(True)
        footer_layout.addWidget(lbl_footer_brand, 0, ALIGN_VCENTER)
        footer_layout.addWidget(lbl_dot, 0, ALIGN_VCENTER)
        footer_layout.addWidget(lbl_footer_copy, 1, ALIGN_VCENTER)
        root_layout.addWidget(footer)

    def _build_left_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        tabs = QTabWidget()
        tabs.setObjectName("moduleTabs")
        tabs.addTab(self._build_tab_import_cad(), "Nhập CAD")
        tabs.addTab(self._build_tab_attributes(), "Thuộc tính")
        tabs.addTab(self._build_tab_quick_rename(), "Đổi tên layer nhanh")
        tabs.addTab(self._build_tab_export(), "Xuất dữ liệu")
        tabs.addTab(self._build_tab_logs(), "Nhật ký xử lý")
        layout.addWidget(tabs, 1)
        return panel

    def _build_tab_logs(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(4, 4, 4, 4)
        tab_layout.setSpacing(8)

        hint = QLabel("Theo dõi toàn bộ thông báo xử lý và trạng thái thao tác.")
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        tab_layout.addWidget(hint)

        log_box = QGroupBox("Nhật ký xử lý")
        log_layout = QVBoxLayout(log_box)
        self.log_edit = QTextEdit()
        self.log_edit.setObjectName("logView")
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)
        tab_layout.addWidget(log_box, 1)
        return tab

    def _build_tab_import_cad(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(4, 4, 4, 4)
        tab_layout.setSpacing(8)

        hint = QLabel("Nhập file CAD (ưu tiên DXF) và tách lớp theo hình học.")
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        tab_layout.addWidget(hint)

        box = QGroupBox("Nguồn dữ liệu CAD")
        box_layout = QVBoxLayout(box)
        row = QHBoxLayout()
        self.input_cad_path = QLineEdit()
        self.input_cad_path.setPlaceholderText("Chọn file CAD (khuyến nghị DXF)")
        btn_browse = QPushButton("Chọn file")
        btn_browse.setObjectName("ghost")
        btn_browse.clicked.connect(self.choose_cad_file)
        row.addWidget(self.input_cad_path, 1)
        row.addWidget(btn_browse)
        box_layout.addLayout(row)

        btn_run = QPushButton("Chạy nhập bản vẽ")
        btn_run.clicked.connect(self.import_and_split_cad)
        box_layout.addWidget(btn_run)
        tab_layout.addWidget(box)
        tab_layout.addStretch(1)
        return tab

    def _build_tab_attributes(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(4, 4, 4, 4)
        tab_layout.setSpacing(8)

        hint = QLabel("Cập nhật trường dữ liệu chuẩn và đồng bộ thuộc tính theo layer được chọn.")
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        tab_layout.addWidget(hint)

        form_box = QGroupBox("Thông tin thuộc tính")
        form_layout = QVBoxLayout(form_box)
        form_grid = QFormLayout()
        form_grid.setLabelAlignment(ALIGN_VCENTER)
        form_grid.setHorizontalSpacing(10)
        form_grid.setVerticalSpacing(6)

        self.input_ma_tt = QLineEdit()
        self.input_ma_hs = QLineEdit()
        self.input_ma_dt = QLineEdit()
        self.input_ten_dt = QLineEdit()
        self.input_phan_loai = QLineEdit()
        self.input_ghi_chu = QLineEdit()

        self._add_labeled_input(form_grid, "Mã thông tin QH", self.input_ma_tt)
        self._add_labeled_input(form_grid, "Mã hồ sơ QH", self.input_ma_hs)
        self._add_labeled_input(form_grid, "Mã đối tượng", self.input_ma_dt)
        self._add_labeled_input(form_grid, "Tên đối tượng", self.input_ten_dt)
        self._add_labeled_input(form_grid, "Phân loại", self.input_phan_loai)
        self._add_labeled_input(form_grid, "Ghi chú", self.input_ghi_chu)
        form_layout.addLayout(form_grid)

        self.chk_delete_old = QCheckBox("Xóa thuộc tính cũ trước khi cập nhật")
        form_layout.addWidget(self.chk_delete_old)
        tab_layout.addWidget(form_box)

        layer_box = QGroupBox("Layer áp dụng")
        layer_layout = QVBoxLayout(layer_box)
        self.list_layers = QListWidget()
        layer_layout.addWidget(self.list_layers)

        btn_row = QHBoxLayout()
        btn_select_all = QPushButton("Chọn tất cả")
        btn_select_all.setObjectName("ghost")
        btn_unselect_all = QPushButton("Bỏ chọn tất cả")
        btn_unselect_all.setObjectName("ghost")
        btn_refresh = QPushButton("Làm mới")
        btn_refresh.setObjectName("ghost")
        btn_select_all.clicked.connect(self.select_all_layers)
        btn_unselect_all.clicked.connect(self.unselect_all_layers)
        btn_refresh.clicked.connect(self.refresh_vector_layers)
        btn_row.addWidget(btn_select_all)
        btn_row.addWidget(btn_unselect_all)
        btn_row.addWidget(btn_refresh)
        layer_layout.addLayout(btn_row)

        excel_row = QHBoxLayout()
        btn_excel_export = QPushButton("Xuất Excel…")
        btn_excel_import = QPushButton("Nhập Excel…")
        btn_excel_export.setObjectName("ghost")
        btn_excel_import.setObjectName("ghost")
        btn_excel_export.clicked.connect(self.export_attributes_excel)
        btn_excel_import.clicked.connect(self.import_attributes_excel)
        excel_row.addWidget(btn_excel_export)
        excel_row.addWidget(btn_excel_import)
        layer_layout.addLayout(excel_row)

        tab_layout.addWidget(layer_box, 1)

        btn_run = QPushButton("Chạy cập nhật thuộc tính")
        btn_run.clicked.connect(self.add_fields_and_data)
        tab_layout.addWidget(btn_run)
        return tab

    def _build_tab_quick_rename(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(4, 4, 4, 4)
        tab_layout.setSpacing(8)

        hint = QLabel(
            "Đổi tên nhanh layer theo danh mục chuẩn. "
            "Chọn tên cũ ở bên trái, rồi chọn tên mới phù hợp (_P/_L/_A) ở bên phải."
        )
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        tab_layout.addWidget(hint)

        box = QGroupBox("Đổi tên nhanh")
        box_layout = QVBoxLayout(box)

        split = QHBoxLayout()
        old_box = QGroupBox("Tên cũ (layer hiện tại)")
        old_layout = QVBoxLayout(old_box)
        self.rename_old_list = QListWidget()
        self.rename_old_list.currentItemChanged.connect(self.on_rename_old_layer_changed)
        old_layout.addWidget(self.rename_old_list)

        new_box = QGroupBox("Tên mới (chọn từ danh mục)")
        new_layout = QVBoxLayout(new_box)
        self.rename_new_list = QListWidget()
        new_layout.addWidget(self.rename_new_list)

        split.addWidget(old_box, 1)
        split.addWidget(new_box, 1)
        box_layout.addLayout(split)

        btn_row = QHBoxLayout()
        btn_refresh = QPushButton("Làm mới danh sách")
        btn_refresh.setObjectName("ghost")
        btn_apply = QPushButton("Đổi tên layer đã chọn")
        btn_refresh.clicked.connect(self.refresh_rename_layers)
        btn_apply.clicked.connect(self.apply_selected_rename)
        btn_row.addWidget(btn_refresh)
        btn_row.addWidget(btn_apply)
        box_layout.addLayout(btn_row)

        tab_layout.addWidget(box, 1)
        return tab

    def _build_tab_export(self):
        tab = QWidget()
        tab_layout = QVBoxLayout(tab)
        tab_layout.setContentsMargins(4, 4, 4, 4)
        tab_layout.setSpacing(8)

        hint = QLabel("Xuất/nhập dữ liệu GDB, GPKG theo cấu trúc nhóm nghiệp vụ.")
        hint.setObjectName("sectionHint")
        hint.setWordWrap(True)
        tab_layout.addWidget(hint)

        import_box = QGroupBox("Nhập dữ liệu GDB")
        import_layout = QVBoxLayout(import_box)
        import_row = QHBoxLayout()
        self.input_gdb_path = QLineEdit()
        self.input_gdb_path.setPlaceholderText("Chọn thư mục File Geodatabase (*.gdb)")
        btn_browse_gdb = QPushButton("Chọn GDB")
        btn_browse_gdb.setObjectName("ghost")
        btn_browse_gdb.clicked.connect(self.choose_gdb_folder)
        import_row.addWidget(self.input_gdb_path, 1)
        import_row.addWidget(btn_browse_gdb)
        import_layout.addLayout(import_row)

        btn_import_gdb = QPushButton("Chạy nhập GDB")
        btn_import_gdb.clicked.connect(self.import_from_gdb)
        import_layout.addWidget(btn_import_gdb)
        tab_layout.addWidget(import_box)

        box = QGroupBox("Xuất dữ liệu")
        box_layout = QVBoxLayout(box)
        btn_row_export = QHBoxLayout()
        btn_run_gdb = QPushButton("Xuất GDB")
        btn_run_gdb.clicked.connect(self.export_to_gdb)
        btn_run_gpkg = QPushButton("Xuất GPKG")
        btn_run_gpkg.clicked.connect(self.export_to_gpkg)
        btn_row_export.addWidget(btn_run_gdb)
        btn_row_export.addWidget(btn_run_gpkg)
        box_layout.addLayout(btn_row_export)
        tab_layout.addWidget(box)
        tab_layout.addStretch(1)
        return tab

    def choose_gdb_folder(self):
        gdb_path = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục File Geodatabase (.gdb)",
            "",
        )
        if gdb_path:
            self.input_gdb_path.setText(gdb_path)

    def _add_labeled_input(self, layout, text, widget):
        if hasattr(layout, "addRow"):
            layout.addRow(f"{text}:", widget)
        else:
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
            f"Quá trình đã hoàn tất.\n\nHoSoGIS · Phát triển & bản quyền © {_COPYRIGHT_YEAR} {_DEV_BRAND}.\nCảm ơn đã sử dụng.",
        )

    def _question_yes_no(self, title, text, default_no=True):
        try:
            yes = QMessageBox.StandardButton.Yes
            no = QMessageBox.StandardButton.No
            mask = yes | no
            default_btn = no if default_no else yes
            reply = QMessageBox.question(self, title, text, mask, default_btn)
            return reply == yes
        except AttributeError:
            reply = QMessageBox.question(
                self,
                title,
                text,
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No if default_no else QMessageBox.Yes,
            )
            return reply == QMessageBox.Yes

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
        self.refresh_rename_layers()
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

    def _geometry_suffix_for_layer(self, layer):
        if layer.geometryType() == QgsWkbTypes.PointGeometry:
            return "_P"
        if layer.geometryType() == QgsWkbTypes.LineGeometry:
            return "_L"
        if layer.geometryType() == QgsWkbTypes.PolygonGeometry:
            return "_A"
        return ""

    def refresh_rename_layers(self):
        if not hasattr(self, "rename_old_list"):
            return
        self.rename_old_list.clear()
        self.rename_new_list.clear()
        all_layers = QgsProject.instance().mapLayers().values()
        vector_layers = [lyr for lyr in all_layers if lyr.type() == QgsMapLayerType.VectorLayer]
        for layer in vector_layers:
            item = QListWidgetItem(layer.name())
            item.setData(USER_ROLE, layer.id())
            self.rename_old_list.addItem(item)
        if self.rename_old_list.count() > 0:
            self.rename_old_list.setCurrentRow(0)

    def on_rename_old_layer_changed(self, current, previous):
        del previous  # tránh cảnh báo biến không dùng
        if not hasattr(self, "rename_new_list"):
            return
        self.rename_new_list.clear()
        if current is None:
            return

        layer_id = current.data(USER_ROLE)
        layer = QgsProject.instance().mapLayers().get(layer_id)
        if not layer:
            return
        suffix = self._geometry_suffix_for_layer(layer)
        if not suffix:
            return

        options = [name for name in _QUICK_LAYER_NAME_OPTIONS if name.endswith(suffix)]
        for name in options:
            self.rename_new_list.addItem(QListWidgetItem(name))

    def apply_selected_rename(self):
        if not hasattr(self, "rename_old_list") or not hasattr(self, "rename_new_list"):
            return
        old_item = self.rename_old_list.currentItem()
        new_item = self.rename_new_list.currentItem()
        if old_item is None:
            self.log("Chưa chọn layer ở cột Tên cũ.")
            return
        if new_item is None:
            self.log("Chưa chọn tên mới ở cột Tên mới.")
            return

        layer_id = old_item.data(USER_ROLE)
        layer = QgsProject.instance().mapLayers().get(layer_id)
        if not layer:
            self.log("Layer đã bị xóa khỏi dự án, vui lòng làm mới danh sách.")
            return

        new_name = new_item.text().strip()
        if not new_name:
            self.log("Tên mới không hợp lệ.")
            return
        old_name = layer.name()
        if old_name == new_name:
            self.log(f"Layer «{old_name}» đã có sẵn tên này.")
            return

        # Cảnh báo khi tên mới đã tồn tại ở layer khác trong project
        duplicated_layers = []
        for lyr in QgsProject.instance().mapLayers().values():
            if lyr.id() == layer.id():
                continue
            if lyr.name() == new_name:
                duplicated_layers.append(lyr)
        if duplicated_layers:
            confirm = self._question_yes_no(
                "Cảnh báo trùng tên layer",
                "Tên mới bạn chọn đang trùng với layer đã tồn tại:\n\n"
                f"«{new_name}»\n\n"
                "Việc trùng tên có thể gây khó phân biệt khi thao tác/xuất dữ liệu.\n"
                "Bạn vẫn muốn tiếp tục đổi tên?",
                default_no=True,
            )
            if not confirm:
                self.log(f"Đã hủy đổi tên do trùng tên layer: «{new_name}».")
                return

        layer.setName(new_name)
        old_item.setText(new_name)
        for i in range(self.list_layers.count()):
            item_attr = self.list_layers.item(i)
            if item_attr.data(USER_ROLE) == layer.id():
                item_attr.setText(new_name)
                break
        self.log(f"Đã đổi tên layer: «{old_name}» -> «{new_name}».")

    def export_attributes_excel(self):
        selected = self._selected_vector_layers()
        if not selected:
            self.log("Bạn chưa chọn lớp vector nào.")
            return
        home = os.path.expanduser("~")
        if len(selected) == 1:
            default_name = sanitize_filename(selected[0].name()) + ".xlsx"
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Lưu thuộc tính ra Excel",
                os.path.join(home, default_name),
                "Excel (*.xlsx)",
            )
        else:
            n = len(selected)
            default_zip = sanitize_filename(selected[0].name()) + f"_{n}lop.zip"
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Lưu thuộc tính (ZIP chứa nhiều file Excel)",
                os.path.join(home, default_zip),
                "ZIP (*.zip)",
            )
        if not path:
            self.log("Đã hủy xuất Excel.")
            return
        self.log(f"Đang xuất thuộc tính của {len(selected)} lớp…")
        out_path, err = export_layers_attributes_excel(selected, path)
        if err:
            self.log(f"Lỗi xuất Excel: {err}")
            return
        self.log(f"Đã lưu: {out_path}")
        self.show_done_message()

    def import_attributes_excel(self):
        selected = self._selected_vector_layers()
        if not selected:
            self.log("Bạn chưa chọn lớp vector nào.")
            return
        if len(selected) > 1:
            QMessageBox.warning(
                self,
                "Nhập thuộc tính — chỉ một lớp",
                "Nhập dữ liệu từ Excel chỉ thực hiện được cho một lớp mỗi lần.\n\n"
                f"Hiện bạn đang chọn {len(selected)} lớp. Hãy dùng «Bỏ chọn tất cả» rồi chỉ đánh dấu đúng một lớp cần nhập.",
            )
            self.log(f"Đã dừng: nhập Excel cần đúng 1 lớp (đang chọn {len(selected)} lớp).")
            return

        layer = selected[0]
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Chọn file Excel để nhập thuộc tính",
            "",
            "Excel (*.xlsx);;ZIP (*.zip)",
        )
        if not path:
            self.log("Đã hủy nhập Excel.")
            return
        if not self._question_yes_no(
            "Xác nhận nhập từ Excel",
            "Sẽ ghi dữ liệu từ file:\n\n"
            f"«{os.path.basename(path)}»\n\n"
            "vào các đối tượng của lớp (khớp theo cột «qgis_fid»):\n\n"
            f"«{layer.name()}»\n\n"
            "Tiếp tục nhập?",
        ):
            self.log("Đã hủy nhập Excel (chưa xác nhận sau khi chọn file).")
            return
        self.log(f"Đang nhập thuộc tính từ «{os.path.basename(path)}» vào lớp «{layer.name()}»…")
        err = import_layers_attributes_excel([layer], path, log_print=self.log)
        if err:
            self.log(f"Lỗi nhập Excel: {err}")
            return
        self.log("--- HOÀN TẤT nhập thuộc tính từ Excel ---")
        self.show_done_message()

    def _get_field_idx(self, layer, field_name):
        idx = layer.fields().indexOf(field_name)
        if idx == -1:
            idx = layer.fields().indexOf(field_name[:10])
        return idx

    def _sanitize_name(self, value):
        text = str(value).strip()
        if not text:
            return "unnamed"
        invalid_chars = '<>:"/\\|?*'
        for ch in invalid_chars:
            text = text.replace(ch, "_")
        text = text.replace(" ", "_").replace("-", "_")
        return text.strip("._") or "unnamed"

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
        # Hậu tố tên lớp theo loại hình học (bước 1 — nhập CAD)
        _cad_geom_suffix = {"Point": "_P", "LineString": "_L", "Polygon": "_A"}

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
                layer_name = clean_name + _cad_geom_suffix.get(geom_type, "")
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
                        safe_layer_name = self._sanitize_name(layer.name())
                        options = QgsVectorFileWriter.SaveVectorOptions()
                        options.driverName = "OpenFileGDB"
                        options.layerName = safe_layer_name
                        if not gdb_exists:
                            options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteFile
                        else:
                            options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteLayer
                        options.layerOptions = [f"FEATURE_DATASET={self._sanitize_name(group_name)}"]

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

    def export_to_gpkg(self):
        base_dir = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục lưu dữ liệu GPKG",
            "",
        )
        if not base_dir:
            self.log("Đã hủy chọn thư mục lưu.")
            return

        folder_name, ok = QInputDialog.getText(
            self,
            "Tên thư mục tổng",
            "Nhập tên thư mục tổng để chứa dữ liệu xuất GPKG:",
            text="export_gpkg",
        )
        if not ok:
            self.log("Đã hủy nhập tên thư mục tổng.")
            return

        folder_name = self._sanitize_name(folder_name)
        output_root = os.path.join(base_dir, folder_name)
        os.makedirs(output_root, exist_ok=True)
        self.log(f"Bắt đầu xuất GPKG vào thư mục: {output_root}")

        exported_layers = 0

        def process_group(group):
            nonlocal exported_layers
            group_name = group.name()
            if group_name.lower() == "base map":
                return

            safe_group_name = self._sanitize_name(group_name)
            group_dir = os.path.join(output_root, safe_group_name)
            os.makedirs(group_dir, exist_ok=True)
            self.log(f"Đang quét nhóm: {group_name}")

            for child in group.children():
                if isinstance(child, QgsLayerTreeGroup):
                    process_group(child)
                elif isinstance(child, QgsLayerTreeLayer):
                    layer = child.layer()
                    if layer and layer.type() == QgsMapLayerType.VectorLayer:
                        safe_layer_name = self._sanitize_name(layer.name())
                        gpkg_path = os.path.join(group_dir, f"{safe_layer_name}.gpkg")

                        options = QgsVectorFileWriter.SaveVectorOptions()
                        options.driverName = "GPKG"
                        options.layerName = safe_layer_name
                        options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteFile

                        error = QgsVectorFileWriter.writeAsVectorFormatV3(
                            layer,
                            gpkg_path,
                            QgsProject.instance().transformContext(),
                            options,
                        )
                        if error[0] == QgsVectorFileWriter.NoError:
                            self.log(f"  + Đã xuất: {layer.name()} -> {gpkg_path}")
                            exported_layers += 1
                        else:
                            self.log(f"  - Lỗi khi xuất {layer.name()}: {error}")

        root = QgsProject.instance().layerTreeRoot()
        for child in root.children():
            if isinstance(child, QgsLayerTreeGroup):
                process_group(child)

        if exported_layers > 0:
            self.log(f"--- HOÀN TẤT --- Đã xuất {exported_layers} layer ra GPKG.")
            self.show_done_message()
        else:
            self.log("--- HOÀN TẤT --- Không có layer nào được xuất ra GPKG.")

    def _list_gdb_sublayers(self, gdb_path):
        probe = QgsVectorLayer(gdb_path, "__gdb_probe__", "ogr")
        if not probe.isValid() or not probe.dataProvider():
            return []
        sublayers = []
        for raw in probe.dataProvider().subLayers():
            text = str(raw)
            # OGR thường trả dạng: "<idx>!!::!!<layer_name>!!::!!<feature_count>..."
            parts = text.split("!!::!!")
            name = parts[1].strip() if len(parts) > 1 else text.strip()
            if name and name not in sublayers:
                sublayers.append(name)
        return sublayers

    def _geometry_group_name(self, geometry_type):
        if geometry_type == QgsWkbTypes.PointGeometry:
            return "Điểm (Point)"
        if geometry_type == QgsWkbTypes.LineGeometry:
            return "Đường (Line)"
        if geometry_type == QgsWkbTypes.PolygonGeometry:
            return "Vùng (Polygon)"
        return "Khác (Other)"

    def _unique_group_name(self, root, base_name):
        name = base_name
        idx = 2
        while root.findGroup(name) is not None:
            name = f"{base_name}_{idx}"
            idx += 1
        return name

    def import_from_gdb(self):
        gdb_path = self.input_gdb_path.text().strip() if hasattr(self, "input_gdb_path") else ""
        if not gdb_path:
            gdb_path = QFileDialog.getExistingDirectory(
                self,
                "Chọn thư mục File Geodatabase (.gdb)",
                "",
            )
        if not gdb_path:
            self.log("Đã hủy chọn thư mục GDB.")
            return
        if hasattr(self, "input_gdb_path"):
            self.input_gdb_path.setText(gdb_path)
        if not gdb_path.lower().endswith(".gdb"):
            self.log("Đường dẫn đã chọn không phải thư mục .gdb.")
            return

        sublayers = self._list_gdb_sublayers(gdb_path)
        if not sublayers:
            self.log("Không đọc được layer nào từ GDB (hoặc GDB rỗng).")
            return

        root = QgsProject.instance().layerTreeRoot()
        gdb_name = os.path.splitext(os.path.basename(gdb_path))[0]
        group_name = self._unique_group_name(root, f"GDB_{self._sanitize_name(gdb_name)}")
        main_group = root.addGroup(group_name)
        subgroups = {}
        loaded = 0

        self.log(f"Bắt đầu nhập GDB: {gdb_path}")
        for sublayer_name in sublayers:
            uri = f"{gdb_path}|layername={sublayer_name}"
            layer = QgsVectorLayer(uri, sublayer_name, "ogr")
            if not layer.isValid():
                self.log(f"  - Bỏ qua (không hợp lệ): {sublayer_name}")
                continue
            if layer.type() != QgsMapLayerType.VectorLayer:
                self.log(f"  - Bỏ qua (không phải vector): {sublayer_name}")
                continue

            geom_group_name = self._geometry_group_name(layer.geometryType())
            if geom_group_name not in subgroups:
                subgroups[geom_group_name] = main_group.addGroup(geom_group_name)

            QgsProject.instance().addMapLayer(layer, False)
            subgroups[geom_group_name].addLayer(layer)
            loaded += 1
            self.log(f"  + Đã nhập: {sublayer_name} -> Nhóm: {geom_group_name}")

        if loaded > 0:
            self.log(f"HOÀN TẤT! Đã nhập {loaded} layer từ GDB và phân nhóm theo hình học.")
            self.refresh_vector_layers()
            self.show_done_message()
        else:
            self.log("KHÔNG THÀNH CÔNG! Không có layer hợp lệ nào được nhập từ GDB.")


def show_hosogis_gui():
    global HOSOGIS_WINDOW
    try:
        HOSOGIS_WINDOW.close()
    except Exception:
        pass

    HOSOGIS_WINDOW = HoSoGISWindow()
    HOSOGIS_WINDOW.show()
    HOSOGIS_WINDOW.raise_()
    HOSOGIS_WINDOW.activateWindow()


show_hosogis_gui()
