import os
from qgis.core import QgsProject, QgsVectorLayer, QgsFeatureRequest
from qgis.gui import QgsProjectionSelectionDialog
from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox

def _extract_dxf_layer_names(file_path: str):
    """
    Trả về list tên layer đầy đủ trong bảng LAYER của DXF (ASCII).
    DXF khi đọc qua OGR/QGIS đôi lúc bị cắt giá trị thuộc tính "Layer" (thường gặp: 10 ký tự hoặc rớt ký tự cuối),
    nên ta cần nguồn tên gốc để khôi phục.
    """
    try:
        with open(file_path, "rb") as f:
            head = f.read(32)
        # DXF nhị phân: khó parse nhanh bằng text -> bỏ qua
        if b"AutoCAD Binary DXF" in head:
            return {}

        # DXF ASCII: parse bảng LAYER theo cặp (group code, value)
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
                # nhìn trước để biết table gì
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

            # Trong entity LAYER: group code 2 là tên layer
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

    # 1) Nếu trùng đúng với tên trong DXF thì dùng luôn
    if cad_layer_str in dxf_full_layer_names:
        return cad_layer_str

    # 2) Nếu cad_layer_str là tiền tố của đúng 1 tên đầy đủ -> khôi phục
    matches = [n for n in dxf_full_layer_names if n.startswith(cad_layer_str)]
    if len(matches) == 1:
        return matches[0]

    # 3) Fallback: theo prefix 10 ký tự (chỉ khi duy nhất)
    key10 = cad_layer_str[:10]
    matches10 = [n for n in dxf_full_layer_names if n.startswith(key10)]
    if len(matches10) == 1:
        return matches10[0]

    return cad_layer_str

def import_and_split_cad():
    # 1. Mở hộp thoại chọn file CAD
    file_path, _ = QFileDialog.getOpenFileName(None, "Chọn file CAD (Nên dùng DXF)", "", "DXF Files (*.dxf);;CAD Files (*.dxf *.dwg)")
    
    if not file_path:
        print("Đã hủy chọn file.")
        return

    # 2. Chọn hệ tọa độ (CRS) cho dữ liệu CAD
    crs_dialog = QgsProjectionSelectionDialog()
    crs_dialog.setWindowTitle("Chọn hệ tọa độ cho bản vẽ CAD")
    if not crs_dialog.exec():
        print("Đã hủy chọn hệ tọa độ.")
        return
    selected_crs = crs_dialog.crs()

    file_name = os.path.basename(file_path)
    file_base_name = os.path.splitext(file_name)[0]
    
    print(f"Đang xử lý file: {file_name} với hệ tọa độ {selected_crs.authid()}...")

    dxf_full_layer_names = _extract_dxf_layer_names(file_path) if file_path.lower().endswith(".dxf") else []

    # Tạo Group cha mang tên file CAD
    root = QgsProject.instance().layerTreeRoot()
    cad_group = root.addGroup(f"CAD_{file_base_name}")

    # 3. Cấu hình kiểu hình học: {Loại: (Hậu tố, Tên Group con)}
    geometry_mapping = {
        'Point': ('', 'Điểm (Point)'),       
        'LineString': ('', 'Đường (Line)'),  
        'Polygon': ('', 'Vùng (Polygon)')      
    }
    _cad_geom_suffix = {'Point': '_P', 'LineString': '_L', 'Polygon': '_A'}

    layer_count = 0
    # Dictionary để lưu trữ các group con (Giúp code biết group nào đã tạo rồi để không tạo trùng)
    subgroups = {}

    # 4. Quét và tách lớp
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

            # Chuẩn hóa tên (bỏ khoảng trắng, dấu gạch ngang)
            clean_cad_layer = restored_layer.strip().replace(" ", "_").replace("-", "_")
            layer_name = clean_cad_layer + _cad_geom_suffix.get(geom_type, '')

            new_layer = QgsVectorLayer(uri, layer_name, "ogr")
            
            # Lọc dữ liệu theo tên lớp bên CAD
            safe_cad_layer = cad_layer_str.replace("'", "''")
            new_layer.setSubsetString(f"\"Layer\" = '{safe_cad_layer}'")
            
            # Nếu layer lọc ra có dữ liệu thực tế
            if new_layer.isValid() and new_layer.featureCount() > 0:
                
                # --- NÂNG CẤP TẠI ĐÂY: CHUYỂN THÀNH LỚP CÓ THỂ CHỈNH SỬA ---
                # Clone layer "chỉ đọc" thành layer nháp trên RAM (Memory layer)
                editable_layer = new_layer.materialize(QgsFeatureRequest())
                editable_layer.setName(layer_name)
                editable_layer.setCrs(selected_crs)
                
                # Đưa layer đã bẻ khóa vào dự án thay vì layer gốc
                QgsProject.instance().addMapLayer(editable_layer, False) 
                
                # Kiểm tra và tạo Group con nếu chưa tồn tại
                if group_name not in subgroups:
                    subgroups[group_name] = cad_group.addGroup(group_name)
                    
                # Thả layer vào đúng Group con kiểu dữ liệu của nó
                subgroups[group_name].addLayer(editable_layer) 
                layer_count += 1
                print(f"  + Đã tách và mở khóa: {layer_name} ---> Đưa vào nhóm: {group_name}")

    if layer_count > 0:
        print(f"\n--- HOÀN TẤT! Đã import và phân loại thành {layer_count} lớp dữ liệu. ---")
        QMessageBox.information(
            None,
            "Hoàn tất",
            "Quá trình đã hoàn tất.\nPhần mềm được phát triển bởi LEDAT.\nCảm ơn đã sử dụng."
        )
    else:
        print("\n--- KHÔNG THÀNH CÔNG! ---")
        print("Không tìm thấy dữ liệu hợp lệ. Nếu bạn đang dùng file DWG, hãy mở AutoCAD và Save As sang định dạng DXF rồi chạy lại mã.")

# Chạy tool
import_and_split_cad()