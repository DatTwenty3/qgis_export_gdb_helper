import os
from qgis.core import QgsProject, QgsVectorLayer, QgsFeatureRequest
from qgis.gui import QgsProjectionSelectionDialog
from qgis.PyQt.QtWidgets import QFileDialog, QMessageBox

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

    # Tạo Group cha mang tên file CAD
    root = QgsProject.instance().layerTreeRoot()
    cad_group = root.addGroup(f"CAD_{file_base_name}")

    # 3. Cấu hình kiểu hình học: {Loại: (Hậu tố, Tên Group con)}
    geometry_mapping = {
        'Point': ('_P', 'Điểm (Point)'),       
        'LineString': ('_L', 'Đường (Line)'),  
        'Polygon': ('_A', 'Vùng (Polygon)')      
    }

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
                
            # Chuẩn hóa tên (bỏ khoảng trắng, dấu gạch ngang)
            clean_cad_layer = str(cad_layer).strip().replace(" ", "_").replace("-", "_")
            layer_name = f"{clean_cad_layer}{suffix}"
            
            new_layer = QgsVectorLayer(uri, layer_name, "ogr")
            
            # Lọc dữ liệu theo tên lớp bên CAD
            safe_cad_layer = str(cad_layer).replace("'", "''") 
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