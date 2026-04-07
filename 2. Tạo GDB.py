import os
from qgis.core import QgsProject, QgsVectorFileWriter, QgsMapLayerType, QgsLayerTreeGroup, QgsLayerTreeLayer
from qgis.PyQt.QtWidgets import QFileDialog

# 1. GỌI HỘP THOẠI CHỌN NƠI LƯU FILE
# Các tham số: Nơi gọi (None), Tiêu đề, Thư mục mặc định, Bộ lọc định dạng file
gdb_path, _ = QFileDialog.getSaveFileName(None, "Chọn nơi lưu và đặt tên file Geodatabase", "", "Esri File Geodatabase (*.gdb)")

# 2. KIỂM TRA NGƯỜI DÙNG CÓ BẤM HỦY (CANCEL) HAY KHÔNG
if not gdb_path:
    print("Đã hủy thao tác lưu file.")
else:
    # Tự động thêm đuôi .gdb nếu người dùng quên gõ
    if not gdb_path.endswith('.gdb'):
        gdb_path += '.gdb'
        
    print(f"Bắt đầu xuất dữ liệu ra: {gdb_path}")
    
    # Biến toàn cục kiểm tra file đã tồn tại chưa
    gdb_exists = os.path.exists(gdb_path)

    def process_group(group):
        global gdb_exists
        group_name = group.name()
        
        # Bỏ qua nhóm bản đồ nền để tránh lỗi
        if group_name == "base map":
            return
            
        print(f"Đang quét nhóm: {group_name}")
        
        for child in group.children():
            # 2.1 Nếu bên trong là một Nhóm con, tiếp tục chui sâu vào trong
            if isinstance(child, QgsLayerTreeGroup):
                process_group(child)
                
            # 2.2 Nếu bên trong là Layer, tiến hành xuất
            elif isinstance(child, QgsLayerTreeLayer):
                layer = child.layer()
                if layer and layer.type() == QgsMapLayerType.VectorLayer:
                    
                    # Chuẩn hóa tên: thay khoảng trắng và gạch ngang bằng gạch dưới
                    safe_layer_name = layer.name().replace(" ", "_").replace("-", "_")
                    
                    options = QgsVectorFileWriter.SaveVectorOptions()
                    options.driverName = "OpenFileGDB"
                    options.layerName = safe_layer_name
                    
                    # Xử lý tự động tạo file mới hoặc ghi đè
                    if not gdb_exists:
                        options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteFile
                    else:
                        options.actionOnExistingFile = QgsVectorFileWriter.CreateOrOverwriteLayer
                        
                    # Gán tên Feature Dataset
                    options.layerOptions = [f"FEATURE_DATASET={group_name}"]
                    
                    # Thực thi xuất
                    error = QgsVectorFileWriter.writeAsVectorFormatV3(
                        layer, 
                        gdb_path, 
                        QgsProject.instance().transformContext(), 
                        options
                    )
                    
                    if error[0] == QgsVectorFileWriter.NoError:
                        print(f"  + Đã xuất: {layer.name()} -> Nhóm GDB: {group_name}")
                        gdb_exists = True 
                    else:
                        print(f"  - Lỗi khi xuất {layer.name()}: {error}")

    # 3. THỰC THI QUÉT CÂY THƯ MỤC LỚP
    root = QgsProject.instance().layerTreeRoot()
    
    for child in root.children():
        if isinstance(child, QgsLayerTreeGroup):
            process_group(child)

    print("\n--- HOÀN TẤT ---")