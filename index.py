import win32com.client

# Tạo đối tượng AutoCAD
acad_app = win32com.client.Dispatch("AutoCAD.Application")

# Lấy tham chiếu đến bản vẽ đang hoạt động
doc = acad_app.ActiveDocument



