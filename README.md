# RTSP Face Attendance System

## Chức năng
- Stream video từ IP camera hoặc webcam
- Tách video thành frame realtime
- Pre-processing: grayscale, equalizeHist, Gaussian blur
- Face detection bằng Haar Cascade
- Face recognition bằng LBPH
- Lưu điểm danh ra file CSV theo ngày
- Giao diện Tkinter để demo trực tiếp

## Cài đặt
```bash
pip install -r requirements.txt
```

## Chạy
```bash
python attendance_app.py
```

## RTSP ví dụ
```text
rtsp://username:password@ip_address:554/stream
```

## Quy trình sử dụng
1. Mở camera hoặc RTSP stream.
2. Nhập MSSV + họ tên.
3. Thu dataset từ camera hoặc import ảnh từ thư mục.
4. Train model LBPH.
5. Bấm nút điểm danh realtime để ghi nhận sinh viên xuất hiện trong frame hiện tại.
6. File điểm danh nằm trong thư mục `attendance/`.

## Cấu trúc thư mục sinh ra
- `dataset/`: ảnh khuôn mặt theo từng sinh viên
- `model/lbph_model.yml`: model đã train
- `model/labels.csv`: ánh xạ label -> MSSV, tên
- `attendance/attendance_YYYY-MM-DD.csv`: log điểm danh
- `snapshots/`: ảnh chụp frame hiện tại

## Lưu ý
- Cần `opencv-contrib-python` để có `cv2.face.LBPHFaceRecognizer_create()`.
- Nếu dùng IP camera trên điện thoại, hãy bật RTSP server hoặc dùng app camera hỗ trợ RTSP.
- Để demo assignment, bạn có thể nhập `0` để dùng webcam laptop thay cho IP camera khi cần.
