import os
import csv
import cv2
import time
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
from datetime import datetime

APP_TITLE = "RTSP Face Attendance System"
WINDOW_SIZE = "1350x820"
CASCADE_PATH = cv2.data.haarcascades + "haarcascade_frontalface_default.xml"

DATASET_DIR = "dataset"
MODEL_DIR = "model"
MODEL_PATH = os.path.join(MODEL_DIR, "lbph_model.yml")
LABELS_PATH = os.path.join(MODEL_DIR, "labels.csv")
ATTENDANCE_DIR = "attendance"
SNAPSHOT_DIR = "snapshots"

SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".bmp"}


def ensure_dirs():
    os.makedirs(DATASET_DIR, exist_ok=True)
    os.makedirs(MODEL_DIR, exist_ok=True)
    os.makedirs(ATTENDANCE_DIR, exist_ok=True)
    os.makedirs(SNAPSHOT_DIR, exist_ok=True)


def preprocess_face(face_bgr):
    gray = cv2.cvtColor(face_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.equalizeHist(gray)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    return gray


class StudentDB:
    def __init__(self, labels_path=LABELS_PATH):
        self.labels_path = labels_path
        self.id_to_info = {}
        self._load()

    def _load(self):
        self.id_to_info = {}
        if not os.path.exists(self.labels_path):
            return

        try:
            with open(self.labels_path, "r", newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    label_id = int(row["label_id"])
                    self.id_to_info[label_id] = {
                        "student_id": row["student_id"],
                        "student_name": row["student_name"],
                    }
        except Exception:
            self.id_to_info = {}

    def save_all(self):
        with open(self.labels_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["label_id", "student_id", "student_name"]
            )
            writer.writeheader()
            for label_id, info in sorted(self.id_to_info.items()):
                writer.writerow({
                    "label_id": label_id,
                    "student_id": info["student_id"],
                    "student_name": info["student_name"],
                })

    def find_label_by_student_id(self, student_id):
        for label_id, info in self.id_to_info.items():
            if info["student_id"] == student_id:
                return label_id
        return None

    def upsert(self, student_id, student_name):
        existing = self.find_label_by_student_id(student_id)
        if existing is not None:
            self.id_to_info[existing] = {
                "student_id": student_id,
                "student_name": student_name,
            }
            self.save_all()
            return existing

        new_id = 0 if not self.id_to_info else max(self.id_to_info.keys()) + 1
        self.id_to_info[new_id] = {
            "student_id": student_id,
            "student_name": student_name,
        }
        self.save_all()
        return new_id

    def get(self, label_id):
        return self.id_to_info.get(label_id)

    def all_students(self):
        rows = []
        for label_id, info in sorted(self.id_to_info.items()):
            rows.append((label_id, info["student_id"], info["student_name"]))
        return rows


class FaceAttendanceApp:
    def __init__(self, root):
        ensure_dirs()

        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.configure(bg="#f3f4f6")

        self.detector = cv2.CascadeClassifier(CASCADE_PATH)
        if self.detector.empty():
            raise RuntimeError("Không tải được Haar Cascade.")

        if not hasattr(cv2, "face") or not hasattr(cv2.face, "LBPHFaceRecognizer_create"):
            raise RuntimeError(
                "OpenCV hiện tại không có module cv2.face.\n"
                "Hãy cài: pip install opencv-contrib-python"
            )

        self.recognizer = cv2.face.LBPHFaceRecognizer_create()
        self.student_db = StudentDB()

        self.cap = None
        self.running = False
        self.current_frame = None
        self.current_faces = []

        self.last_mark_times = {}
        self.cooldown_seconds = 20

        self.auto_attendance = False
        self.flip_camera = False

        self.rtsp_var = tk.StringVar(value="0")
        self.capture_interval_var = tk.IntVar(value=15)
        self.student_id_var = tk.StringVar()
        self.student_name_var = tk.StringVar()
        self.samples_var = tk.IntVar(value=30)
        self.lbph_threshold_var = tk.DoubleVar(value=70.0)
        self.status_var = tk.StringVar(value="Sẵn sàng.")

        self.model_ready = os.path.exists(MODEL_PATH)
        if self.model_ready:
            self._safe_load_model()

        self.build_ui()
        self.refresh_student_table()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)

        # ===== LEFT WRAPPER WITH SCROLLBAR =====
        left_wrapper = ttk.Frame(main, width=400)
        left_wrapper.pack(side="left", fill="y", padx=(0, 10))
        left_wrapper.pack_propagate(False)

        self.left_canvas = tk.Canvas(left_wrapper, highlightthickness=0)
        self.left_canvas.pack(side="left", fill="both", expand=True)

        left_scrollbar = ttk.Scrollbar(left_wrapper, orient="vertical", command=self.left_canvas.yview)
        left_scrollbar.pack(side="right", fill="y")

        self.left_canvas.configure(yscrollcommand=left_scrollbar.set)

        self.left_inner = ttk.Frame(self.left_canvas)
        self.left_window = self.left_canvas.create_window((0, 0), window=self.left_inner, anchor="nw")

        self.left_inner.bind("<Configure>", self._on_left_frame_configure)
        self.left_canvas.bind("<Configure>", self._on_left_canvas_configure)

        # mouse wheel for left panel
        self.left_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # ===== RIGHT PANEL =====
        right = ttk.Frame(main)
        right.pack(side="right", fill="both", expand=True)

        # ===== LEFT CONTENT =====
        source_box = ttk.LabelFrame(self.left_inner, text="1) Nguồn camera / RTSP", padding=10)
        source_box.pack(fill="x", pady=(0, 10))

        ttk.Label(source_box, text="RTSP URL hoặc camera index (0, 1, ...)").pack(anchor="w")
        ttk.Entry(source_box, textvariable=self.rtsp_var).pack(fill="x", pady=5)

        btns = ttk.Frame(source_box)
        btns.pack(fill="x", pady=5)

        ttk.Button(btns, text="Mở camera", command=self.start_camera).pack(
            side="left", fill="x", expand=True, padx=(0, 4)
        )
        ttk.Button(btns, text="Dừng", command=self.stop_camera).pack(
            side="left", fill="x", expand=True, padx=(4, 4)
        )
        ttk.Button(btns, text="Lật camera", command=self.toggle_flip_camera).pack(
            side="left", fill="x", expand=True, padx=(4, 0)
        )

        ttk.Label(source_box, text="Chu kỳ auto capture khi thu dataset (mỗi N frame)").pack(anchor="w", pady=(8, 0))
        ttk.Spinbox(source_box, from_=1, to=120, textvariable=self.capture_interval_var, width=10).pack(anchor="w", pady=5)

        reg_box = ttk.LabelFrame(self.left_inner, text="2) Đăng ký sinh viên / thu dataset", padding=10)
        reg_box.pack(fill="x", pady=(0, 10))

        ttk.Label(reg_box, text="Mã số sinh viên").pack(anchor="w")
        ttk.Entry(reg_box, textvariable=self.student_id_var).pack(fill="x", pady=4)

        ttk.Label(reg_box, text="Họ tên").pack(anchor="w")
        ttk.Entry(reg_box, textvariable=self.student_name_var).pack(fill="x", pady=4)

        ttk.Label(reg_box, text="Số ảnh khuôn mặt cần chụp").pack(anchor="w")
        ttk.Spinbox(reg_box, from_=10, to=200, textvariable=self.samples_var, width=10).pack(anchor="w", pady=4)

        ttk.Button(reg_box, text="Thu dataset từ camera hiện tại", command=self.capture_dataset).pack(fill="x", pady=(8, 4))
        ttk.Button(reg_box, text="Nhập ảnh từ thư mục", command=self.import_images_from_folder).pack(fill="x")

        train_box = ttk.LabelFrame(self.left_inner, text="3) Huấn luyện & nhận diện", padding=10)
        train_box.pack(fill="x", pady=(0, 10))

        ttk.Button(train_box, text="Train model LBPH", command=self.train_model).pack(fill="x", pady=(0, 6))
        ttk.Button(train_box, text="Nạp lại model đã lưu", command=self.reload_model).pack(fill="x", pady=(0, 6))
        ttk.Button(train_box, text="Bật / Tắt điểm danh realtime", command=self.toggle_auto_attendance).pack(fill="x", pady=(0, 6))
        ttk.Button(train_box, text="Ghi attendance frame hiện tại", command=self.mark_attendance_current_frame).pack(fill="x", pady=(0, 6))
        ttk.Button(train_box, text="Lưu snapshot hiện tại", command=self.save_current_snapshot).pack(fill="x")

        ttk.Label(train_box, text="Ngưỡng LBPH (càng thấp càng chặt)").pack(anchor="w", pady=(10, 0))
        ttk.Scale(train_box, from_=30, to=120, variable=self.lbph_threshold_var, orient="horizontal").pack(fill="x")
        self.lbph_label = ttk.Label(train_box, text="Threshold hiện tại: 70.0")
        self.lbph_label.pack(anchor="w", pady=(4, 0))

        table_box = ttk.LabelFrame(self.left_inner, text="Sinh viên đã đăng ký", padding=10)
        table_box.pack(fill="both", expand=True, pady=(0, 10))

        self.student_tree = ttk.Treeview(table_box, columns=("label", "sid", "name"), show="headings", height=10)
        self.student_tree.heading("label", text="Label")
        self.student_tree.heading("sid", text="Student ID")
        self.student_tree.heading("name", text="Name")
        self.student_tree.column("label", width=55, anchor="center")
        self.student_tree.column("sid", width=95, anchor="center")
        self.student_tree.column("name", width=180)
        self.student_tree.pack(fill="both", expand=True)

        # ===== RIGHT CONTENT =====
        self.video_label = ttk.Label(right)
        self.video_label.pack(fill="both", expand=True)

        bottom = ttk.Frame(right)
        bottom.pack(fill="x", pady=(8, 0))

        self.info_text = tk.Text(bottom, height=9, wrap="word")
        self.info_text.pack(fill="x")
        self.info_text.insert("end", self.get_help_text())
        self.info_text.configure(state="disabled")

        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(side="bottom", fill="x")

        self.lbph_threshold_var.trace_add("write", self.update_threshold_label)

    def _on_left_frame_configure(self, event):
        self.left_canvas.configure(scrollregion=self.left_canvas.bbox("all"))

    def _on_left_canvas_configure(self, event):
        self.left_canvas.itemconfig(self.left_window, width=event.width)

    def _on_mousewheel(self, event):
        try:
            self.left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    def update_threshold_label(self, *args):
        self.lbph_label.config(text=f"Threshold hiện tại: {self.lbph_threshold_var.get():.1f}")

    def get_help_text(self):
        return (
            "Quy trình demo:\n"
            "1. Nhập RTSP URL hoặc để 0 để dùng webcam.\n"
            "2. Mở camera.\n"
            "3. Nhập MSSV + Họ tên, rồi thu dataset hoặc import ảnh.\n"
            "4. Train model LBPH hoặc nạp model cũ.\n"
            "5. Bật điểm danh realtime để hệ thống tự ghi attendance liên tục.\n"
            "6. Có thể bấm 'Ghi attendance frame hiện tại' nếu chỉ muốn ghi một lần.\n"
            "7. Dữ liệu attendance nằm trong thư mục attendance/.\n"
            "8. Nút 'Lật camera' sẽ lật ngang hình để dễ thao tác như camera trước."
        )

    def log_status(self, text):
        if hasattr(self, "status_var"):
            self.status_var.set(text)
            self.root.update_idletasks()

    def refresh_student_table(self):
        for item in self.student_tree.get_children():
            self.student_tree.delete(item)
        for row in self.student_db.all_students():
            self.student_tree.insert("", "end", values=row)

    def _open_capture(self, source_text):
        source_text = source_text.strip()
        source = int(source_text) if source_text.isdigit() else source_text

        cap = cv2.VideoCapture(source)
        if not cap.isOpened():
            return None
        return cap

    def start_camera(self):
        if self.running:
            self.log_status("Camera đang chạy.")
            return

        cap = self._open_capture(self.rtsp_var.get())
        if cap is None:
            messagebox.showerror("Lỗi", "Không mở được camera / RTSP stream.")
            return

        self.cap = cap
        self.running = True
        self.log_status("Đã mở camera thành công.")
        self.update_frame()

    def stop_camera(self):
        self.running = False
        if self.cap is not None:
            self.cap.release()
            self.cap = None
        self.video_label.configure(image="")
        self.log_status("Đã dừng camera.")

    def toggle_flip_camera(self):
        self.flip_camera = not self.flip_camera
        state = "BẬT" if self.flip_camera else "TẮT"
        self.log_status(f"Đã {state} lật camera.")

    def toggle_auto_attendance(self):
        if not self.running:
            messagebox.showwarning("Chưa mở camera", "Hãy mở camera trước.")
            return
        if not self.model_ready:
            messagebox.showwarning("Chưa có model", "Hãy train model hoặc nạp model trước.")
            return

        self.auto_attendance = not self.auto_attendance
        state = "BẬT" if self.auto_attendance else "TẮT"
        self.log_status(f"Đã {state} điểm danh realtime.")

    def update_frame(self):
        if not self.running or self.cap is None:
            return

        ok, frame = self.cap.read()
        if not ok or frame is None:
            self.log_status("Mất frame từ camera. Đang dừng.")
            self.stop_camera()
            return

        if self.flip_camera:
            frame = cv2.flip(frame, 1)

        self.current_frame = frame.copy()
        display = frame.copy()

        faces = self.detect_faces(display)
        self.current_faces = faces

        if self.auto_attendance:
            self.auto_mark_attendance_from_faces(faces)

        rgb = cv2.cvtColor(display, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(rgb)

        w = max(self.video_label.winfo_width(), 800)
        h = max(self.video_label.winfo_height(), 550)
        img.thumbnail((w, h))

        imgtk = ImageTk.PhotoImage(img)
        self.video_label.imgtk = imgtk
        self.video_label.configure(image=imgtk)

        self.root.after(15, self.update_frame)

    def detect_faces(self, display_frame):
        gray = cv2.cvtColor(display_frame, cv2.COLOR_BGR2GRAY)
        faces = self.detector.detectMultiScale(
            gray,
            scaleFactor=1.2,
            minNeighbors=5,
            minSize=(90, 90)
        )

        results = []

        for (x, y, w, h) in faces:
            face_bgr = display_frame[y:y+h, x:x+w]
            face_gray = preprocess_face(face_bgr)

            pred_text = "Face"
            color = (0, 255, 0)
            recognized_info = None
            lbph_conf = None

            if self.model_ready:
                try:
                    resized = cv2.resize(face_gray, (160, 160))
                    label, conf = self.recognizer.predict(resized)
                    lbph_conf = float(conf)
                    info = self.student_db.get(label)

                    if info and lbph_conf <= self.lbph_threshold_var.get():
                        pred_text = f"{info['student_id']} - {info['student_name']} (LBPH: {lbph_conf:.2f})"
                        color = (0, 200, 255)
                        recognized_info = info
                    else:
                        pred_text = f"Unknown (LBPH: {lbph_conf:.2f})"
                        color = (0, 0, 255)

                except Exception:
                    pred_text = "Unknown"
                    color = (0, 0, 255)
                    recognized_info = None
                    lbph_conf = None

            cv2.rectangle(display_frame, (x, y), (x + w, y + h), color, 2)
            cv2.putText(
                display_frame,
                pred_text,
                (x, max(30, y - 10)),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.65,
                color,
                2
            )

            results.append({
                "bbox": (x, y, w, h),
                "face_gray": face_gray,
                "recognized_info": recognized_info,
                "lbph_conf": lbph_conf,
            })

        return results

    def capture_dataset(self):
        sid = self.student_id_var.get().strip()
        name = self.student_name_var.get().strip()
        total_samples = int(self.samples_var.get())

        if not sid or not name:
            messagebox.showwarning("Thiếu dữ liệu", "Hãy nhập MSSV và Họ tên.")
            return

        if not self.running or self.current_frame is None:
            messagebox.showwarning("Chưa có camera", "Hãy mở camera trước khi thu dataset.")
            return

        label_id = self.student_db.upsert(sid, name)
        save_dir = os.path.join(DATASET_DIR, f"{label_id}_{sid}_{self.safe_name(name)}")
        os.makedirs(save_dir, exist_ok=True)

        captured = 0
        frame_count = 0
        interval = max(1, int(self.capture_interval_var.get()))
        self.log_status("Bắt đầu thu dataset...")

        while self.running and captured < total_samples:
            self.root.update()

            if self.current_frame is None:
                continue

            frame_count += 1
            if frame_count % interval != 0:
                continue

            frame = self.current_frame.copy()
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = self.detector.detectMultiScale(
                gray,
                scaleFactor=1.2,
                minNeighbors=5,
                minSize=(90, 90)
            )

            if len(faces) == 0:
                continue

            areas = [(fw * fh, (fx, fy, fw, fh)) for (fx, fy, fw, fh) in faces]
            _, (x, y, w, h) = max(areas, key=lambda t: t[0])

            face = frame[y:y+h, x:x+w]
            face = preprocess_face(face)
            face = cv2.resize(face, (160, 160))

            filename = os.path.join(save_dir, f"img_{captured + 1:03d}.jpg")
            cv2.imwrite(filename, face)

            captured += 1
            self.log_status(f"Đang thu dataset cho {name}: {captured}/{total_samples}")
            time.sleep(0.05)

        self.refresh_student_table()
        messagebox.showinfo("Hoàn tất", f"Đã lưu {captured} ảnh cho {name}.")
        self.log_status("Thu dataset hoàn tất.")

    def import_images_from_folder(self):
        sid = self.student_id_var.get().strip()
        name = self.student_name_var.get().strip()

        if not sid or not name:
            messagebox.showwarning("Thiếu dữ liệu", "Hãy nhập MSSV và Họ tên trước.")
            return

        folder = filedialog.askdirectory(title="Chọn thư mục ảnh khuôn mặt")
        if not folder:
            return

        label_id = self.student_db.upsert(sid, name)
        save_dir = os.path.join(DATASET_DIR, f"{label_id}_{sid}_{self.safe_name(name)}")
        os.makedirs(save_dir, exist_ok=True)

        count = 0
        for fn in os.listdir(folder):
            ext = os.path.splitext(fn)[1].lower()
            if ext not in SUPPORTED_EXTS:
                continue

            path = os.path.join(folder, fn)
            img = cv2.imread(path)
            if img is None:
                continue

            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            faces = self.detector.detectMultiScale(
                gray,
                scaleFactor=1.2,
                minNeighbors=5,
                minSize=(90, 90)
            )

            if len(faces) == 0:
                continue

            areas = [(fw * fh, (fx, fy, fw, fh)) for (fx, fy, fw, fh) in faces]
            _, (x, y, w, h) = max(areas, key=lambda t: t[0])

            face = img[y:y+h, x:x+w]
            face = preprocess_face(face)
            face = cv2.resize(face, (160, 160))

            out = os.path.join(save_dir, f"import_{count + 1:03d}.jpg")
            cv2.imwrite(out, face)
            count += 1

        self.refresh_student_table()
        messagebox.showinfo("Hoàn tất", f"Đã import {count} ảnh cho {name}.")
        self.log_status("Import ảnh hoàn tất.")

    def collect_training_data(self):
        faces = []
        labels = []

        if not os.path.exists(DATASET_DIR):
            return faces, labels

        for folder_name in os.listdir(DATASET_DIR):
            folder_path = os.path.join(DATASET_DIR, folder_name)
            if not os.path.isdir(folder_path):
                continue

            try:
                label_id = int(folder_name.split("_")[0])
            except Exception:
                continue

            for fn in os.listdir(folder_path):
                ext = os.path.splitext(fn)[1].lower()
                if ext not in SUPPORTED_EXTS:
                    continue

                path = os.path.join(folder_path, fn)
                img = cv2.imread(path, cv2.IMREAD_GRAYSCALE)
                if img is None:
                    continue

                img = cv2.resize(img, (160, 160))
                faces.append(img)
                labels.append(label_id)

        return faces, labels

    def train_model(self):
        faces, labels = self.collect_training_data()

        if len(faces) < 2:
            messagebox.showwarning("Thiếu dữ liệu", "Dataset chưa đủ ảnh để train.")
            return

        if len(set(labels)) < 1:
            messagebox.showwarning("Thiếu dữ liệu", "Không có label hợp lệ để train.")
            return

        try:
            self.recognizer = cv2.face.LBPHFaceRecognizer_create()
            self.recognizer.train(faces, np.array(labels))
            self.recognizer.save(MODEL_PATH)
            self.model_ready = True
            self.log_status(f"Train hoàn tất với {len(faces)} ảnh.")
            messagebox.showinfo("Thành công", f"Đã train model với {len(faces)} ảnh.")
        except Exception as e:
            messagebox.showerror("Lỗi train", str(e))

    def reload_model(self):
        if not os.path.exists(MODEL_PATH):
            messagebox.showwarning("Chưa có model", "Chưa tìm thấy file model đã lưu.")
            return

        self._safe_load_model()
        if self.model_ready:
            messagebox.showinfo("Model", "Đã nạp model đã lưu thành công.")
        else:
            messagebox.showerror("Model", "Không nạp được model đã lưu.")

    def _safe_load_model(self):
        try:
            self.recognizer = cv2.face.LBPHFaceRecognizer_create()
            self.recognizer.read(MODEL_PATH)
            self.model_ready = True
            self.log_status("Đã nạp model đã train.")
        except Exception:
            self.model_ready = False
            self.log_status("Không nạp được model đã train.")

    def auto_mark_attendance_from_faces(self, faces):
        marked = 0

        for item in faces:
            info = item.get("recognized_info")
            lbph_conf = item.get("lbph_conf")

            if not info or lbph_conf is None:
                continue

            key = info["student_id"]
            now_ts = time.time()

            if key in self.last_mark_times and now_ts - self.last_mark_times[key] < self.cooldown_seconds:
                continue

            self.write_attendance(info["student_id"], info["student_name"], lbph_conf)
            self.last_mark_times[key] = now_ts
            marked += 1

        if marked > 0:
            self.log_status(f"Đã tự động ghi nhận {marked} lượt điểm danh.")

    def mark_attendance_current_frame(self):
        if not self.running:
            messagebox.showwarning("Chưa mở camera", "Hãy mở camera trước.")
            return

        if not self.model_ready:
            messagebox.showwarning("Chưa có model", "Hãy train model hoặc nạp model trước.")
            return

        marked = 0

        for item in self.current_faces:
            info = item.get("recognized_info")
            lbph_conf = item.get("lbph_conf")

            if not info or lbph_conf is None:
                continue

            key = info["student_id"]
            now_ts = time.time()

            if key in self.last_mark_times and now_ts - self.last_mark_times[key] < self.cooldown_seconds:
                continue

            self.write_attendance(info["student_id"], info["student_name"], lbph_conf)
            self.last_mark_times[key] = now_ts
            marked += 1

        if marked == 0:
            self.log_status("Không có sinh viên mới nào được ghi attendance trong frame hiện tại.")
            messagebox.showinfo("Attendance", "Không có sinh viên mới nào được ghi trong frame hiện tại.")
        else:
            self.log_status(f"Đã ghi nhận {marked} lượt điểm danh.")
            messagebox.showinfo("Attendance", f"Đã ghi nhận {marked} lượt điểm danh.")

    def write_attendance(self, student_id, student_name, lbph_confidence):
        today = datetime.now().strftime("%Y-%m-%d")
        now = datetime.now()
        csv_path = os.path.join(ATTENDANCE_DIR, f"attendance_{today}.csv")
        file_exists = os.path.exists(csv_path)

        with open(csv_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)

            if not file_exists:
                writer.writerow([
                    "student_id",
                    "student_name",
                    "date",
                    "time",
                    "datetime",
                    "lbph_confidence"
                ])

            writer.writerow([
                student_id,
                student_name,
                now.strftime("%Y-%m-%d"),
                now.strftime("%H:%M:%S"),
                now.strftime("%Y-%m-%d %H:%M:%S"),
                f"{lbph_confidence:.2f}",
            ])

    def save_current_snapshot(self):
        if self.current_frame is None:
            messagebox.showwarning("Không có frame", "Chưa có frame hiện tại.")
            return

        path = os.path.join(
            SNAPSHOT_DIR,
            f"snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
        )
        cv2.imwrite(path, self.current_frame)
        self.log_status(f"Đã lưu snapshot: {path}")
        messagebox.showinfo("Snapshot", f"Đã lưu ảnh:\n{path}")

    @staticmethod
    def safe_name(name):
        return "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in name)

    def on_close(self):
        try:
            self.left_canvas.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self.stop_camera()
        self.root.destroy()


def main():
    ensure_dirs()

    root = tk.Tk()
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    app = FaceAttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
