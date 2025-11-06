import warnings
warnings.filterwarnings("ignore", category=UserWarning, module='face_recognition_models')

import cv2
import face_recognition
import os
import numpy as np
import datetime
import csv
import time
import ssl
import smtplib
from email.message import EmailMessage
import winsound
import threading
from PIL import Image, ImageTk
import customtkinter as ctk
import subprocess
import sys
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# =============================================================================
# CONFIGURATION
# =============================================================================
KNOWN_FACE_DIR = 'KNOWN_FACE'
ATTENDANCE_RECORD_DIR = 'ATTENDANCE RECORD'
os.makedirs(ATTENDANCE_RECORD_DIR, exist_ok=True)

END_TIME_SECONDS = 15 * 3600 + 30 * 60
ONTIME_LIMIT_SECONDS = 8 * 3600 + 45 * 60
PROCESS_EVERY_N_FRAME = 5 
SENDER_EMAIL = "203alamtree@gmail.com"
APP_PASSWORD = "meluhztxcstffiit"
RECIPIENT_EMAIL = "gsvelvishal@gmail.com"

# =============================================================================
# APPLICATION CLASS
# =============================================================================

class AttendanceApp:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.is_running = False
        self.recognition_thread = None
        self.stop_event = threading.Event()
        self.cap = None
        self.known_encodings, self.known_names, self.register_numbers = [], [], []
        self.recognized_today, self.late_entries = set(), set()
        self.video_label, self.status_label, self.time_label = None, None, None
        self.present_label, self.absent_label, self.late_label = None, None, None
        self.start_button, self.stop_button = None, None
        self.create_widgets()
        self.load_known_faces()
        self.update_time_label()

    def upload_to_google_drive(self, file_to_upload):
        try:
            self.schedule_status_update("Authenticating with Google Drive...")
            
            # --- FINAL OAUTH 2.0 AUTHENTICATION WITH REFRESH TOKEN ---
            creds_path = os.path.join(os.path.expanduser('~'), 'mycreds.txt')

            # Define the settings to force offline access for the refresh token
            gauth_settings = {
                "client_config_file": "client_secrets.json",
                "save_credentials": True,
                "save_credentials_file": creds_path,
                "get_refresh_token": True,
                "save_credentials_backend": "file",  # <-- THIS IS THE NEW, REQUIRED LINE
                "oauth_scope": ["https://www.googleapis.com/auth/drive"]
            }
            gauth = GoogleAuth(settings=gauth_settings)
            
            # Use LocalWebserverAuth for the one-time login
            gauth.LocalWebserverAuth()
            # --- End of update ---

            drive = GoogleDrive(gauth)

            # --- Create/Find folders and upload file (logic is unchanged) ---
            root_folder_name = "ATTENDANCE_REPORTS"
            root_folder_list = drive.ListFile({'q': f"title='{root_folder_name}' and 'root' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"}).GetList()
            
            if not root_folder_list:
                self.schedule_status_update(f"Creating new folder '{root_folder_name}' in your Google Drive...")
                folder_metadata = {'title': root_folder_name, 'mimeType': 'application/vnd.google-apps.folder'}
                folder = drive.CreateFile(folder_metadata)
                folder.Upload()
                root_folder_id = folder['id']
            else:
                root_folder_id = root_folder_list[0]['id']

            section_folder_id = self._find_or_create_folder(drive, "ECE-B", root_folder_id)
            month_name = datetime.datetime.now().strftime('%B')
            month_folder_id = self._find_or_create_folder(drive, month_name, section_folder_id)

            self.schedule_status_update(f"Uploading '{os.path.basename(file_to_upload)}' to Google Drive...")
            file_metadata = {'title': os.path.basename(file_to_upload), 'parents': [{'id': month_folder_id}]}
            gfile = drive.CreateFile(file_metadata)
            gfile.SetContentFile(file_to_upload)
            gfile.Upload()
            self.schedule_status_update("File uploaded successfully to Google Drive.")
        except Exception as e:
            error_message = f"[GDRIVE ERROR] {str(e)}"
            self.schedule_status_update(error_message)


    # The rest of the script is identical to the last version
    def _update_status_on_main_thread(self, text):
        if self.status_label: self.status_label.configure(text=f"Status: {text}")
        print(f"[STATUS] {text}")

    def schedule_status_update(self, text):
        if self.root: self.root.after(0, self._update_status_on_main_thread, text)

    def _find_or_create_folder(self, drive, folder_name, parent_id):
        query = f"title='{folder_name}' and '{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
        folder_list = drive.ListFile({'q': query}).GetList()
        if folder_list:
            return folder_list[0]['id']
        else:
            self.schedule_status_update(f"Creating new folder '{folder_name}' in Google Drive...")
            folder_metadata = {'title': folder_name, 'parents': [{'id': parent_id}], 'mimeType': 'application/vnd.google-apps.folder'}
            folder = drive.CreateFile(folder_metadata)
            folder.Upload()
            return folder['id']
            
    def run_recognition_loop(self):
        frame_counter = 0
        last_known_locations, last_known_names, last_known_colors = [], [], []
        while not self.stop_event.is_set():
            ret, frame = self.cap.read()
            if not ret: break
            frame_counter += 1
            if frame_counter % PROCESS_EVERY_N_FRAME == 0:
                last_known_locations, last_known_names, last_known_colors = [], [], []
                small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                rgb_small_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                face_locations = face_recognition.face_locations(rgb_small_frame)
                face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
                for face_encoding, face_location in zip(face_encodings, face_locations):
                    matches = face_recognition.compare_faces(self.known_encodings, face_encoding, tolerance=0.5)
                    name, color = "Unknown", (0, 0, 255)
                    if True in matches:
                        face_distances = face_recognition.face_distance(self.known_encodings, face_encoding)
                        best_match_index = np.argmin(face_distances)
                        if matches[best_match_index]:
                            name, color = self.known_names[best_match_index], (0, 255, 0)
                            if name not in self.recognized_today:
                                now = datetime.datetime.now()
                                current_seconds = now.hour*3600 + now.minute*60 + now.second
                                remark = "LATE" if current_seconds > ONTIME_LIMIT_SECONDS else "ON-TIME"
                                if remark == "LATE": self.late_entries.add(name)
                                winsound.Beep(1000, 200)
                                self.recognized_today.add(name)
                                self.schedule_status_update(f"Recognized: {self.register_numbers[best_match_index]}_{name} | {remark}")
                    last_known_locations.append(face_location)
                    last_known_names.append(name)
                    last_known_colors.append(color)
            for face_location, name, color in zip(last_known_locations, last_known_names, last_known_colors):
                top, right, bottom, left = (i * 4 for i in face_location)
                cv2.rectangle(frame, (left, top), (right, bottom), color, 2)
                cv2.putText(frame, name, (left + 6, bottom - 6), cv2.FONT_HERSHEY_DUPLEX, 0.7, (255, 255, 255), 1)
            current_time_of_day = datetime.datetime.now().hour * 3600 + datetime.datetime.now().minute * 60
            if current_time_of_day >= END_TIME_SECONDS:
                self.schedule_status_update(f"Time limit reached ({datetime.timedelta(seconds=END_TIME_SECONDS)}). Shutting down...")
                self.stop_event.set()
                break
            self.latest_frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        self.cap.release()
        self.is_running = False
        self.save_and_email_final_report()

    def save_and_email_final_report(self):
        date_str = datetime.datetime.now().strftime('%d-%m-%Y')
        filename = os.path.join(ATTENDANCE_RECORD_DIR, f"Attendance_{date_str}.csv")
        attendance_data = {}
        for reg_no, name in zip(self.register_numbers, self.known_names):
             attendance_data[name] = [reg_no, name, date_str, "-", "Absent", "-"]
        for name in self.recognized_today:
            reg_no = self.register_numbers[self.known_names.index(name)]
            remark = "LATE" if name in self.late_entries else "ON-TIME"
            time_of_recognition = datetime.datetime.now().strftime('%H:%M:%S') 
            attendance_data[name] = [reg_no, name, date_str, time_of_recognition, "Present", remark]
        try:
            with open(filename, 'w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["REGISTER NO", "NAME", "DATE", "TIME", "STATUS", "REMARKS"])
                for record in attendance_data.values():
                    writer.writerow(record)
            self.schedule_status_update(f"Attendance saved to {filename}")
            threading.Thread(target=self.run_post_processing, args=(filename,), daemon=True).start()
        except Exception as e:
            self.schedule_status_update(f"[ERROR] Could not save file: {e}")

    def run_post_processing(self, filename):
        self.upload_to_google_drive(filename)
        self.send_email_report(filename)

    def setup_window(self):
        self.root.title("Smart Attendance System")
        self.root.geometry("1200x850")
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.bind("<KeyPress>", self.key_press_handler)

    def load_known_faces(self):
        self.schedule_status_update(f"Loading known faces from '{KNOWN_FACE_DIR}'...")
        for filename in os.listdir(KNOWN_FACE_DIR):
            if filename.lower().endswith(('.jpg', '.png', '.jpeg')):
                try:
                    path = os.path.join(KNOWN_FACE_DIR, filename)
                    image = face_recognition.load_image_file(path)
                    encodings = face_recognition.face_encodings(image)
                    if encodings:
                        self.known_encodings.append(encodings[0])
                        reg_no, name = os.path.splitext(filename)[0].split('_')
                        self.known_names.append(name)
                        self.register_numbers.append(reg_no)
                    else: print(f"[WARNING] No face found in {filename}. Skipping.")
                except Exception as e: print(f"[ERROR] Could not process {filename}: {e}")
        self.absent_label.configure(text=f"NOT DETECTED: {len(self.known_names)}")
        self.schedule_status_update(f"Loaded {len(self.known_names)} known faces. System ready.")

    def create_widgets(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        header_frame = ctk.CTkFrame(self.root, corner_radius=0)
        header_frame.grid(row=0, column=0, sticky="ew", padx=0, pady=0)
        ctk.CTkLabel(header_frame, text="✅ Smart Attendance System", font=ctk.CTkFont(size=20, weight="bold")).pack(side="left", padx=20, pady=10)
        self.time_label = ctk.CTkLabel(header_frame, text="", font=ctk.CTkFont(size=16))
        self.time_label.pack(side="right", padx=20, pady=10)
        self.video_label = ctk.CTkLabel(self.root, text="Press 'Start' to begin camera feed.", font=ctk.CTkFont(size=20))
        self.video_label.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        info_frame = ctk.CTkFrame(self.root)
        info_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=10)
        info_frame.grid_columnconfigure((0, 1, 2), weight=1)
        self.present_label = ctk.CTkLabel(info_frame, text="PRESENT: 0", font=ctk.CTkFont(size=16, weight="bold"), text_color="#34eb46")
        self.present_label.grid(row=0, column=0, pady=10)
        self.absent_label = ctk.CTkLabel(info_frame, text=f"NOT DETECTED: {len(self.known_names)}", font=ctk.CTkFont(size=16, weight="bold"), text_color="#ebb434")
        self.absent_label.grid(row=0, column=1, pady=10)
        self.late_label = ctk.CTkLabel(info_frame, text="LATE: 0", font=ctk.CTkFont(size=16, weight="bold"), text_color="#eb344f")
        self.late_label.grid(row=0, column=2, pady=10)
        control_frame = ctk.CTkFrame(self.root)
        control_frame.grid(row=3, column=0, sticky="ew", padx=20, pady=10)
        control_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.start_button = ctk.CTkButton(control_frame, text="▶ Start (S)", command=self.start_recognition, height=40)
        self.start_button.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.stop_button = ctk.CTkButton(control_frame, text="■ Stop (Q)", command=self.stop_recognition, height=40, state="disabled", fg_color="#D32F2F", hover_color="#B71C1C")
        self.stop_button.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(control_frame, text="📁 Open Folder (O)", command=self.open_folder, height=40).grid(row=0, column=2, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(control_frame, text="✉ Send Email (E)", command=self.send_email_manually, height=40).grid(row=0, column=3, padx=10, pady=10, sticky="ew")
        self.status_label = ctk.CTkLabel(self.root, text="System Ready.", anchor="w")
        self.status_label.grid(row=4, column=0, sticky="ew", padx=20, pady=(5, 10))

    def start_recognition(self):
        if self.is_running: return
        now_seconds = datetime.datetime.now().hour * 3600 + datetime.datetime.now().minute * 60
        if now_seconds >= END_TIME_SECONDS:
            self.schedule_status_update(f"Cannot start: It's past the scheduled stop time ({datetime.timedelta(seconds=END_TIME_SECONDS)}).")
            return
        self.is_running = True
        self.stop_event.clear()
        self.recognized_today.clear()
        self.late_entries.clear()
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            self.schedule_status_update("[ERROR] Cannot open webcam.")
            self.is_running = False
            return
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.schedule_status_update("System started. Recognizing faces...")
        self.recognition_thread = threading.Thread(target=self.run_recognition_loop, daemon=True)
        self.recognition_thread.start()
        self.update_gui_frame()
    
    def stop_recognition(self, manual=True):
        if not self.is_running: return
        self.is_running = False
        self.reset_gui_on_stop()
        if manual: self.schedule_status_update("Stopping system... saving final report.")
        self.stop_event.set()

    def update_gui_frame(self):
        if self.is_running and hasattr(self, 'latest_frame'):
            img = Image.fromarray(self.latest_frame)
            container_w, container_h = self.video_label.winfo_width(), self.video_label.winfo_height()
            if container_w > 1 and container_h > 1:
                img_w, img_h = img.size
                aspect_ratio = img_w / img_h
                new_w, new_h = container_w, int(container_w / aspect_ratio)
                if new_h > container_h: new_h, new_w = container_h, int(container_h * aspect_ratio)
                if (new_w, new_h) != (img_w, img_h): img = img.resize((new_w, new_h), Image.LANCZOS)
            ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(img.width, img.height))
            self.video_label.configure(image=ctk_img, text="")
            self.video_label.image = ctk_img
        present_count = len(self.recognized_today)
        absent_count = len(self.known_names) - present_count
        late_count = len(self.late_entries)
        self.present_label.configure(text=f"PRESENT: {present_count}")
        self.absent_label.configure(text=f"NOT DETECTED: {absent_count}")
        self.late_label.configure(text=f"LATE: {late_count}")
        if self.is_running: self.root.after(20, self.update_gui_frame)

    def reset_gui_on_stop(self):
        self.video_label.configure(image=None, text="System stopped. Press 'Start' to begin.")
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")

    def send_email_report(self, csv_file_path):
        if not os.path.exists(csv_file_path):
            self.schedule_status_update(f"[EMAIL ERROR] File not found: {csv_file_path}")
            return
        self.schedule_status_update("Preparing to send email report...")
        try:
            msg = EmailMessage()
            msg["Subject"] = f"Daily Attendance Report - {datetime.datetime.now().strftime('%d-%m-%Y')}"
            msg["From"] = SENDER_EMAIL
            msg["To"] = RECIPIENT_EMAIL
            msg.set_content("Attached is the automated attendance report.")
            with open(csv_file_path, "rb") as f:
                msg.add_attachment(f.read(), maintype="application", subtype="octet-stream", filename=os.path.basename(csv_file_path))
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
                smtp.login(SENDER_EMAIL, APP_PASSWORD)
                smtp.send_message(msg)
            self.schedule_status_update(f"Email report sent successfully to {RECIPIENT_EMAIL}")
        except Exception as e:
            self.schedule_status_update(f"[EMAIL ERROR] {e}")

    def send_email_manually(self):
        if self.is_running:
            self.schedule_status_update("Please stop the system before sending a manual report.")
            return
        self.schedule_status_update("Manual report generation requested...")
        threading.Thread(target=self.save_and_email_final_report, daemon=True).start()

    def open_folder(self):
        path = os.path.realpath(ATTENDANCE_RECORD_DIR)
        try:
            if sys.platform == "win32": os.startfile(path)
            elif sys.platform == "darwin": subprocess.run(["open", path])
            else: subprocess.run(["xdg-open", path])
            self.schedule_status_update(f"Opened folder: {path}")
        except Exception as e: self.schedule_status_update(f"Error opening folder: {e}")

    def key_press_handler(self, event):
        key = event.char.lower()
        if key == 's' and self.start_button.cget('state') == 'normal': self.start_recognition()
        elif key == 'q' and self.stop_button.cget('state') == 'normal': self.stop_recognition()
        elif key == 'o': self.open_folder()
        elif key == 'e': self.send_email_manually()
        elif event.keysym == 'F11': self.root.attributes("-fullscreen", not self.root.attributes("-fullscreen"))

    def on_closing(self):
        if self.is_running:
            self.stop_recognition(manual=False)
            self.root.after(1000, self.root.destroy) 
        else: self.root.destroy()
            
    def update_time_label(self):
        current_time = datetime.datetime.now().strftime("%d/%m/%Y | %I:%M:%S %p")
        self.time_label.configure(text=current_time)
        self.root.after(1000, self.update_time_label)

if __name__ == "__main__":
    root = ctk.CTk()
    app = AttendanceApp(root)
    root.mainloop()
