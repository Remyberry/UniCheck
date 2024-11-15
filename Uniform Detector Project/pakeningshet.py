import cv2
import threading
import tkinter as tk
from tkinter import Canvas, Button
from PIL import Image, ImageTk
import pandas as pd
import time
import csv

class ObjectDetectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Object Detection and QR Capture")

        # Set up flags and canvas
        self.cam1_running = False
        self.cam2_running = False
        self.cam1_frame = None
        self.canvas = Canvas(root, width=640, height=480)
        self.canvas.grid(row=0, column=0)

        # Buttons for Camera Control
        self.start_button = Button(root, text="Start", command=self.start_camera)
        self.start_button.grid(row=1, column=0)
        self.capture_button = Button(root, text="Capture", command=self.capture)
        self.capture_button.grid(row=2, column=0)

        # CSV file for logging data
        self.csv_file = 'uniform_check.csv'

        # Placeholder for YOLO model (Assumed to be loaded)
        self.yolo_model = self.load_model()

        # For QR detection, second camera feed
        self.cam2_label = Canvas(root, width=320, height=240)
        self.cam2_label.grid(row=0, column=1)
    
    def load_model(self):
        # Placeholder to load your trained YOLO model
        pass

    def start_camera(self):
        if not self.cam1_running:
            self.cam1_running = True
            self.cam1_thread = threading.Thread(target=self.camera_feed, args=(0,))
            self.cam1_thread.start()

    def camera_feed(self, cam_index):
        cap = cv2.VideoCapture(cam_index, cv2.CAP_DSHOW)
        while self.cam1_running:
            ret, frame = cap.read()
            if ret:
                # Run object detection on the frame
                detections = self.run_object_detection(frame)
                
                # Convert to Tkinter compatible format
                frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(frame_rgb)
                imgtk = ImageTk.PhotoImage(image=img)

                # Update the canvas with the new frame
                self.canvas.create_image(0, 0, anchor=tk.NW, image=imgtk)
                
                # Draw detection bounding boxes and labels
                self.draw_detections(detections)

            time.sleep(0.03)
        cap.release()

    def run_object_detection(self, frame):
        # Placeholder: run the YOLO model on the frame
        # Return mock detections for now: [(x1, y1, x2, y2, label)]
        return [(50, 50, 200, 200, 'Student')]  # Example mock detection

    def draw_detections(self, detections):
        # Clear previous annotations
        self.canvas.delete('all')

        # Draw bounding boxes and labels
        for (x1, y1, x2, y2, label) in detections:
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=2)
            self.canvas.create_text(x1, y1, anchor=tk.NW, text=label, fill="red", font=("Arial", 16))

    def capture(self):
        # Capture the frame and analyze it
        if self.cam1_frame is not None:
            # Perform capture, process student uniform status
            qr_code = self.detect_qr_code()
            if self.student_not_in_uniform():
                # Record in CSV
                self.record_to_csv(qr_code, "Not in uniform")

    def detect_qr_code(self):
        # Placeholder: detect QR code from second camera
        return "Student QR 12345"  # Example mock QR code

    def student_not_in_uniform(self):
        # Placeholder: Logic to determine if student is not in uniform
        return True  # Example: Assume student is not in uniform

    def record_to_csv(self, qr_code, remarks):
        # Record QR code, remarks, and image in CSV
        with open(self.csv_file, 'a', newline='') as file:
            writer = csv.writer(file)
            writer.writerow([qr_code, remarks, 'captured_image.jpg'])  # Save image separately

    def on_close(self):
        self.cam1_running = False
        self.cam2_running = False
        self.root.quit()

if __name__ == '__main__':
    root = tk.Tk()
    app = ObjectDetectionApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
