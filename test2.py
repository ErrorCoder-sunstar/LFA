import face_recognition
import cv2
import numpy as np
import csv
import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import threading

# Initialize Tkinter root window
root = tk.Tk()
root.title("Attendance System")

video_capture = cv2.VideoCapture(0)

# Initialize lists to store known face encodings and corresponding names
known_face_encodings = []
known_face_names = []
students = []

attendance_records = []

def save_attendance_to_excel(records, filename):
    if os.path.exists(filename):
        # Load the existing attendance data from the Excel sheet
        df = pd.read_excel(filename)
        existing_records = df.values.tolist()
        records.extend(existing_records)

    df = pd.DataFrame(records, columns=["Name", "Time"])
    df.to_excel(filename, index=False)
    print(f"Attendance data appended to {filename}")

def load_registered_students():
    for name in os.listdir("photo"):
        if name.endswith(".jpg"):
            image = face_recognition.load_image_file(os.path.join("photo", name))
            face_encoding = face_recognition.face_encodings(image)
            if len(face_encoding) > 0:
                known_face_encodings.append(face_encoding[0])
                known_face_names.append(name.split("_")[0])
                students.append(name.split("_")[0])

def add_new_registration():
    name = entry_name.get()
    if name:
        # Capture an image from the video stream
        ret, frame = video_capture.read()

        # Save the image for the new registration
        now = datetime.now()
        img_filename = f"photo/{name}_{now.strftime('%Y%m%d_%H%M%S')}.jpg"
        cv2.imwrite(img_filename, frame)

        # Add the new student to the known_face_encodings and known_face_names lists
        image = face_recognition.load_image_file(img_filename)
        face_encoding = face_recognition.face_encodings(image)
        if len(face_encoding) > 0:
            known_face_encodings.append(face_encoding[0])
            known_face_names.append(name)
            students.append(name)
            lbl_status.config(text=f"New registration added for {name}")
        else:
            lbl_status.config(text="No face found in the captured image")
    else:
        lbl_status.config(text="Please enter a name")

def take_attendance():
    global students, attendance_records
    # Capture an image from the video stream
    ret, frame = video_capture.read()

    # Find all face locations in the frame
    face_locations = face_recognition.face_locations(frame)
    face_encodings = face_recognition.face_encodings(frame, face_locations)

    for face_encoding, (top, right, bottom, left) in zip(face_encodings, face_locations):
        # Compare the current face encoding with the known face encodings
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
        name = "Unknown"

        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = np.argmin(face_distances)
        if matches[best_match_index]:
            name = known_face_names[best_match_index]

            # Record attendance for the recognized student
            if name in students:
                students.remove(name)
                now = datetime.now()
                current_time = now.strftime("%H:%M:%S")

                # Add the attendance record to the list
                attendance_records.append([name, current_time])

    cv2.imshow("Attendance System", frame)

def video_capture_thread():
    global students, attendance_records

    while True:
        take_attendance()
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Save the attendance records to the same Excel sheet after the loop ends
    save_attendance_to_excel(attendance_records, "attendance_records.xlsx")

    # Release video capture and close all windows
    video_capture.release()
    cv2.destroyAllWindows()

def start_video_capture_thread():
    threading.Thread(target=video_capture_thread).start()

# Load registered students' photos and face encodings
load_registered_students()

# Create Tkinter elements
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label_name = tk.Label(frame, text="Enter Name:")
label_name.grid(row=0, column=0, padx=5, pady=5)

entry_name = tk.Entry(frame)
entry_name.grid(row=0, column=1, padx=5, pady=5)

btn_register = tk.Button(frame, text="Register", command=add_new_registration)
btn_register.grid(row=0, column=2, padx=5, pady=5)

btn_take_attendance = tk.Button(frame, text="Take Attendance", command=start_video_capture_thread)
btn_take_attendance.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

lbl_status = tk.Label(frame, text="")
lbl_status.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

# Start the Tkinter main loop
root.mainloop()
