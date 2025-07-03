import cv2
import numpy as np
import face_recognition
import os
import time
import pandas as pd
from datetime import datetime


# Path to images folder

IMAGE_PATH = "C:/Users/binay/Desktop/attendance_web_app/Image"
file_path = "Attendance.xlsx"

images = []
classNames = []
myList = os.listdir(IMAGE_PATH)
for cl in myList:
    curImg = cv2.imread(os.path.join(IMAGE_PATH, cl))
    if curImg is not None:
        images.append(curImg)
        classNames.append(os.path.splitext(cl)[0])

def findEncodings(images):
    encodeList = []
    for img in images:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encodings = face_recognition.face_encodings(img)
        if encodings:
            encodeList.append(encodings[0])
    return encodeList

encodeListKnown = findEncodings(images)
attendance_data = {}
last_seen = {}
time_threshold = 5

def markAttendance(name, status="Entry"):
    time_now = datetime.now()
    tString = time_now.strftime('%H:%M:%S')
    dString = time_now.strftime('%d/%m/%Y')

    if not os.path.exists(file_path):
        df = pd.DataFrame(columns=["Name", "Entry Time", "Exit Time", "Duration (sec)", "Date"])
        df.to_excel(file_path, index=False)

    df = pd.read_excel(file_path)

    if status == "Entry":
        if name not in attendance_data:
            attendance_data[name] = time.time()
            new_entry = pd.DataFrame([[name, tString, "", "", dString]], columns=df.columns)
            df = pd.concat([df, new_entry], ignore_index=True)
    elif status == "Exit" and name in attendance_data:
        entry_time = attendance_data.pop(name)
        duration = int(time.time() - entry_time)
        for i in range(len(df)):
            if df.loc[i, "Name"] == name and pd.isna(df.loc[i, "Exit Time"]):
                df.at[i, "Exit Time"] = tString
                df.at[i, "Duration (sec)"] = duration
                break

    df.to_excel(file_path, index=False)

def process_frame_stream():
    cap = cv2.VideoCapture(0)
    while True:
        success, img = cap.read()
        if not success:
            break

        imgS = cv2.resize(img, (320, 240))
        imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)

        facesCurFrame = face_recognition.face_locations(imgS)
        encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)

        detected_names = []

        for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
            matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
            faceDis = face_recognition.face_distance(encodeListKnown, encodeFace)

            if len(faceDis) > 0:
                matchIndex = np.argmin(faceDis)
                if matches[matchIndex]:
                    name = classNames[matchIndex].upper()
                    detected_names.append(name)

                    y1, x2, y2, x1 = [v * 2 for v in faceLoc]
                    cv2.rectangle(img, (x1, y1), (x2, y2), (0, 255, 0), 2)
                    cv2.putText(img, name, (x1 + 6, y2 - 6), cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 2)

                    markAttendance(name, "Entry")
                    last_seen[name] = time.time()

        for name, last_time in list(last_seen.items()):
            if name not in detected_names and time.time() - last_time > time_threshold:
                markAttendance(name, "Exit")
                del last_seen[name]

        _, jpeg = cv2.imencode('.jpg', img)
        frame = jpeg.tobytes()
        yield frame

    cap.release()