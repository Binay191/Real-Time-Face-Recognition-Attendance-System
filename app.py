from flask import Flask, render_template, Response, send_file, redirect, url_for, jsonify
from attendance_core import process_frame_stream, markAttendance
import os
import pandas as pd
import smtplib
from email.message import EmailMessage
import time
import webbrowser
import threading

app = Flask(__name__)

ATTENDANCE_FILE = 'Attendance.xlsx'
EMAIL_SENDER = ' ' # write email id of sender
EMAIL_PASSWORD = ' ' # email 16 digit stmp password 

webcam_active = False
start_time = None

@app.route('/')
def index():
    return render_template('index.html')

def gen():
    global start_time, webcam_active
    start_time = time.time()
    webcam_active = True
    for frame in process_frame_stream():
        if not webcam_active or (time.time() - start_time > 1800):
            print("Webcam stream stopped.")
            webcam_active = False
            break
        yield (b'--frame\r\n'
               b'Content-Type: image/jpeg\r\n\r\n' + frame + b'\r\n')

@app.route('/video_feed')
def video_feed():
    return Response(gen(), mimetype='multipart/x-mixed-replace; boundary=frame')

@app.route('/start_class')
def start_class():
    global webcam_active
    webcam_active = True
    return redirect(url_for('index'))

@app.route('/end_class')
def end_class():
    global webcam_active
    webcam_active = False

    if os.path.exists(ATTENDANCE_FILE):
        df = pd.read_excel(ATTENDANCE_FILE)
        pending_exit = df[(df["Exit Time"].isna()) | (df["Exit Time"] == "")]["Name"].unique()
        for name in pending_exit:
            markAttendance(name, status="Exit")

    record_absentees()
    send_attendance_email()
    return jsonify({"status": "Class ended and notifications sent."})

@app.route('/download')
def download_attendance():
    if os.path.exists(ATTENDANCE_FILE):
        return send_file(ATTENDANCE_FILE, as_attachment=True)
    return "Attendance file not found.", 404

@app.route('/view')
def view_attendance():
    if os.path.exists(ATTENDANCE_FILE):
        df = pd.read_excel(ATTENDANCE_FILE)
        return render_template('attendance.html', tables=[df.fillna("").to_html(classes='data', header="true", index=False)])
    return "Attendance file not found.", 404

def record_absentees():
    if not os.path.exists(ATTENDANCE_FILE):
        df = pd.DataFrame(columns=["Name", "Entry Time", "Exit Time", "Duration (sec)", "Date"])
    else:
        df = pd.read_excel(ATTENDANCE_FILE)

    email_map = {
        "SHUBHA": "..........@gmail.com", # write receiver password
        "BINAY": "...........@gmail.com"
    
    }

    present_names = set(df['Name'].str.upper())
    today = pd.Timestamp.today().strftime('%d/%m/%Y')

    for name in email_map:
        if name not in present_names:
            new_row = pd.DataFrame([[name, '', '', '', today]], columns=df.columns)
            df = pd.concat([df, new_row], ignore_index=True)

    df.to_excel(ATTENDANCE_FILE, index=False)

def send_attendance_email():
    if not os.path.exists(ATTENDANCE_FILE):
        return

    df = pd.read_excel(ATTENDANCE_FILE)
    email_map = {
        "SHUBHA": ".......@gmail.com", # write receiver password
        "BINAY": "........@gmai.com"
    }

    all_names = set(email_map.keys())
    today = pd.Timestamp.today().strftime('%d/%m/%Y')

    for upper_name in all_names:
        email = email_map[upper_name]
        person_data = df[(df['Name'].str.upper() == upper_name) & (df['Date'] == today)]

        if not person_data.empty and pd.notna(person_data.iloc[0]['Entry Time']) and person_data.iloc[0]['Entry Time'] != '':
            body = f"Hello {upper_name},\n\nYou were marked PRESENT today. Here is your attendance record:\n\n"
            for _, row in person_data.iterrows():
                exit_time = row['Exit Time'] if pd.notna(row['Exit Time']) else 'N/A'
                duration = row['Duration (sec)'] if pd.notna(row['Duration (sec)']) else 'N/A'
                body += f"Date: {row['Date']}\nEntry: {row['Entry Time']}\nExit: {exit_time}\nDuration: {duration} seconds\n\n"
            body += "\n\nBest regards,\nAttendance System\nUsha Martin University, Ranchi\n\nThank You"

        else:
            body = f"Hello {upper_name},\n\nYou were marked ABSENT today.\n\nBest regards,\nAttendance System\nUsha Martin University, Ranchi\n\nThank You"

        msg = EmailMessage()
        msg['Subject'] = "Your Attendance Report"
        msg['From'] = EMAIL_SENDER
        msg['To'] = email
        msg.set_content(body)

        try:
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
                smtp.send_message(msg)
                print(f"Email sent to {email}")
        except Exception as e:
            print(f"Failed to send email to {email}: {e}")

if __name__ == '__main__':
    port = 5000
    threading.Timer(1.0, lambda: webbrowser.open(f"http://localhost:{port}")).start()
    app.run(host='127.0.0.1', port=port, debug=True, use_reloader=False)