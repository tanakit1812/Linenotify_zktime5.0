import pyodbc
import requests
import time
import os
import psutil
from datetime import datetime

def read_config_file(file_path):
    with open(file_path, 'r') as f:
        lines = f.readlines()
        database_path = lines[0].strip()
        line_notify_token = lines[1].strip()
    return database_path, line_notify_token

def get_config_path():
    config_file_path = 'config.txt'
    if not os.path.exists(config_file_path):
        print(f"Not found '{config_file_path}' in the same folder.")
        while True:
            new_path = input("Please enter the path to the config file: ")
            if os.path.exists(new_path):
                return new_path
            else:
                print(f"File '{new_path}' not found. Please try again.")
    return config_file_path

def check_program_running(program_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == program_name:
            return True
    return False

# ตรวจสอบว่าโปรแกรม Att.exe รันอยู่หรือไม่
program_name = 'Att.exe'
if not check_program_running(program_name):
    print("Must run ZKTime5.0 first.")
    input("Press Enter to exit...")  # รอให้ผู้ใช้กด Enter เพื่อปิดโปรแกรม
    exit()  # ปิดโปรแกรม

config_file_path = get_config_path()
database_path, line_notify_token = read_config_file(config_file_path)

# สร้างการเชื่อมต่อ
conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    rf'DBQ={database_path};'
)

line_notify_api = 'https://notify-api.line.me/api/notify'

headers = {
    'Content-Type': 'application/x-www-form-urlencoded',
    'Authorization': f'Bearer {line_notify_token}'
}

def get_latest_records(cursor, latest_checktime):
    query = """
    SELECT CHECKINOUT.USERID, USERINFO.Badgenumber, USERINFO.Name, Format(CHECKINOUT.CHECKTIME, 'dd/mm/yyyy hh:nn:ss') AS FormattedCheckTime
    FROM CHECKINOUT
    INNER JOIN USERINFO ON CHECKINOUT.USERID = USERINFO.USERID
    WHERE CHECKINOUT.CHECKTIME > ?
    ORDER BY CHECKINOUT.CHECKTIME ASC
    """
    cursor.execute(query, latest_checktime)
    return cursor.fetchall()

def send_line_notify(message):
    payload = {'message': message}
    response = requests.post(line_notify_api, headers=headers, data=payload)
    if response.status_code == 200:
        print("send attendance successfully")
    else:
        print(f"Failed to send attendance. Status code: {response.status_code}")

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ดึงข้อมูลล่าสุดใน database เพื่อตั้งค่าเริ่มต้นของ latest_checktime
cursor.execute("""
    SELECT TOP 1 Format(CHECKTIME, 'yyyy-mm-dd hh:nn:ss')
    FROM CHECKINOUT
    ORDER BY CHECKTIME DESC
""")
latest_checktime_record = cursor.fetchone()
latest_checktime = latest_checktime_record[0] if latest_checktime_record else "2000-01-01 00:00:00"
print("LineNotify Ready to use")  # แสดงข้อความเมื่อผ่านทุกเงื่อนไข

# วนลูปเพื่อเช็คการเปลี่ยนแปลง
try:
    while True:
        new_records = get_latest_records(cursor, latest_checktime)

        for record in new_records:
            formatted_time_str = record.FormattedCheckTime
            message = f'รหัสพนักงาน: {record.Badgenumber}\nชื่อ: {record.Name}\nบันทึกเวลา: {formatted_time_str}'
            send_line_notify(message)
            latest_checktime = formatted_time_str  # อัพเดท latest_checktime หลังจากส่งข้อมูล

            # เพิ่ม delay เล็กน้อยระหว่างการส่งข้อความแต่ละครั้งเพื่อไม่ให้เกินขีดจำกัดของ LINE Notify
            time.sleep(0.5)

        time.sleep(0.1)  # เช็คทุก ๆ 1 วินาที

except KeyboardInterrupt:
    print("Exit Linenotify...")
    time.sleep(3)

finally:
    cursor.close()
    conn.close()
