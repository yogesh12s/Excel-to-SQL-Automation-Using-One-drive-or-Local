import time
import pandas as pd
import pyodbc
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os

# 🔁 Folder to monitor (your synced OneDrive folder path)
WATCH_FOLDER = r"C:\Users\YourName\OneDrive - Your Company\FolderName"

# SQL Server connection details
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=YOUR_SERVER;'
    r'DATABASE=YOUR_DATABASE;'
    r'UID=YOUR_USERNAME;'
    r'PWD=YOUR_PASSWORD'
)

class ExcelHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.src_path.endswith(".xlsx") or event.src_path.endswith(".xls"):
            print(f"📁 New file detected: {event.src_path}")
            try:
                df = pd.read_excel(event.src_path, engine='openpyxl')

                # 🔁 Rename columns if needed to match SQL table
                df = df.rename(columns={
                    'EmpID': 'EmployeeID',
                    'FullName': 'Name',
                    'Department': 'Dept'
                })

                # Insert into SQL Server
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()
                for _, row in df.iterrows():
                    cursor.execute(
                        "INSERT INTO Employee (EmployeeID, Name, Dept, Age) VALUES (?, ?, ?, ?)",
                        row.EmployeeID, row.Name, row.Dept, row.Age
                    )
                conn.commit()
                cursor.close()
                conn.close()
                print("✅ Data inserted successfully.")
            except Exception as e:
                print(f"❌ Error: {e}")

# Start watching the folder
observer = Observer()
observer.schedule(ExcelHandler(), path=WATCH_FOLDER, recursive=False)
observer.start()
print(f"👀 Watching folder: {WATCH_FOLDER}")

try:
    while True:
        time.sleep(5)
except KeyboardInterrupt:
    observer.stop()
observer.join()
