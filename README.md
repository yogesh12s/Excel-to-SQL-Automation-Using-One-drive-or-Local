# 📊 Excel to SQL Server Auto Import via OneDrive

This Python script watches a **OneDrive-synced folder** and automatically imports any new Excel file into a **SQL Server database**.

---

## 🔧 Features

- 🕵️ Monitors a OneDrive folder for new `.xlsx` files
- 📥 Automatically reads and inserts Excel data into SQL Server
- 🧠 Smart column mapping support
- 🔁 Runs continuously in the background

---

## 📁 Folder Structure


---

## ⚙️ Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

🛠️ Setup
1. Configure OneDrive Folder Path
Update the WATCH_FOLDER in main.py:

python
Copy
Edit
WATCH_FOLDER = r"C:\Users\YourName\OneDrive - Your Company\FolderName"
2. Set SQL Server Connection
Update your connection string:

python
Copy
Edit
conn_str = (
    r'DRIVER={ODBC Driver 17 for SQL Server};'
    r'SERVER=YOUR_SERVER;'
    r'DATABASE=YOUR_DATABASE;'
    r'UID=YOUR_USERNAME;'
    r'PWD=YOUR_PASSWORD'
)


3. Adjust Column Mapping (if needed)
Update this block to match your Excel and SQL column names:

python
Copy
Edit
df = df.rename(columns={
    'EmpID': 'EmployeeID',
    'FullName': 'Name',
    'Department': 'Dept'
})


▶️ Run the Script


```bash
python main.py
```

The script will start watching for new Excel files and auto-insert them into SQL Server.

✅ Example SQL Table
sql
Copy
Edit
CREATE TABLE Employee (
    EmployeeID INT,
    Name NVARCHAR(100),
    Dept NVARCHAR(100),
    Age INT
);
📌 Tips
Ensure OneDrive is fully synced to your local machine.

Archive or move processed files to prevent duplicate insertion.

Use Task Scheduler or systemd to run this script on startup (optional).


