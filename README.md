# Auto Packing & DWS Processor

<img width="899" height="748" alt="image" src="https://github.com/user-attachments/assets/e3cbdc07-972f-47fe-a957-f2cb3fa28927" />
<img width="898" height="746" alt="image" src="https://github.com/user-attachments/assets/61ce1f1b-090a-4eda-807e-4e4d35d985d4" />
<img width="900" height="748" alt="image" src="https://github.com/user-attachments/assets/12eae2c8-23f3-4fbf-aa4e-7e28abbb01ac" />

Python desktop application สำหรับประมวลผลข้อมูล Auto Packing และ DWS แบบครบ workflow พร้อมระบบ Merge, VLOOKUP, Filtering, Data Validation และ Export Excel

โปรแกรมถูกออกแบบสำหรับงาน warehouse operation และ sorting center เพื่อช่วยลดเวลาการ prepare data ก่อนใช้งานจริง

รองรับ:

- Auto Packing Processing
- DWS Data Processing
- Rule-based Filtering
- Excel Automation
- Multi-tab GUI
- Progress Tracking
- ETA Estimation
- Background Thread Processing

---

# Features

- Multi-tab processing system
- Auto Packing workflow
- DWS workflow
- Filter rule management
- Excel merge automation
- VLOOKUP automation
- Duplicate removal
- Weight & dimension filtering
- Province extraction
- Progress bar + ETA
- Real-time log viewer
- Cancel processing
- Config persistence
- Auto remember last path
- Styled Excel export
- Auto open output folder
- Background threading
- Windows desktop GUI

---

# Application Overview

โปรแกรมนี้ใช้สำหรับ:

1. ประมวลผลข้อมูล Auto Packing
2. รวมข้อมูล DWS หลาย source
3. กรองข้อมูลตาม business rules
4. Generate Excel output พร้อม format

เหมาะสำหรับ:

- Warehouse Operation
- Logistics Hub
- Sorting Center
- DWS Validation
- Auto Packing Validation
- Parcel Data Processing

---

# Tech Stack

- Python
- Tkinter
- Pandas
- NumPy
- OpenPyXL
- Threading
- ConfigParser

---

# Project Structure

```text
project/
│
├── Main.py
├── Auto.py
├── Dws.py
├── Filter.py
├── path_manager.py
├── excel_utils.py
├── config.ini
├── filter_rules.ini
│
└── output/
```

---

# Application Modules

## 1. AUTO Tab

ใช้สำหรับ process ข้อมูล Auto Packing

Input:

- JMS-AWB File
- Billcode File

Workflow:

```text
Load Billcode
    ↓
Clean Sort Code
    ↓
VLOOKUP JMS-AWB
    ↓
Extract Province
    ↓
Remove Duplicate
    ↓
Filter Rules
    ↓
Export Excel
```

---

## 2. DWS Tab

ใช้สำหรับ process ข้อมูล DWS

Input:

- JMS-AWB File
- DWS 1-8 Folder
- DWS 9-11 File

Workflow:

```text
Load DWS 1-8
    ↓
Load DWS 9-11
    ↓
Merge Dataset
    ↓
VLOOKUP JMS-AWB
    ↓
Extract Province
    ↓
Remove Duplicate
    ↓
Filter Rules
    ↓
Export Excel
```

---

## 3. Filter Tab

ใช้สำหรับ filter ข้อมูลตาม DP Rules

รองรับ:

- Autopacking Filter
- DWS Filter

สามารถ:

- Add Rule
- Remove Rule
- Save Rule
- Load Rule

โดย rules จะถูกเก็บใน:

```text
filter_rules.ini
```

---

# Auto Packing Logic

## Step 1 — Load Billcode

อ่านข้อมูลจากทุก sheet ของ Billcode file

ใช้ columns:

```text
Billcode
Sort Code
```

---

## Step 2 — Clean Sort Code

ระบบจะ:

- trim space
- ตัดเหลือ 2 ตัวแรก
- remove invalid value

Invalid:

```text
na
-
S1
```

---

## Step 3 — VLOOKUP JMS-AWB

ระบบจะ merge ข้อมูลจาก:

```text
ส่งออกAWB
```

เพื่อดึง:

```text
OriginCode
```

---

## Step 4 — Province Extraction

ระบบจะสร้าง column:

```text
จังหวัด
```

โดย extract จาก:

```text
OriginCode[1:3]
```

---

## Step 5 — Filtering

Filter rules:

| Rule | Condition |
|---|---|
| Province | >= 27 |
| Duplicate | Remove |
| Empty DP | Remove |

---

# DWS Processing Logic

## Source Files

### DWS 1-8

อ่านทุกไฟล์ `.xlsx` ภายใน folder

Columns:

```text
Barcode
Weight
Length
Width
Height
```

---

### DWS 9-11

อ่านทุก sheet ภายใน workbook

แล้ว append เข้ากับ DWS 1-8

---

# DWS Filtering Rules

ระบบจะกรอง:

| Rule | Condition |
|---|---|
| Barcode | not noread |
| Weight | <= 3.00 |
| Length | <= 29.99 |
| Width | <= 29.99 |
| Height | <= 29.99 |
| Province | >= 27 |

---

# Filter Tab Logic

ระบบ filter ใช้ DP blacklist

ตัวอย่าง:

```text
829155
829156
829157
```

เมื่อ process:

- ถ้า DP ตรง blacklist
- row จะถูก remove

---

# Config System

## config.ini

ใช้เก็บ:

- last opened path
- output path
- selected file path

---

## filter_rules.ini

ใช้เก็บ:

```ini
[FILTER]
block_dp=829155,829156,829157
```

---

# Progress System

ทุก workflow มี:

- Progress Bar
- ETA Calculation
- Real-time Log
- Cancel Support

ETA คำนวณจาก:

```python
elapsed_time / progress
```

---

# Excel Export

ระบบ export ด้วย:

```python
openpyxl
```

และ apply:

- auto style
- number formatting
- barcode formatting
- DP formatting

---

# Excel Styling

ใช้ utility:

```python
apply_excel_style()
```

เพื่อ:

- auto column width
- header style
- alignment
- border
- readable layout

---

# Path Management

ระบบใช้:

```python
path_manager.py
```

เพื่อ:

- remember last directory
- auto initial folder
- persistent config

---

# GUI Design

โปรแกรมใช้:

```python
Tkinter + ttk
```

แบ่งเป็น:

- Header
- Notebook Tabs
- Input Panel
- Progress Panel
- Log Panel
- Action Panel

---

# Threading Design

โปรแกรมใช้:

```python
threading.Thread()
```

เพื่อ:

- ป้องกัน GUI freeze
- รองรับ long-running process
- ทำ background processing
- cancel task ได้

---

# Real-time Queue System

DWS pipeline ใช้:

```python
queue.Queue()
```

สำหรับ:

- progress update
- ETA update
- result transfer
- log message

ระหว่าง worker thread และ GUI thread

---

# Installation

## 1. Clone Repository

```bash
git clone https://github.com/yourname/auto-packing-dws-processor.git
```

---

## 2. Install Dependencies

```bash
pip install pandas numpy openpyxl xlrd
```

---

## 3. Run Application

```bash
python Main.py
```

---

# Build EXE

ใช้ PyInstaller:

```bash
pyinstaller --onefile --windowed Main.py
```

หรือ:

```bash
pyinstaller --noconsole --onefile Main.py
```

---

# Output Files

## Auto Packing

```text
AutopackingData.xlsx
```

---

## DWS

```text
DwsQueryData.xlsx
```

---

## Filter

```text
Autopacking_Filtered.xlsx
DWS_Filtered.xlsx
```

---

# Error Handling

ระบบรองรับ:

- Invalid Excel structure
- Missing file
- Invalid DP
- Cancel process
- Save failure
- Excel format error
- Duplicate handling
- Invalid numeric conversion

---

# User Experience Features

- Auto remember path
- Auto open output folder
- Real-time processing log
- Large progress bar
- ETA display
- Disable UI while processing
- Cancel button
- Styled output

---

# Future Improvements

- Multi-thread batch processing
- Parallel Excel loader
- SQLite cache
- Rule editor GUI
- Drag & drop file support
- Auto export scheduler
- Database integration
- CSV support
- Report dashboard
- Packaging statistics

---

# License

MIT License

---

# Author

Developed for warehouse automation and DWS / Auto Packing data processing workflow.



