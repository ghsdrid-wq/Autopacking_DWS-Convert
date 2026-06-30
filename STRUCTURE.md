# STRUCTURE — Autopacking_DWS-Convert (AD:P)

> ⚠️ **กฎการดูแลไฟล์นี้ (สำคัญ)**
> ทุกครั้งที่แก้ไขโค้ดใน repo นี้ — เพิ่ม/ลบ/ย้ายไฟล์, เปลี่ยน logic การกรอง/คำนวณ, เปลี่ยนคอลัมน์ที่อ้างอิง (index), เปลี่ยน rule/threshold, หรือเพิ่มแท็บ — **ต้องอัปเดต STRUCTURE.md นี้ให้ตรงกับโค้ดเสมอ**

## ภาพรวม
แอป GUI (**Tkinter**) ชื่อ *"Auto Packing & DWS Processor"* — เครื่องมือ **แปลง/รวม/กรองข้อมูล Excel** ของงาน Autopacking และ DWS แล้ว export เป็นไฟล์ผลลัพธ์ มี 3 แท็บ

## วิธีรัน / Entry point
- รัน: `python Main.py` → คลาส `MainApp(tk.Tk)` ประกอบ 3 แท็บ

## โครงสร้างไฟล์
| ไฟล์ | หน้าที่ |
|------|---------|
| `Main.py` | หน้าต่างหลัก + ประกอบ 3 แท็บ (`AutoTab`, `DwsTab`, `FilterTab`) |
| `Auto.py` | **แท็บ AUTO** — รวม Billcode detail + VLOOKUP กับ JMS-AWB + สร้างคอลัมน์จังหวัด + ลบซ้ำ + กรอง → export `AutopackingData.xlsx` |
| `Dws.py` | **แท็บ DWS** — pipeline แยก (`process_excel_pipeline`) รวมไฟล์ DWS 1-8 (โฟลเดอร์) + DWS 9-11 (หลาย sheet) + VLOOKUP + กรองตาม rule น้ำหนัก/มิติ/จังหวัด → export `DwsQueryData.xlsx` (ใช้ queue คุย thread↔GUI) |
| `Filter.py` | **แท็บ Filter ค่าปรับ** — กรอง DP ที่อยู่ใน block list ออกจากไฟล์ Autopacking/DWS, จัดการ block list ใน `filter_rules.ini` |
| `path_manager.py` | จำ path ล่าสุดของแต่ละช่อง ใน `config.ini` (`load_path`/`save_path`/`get_initial_dir`) |

## ⚠️ Dependency ที่ขาดใน repo (ต้องมี)
- ทั้ง `Auto.py`, `Dws.py`, `Filter.py` import `from excel_utils import apply_excel_style` แต่ **ไม่มีไฟล์ `excel_utils.py` ใน repo** → ต้องมีไฟล์นี้ (จัด style ให้ sheet) ไม่งั้นรันไม่ขึ้น. *(ถ้าเพิ่ม/แก้ไฟล์นี้ ให้บันทึกในตารางด้านบนด้วย)*

## Logic/Rule ที่ควรรู้ (อาจต้องแก้เมื่อ requirement เปลี่ยน)
- **Auto:** อ่าน Billcode (col 0,8) → clean Sort Code 2 ตัวแรก (ตัด `na/-/S1`) → VLOOKUP JMS sheet `ส่งออกAWB` → จังหวัด = `int(OriginCode[1:3])` → กรอง `จังหวัด >= 27`
- **Dws:** base files cols [0,2,3,4,5], append cols [3,12,13,14,15]; กรอง weight ≤ 3.00, มิติ ≤ 29.99, จังหวัด ≥ 27
- **Filter:** block DP เก็บใน `filter_rules.ini` `[FILTER] block_dp`; Auto กรองคอลัมน์ index 2, DWS กรอง index 1 + weight 1–2.99, มิติ 1–29

## Config
- `config.ini` `[PATHS]` (จาก `path_manager.py`) — เก็บ path ล่าสุดของทุกช่อง
- `filter_rules.ini` `[FILTER] block_dp` — รายการ DP ที่บล็อก

## Dependencies
- `pandas`, `numpy`, `openpyxl`, `xlrd` (อ่าน .xls ใน Auto), `tkinter`, + `excel_utils` (ต้องเพิ่มเอง)

## ข้อควรระวัง
- โค้ดอ้างอิงคอลัมน์ด้วย **ตำแหน่ง index** เป็นหลัก → ถ้าฟอร์แมต Excel ต้นทางเปลี่ยนลำดับคอลัมน์ จะพังเงียบ
