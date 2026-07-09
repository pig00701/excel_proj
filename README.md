# Excel Power Query — Clean Merged Headers & Null Rows

โปรเจกต์สำหรับแก้ไขปัญหา Excel ที่มี:
1. **Header แบบ Merge Cell 3 แถว** — รวม header หลายชั้นให้เป็นชื่อคอลัมน์เดียว
2. **แถวที่เป็น null ทั้งหมด** — ลบแถวว่างที่ไม่มีข้อมูล

ทั้งหมดใช้ **Power Query M** เท่านั้น ไม่มี VBA, Python ใน production

---

## โครงสร้างโปรเจกต์

```
excel_proj/
├── sample_data.xlsx                    # ไฟล์ตัวอย่างที่มีปัญหา
├── generate_sample.py                  # สคริปต์สร้าง sample_data.xlsx
├── Daily_Update_2566.xlsx              # ไฟล์ตัวอย่างรายปีสำหรับทดสอบ Daily_Update_Combined
├── Daily_Update_2567.xlsx              # (3 แถวขยะ + 3 แถว merged header ที่ row 4-6, ปี พ.ศ. ท้ายชื่อ)
├── generate_daily_update_samples.py    # สคริปต์สร้างไฟล์ตัวอย่างข้างต้น (.xlsx — ดูหมายเหตุด้านล่าง)
├── queries/
│   ├── fnCleanMergedHeaders.pq     # ฟังก์ชันหลัก (generic)
│   ├── Sales_Report_Clean.pq       # Query สำหรับ Sheet Sales_Report
│   ├── Employee_Data_Clean.pq      # Query สำหรับ Sheet Employee_Data
│   ├── fnCleanMergedHeadersSkipJunk.pq  # wrapper: ข้ามแถวขยะก่อนเรียก fnCleanMergedHeaders
│   ├── fnConfigValue.pq            # ฟังก์ชัน lookup ค่า config จากตาราง (แก้ปัญหา Formula.Firewall — ดูหัวข้อด้านล่าง)
│   ├── fnColumnLetterToIndex.pq    # แปลงตัวอักษรคอลัมน์ → เลขตำแหน่ง (ใช้เฉพาะ Daily_Update_HeaderList — ดูหัวข้อด้านล่าง)
│   ├── Config.pq                   # อ่านการตั้งค่าทั้งหมดจาก sheet "config" (Table: ConfigTable)
│   ├── Daily_Update_Combined.pq    # รวม sheet "Daily Update" จากทุกไฟล์ .xlsm ในโฟลเดอร์
│   └── Daily_Update_HeaderList.pq  # ตาราง reference: ชื่อ header ทั้งหมด + ตำแหน่งคอลัมน์ (A, B, C, ...)
├── tests/
│   ├── test_harness.py           # Python test harness (รันได้ทันที)
│   ├── Test_Sales_Report.pq      # M test query สำหรับ Sales
│   ├── Test_Employee_Data.pq     # M test query สำหรับ Employee
│   └── fixtures/
│       ├── build_fixtures.py     # สคริปต์สร้างไฟล์ทดสอบชุดล่าง
│       ├── ReportWorkbook.xlsx   # workbook พร้อม ConfigTable + SelectColumnTable สำหรับวางคิวรีทดสอบ
│       └── daily_files/          # ไฟล์ต้นทางตัวอย่าง (Daily_Update_2568/2569.xlsx)
└── README.md
```

---

## วิธีใช้ใน Power Query (Excel / Power BI)

### Step 1: โหลดฟังก์ชัน `fnCleanMergedHeaders`

1. เปิด **Power Query Editor**
2. สร้าง **Blank Query** → ตั้งชื่อว่า `fnCleanMergedHeaders`
3. เปิด **Advanced Editor** → วางโค้ดจาก `queries/fnCleanMergedHeaders.pq`
4. กด **Done**

### Step 2: สร้าง Query สำหรับแต่ละ Sheet

1. สร้าง **Blank Query** → ตั้งชื่อว่า `Sales_Report_Clean`
2. วางโค้ดจาก `queries/Sales_Report_Clean.pq`
3. แก้ไข path ไฟล์ให้ตรงกับเครื่องของคุณ
4. ทำซ้ำสำหรับ `Employee_Data_Clean`

### Step 3: โหลดข้อมูล

กด **Close & Load** — ข้อมูลจะถูก clean อัตโนมัติ

### (Optional) รวมหลายไฟล์ .xlsm ด้วย `Daily_Update_Combined`

สำหรับกรณีที่ต้องรวม sheet ชื่อ **"Daily Update"** จากไฟล์ `.xlsm` หลายไฟล์ (แยกเป็นไฟล์รายปี) ในโฟลเดอร์
เดียวกัน มีไฟล์ตัวอย่าง `Daily_Update_2566.xlsx` และ `Daily_Update_2567.xlsx` ให้ทดสอบได้ (สร้างจาก
`generate_daily_update_samples.py`) — มีโครงสร้าง 3 แถวขยะ (title/generated-on/department) ตามด้วย
merged header 3 แถวที่ row 4-6 แล้วข้อมูลเริ่ม row 7 ตรงกับที่ `Daily_Update_Combined.pq` คาดหวังไว้

> **ทำไมเป็น `.xlsx` ไม่ใช่ `.xlsm`**: `openpyxl` (ไลบรารีที่ใช้ generate ไฟล์) เขียนได้แค่เนื้อไฟล์แบบ
> xlsx ธรรมดา ไม่สามารถฝัง VBA project จริงได้ ถ้าตั้งนามสกุลเป็น `.xlsm` เนื้อไฟล์ข้างในจะไม่ตรงกับ
> extension แล้ว Excel จะปฏิเสธไม่ยอมเปิดเลย (error "file format or file extension is not valid")
> จึงต้อง save เป็น `.xlsx` แทน — ไฟล์การผลิตจริงของคุณที่เป็น macro-enabled workbook จาก Excel จริงๆ
> ไม่ได้รับผลกระทบอะไร ยังเป็น `.xlsm` ปกติ
>
> **วิธีทดสอบกับ `Daily_Update_Combined.pq`** (ซึ่ง filter เฉพาะนามสกุล `.xlsm`): copy ไฟล์ตัวอย่างแล้ว
> เปลี่ยนนามสกุล copy นั้นเป็น `.xlsm` ได้เลย — Power Query อ่านไฟล์จาก content ข้างใน ไม่ได้เช็คว่า
> extension ตรงกับ content type แบบที่ Excel UI เช็ค จึงเปิดอ่านได้ปกติแม้ Excel จะไม่ยอมเปิดไฟล์นั้น
> ตรงๆ ก็ตาม (ใช้ trick นี้เฉพาะตอนทดสอบเท่านั้น ไม่ใช่กับไฟล์จริง)

1. ทำ **Step 1** (โหลด `fnCleanMergedHeaders`) ให้เสร็จก่อน
2. โหลด `fnCleanMergedHeadersSkipJunk` และ `fnConfigValue` เพิ่มอีกสองตัว:
   - สร้าง **Blank Query** → ตั้งชื่อว่า `fnCleanMergedHeadersSkipJunk`
   - เปิด **Advanced Editor** → วางโค้ดจาก `queries/fnCleanMergedHeadersSkipJunk.pq`
   - ทำซ้ำ: **Blank Query** → ตั้งชื่อว่า `fnConfigValue` → วางโค้ดจาก `queries/fnConfigValue.pq`
     (ฟังก์ชันนี้จำเป็น — `Daily_Update_Combined` / `Daily_Update_HeaderList` เรียกใช้ตรงๆ
     เพื่อเลี่ยง Formula.Firewall ดูหัวข้อ "ปัญหา Formula.Firewall และวิธีแก้" ด้านล่าง)
3. ทำ **sheet "config"** เก็บการตั้งค่าทั้งหมด (แก้ค่าที่นี่แล้ว Refresh — ไม่ต้องเปิด M code อีกเลย):
   - เพิ่ม sheet ใหม่ ตั้งชื่อ tab ว่าอะไรก็ได้ เช่น "config"
   - พิมพ์ตาราง 3 คอลัมน์ตามตัวอย่างนี้ (คอลัมน์ `หมายเหตุ` มีไว้ให้คนอ่าน query ไม่ใช้ จะไม่ใส่ก็ได้):

     | Setting | Value | หมายเหตุ |
     |---|---|---|
     | FolderPath | `C:\Users\User\Documents\DailyFiles` | โฟลเดอร์เก็บไฟล์ต้นทาง — ย้ายโฟลเดอร์เมื่อไหร่แก้ช่องนี้ |
     | SheetName | `Daily Update` | ชื่อ sheet ในไฟล์ต้นทาง ต้องสะกดตรงเป๊ะ |
     | JunkRows | `3` | จำนวนแถวขยะบนสุดที่ข้ามทิ้ง |
     | HeaderRows | `3` | จำนวนแถวหัวตารางที่ merge กัน |
     | Separator | `_` | ตัวคั่นตอนรวมชื่อหัวตาราง |
     | FileExtension | `.xlsm` | นามสกุลไฟล์ที่สแกน (เปลี่ยนเป็น `.xlsx` ได้ตอนทดสอบ) |

   - เลือกทั้งตาราง (รวมคอลัมน์หมายเหตุ) → **Insert > Table** (Ctrl+T)
   - **Table Design > Table Name** → เปลี่ยนชื่อเป็น `ConfigTable` (สำคัญ — query อ้างชื่อนี้)
   - ทุก Setting ยกเว้น `FolderPath` มีค่า default ในตัว ถ้าลบแถวไหนทิ้ง query จะใช้ค่าตามตารางข้างบน
     แต่ `FolderPath` บังคับต้องมี ไม่งั้น error พร้อมข้อความบอกว่าขาดอะไร
4. (ไม่บังคับ) โหลด `Config`:
   - สร้าง **Blank Query** → ตั้งชื่อว่า `Config`
   - วางโค้ดจาก `queries/Config.pq`
   - `Daily_Update_Combined` / `Daily_Update_HeaderList` **ไม่ได้ใช้คิวรีนี้แล้ว** (อ่าน ConfigTable
     เองผ่าน `fnConfigValue` เพื่อเลี่ยง Formula.Firewall) — โหลดไว้เผื่อคิวรีอื่นที่ไม่แตะ data source
     อยากอ่านค่า config แบบสะดวกๆ เท่านั้น
5. ทำ sheet สำหรับเลือกคอลัมน์:
   - เพิ่ม sheet ใหม่ (เช่น "select column")
   - ใส่หัวคอลัมน์ `ColumnName` แล้วพิมพ์ชื่อคอลัมน์ที่ต้องการเก็บ ทีละแถว (เช่น `ID`, `Date`, `Status`)
     จะเพิ่มคอลัมน์ `หมายเหตุ` ต่อท้ายไว้จดว่าแต่ละคอลัมน์คืออะไรก็ได้ query ไม่สนใจคอลัมน์เกิน
   - เลือกช่วงข้อมูลนี้ → **Insert > Table** (Ctrl+T)
   - ไปที่ **Table Design > Table Name** เปลี่ยนชื่อ Table เป็น `SelectColumnTable` (สำคัญ — query จะอ้างชื่อนี้)
6. สร้าง **Blank Query** → ตั้งชื่อว่า `Daily_Update_Combined` → วางโค้ดจาก `queries/Daily_Update_Combined.pq` (ไม่ต้องแก้อะไรในโค้ดแล้ว — ทุกการตั้งค่าอ่านจาก ConfigTable)

Query นี้จะสแกนทุกไฟล์ `.xlsm` ในโฟลเดอร์ ข้าม 3 แถวบนสุด (ไม่ใช่ header — เป็นแถวขยะ/ชื่อเรื่อง) ผ่าน `fnCleanMergedHeadersSkipJunk` ซึ่งจะเรียก `fnCleanMergedHeaders` ต่อให้อัตโนมัติเพื่อทำความสะอาด merged header อีก 3 แถวถัดมา จากนั้น**กรองเหลือเฉพาะคอลัมน์ที่ระบุใน sheet "select column"** ก่อนรวมทุกไฟล์เป็นตารางเดียว พร้อมคอลัมน์ `SourceFile` ไว้ตรวจสอบย้อนกลับว่าแต่ละแถวมาจากไฟล์ไหน ไฟล์ที่ไม่มี sheet "Daily Update" จะถูกข้ามโดยไม่ error

**คอลัมน์ `ปี` (ปีหมวดหมู่จากชื่อไฟล์)**: ไฟล์ต้นทางเป็นไฟล์รายปีที่มีปี พ.ศ. ต่อท้ายชื่อ (เช่น `Daily_Update_2566.xlsm`) query จะดึงเลขปีจากชื่อไฟล์มาใส่คอลัมน์ `ปี` ให้อัตโนมัติทุกแถว — วิธีดึงคือเอา **4 ตัวอักษรสุดท้ายก่อนนามสกุลไฟล์** มาแปลงเป็นตัวเลขตรงๆ (ไม่ได้ค้นหารูปแบบปีในชื่อไฟล์)

> ⚠️ **ต้องตั้งชื่อไฟล์ให้ลงท้ายด้วยเลขปี 4 หลักเสมอ** เช่น `Daily_Update_2566.xlsm`, `name_2556.xlsm` — ถ้าชื่อไฟล์ลงท้ายด้วยอย่างอื่น (เช่น `Daily_Update_2026-07-01.xlsm` ที่ลงท้ายด้วยวันที่ ไม่ใช่ปีเปล่าๆ) คอลัมน์ `ปี` จะได้ค่า null แทน ไม่ error แต่ก็จะไม่มีปีให้ใช้งาน

**ทำไมแยกเป็น `fnCleanMergedHeadersSkipJunk` ต่างหาก**: `fnCleanMergedHeaders` ตัวหลักยังใช้กับ `Sales_Report_Clean` / `Employee_Data_Clean` ที่ไม่มีแถวขยะเหมือนเดิมทุกอย่าง ไม่ต้องแก้ไขอะไร ส่วน sheet ที่มีแถวขยะเหนือ header (อย่าง "Daily Update") จะเรียกผ่าน wrapper ตัวนี้แทน ซึ่งข้างในแค่ `Table.Skip` แถวขยะออกก่อนแล้วส่งต่อให้ `fnCleanMergedHeaders` เหมือนเดิม — แยกไฟล์ให้ชัดว่า sheet ไหนต้องใช้ตัวไหน

**การเลือกคอลัมน์แบบ dynamic**: query อ่านรายชื่อคอลัมน์ที่ต้องการเก็บจาก Excel Table ชื่อ `SelectColumnTable` (คอลัมน์ `ColumnName`) ที่อยู่ในไฟล์เดียวกับ query เอง ผ่าน `Excel.CurrentWorkbook()` — เวลาต้องการเปลี่ยนว่าจะเก็บคอลัมน์ไหนบ้าง แค่แก้ค่าใน sheet "select column" ตรงๆ ใน Excel แล้วกด **Refresh** ไม่ต้องเปิด M code เลย คอลัมน์ที่ไม่อยู่ใน list จะถูกตัดทิ้งเพื่อไม่ให้ตารางผลลัพธ์มีข้อมูลเยอะเกินความจำเป็น ถ้าไฟล์ไหนไม่มีคอลัมน์ที่ระบุ (ชื่อไม่ตรง) จะข้ามเฉยๆ ไม่ error เพราะใช้ `MissingField.Ignore`

> หมายเหตุ: `Excel.CurrentWorkbook()` อ่านได้เฉพาะ **Excel Table** หรือ Named Range เท่านั้น — sheet ธรรมดาที่ไม่ได้แปลงเป็น Table จะไม่โผล่มาให้ query เห็น ต้องทำ Ctrl+T ตามขั้นตอนข้างบนเสมอ

### (Optional) ดูรายชื่อ header ทั้งหมดด้วย `Daily_Update_HeaderList`

ก่อนจะไปกรอก sheet "select column" ว่าจะเก็บคอลัมน์ไหนบ้าง จะได้รู้ก่อนว่ามี header อะไรให้เลือกบ้าง:

1. สร้าง **Blank Query** → ตั้งชื่อว่า `Daily_Update_HeaderList`
2. วางโค้ดจาก `queries/Daily_Update_HeaderList.pq` (ไม่ต้องแก้อะไรในโค้ด — อ่านการตั้งค่าจาก `ConfigTable` ชุดเดียวกับ `Daily_Update_Combined`)

ต้องโหลด `fnCleanMergedHeadersSkipJunk` และ `fnConfigValue` ไว้ก่อนเหมือนกัน (ดูขั้นตอนใน `Daily_Update_Combined` ด้านบน) Query นี้จะอ่าน sheet "Daily Update" จากไฟล์ `.xlsm` **ไฟล์เดียว** ในโฟลเดอร์ — ไฟล์ที่ชื่อเรียงมาเป็นอันดับแรกหลัง sort จากมาก→น้อย (Z→A) ตามชื่อไฟล์ ซึ่งตรงกับไฟล์ปีล่าสุด เพราะชื่อไฟล์ลงท้ายด้วยปี 4 หลักเสมอ (ไม่ใช้ทุกไฟล์ เพราะแค่ต้องการดู header ว่ามีอะไรบ้าง) ทำความสะอาด header ผ่าน `fnCleanMergedHeadersSkipJunk` เหมือน `Daily_Update_Combined` แต่**ไม่กรองคอลัมน์ตาม SelectColumnTable** — เพื่อให้เห็นชื่อ header ครบทุกตัวเสมอ ไม่ใช่แค่ตัวที่เลือกไว้แล้ว ผลลัพธ์เป็นตาราง 2 คอลัมน์:

| ชื่อ header | ตำแหน่ง column |
|---|---|
| ID | A |
| Date | B |
| Status | C |
| ... | ... |

`ตำแหน่ง column` คือตัวอักษรคอลัมน์ Excel ดั้งเดิม (A, B, C, ...) นับจากตำแหน่งจริงใน sheet "Daily Update" — ใช้ตารางนี้เป็น reference ตอนกรอกชื่อคอลัมน์ลงใน `SelectColumnTable`

---

## ปัญหา Formula.Firewall และวิธีแก้

### อาการ

ตอน Refresh `Daily_Update_HeaderList` / `Daily_Update_Combined` เจอ error:

```
Formula.Firewall: Query 'Daily_Update_HeaderList' (step 'FirstFile') references
other queries or steps, so it may not directly access a data source.
Please rebuild this data combination.
```

### สาเหตุ

โครงสร้างเดิมให้คิวรีหลักดึงค่า config ผ่าน**คิวรี `Config` แยกต่างหาก** (`Config[FolderPath]`)
แล้วเอาค่านั้นไปป้อน `Folder.Files(...)` ในคิวรีเดียวกันทันที:

```
โครงสร้างเดิม (โดน firewall บล็อก)
═══════════════════════════════════

  ┌─────────────────┐
  │  ConfigTable    │  data source #1 (Excel.CurrentWorkbook)
  │  (sheet config) │
  └────────┬────────┘
           │ อ่านโดย
           ▼
  ┌─────────────────┐
  │  Query: Config  │  ← คิวรีแยก มีการแตะ data source ของตัวเอง
  └────────┬────────┘
           │ Config[FolderPath] ข้ามขอบเขตคิวรี ✗
           ▼
  ┌──────────────────────────────┐
  │  Daily_Update_Combined /     │
  │  Daily_Update_HeaderList     │
  │                              │
  │  Folder.Files(FolderPath)    │  data source #2 (โฟลเดอร์)
  │        ▲                     │
  │        └── ค่าที่มาจากคิวรีอื่น │  ← 🔥 Formula.Firewall บล็อกตรงนี้
  └──────────────────────────────┘
```

Power Query's Formula Firewall ไม่ยอมให้ **ค่าที่ไหลมาจากคิวรีอื่น (ที่แตะ data source ของตัวเอง)
ถูกใช้เป็น input ของ data source function ในอีกคิวรี** — เพราะมันตรวจสอบ privacy isolation
ระหว่างสอง data source ข้ามขอบเขตคิวรีแบบนี้ไม่ได้ เลยบล็อกไว้ก่อน

### วิธีแก้ (rebuild data combination)

ย้ายการอ่าน `ConfigTable` เข้ามาทำ**ในคิวรีเดียวกัน**กับ `Folder.Files` โดยตรง แล้วแยก logic
การ lookup ค่า config ออกเป็น**ฟังก์ชันเปล่า** `fnConfigValue` (ฟังก์ชันที่ไม่แตะ data source
เองไม่โดน firewall — มันแค่ประมวลผลตารางที่ถูกส่งเข้ามา):

```
โครงสร้างใหม่ (ผ่าน firewall)
═════════════════════════════

  ┌──────────────────────────────────────────────┐
  │  Daily_Update_Combined / Daily_Update_HeaderList │
  │                                              │
  │  1. ConfigTable = Excel.CurrentWorkbook()... │  data source #1 ← อ่านเองในคิวรีนี้
  │  2. FolderPath  = fnConfigValue(ConfigTable, │
  │                     "FolderPath", null)      │  ← เรียกฟังก์ชันเปล่า ✓ ไม่ข้ามคิวรี
  │  3. Folder.Files(FolderPath)                 │  data source #2 ← อยู่คิวรีเดียวกัน ✓
  └──────────────────────────────────────────────┘
                      │ เรียกใช้ (ฟังก์ชัน ไม่ใช่ data source)
                      ▼
            ┌──────────────────┐
            │  fnConfigValue   │  ฟังก์ชันเปล่า — รับตาราง + ชื่อ setting
            │  (pure function) │  คืนค่า ไม่แตะ data source ใดๆ เอง
            └──────────────────┘
```

data source ทั้งสองตัว (workbook + โฟลเดอร์) ถูกเรียก**จากภายในคิวรีเดียวกัน** firewall
จึงวิเคราะห์และอนุญาตได้ ส่วน `fnConfigValue` เป็นแค่ฟังก์ชันช่วย lookup — ไม่นับเป็น data source

### สรุปสิ่งที่เปลี่ยน

| ไฟล์ | การเปลี่ยนแปลง |
|---|---|
| `fnConfigValue.pq` | **ใหม่** — ฟังก์ชัน lookup ค่า config จากตารางที่ส่งเข้ามา (case/space-insensitive, มี default) |
| `Daily_Update_HeaderList.pq` | อ่าน `ConfigTable` เอง + เรียก `fnConfigValue` แทนการอ้าง `Config[...]` |
| `Daily_Update_Combined.pq` | เหมือนกัน — เลิกอ้าง `Config[...]` ทั้งหมด |
| `Config.pq` | ยังอยู่ (เผื่อคิวรีอื่นใช้) แต่ข้างในเปลี่ยนมาเรียก `fnConfigValue` ร่วมกัน — **ห้าม**ให้คิวรีที่เรียก data source อื่นมาอ้างอิงมัน |

### ทางเลือก: ปิด privacy check แทนการแก้โค้ด (สำหรับปัญหา Formula.Firewall)

ถ้าไม่อยากแก้โค้ด ปิด firewall ที่ระดับเครื่อง/workbook ได้:
**File → Options and Settings → Query Options → Privacy → "Always ignore Privacy Level settings"**

ข้อควรระวัง: เป็นการตั้งค่าของ Power Query engine **บนเครื่องนั้น** ไม่ติดไปกับไฟล์ Excel —
คนอื่นที่เปิดไฟล์เดียวกันในเครื่องอื่นต้องตั้งเองด้วย ไม่งั้นเจอ error เดิม โค้ดเวอร์ชันที่แก้แล้ว
(`fnConfigValue`) ใช้ได้โดยไม่พึ่งการตั้งค่านี้เลย จึงเสถียรกว่าถ้าต้องแชร์ไฟล์ให้หลายคน

---

## ปัญหา used-range บวม (Refresh ไฟล์ไม่กี่สิบ MB แต่กิน RAM เป็น GB)

### อาการ

Refresh `Daily_Update_HeaderList` (หรือ `Daily_Update_Combined`) ทีเดียว กิน RAM ไปเป็น GB
ทั้งที่ไฟล์ `.xlsm` ต้นทางมีขนาดแค่หลักสิบ MB

### สาเหตุ

`Excel.Workbook(...)[Data]` ดึงข้อมูลของทั้ง sheet มาทีเดียว ไม่มีทาง "ขอแค่บาง range" ได้
ถ้า sheet นั้นมี **used range บวม** — คือ Excel จำขอบเขตที่เคยมีการพิมพ์/ใส่ format ไว้ (แม้จะลบ
เนื้อหาทิ้งไปแล้ว) เช่น ข้อมูลจริงมีแค่ถึงคอลัมน์ J แต่ used range ยังลากไปถึงคอลัมน์ ZZ — คอลัมน์
ว่างเปล่าหลายพันคอลัมน์เหล่านั้นจะถูก parse เข้ามาด้วย แล้วยังถูกลากผ่านทุก step ของ
`fnCleanMergedHeaders` (FillDown, รวม header, กรองแถว null) ทั้งที่ไม่มีข้อมูลอะไรอยู่เลย

### วิธีแก้

**แก้ที่ไฟล์ต้นทางโดยตรง (ทางแก้จริง — ต้องทำ)**: เปิดไฟล์ `.xlsm` ทุก sheet → กด **Ctrl+End**
ดูว่า cursor กระโดดไปไกลเกินข้อมูลจริงไหม ถ้าใช่ → เลือกแถว/คอลัมน์ส่วนเกินทั้งแถว/คอลัมน์
(ไม่ใช่แค่ Clear Contents) → Delete → Save — `Excel.Workbook(...)[Data]` ดึงข้อมูลทั้ง sheet
มาทีเดียว ไม่มีทาง "ขอแค่บาง range" ได้จากฝั่ง M code เลย ดังนั้นถ้า sheet มี used range บวมจริง
ทางแก้ที่ตัดปัญหาที่ต้นตอมีทางเดียวคือไปเคลียร์ที่ไฟล์ต้นทาง

**ข้อยกเว้น: `Daily_Update_HeaderList` ไม่กระทบ (แก้ที่ M code แล้ว)** — query นี้ใช้แค่ **ชื่อ
คอลัมน์** ไม่แตะข้อมูลจริงเลยสักแถว จึงตัดข้อมูลทิ้งได้ตั้งแต่ต้น: หลังอ่าน `RawTable` มา จะทำ
`Table.FirstN(RawTable, JunkRows + HeaderRows)` ตัดเหลือแค่แถว junk+header (6 แถว) ก่อนส่งเข้า
`fnCleanMergedHeadersSkipJunk` ทันที — ไม่ว่า sheet จะบวมกี่พันแถวก็ไม่กระทบ เพราะไม่มีแถว
ข้อมูลให้ประมวลผลอยู่แล้ว (`Daily_Update_HeaderList.pq` ตัวแปร `HeaderRowsOnly`)

`Daily_Update_Combined` ทำแบบเดียวกันไม่ได้ เพราะต้องเก็บข้อมูลจริงทุกแถวไว้ใช้งาน — ถ้าไฟล์
ต้นทางมี used range บวม query นี้ยังต้องพึ่งการเคลียร์ที่ไฟล์ต้นทางเท่านั้น

---

## ตำแหน่งคอลัมน์ (`ตำแหน่ง column`) คลาดเคลื่อนจากไฟล์จริง

### อาการ

`Daily_Update_HeaderList` แสดงตำแหน่งคอลัมน์ผิดจากที่เปิดไฟล์ Excel จริงดู — เช่น query บอกว่า
header อยู่คอลัมน์ A แต่ไฟล์จริงอยู่คอลัมน์ B (เลื่อนเท่ากันทุกคอลัมน์)

### สาเหตุ

`Excel.Workbook()[Data]` คืนตารางตาม **"used range"** ของ sheet เท่านั้น ไม่ใช่คอลัมน์ A เสมอไป —
ถ้าคอลัมน์ A (หรือมากกว่านั้น) ของ sheet ไม่เคยมีข้อมูล/format อะไรเลยจริงๆ Excel จะไม่นับเป็นส่วน
หนึ่งของ used range เลย ทำให้คอลัมน์แรกที่ Power Query อ่านมาได้ (`Column1`) แท้จริงคือคอลัมน์อื่น
(เช่น B) ของไฟล์จริง — และไม่มี field ไหนใน `Excel.Workbook()` บอกตรงๆ ว่า used range เริ่มที่
คอลัมน์ไหน จึงคำนวณ offset นี้เองจากโค้ดไม่ได้ ต้องให้ผู้ใช้ระบุเอง

### วิธีแก้ (เฉพาะ `Daily_Update_HeaderList` — ไม่กระทบ query อื่น)

เพิ่ม setting ใหม่ใน `ConfigTable`:

| Setting | Value | หมายเหตุ |
|---|---|---|
| FirstColumnLetter | `B` (ตัวอย่าง) | ตัวอักษรคอลัมน์จริงที่ข้อมูลเริ่มต้น (เปิดไฟล์ต้นทางดูเอง) — ค่า default คือ `A` (ไม่มี offset) |

`fnColumnLetterToIndex.pq` แปลงค่านี้เป็นตัวเลข offset แล้ว `Daily_Update_HeaderList.pq` บวกเข้ากับ
ตำแหน่งที่คำนวณได้ก่อนแปลงเป็นตัวอักษร (`ColIndexToLetter(_ + FirstColumnOffset)`) — ทำให้
`ตำแหน่ง column` ที่แสดงตรงกับตำแหน่งจริงในไฟล์

**ข้อจำกัด**: ต้องตรวจสอบและตั้งค่า `FirstColumnLetter` เองครั้งแรก (เปิดไฟล์ดูว่าคอลัมน์ A ว่างจริง
หรือแค่ถูกซ่อน — คอลัมน์ที่ **ซ่อน** ไว้ (hidden) ไม่กระทบ เพราะ `Excel.Workbook()` ยังอ่านคอลัมน์ที่
ซ่อนได้ปกติ มีผลเฉพาะคอลัมน์ที่ **ว่างเปล่าจริง** ไม่เคยมีข้อมูล/format เลยเท่านั้น) — ถ้าไฟล์ต้นทาง
เปลี่ยนโครงสร้างในอนาคต (ข้อมูลเริ่มขยับคอลัมน์) ต้องกลับมาแก้ค่านี้ใหม่

---

## หลักการทำงานของ `fnCleanMergedHeaders`

```
SourceTable (raw)
    │
    ├─ Step 1: แยก header แถวบนสุด (N แถว)
    │
    ├─ Step 2: คอลัมน์ที่ header null/ว่างทั้งหมด → เปลี่ยนชื่อเป็น "Column_A", "Column_B", ...
    │     (อิงตำแหน่งคอลัมน์ Excel) แทนที่จะ drop ทิ้ง
    │     ต้องทำก่อน FillRight ไม่งั้น FillRight จะเติมค่าจากซ้ายมา
    │
    ├─ Step 3: Fill Right — เติมค่าในแนวนอน (merge cell ซ้าย→ขวา)
    │     [A, null, null] → [A, A, A]
    │
    ├─ Step 4: Fill Down — เติมค่าในแนวตั้ง (merge cell บน→ล่าง)
    │     [A, B, C]
    │     [null, null, C] → [A, B, C]
    │
    ├─ Step 5: รวม header หลายชั้นด้วย "_"
    │     "Product Info" + "ID" → "Product Info_ID"
    │     (ตัดคำซ้ำติดกันออก เช่น "ID" + "ID" → "ID")
    │
    ├─ Step 5b: กันชื่อคอลัมน์ว่าง/ซ้ำ
    │     ชื่อว่างทั้งหมด → fallback "Column1", "Column2", ...
    │     ชื่อซ้ำกัน → เติม suffix "_2", "_3", ...
    │
    ├─ Step 6: เปลี่ยนชื่อคอลัมน์ + ตัด header rows ทิ้ง
    │
    ├─ Step 7: ลบแถวที่เป็น null ทั้งหมด
    │
    └─ Step 8: Auto-detect type (text → number)
          เช็คค่าทุกแถวในคอลัมน์ ไม่ใช่แค่แถวแรก — คอลัมน์ที่มีค่า
          text แปลกปน (เช่น "N/A") จะถูกจัดเป็น text ทั้งคอลัมน์
```

### Parameters

| Parameter | Type | Default | คำอธิบาย |
|-----------|------|---------|----------|
| `SourceTable` | table | required | ตารางดิบจาก Excel |
| `HeaderRows` | number | required | จำนวนแถว header (เช่น 3) |
| `Separator` | text | `"_"` | ตัวคั่นระหว่างชั้น header |

---

## การรัน Test

### Python Test Harness (แนะนำ — รันได้ทันที)

```powershell
python tests/test_harness.py
```

ผลลัพธ์:
```
============================================================
TEST: Sales_Report (3-row merged headers + null rows)
============================================================
  [PASS] Column count = 10
  [PASS] Row count = 7
  [PASS] No null rows
  [PASS] Column names match
  [PASS] First row ID = P001
  [PASS] Last row ID = P006
  [PASS] All IDs present (incl row 6 junk)

============================================================
TEST: Employee_Data (3-row merged headers + null rows)
============================================================
  [PASS] Column count = 6
  [PASS] Row count = 4
  [PASS] No null rows
  [PASS] Column names match
  [PASS] First row ID = E001
  [PASS] All IDs present

============================================================
TEST: Reference (clean sheet, 1-row header, no nulls)
============================================================
  [PASS] Column count = 3
  [PASS] Row count = 6
  [PASS] No null rows
  [PASS] Column names match

============================================================
TEST: Edge_NullHeaders (columns with null headers across all rows)
============================================================
  [PASS] Column count = 4
  [PASS] Row count = 3
  [PASS] No null rows
  [PASS] No empty column names
  [PASS] Column names match
  [PASS] All column names unique

============================================================
TEST: Edge_DupHeaders (duplicate header names after combine)
============================================================
  [PASS] Column count = 4
  [PASS] Row count = 2
  [PASS] No null rows
  [PASS] All column names unique
  [PASS] Column names match

============================================================
TEST: Edge_MixedTypes (numeric column with a stray text value)
============================================================
  [PASS] Old logic misdetects Amount as number (regression check)
  [PASS] Fixed logic detects Amount as text

============================================================
ALL TESTS PASSED
============================================================
```

Sheet `Sales_Report` และ `Employee_Data` เป็นตัวอย่างข้อมูลจริง ส่วน 3 sheet ต่อไปนี้เป็น **edge case** ที่เพิ่มเข้ามาเพื่อทดสอบ behavior เฉพาะทาง:

| Sheet | ทดสอบอะไร |
|---|---|
| `Edge_NullHeaders` | คอลัมน์ที่ header เป็น null ทั้ง 3 แถว → ต้อง rename เป็น `Column_A`, `Column_C` แทนที่จะพัง |
| `Edge_DupHeaders` | header รวมกันแล้วชื่อซ้ำ (เช่น `Value`, `Value`) → ต้อง dedupe เป็น `Value`, `Value_2` |
| `Edge_MixedTypes` | คอลัมน์ตัวเลขที่มีค่า text ปนอยู่ (`"N/A"`) → ต้องจัดเป็น text ไม่ error ตอน convert type |

### M Test Queries (ใน Power Query)

หลังจากโหลด `Sales_Report_Clean` และ `Employee_Data_Clean` แล้ว:
- โหลด `tests/Test_Sales_Report.pq` → จะแสดงตารางผล test
- โหลด `tests/Test_Employee_Data.pq` → จะแสดงตารางผล test

### ทดสอบ Daily_Update_* แบบ end-to-end (tests/fixtures)

สร้างชุดไฟล์ทดสอบด้วย `python tests/fixtures/build_fixtures.py` แล้วจะได้:
- `tests/fixtures/ReportWorkbook.xlsx` — workbook ที่มี `ConfigTable` (ชี้ `FolderPath` ไปที่
  โฟลเดอร์ fixture แล้ว, `FileExtension = .xlsx`) และ `SelectColumnTable` พร้อมใช้
- `tests/fixtures/daily_files/Daily_Update_2568.xlsx`, `Daily_Update_2569.xlsx` — ไฟล์ต้นทาง
  ที่มี sheet "Daily Update" โครงสร้างตรงกับของจริง (3 แถวขยะ + merged header 3 แถว + ข้อมูล)

วิธีทดสอบ: เปิด `ReportWorkbook.xlsx` → วางคิวรีทั้ง 6 ตัวตามลำดับในหัวข้อ setup ด้านบน →
**Refresh All** → `Daily_Update_HeaderList` ต้องแสดง header 3 ตัว (`ข้อมูลพนักงาน_รหัส/ชื่อ/จำนวนเงิน`)
และ `Daily_Update_Combined` ต้องได้ 6 แถว (ปี 2568 สามแถว + 2569 สามแถว) โดยไม่มี
Formula.Firewall error

---

## ตัวอย่างข้อมูล

### ก่อน Clean (Sales_Report)

| Product Info | | Q1 Sales | | | Q2 Sales | | |
|---|---|---|---|---|---|---|---|
| ID | Name | Revenue | | Units | Revenue | | Units |
| ID | Name | Jan | Feb | Mar | Apr | May | Jun |
| P001 | Widget A | 1000 | 1200 | 1100 | 1300 | 1400 | 1250 |
| null | null | null | null | null | null | null | null |
| P002 | Gadget B | 800 | 900 | 850 | 950 | 1000 | 920 |

### หลัง Clean

| Product Info_ID | Product Info_Name | Q1 Sales_Revenue_Jan | ... | Q2 Sales_Units_Jun |
|---|---|---|---|---|
| P001 | Widget A | 1000 | ... | 1250 |
| P002 | Gadget B | 800 | ... | 920 |

---

## หมายเหตุ

- ฟังก์ชัน `fnCleanMergedHeaders` เป็น **generic** — ใช้ได้กับทุก sheet ที่มี merge header และ null rows
- ถ้า header มีจำนวนแถวไม่เท่ากัน ให้เปลี่ยน parameter `HeaderRows`
- ถ้าต้องการตัวคั่นอื่นนอกจาก `_` ให้เปลี่ยน parameter `Separator`
- M language เป็น **case-sensitive** — ระวัง `Table.ColumnNames` (C, N ใหญ่) และ `Text.Combine` (C ใหญ่)
- คอลัมน์ที่ header ว่างทั้งหมด (ทุกแถว header เป็น null) จะไม่ถูก drop แต่จะได้ชื่อ fallback
  ตามตำแหน่งคอลัมน์ Excel เช่น `Column_A`, `Column_I` — กันไม่ให้ข้อมูลในคอลัมน์นั้นหายไปเงียบๆ
- ถ้า header รวมกันแล้วชื่อซ้ำกัน (เช่น sheet มีสอง section ชื่อ "Value") จะเติม suffix อัตโนมัติ
  เป็น `Value`, `Value_2`, `Value_3`, ...
- การ auto-detect type (Step 8) เช็คค่า **ทุกแถว** ในคอลัมน์ ไม่ใช่แค่แถวแรก — คอลัมน์ที่ควรเป็นตัวเลข
  แต่มีค่า text ปนอยู่ (เช่น `"N/A"`) จะถูกจัดเป็น text ทั้งคอลัมน์แทนที่จะ error ตอน convert
- ถ้า sheet มีแถวขยะ/ชื่อเรื่องอยู่เหนือ merged header จริง (เช่น "Daily Update" ที่ข้าม 3 แถวบนสุด
  ก่อนถึง header) ให้ใช้ `fnCleanMergedHeadersSkipJunk(SourceTable, JunkRows, HeaderRows, Separator)`
  แทน `fnCleanMergedHeaders` ตรงๆ — มันแค่ `Table.Skip` แถวขยะออกก่อนแล้วส่งต่อให้ `fnCleanMergedHeaders`
  เหมือนเดิม ไม่ต้องแก้ตัวฟังก์ชันหลักเลย sheet ที่ไม่มีแถวขยะ (Sales_Report, Employee_Data) ยังเรียก
  `fnCleanMergedHeaders` ตรงๆ เหมือนเดิม
- การตั้งค่าทั้งหมดของ `Daily_Update_Combined` / `Daily_Update_HeaderList` (โฟลเดอร์, ชื่อ sheet,
  จำนวนแถวขยะ/header, ตัวคั่น, นามสกุลไฟล์) อยู่ใน sheet "config" (Table `ConfigTable`) — แต่ละคิวรี
  อ่านตารางนี้เองแล้ว lookup ค่าผ่านฟังก์ชัน `fnConfigValue` (ไม่ผ่านคิวรี `Config` เพื่อเลี่ยง
  Formula.Firewall — ดูหัวข้อ "ปัญหา Formula.Firewall และวิธีแก้") — แก้ค่าใน Excel แล้ว Refresh
  ไม่ต้องเปิด M code (ส่วน `Sales_Report_Clean.pq` / `Employee_Data_Clean.pq` ยัง hardcode path
  ในโค้ดเหมือนเดิม ถ้าย้ายไฟล์ต้องแก้ในโค้ด)
- query ที่สแกนโฟลเดอร์ (`Daily_Update_Combined`, `Daily_Update_HeaderList`) กรองไฟล์ lock ของ Excel
  (`~$xxx.xlsm` ที่โผล่มาตอนมีคนเปิดไฟล์ค้างไว้) ออกอัตโนมัติ — refresh ได้แม้ไฟล์ต้นทางถูกเปิดอยู่
- ถ้าโฟลเดอร์ไม่มีไฟล์ `.xlsm` เลย หรือไม่มีไฟล์ไหนมี sheet "Daily Update" จะได้ error ข้อความชัดเจน
  พร้อมบอก `FolderPath` ที่ใช้อยู่ แทน error ปริศนาจาก `Table.Combine`