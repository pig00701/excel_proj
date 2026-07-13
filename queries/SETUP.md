# SETUP — ติดตั้ง Daily Update pipeline จากศูนย์ (v2)

ใช้เมื่อล้าง query ใน workbook ทิ้งทั้งหมดแล้วลงใหม่ ทำตามลำดับเป๊ะ ๆ — ลำดับสำคัญเพราะ query หลังอ้างฟังก์ชันของ query ก่อน

## ขั้นที่ 0: ล้างของเก่า

ใน Excel: Data > Queries & Connections > ลบ query เก่าที่เกี่ยวกับ Daily Update ทั้งหมด (ตาราง output เก่าบน sheet จะกลายเป็นตารางธรรมดา ลบทิ้งได้)

## ขั้นที่ 1: สร้างตารางใน workbook (2 ตาราง)

ทั้งคู่สร้างด้วย: พิมพ์หัวตาราง + ข้อมูล → เลือกช่วง → **Ctrl+T** (มี header) → ตั้งชื่อที่ **Table Design > Table Name**

**ตาราง `ConfigTable`** (ชีท config):

| Setting | Value | หมายเหตุ |
|---|---|---|
| FolderPath | C:\...\โฟลเดอร์ไฟล์ต้นทาง | จำเป็น |
| SheetName | Daily Update | ไม่ใส่ = ค่านี้ |
| JunkRows | 3 | แถวขยะเหนือ header |
| HeaderRows | 3 | แถว merged header |
| Separator | _ | ตัวคั่นชื่อ header |
| FileExtension | .xlsm | |
| CurrentYear | 2569 | ใส่ = เปิดโหมด archive/current (แนะนำ) |

**ตาราง `SelectColumnTable`** (ชีท select column): คอลัมน์เดียวหัวชื่อ `ColumnName` หนึ่งแถวต่อชื่อคอลัมน์ (ชื่อหลังรวม header แล้ว — ได้จาก Daily_Update_HeaderList ในขั้นที่ 4)

## ขั้นที่ 2: ลง query ตามลำดับ

แต่ละตัว: Data > Get Data > From Other Sources > **Blank Query** > Advanced Editor > วางโค้ดจากไฟล์ .pq > ตั้งชื่อ query ให้**ตรงกับชื่อไฟล์เป๊ะ** (ไม่มี .pq) > Close & Load To… ตามตาราง:

| ลำดับ | Query | Load แบบ |
|---|---|---|
| 1 | `fnConfigValue` | Connection Only |
| 2 | `fnColumnLetterToIndex` | Connection Only |
| 3 | `fnCleanMergedHeaders` | Connection Only |
| 4 | `fnCleanMergedHeadersSkipJunk` | Connection Only |
| 5 | `fnDailyUpdateConfig` | Connection Only |
| 6 | `fnDailyUpdateCombine` | Connection Only |
| 7 | `Daily_Update_FileCheck` | Table (ชีทใหม่) |
| 8 | `Daily_Update_HeaderList` | Table (ชีทใหม่) |
| 9 | `Daily_Update_Archive` | **Table (ชีทใหม่)** — จำเป็น ห้าม Connection Only |
| 10 | `Daily_Update_Combined` | Table (ชีทใหม่) |

## ขั้นที่ 3: ปิด refresh อัตโนมัติ 2 ตัว (สำคัญที่สุด — พลาดตรงนี้ = ช้าเท่าเดิม)

Data > Queries & Connections > คลิกขวา query > **Properties…** > แท็บ Usage > ติ๊กออก **"Refresh this connection on Refresh All"**:

- [ ] `Daily_Update_Archive` — รันเองปีละครั้ง (ตอนเปลี่ยน CurrentYear) หรือตอนแก้ SelectColumnTable
- [ ] `Daily_Update_HeaderList` — รันเองเฉพาะตอนไฟล์ต้นทางเปลี่ยนโครง header

## ขั้นที่ 4: เติม SelectColumnTable

ดูตาราง output ของ `Daily_Update_HeaderList` → copy ชื่อจากคอลัมน์ "ชื่อ header" ที่ต้องการเก็บ ไปวางใน SelectColumnTable (ลำดับแถว = ลำดับคอลัมน์ในผลลัพธ์ / ไม่ต้องใส่ SourceFile กับ ปี — ระบบเพิ่มให้เอง)

## ขั้นที่ 5: ตรวจก่อนรันจริง

Refresh `Daily_Update_FileCheck` (เร็วมาก ไม่เปิดไฟล์) — เช็คคอลัมน์ "สถานะ" ว่า:
- ไฟล์ปีเก่าขึ้น `Archive (ปีเก่า)`
- ไฟล์ปีนี้ขึ้น `Combined (ปีปัจจุบัน)`
- ไม่มีไฟล์ที่ควรดึงไปโผล่ฝั่ง "ข้าม"

ถ้าสถานะไม่ตรงที่คาด แก้ ConfigTable ก่อน — ยังไม่ต้องเสียเวลา refresh ตัวจริง

## ขั้นที่ 6: รันครั้งแรก

1. Refresh `Daily_Update_Archive` เอง 1 ครั้ง — **รอบนี้ช้าเป็นรอบสุดท้าย** (parse ไฟล์ปีเก่าทุกไฟล์)
2. Refresh All — ตั้งแต่รอบนี้ parse เฉพาะไฟล์ปีปัจจุบัน

## งานประจำ

| เหตุการณ์ | ทำอะไร |
|---|---|
| ใช้งานปกติ | Refresh All อย่างเดียว |
| ขึ้นปีใหม่ | แก้ CurrentYear ใน ConfigTable → refresh `Daily_Update_Archive` เอง 1 ครั้ง |
| แก้รายชื่อคอลัมน์ | แก้ SelectColumnTable → refresh `Daily_Update_Archive` เอง 1 ครั้ง → Refresh All |
| header ต้นทางเปลี่ยน | refresh `Daily_Update_HeaderList` เอง → ปรับ SelectColumnTable |
| refresh ช้าผิดปกติ | ดู `Daily_Update_FileCheck` ก่อนว่าไฟล์ไหนอยู่ฝั่ง Combined บ้าง |

## Troubleshooting

- **error "ยังไม่มีตาราง Daily_Update_Archive"** → ขั้นที่ 2 ลำดับ 9 ต้อง Load เป็น Table ไม่ใช่ Connection Only และชื่อตารางต้องเป็น `Daily_Update_Archive`
- **Refresh All ยังช้า** → เกือบทุกครั้งคือขั้นที่ 3 ไม่ครบ เช็ค Properties ของ Archive/HeaderList อีกรอบ แล้วดู FileCheck ว่าฝั่ง `Combined` มีไฟล์เกินคาดไหม
- **firewall error "references other queries..."** → File > Options and Settings > Query Options > Privacy > Ignore Privacy Levels (ข้อมูลเป็นไฟล์ local ทั้งหมด)
- **คอลัมน์หายจากบางไฟล์** → ชื่อใน SelectColumnTable สะกดไม่ตรงกับ header ไฟล์นั้น (MissingField.Ignore ข้ามให้เงียบ ๆ) — เทียบกับ HeaderList
