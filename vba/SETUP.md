# Daily Update Pipeline — ฉบับ VBA

Port ของ pipeline เดิม (Power Query) มาเป็น VBA ทั้งชุด — แนวคิด สถาปัตยกรรม และกติกาทุกข้อเหมือนเดิม
(ดูแผนภาพใน `daily_update_pipeline.html`) ต่างแค่ตัวขับเคลื่อน: จาก Refresh query → กดปุ่มรัน macro

## แนวคิด (ตาม diagram เดิม)

- **โฟลเดอร์ต้นทาง** กับ **ไฟล์ report** แยกกันคนละที่ เชื่อมกันผ่านค่า `FolderPath` ใน ConfigTable
- ลำดับจริงคือ **clean header → เลือกคอลัมน์ → รวมไฟล์** (ทำต่อไฟล์ก่อน แล้วค่อยรวม — ตอน
  รวมจึงไม่ต้องลากคอลัมน์ที่ไม่ใช้จากทุกไฟล์)
- lock file `~$...` ถูกกรองทิ้งอัตโนมัติ
- ทุกแถวมีคอลัมน์ `SourceFile` บอกว่ามาจากไฟล์ไหน (คอลัมน์ `ปี` ของฉบับ PQ ถูกตัดออกแล้ว
  — helper `YearFromFileName` ยังอยู่ใน modUtils เผื่ออยากใส่กลับ)
- ถ้าไม่มีไฟล์ที่ใช้ได้เลย macro จะ error พร้อมข้อความบอก `FolderPath` ที่ใช้อยู่
  — แปลว่า path ผิดหรือไม่มีไฟล์ไหนมี sheet เป้าหมาย ไม่ใช่โค้ดพัง
- การตั้งค่าทั้งหมดแก้จาก sheet แล้วรันใหม่ — ไม่ต้องแตะโค้ด (parameter sheet pattern)

## โมดูล

| ไฟล์ | หน้าที่ |
|---|---|
| `modUtils.bas` | helpers กลาง: เช็คค่าว่าง, เลขคอลัมน์→ตัวอักษร, หา table/sheet, list ไฟล์, อ่าน grid, ดึงปีจากชื่อไฟล์ |
| `modConfig.bas` | อ่าน `ConfigTable` (พร้อม default ทุกค่า ยกเว้น `FolderPath` ที่บังคับ) และ `SelectColumnTable` |
| `modCleanHeaders.bas` | `CleanMergedHeaders` — port ของ `fnCleanMergedHeaders(SkipJunk)`: ข้ามแถวขยะ, fill right/down merged header, รวมชื่อด้วย separator, fallback `Column_A`, dedupe `_2 _3`, ตัดแถวว่าง |
| `modCombine.bas` | **`RunDailyUpdateCombine`** — entrypoint หลัก: สแกนโฟลเดอร์ → clean ทีละไฟล์ (targeted read: อ่านเฉพาะ header block + คอลัมน์ที่เลือก) → เติม `SourceFile` → รวมลง sheet `Daily_Update_Combined` / ถ้าตั้ง `CurrentFilePath` ไว้ = archive mode |
| `modArchive.bas` | **`RunDailyUpdateArchive`** — แช่ข้อมูลไฟล์ปีเก่าทั้งโฟลเดอร์ลง sheet ซ่อน `Daily_Update_Archive` ครั้งเดียว เพื่อให้ combine รายวันเปิดแค่ไฟล์ปัจจุบัน |
| `modHeaderList.bas` | **`RunDailyUpdateHeaderList`** — ตัวช่วยดู header: อ่านไฟล์แรกที่ใช้ได้ (ถ้าตั้ง `CurrentFilePath` จะใช้ไฟล์นั้นก่อน) แล้ว list ชื่อ header + ตำแหน่งคอลัมน์ลง sheet `Daily_Update_HeaderList` |

## ติดตั้ง

1. เปิดไฟล์ report (บันทึกเป็น **.xlsm**) → `Alt+F11` เปิด VBA Editor
2. สร้าง module ใหม่ 5 อัน (Insert → Module) ตั้งชื่อตามตาราง แล้ว **copy-paste เนื้อหาแต่ละไฟล์ลงไป**
   (ข้ามบรรทัด `Attribute VB_Name = ...` บนสุด — VBE ใส่ให้เองตอนตั้งชื่อ module)
   > ใช้วิธี copy-paste แทน File → Import เพราะไฟล์ .bas ใน repo เก็บเป็น UTF-8
   > แต่ VBE import แบบ ANSI — ข้อความภาษาไทยในโค้ดจะเพี้ยนถ้า import ตรง ๆ
3. สร้าง sheet **config** → ทำตารางคอลัมน์ `Setting` / `Value` (Ctrl+T, ตั้งชื่อตาราง
   `ConfigTable`) ใส่อย่างน้อยแถว `FolderPath`
4. สร้าง sheet **select column** → ทำตารางคอลัมน์ `ColumnName` (Ctrl+T, ตั้งชื่อตาราง
   `SelectColumnTable`)
5. รัน `RunDailyUpdateHeaderList` → ดูชื่อ header ที่มีทั้งหมดใน sheet `Daily_Update_HeaderList`
6. พิมพ์ชื่อคอลัมน์ที่ต้องการลง `SelectColumnTable` → รัน `RunDailyUpdateCombine`
7. (แนะนำ) วางปุ่ม 2 ปุ่มบน sheet config ผูกกับ macro ทั้งสองตัว

## ค่าใน ConfigTable

| Setting | Default | ความหมาย |
|---|---|---|
| `FolderPath` | — (บังคับ) | โฟลเดอร์ไฟล์ต้นทาง |
| `SheetName` | `Daily Update` | ชื่อ sheet ที่ดึงจากทุกไฟล์ |
| `JunkRows` | `3` | แถวขยะบนสุดที่ข้าม |
| `HeaderRows` | `3` | จำนวนแถว merged header |
| `Separator` | `_` | ตัวคั่นตอนรวมชื่อ header หลายชั้น |
| `FileExtension` | `.xlsm` | นามสกุลไฟล์ที่สแกน |
| `CurrentFilePath` | — (ไม่ใส่ = โหมดปกติ) | full path ไฟล์ปีปัจจุบัน (ชื่อ/ที่อยู่อะไรก็ได้) — ใส่เมื่อไหร่เปิด archive mode |

ชื่อ Setting จับแบบไม่สนตัวพิมพ์เล็ก-ใหญ่และตัดช่องว่างหัวท้าย แถวไหนใช้ default ได้ลบทิ้งได้เลย
(ยกเว้น `FolderPath`)

## Archive mode — รวมเร็วขึ้นเมื่อไฟล์ปีเก่าไม่เปลี่ยนแล้ว

ปัญหา: โหมดปกติเปิดทุกไฟล์ทุกรอบ ทั้งที่ไฟล์ปีเก่า (ปิดปีแล้ว) ให้ผลเหมือนเดิมเป๊ะ
— ค่าเปิดไฟล์ (`Workbooks.Open`) คือส่วนที่แพงที่สุดของทั้ง pipeline

วิธีใช้:

1. จัดไฟล์: ไฟล์ปีเก่าทั้งหมดอยู่ใน `FolderPath` / ไฟล์ปีปัจจุบันอยู่ที่ไหนก็ได้ ชื่ออะไรก็ได้
2. เพิ่มแถว `CurrentFilePath` ใน ConfigTable ชี้ full path ของไฟล์ปีปัจจุบัน
3. รัน **`RunDailyUpdateArchive`** 1 ครั้ง — เปิดไฟล์ปีเก่าทั้งหมด clean แล้วแช่ผลลง
   sheet ซ่อน `Daily_Update_Archive` (ใช้ชุดคอลัมน์จาก `SelectColumnTable` ณ ตอนนั้น)
4. จากนี้ **`RunDailyUpdateCombine`** จะเปิดแค่ไฟล์ `CurrentFilePath` ไฟล์เดียว
   แล้วต่อท้ายข้อมูล archive → เร็วขึ้นตามจำนวนไฟล์ปีเก่าที่ไม่ต้องเปิด

ต้องรัน `RunDailyUpdateArchive` ใหม่เมื่อ:

- **แก้รายชื่อคอลัมน์ใน `SelectColumnTable`** — combine ตรวจเจอเองและจะ error
  บอกให้ rebuild ก่อน (กันตารางมีรูเงียบ ๆ: คอลัมน์ใหม่ว่างเปล่าฝั่งปีเก่า)
- **มีคนย้อนแก้ข้อมูลไฟล์ปีเก่า** — ไม่มีอะไรตรวจได้ ต้องจำเอง
- **ขึ้นปีใหม่** — ย้ายไฟล์ปีที่เพิ่งปิดเข้า `FolderPath`, ชี้ `CurrentFilePath` ไปไฟล์ใหม่,
  รัน archive อีกรอบ

ลบแถว `CurrentFilePath` ออกเมื่อไหร่ = กลับโหมดปกติ (สแกนทุกไฟล์ใน `FolderPath`)
โดย sheet archive ที่ซ่อนไว้ไม่ถูกใช้แต่ไม่รบกวนอะไร

## ข้อแตกต่างจากฉบับ Power Query

- ไม่มีขั้น auto-detect type: VBA อ่านค่าจาก cell ได้ native type อยู่แล้ว (ตัวเลขเป็นตัวเลข
  วันที่เป็นวันที่) ไม่ต้องเดา type จาก text เหมือน M
- ไฟล์ที่เปิดไม่ได้/ไม่มี sheet เป้าหมาย/ไม่มีคอลัมน์ที่เลือกเลย จะถูก **ข้ามพร้อมแจ้งเตือน**
  ตอนจบ (แสดงใน MsgBox สรุป) แทนที่จะทำให้ทั้ง run ล้ม
- โหมด Archive/CurrentYear (incremental refresh ของฉบับ PQ v2) **ยังไม่ port** — ฉบับ VBA
  ตามขอบเขตของ diagram เดิม: รวมทุกไฟล์ในโฟลเดอร์ทุกครั้งที่รัน
