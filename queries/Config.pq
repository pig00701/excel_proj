// ============================================================================
// Config.pq  (v2 — thin caller over fnDailyUpdateConfig)
// ============================================================================
// Exposes the parsed ConfigTable as a record for queries that do NOT touch
// a data source themselves (browse it in the editor to sanity-check what
// each setting resolved to). Queries that call Folder.Files/Excel.Workbook
// (Daily_Update_Combined / _Archive / _HeaderList / _FileCheck) must NOT
// reference this query — they call fnDailyUpdateConfig directly instead,
// or the Formula.Firewall raises "references other queries... may not
// directly access a data source". See fnConfigValue.pq for details.
//
// ConfigTable layout (Ctrl+T, name it "ConfigTable"):
//   Setting        | Value                 | หมายเหตุ (optional)
//   FolderPath     | C:\...\DailyFiles     | โฟลเดอร์ไฟล์ต้นทาง (จำเป็น)
//   SheetName      | Daily Update          | ชื่อ sheet ที่ดึง
//   JunkRows       | 3                     | แถวขยะบนสุดที่ข้าม
//   HeaderRows     | 3                     | แถว merged header
//   Separator      | _                     | ตัวคั่นชื่อ header
//   FileExtension  | .xlsm                 | นามสกุลไฟล์ที่สแกน
//   YearFrom       | 2566                  | (optional) ปีแรกที่ดึง
//   YearTo         | 2567                  | (optional) ปีสุดท้ายที่ดึง
//   CurrentYear    | 2569                  | (optional) เปิดโหมด archive/current
//
// Rules:
//   - Setting names match case-insensitively, ignoring surrounding spaces.
//   - Only FolderPath is required; everything else falls back to the
//     defaults above (YearFrom/YearTo/CurrentYear default to "not set").
//   - Years compare against the 4-digit year at the END of each file name
//     (e.g. "Daily_Update_2566.xlsm" -> 2566).
//   - CurrentYear activates the archive/current split — see
//     Daily_Update_Archive.pq and SETUP.md.
//
// PREREQUISITE: fnConfigValue and fnDailyUpdateConfig loaded as queries.
// ============================================================================

let
    Result = fnDailyUpdateConfig(Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content])
in
    Result
