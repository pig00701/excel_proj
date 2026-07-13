// ============================================================================
// Config.pq
// ============================================================================
// Reads all pipeline settings from an Excel Table named "ConfigTable" in the
// SAME workbook (report file) as the queries, so everything is adjustable
// from a sheet — edit a cell, hit Refresh, no M code involved.
//
// ConfigTable layout (make it with Ctrl+T, name it "ConfigTable"):
//   Setting        | Value                          | หมายเหตุ (optional)
//   FolderPath     | C:\...\DailyFiles              | โฟลเดอร์ไฟล์ต้นทาง
//   SheetName      | Daily Update                   | ชื่อ sheet ที่ดึง
//   JunkRows       | 3                              | แถวขยะบนสุดที่ข้าม
//   HeaderRows     | 3                              | แถว merged header
//   Separator      | _                              | ตัวคั่นชื่อ header
//   FileExtension  | .xlsm                          | นามสกุลไฟล์ที่สแกน
//   YearFrom       | 2566                           | (optional) ปีแรกที่ดึง
//   YearTo         | 2567                           | (optional) ปีสุดท้ายที่ดึง
//   CurrentYear    | 2569                           | (optional) เปิดโหมด archive/current
//
// YearFrom/YearTo filter source FILES by the 4-digit year at the end of the
// file name (e.g. "Daily_Update_2566.xlsm") before the file is ever opened —
// use them to skip old yearly files entirely. Either bound may be set alone;
// omit both to combine every file (the default). Files whose name doesn't
// end in a 4-digit year are still included even when a filter is set.
//
// CurrentYear activates the archive/current split (incremental-refresh
// pattern): Daily_Update_Archive combines years BELOW CurrentYear into a
// static table (loaded once, refresh disabled), and Daily_Update_Combined
// then only re-opens files of CurrentYear onward, prepending that table.
// Omit CurrentYear to keep the original everything-per-refresh behavior.
// See Daily_Update_Archive.pq for setup + new-year rollover steps.
//
// Rules:
//   - Setting names match case-insensitively and ignore surrounding spaces.
//   - Every setting except FolderPath has a default (the values shown
//     above; YearFrom/YearTo default to "no filter"), so a missing row just
//     falls back — only FolderPath is required and errors with a clear
//     message when absent.
//   - Extra columns (like หมายเหตุ) are ignored — add as many as you like.
//   - FileExtension may be typed with or without the leading dot.
//
// Load this as a query named "Config". Other queries read settings as
// record fields: Config[FolderPath], Config[SheetName], Config[JunkRows],
// Config[HeaderRows], Config[Separator], Config[FileExtension].
//
// IMPORTANT: this query is safe to reference ONLY from queries that do not
// also call a data-source function (Folder.Files, Excel.Workbook, etc.) in
// the same step chain. Daily_Update_HeaderList and Daily_Update_Combined do
// call Folder.Files, so they do NOT use this query — they read ConfigTable
// locally and call fnConfigValue directly instead, to avoid tripping the
// Formula.Firewall "references other queries... may not directly access a
// data source" error. See fnConfigValue.pq for why.
//
// PREREQUISITE: Load fnConfigValue.pq (as "fnConfigValue") as a query first.
// ============================================================================

let
    RawTable = Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content],

    FolderPathRaw = fnConfigValue(RawTable, "FolderPath", null),
    FolderPath = if FolderPathRaw = null then
        error [
            Reason = "MissingConfig",
            Message = "ยังไม่ได้ตั้งค่า FolderPath ใน ConfigTable — เพิ่มแถว Setting = FolderPath, Value = โฟลเดอร์ที่เก็บไฟล์ต้นทาง"
        ]
    else
        Text.Trim(Text.From(FolderPathRaw)),

    ExtensionRaw = Text.Lower(Text.Trim(Text.From(fnConfigValue(RawTable, "FileExtension", ".xlsm")))),
    FileExtension = if Text.StartsWith(ExtensionRaw, ".") then ExtensionRaw else "." & ExtensionRaw,

    YearFromRaw = fnConfigValue(RawTable, "YearFrom", null),
    YearToRaw = fnConfigValue(RawTable, "YearTo", null),
    CurrentYearRaw = fnConfigValue(RawTable, "CurrentYear", null),

    Result = [
        FolderPath = FolderPath,
        SheetName = Text.From(fnConfigValue(RawTable, "SheetName", "Daily Update")),
        JunkRows = Number.From(fnConfigValue(RawTable, "JunkRows", 3)),
        HeaderRows = Number.From(fnConfigValue(RawTable, "HeaderRows", 3)),
        Separator = Text.From(fnConfigValue(RawTable, "Separator", "_")),
        FileExtension = FileExtension,
        YearFrom = if YearFromRaw = null then null else Number.From(YearFromRaw),
        YearTo = if YearToRaw = null then null else Number.From(YearToRaw),
        CurrentYear = if CurrentYearRaw = null then null else Number.From(CurrentYearRaw)
    ]
in
    Result
