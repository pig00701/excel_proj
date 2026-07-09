// ============================================================================
// Daily_Update_Combined.pq
// ============================================================================
// Combines the "Daily Update" sheet from every .xlsm file in a folder into
// one table. Rows 1-3 of the sheet are junk/title rows (not header data),
// rows 4-6 are the real 3-row merged header, data starts at row 7 — so the
// first 3 rows are skipped before handing off to fnCleanMergedHeaders. That
// skip-then-clean step lives in fnCleanMergedHeadersSkipJunk, a small
// wrapper kept separate from fnCleanMergedHeaders so Sales_Report_Clean /
// Employee_Data_Clean (which have no junk rows) are unaffected.
//
// PREREQUISITE: Load fnCleanMergedHeaders.pq (as "fnCleanMergedHeaders"),
// fnCleanMergedHeadersSkipJunk.pq (as "fnCleanMergedHeadersSkipJunk"), and
// fnConfigValue.pq (as "fnConfigValue") as queries first.
//
// All settings (folder path, sheet name, junk/header row counts, separator,
// file extension) come from the "ConfigTable" Excel Table in this workbook —
// edit the config sheet and Refresh; no M code changes needed. See
// Config.pq for the table layout and defaults.
//
// Config values are resolved LOCALLY here (via fnConfigValue) rather than
// through the "Config" query, because this query also calls
// Folder.Files/Excel.Workbook — mixing a data-source call with a value
// pulled from another query trips Power Query's Formula.Firewall
// ("references other queries... may not directly access a data source").
// See fnConfigValue.pq for details.
//
// A "SourceFile" column is added to each row so you can trace which .xlsm
// file it came from after combining. Files missing a "Daily Update" sheet
// are skipped rather than erroring out.
//
// A "ปี" (year) column is also added, extracted from the file name: source
// files are yearly, and MUST end with a 4-digit year right before the
// extension (e.g. "Daily_Update_2566.xlsm" → 2566, "name_2556.xlsm" →
// 2556). Extraction is just "take the last 4 characters of the base
// name" — this relies on that naming convention being followed exactly.
// A file whose last 4 characters aren't digits gets null in this column
// rather than an error.
//
// The list of columns to keep is read from an Excel Table named
// "SelectColumnTable" (a single column "ColumnName", one row per header
// name to keep) living in the SAME workbook as this query. Edit that table
// directly in Excel (e.g. on a sheet tab called "select column") — no need
// to touch this M code to change which columns come through. Anything not
// listed is dropped, so the combined table doesn't balloon with columns
// you don't need. MissingField.Ignore means a file missing one of these
// columns just won't have it — no error.
//
// SETUP: on a sheet in this workbook, list header names in a column with
// a "ColumnName" heading, select the range, Insert > Table (Ctrl+T), and
// name the table "SelectColumnTable" (Table Design > Table Name).
// ============================================================================

let
    ConfigTable = Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content],
    FolderPathRaw = fnConfigValue(ConfigTable, "FolderPath", null),
    FolderPath = if FolderPathRaw = null then
        error [
            Reason = "MissingConfig",
            Message = "ยังไม่ได้ตั้งค่า FolderPath ใน ConfigTable — เพิ่มแถว Setting = FolderPath, Value = โฟลเดอร์ที่เก็บไฟล์ต้นทาง"
        ]
    else
        Text.Trim(Text.From(FolderPathRaw)),
    SheetName = Text.From(fnConfigValue(ConfigTable, "SheetName", "Daily Update")),
    ExtensionRaw = Text.Lower(Text.Trim(Text.From(fnConfigValue(ConfigTable, "FileExtension", ".xlsm")))),
    FileExtension = if Text.StartsWith(ExtensionRaw, ".") then ExtensionRaw else "." & ExtensionRaw,
    JunkRows = Number.From(fnConfigValue(ConfigTable, "JunkRows", 3)),
    HeaderRows = Number.From(fnConfigValue(ConfigTable, "HeaderRows", 3)),
    Separator = Text.From(fnConfigValue(ConfigTable, "Separator", "_")),

    SelectColumnTable = Excel.CurrentWorkbook(){[Name = "SelectColumnTable"]}[Content],
    SelectedColumns = List.Buffer(List.RemoveNulls(Table.Column(SelectColumnTable, "ColumnName"))),

    // Extract the year from a file name that MUST end in a 4-digit year
    // right before the extension, e.g. "Daily_Update_2566.xlsm" → 2566.
    // Just takes the last 4 characters of the base name — no year-shaped
    // scanning needed since the naming convention guarantees the position.
    ExtractYear = (fileName as text) as nullable number =>
        let
            baseName = Text.BeforeDelimiter(fileName, ".", {0, RelativePosition.FromEnd}),
            last4 = Text.End(baseName, 4),
            year = try Number.From(last4) otherwise null
        in
            year,

    Source = Folder.Files(FolderPath),
    // "~$xxx.xlsm" are hidden Excel lock files that appear while a workbook
    // is open — same extension, but not real workbooks: Excel.Workbook()
    // errors on them. Exclude by name prefix so a refresh doesn't fail just
    // because someone has a source file open.
    XlsmFiles = Table.SelectRows(
        Source,
        each Text.Lower([Extension]) = FileExtension and not Text.StartsWith([Name], "~$")
    ),

    AddedSheetData = Table.AddColumn(XlsmFiles, "SheetData", each
        let
            FileName = [Name],
            Workbook = Excel.Workbook([Content], null, true),
            DailyUpdateRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
            RawTable = if Table.RowCount(DailyUpdateRow) > 0 then DailyUpdateRow{0}[Data] else null,
            Cleaned = if RawTable <> null then fnCleanMergedHeadersSkipJunk(RawTable, JunkRows, HeaderRows, Separator) else null,
            Trimmed = if Cleaned <> null then Table.SelectColumns(Cleaned, SelectedColumns, MissingField.Ignore) else null,
            WithSource = if Trimmed <> null then Table.AddColumn(Trimmed, "SourceFile", each FileName) else null,
            FileYear = ExtractYear(FileName),
            WithYear = if WithSource <> null then Table.AddColumn(WithSource, "ปี", each FileYear, type nullable number) else null
        in
            WithYear
    ),

    ValidOnly = Table.SelectRows(AddedSheetData, each [SheetData] <> null),
    // Fail with a readable message instead of Table.Combine's cryptic error
    // when the folder has no usable files (wrong FolderPath, or no file has
    // a "Daily Update" sheet).
    Combined = if List.Count(ValidOnly[SheetData]) = 0 then
        error [
            Reason = "NoData",
            Message = "ไม่พบไฟล์ " & FileExtension & " ที่มี sheet """ & SheetName & """ ในโฟลเดอร์ — ตรวจสอบ ConfigTable: FolderPath = " & FolderPath
        ]
    else
        Table.Combine(ValidOnly[SheetData])
in
    Combined
