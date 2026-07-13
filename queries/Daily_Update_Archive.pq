// ============================================================================
// Daily_Update_Archive.pq
// ============================================================================
// The STATIC half of the archive/current split (see the incremental-refresh
// pattern: separate historical static data from current refresh data).
// Combines the "Daily Update" sheet from every CLOSED-year .xlsm file —
// years BEFORE ConfigTable's "CurrentYear" — using the exact same cleaning
// pipeline as Daily_Update_Combined.
//
// WHY THIS EXISTS: parsing an 80-100 MB .xlsm costs minutes per file per
// refresh, but closed years never change. Load this query's result to a
// worksheet ONCE, then stop refreshing it — Daily_Update_Combined reads the
// loaded table back (cheap, no file parsing) and only ever re-opens the
// current-year file(s).
//
// SETUP (one time):
//   1. Add a "CurrentYear" row to ConfigTable (e.g. Setting = CurrentYear,
//      Value = 2569) — years STRICTLY BELOW this are archived here.
//   2. Load this query TO A TABLE on its own worksheet. The Excel table
//      must be named "Daily_Update_Archive" (Excel names it after the
//      query automatically — check Table Design > Table Name).
//   3. Data > Queries & Connections > right-click this query > Properties >
//      UNTICK "Refresh this connection on Refresh All".
//
// NEW YEAR ROLLOVER (once a year):
//   1. Change ConfigTable's CurrentYear to the new year.
//   2. Manually refresh THIS query once (it now swallows the year that just
//      closed), then Refresh All as usual.
//
// IF SelectColumnTable CHANGES: the archived table keeps its old columns —
// manually refresh this query once to rebuild it with the new column list.
//
// Files whose name does NOT end in a 4-digit year are NOT archived — they
// stay on the live side (Daily_Update_Combined) so a changing off-convention
// file can never go silently stale inside a frozen archive.
//
// PREREQUISITE: same as Daily_Update_Combined — fnCleanMergedHeaders,
// fnCleanMergedHeadersSkipJunk, fnConfigValue loaded as queries. Config is
// resolved locally via fnConfigValue for the same Formula.Firewall reason.
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
    YearFromRaw = fnConfigValue(ConfigTable, "YearFrom", null),
    YearFrom = if YearFromRaw = null then null else Number.From(YearFromRaw),
    // CurrentYear is REQUIRED here — without it "closed year" has no meaning.
    CurrentYearRaw = fnConfigValue(ConfigTable, "CurrentYear", null),
    CurrentYear = if CurrentYearRaw = null then
        error [
            Reason = "MissingConfig",
            Message = "ยังไม่ได้ตั้งค่า CurrentYear ใน ConfigTable — เพิ่มแถว Setting = CurrentYear, Value = ปีปัจจุบัน (พ.ศ. ตามชื่อไฟล์ เช่น 2569) แล้วปีที่ต่ำกว่านี้จะถูกเก็บเข้า archive"
        ]
    else
        Number.From(CurrentYearRaw),

    SelectColumnTable = Excel.CurrentWorkbook(){[Name = "SelectColumnTable"]}[Content],
    SelectedColumns = List.Buffer(List.RemoveNulls(Table.Column(SelectColumnTable, "ColumnName"))),

    ExtractYear = (fileName as text) as nullable number =>
        let
            baseName = Text.BeforeDelimiter(fileName, ".", {0, RelativePosition.FromEnd}),
            last4 = Text.End(baseName, 4),
            year = try Number.From(last4) otherwise null
        in
            year,

    Source = Folder.Files(FolderPath),
    XlsmFiles = Table.SelectRows(
        Source,
        each Text.Lower([Extension]) = FileExtension and not Text.StartsWith([Name], "~$")
    ),

    // Archive scope: named-year files STRICTLY BELOW CurrentYear only
    // (respecting YearFrom as the lower bound when set). Null-year files are
    // deliberately excluded — they belong to the live side.
    ArchiveFiles = Table.SelectRows(XlsmFiles, each
        let
            y = ExtractYear([Name])
        in
            y <> null
            and y < CurrentYear
            and (YearFrom = null or y >= YearFrom)
    ),

    AddedSheetData = Table.AddColumn(ArchiveFiles, "SheetData", each
        let
            FileName = [Name],
            Workbook = Excel.Workbook(Binary.Buffer([Content]), null, true),
            DailyUpdateRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
            RawTable = if Table.RowCount(DailyUpdateRow) > 0 then DailyUpdateRow{0}[Data] else null,
            ColumnCappedTable = if RawTable <> null then
                Table.SelectColumns(RawTable, List.FirstN(Table.ColumnNames(RawTable), 300))
            else
                null,
            Cleaned = if ColumnCappedTable <> null then fnCleanMergedHeadersSkipJunk(ColumnCappedTable, JunkRows, HeaderRows, Separator, SelectedColumns) else null,
            WithSource = if Cleaned <> null then Table.AddColumn(Cleaned, "SourceFile", each FileName) else null,
            FileYear = ExtractYear(FileName),
            WithYear = if WithSource <> null then Table.AddColumn(WithSource, "ปี", each FileYear, type nullable number) else null
        in
            WithYear
    ),

    ValidOnly = Table.SelectRows(AddedSheetData, each [SheetData] <> null),
    Combined = if List.Count(ValidOnly[SheetData]) = 0 then
        error [
            Reason = "NoData",
            Message = "ไม่พบไฟล์ปีเก่า (ปี < " & Text.From(CurrentYear) & ") ที่มี sheet """ & SheetName & """ ในโฟลเดอร์ — ตรวจสอบ FolderPath / CurrentYear ใน ConfigTable"
        ]
    else
        Table.Combine(ValidOnly[SheetData])
in
    Combined
