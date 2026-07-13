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
// file extension, optional YearFrom/YearTo file filter, optional
// CurrentYear) come from the "ConfigTable" Excel Table in this workbook —
// edit the config sheet and Refresh; no M code changes needed. See
// Config.pq for the table layout and defaults.
//
// ARCHIVE/CURRENT SPLIT (optional; set "CurrentYear" in ConfigTable):
// closed-year files never change, so re-parsing 80-100 MB .xlsm every
// refresh is wasted time. With CurrentYear set, this query only opens
// files of that year onward and prepends the static "Daily_Update_Archive"
// table (built once by the Daily_Update_Archive query — see that file for
// setup and new-year rollover steps). Without CurrentYear everything works
// exactly as before.
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
// columns just won't have it — no error. The list is handed INTO
// fnCleanMergedHeadersSkipJunk so unwanted columns are dropped right after
// header names are built, before the expensive null-row / type-detection
// scans — not trimmed afterwards.
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
    // Optional year range: files are yearly and named "..._<year>.xlsm", so
    // files outside [YearFrom, YearTo] are excluded BEFORE Excel.Workbook
    // ever opens them — old years cost nothing. Both default to null = no
    // filter (original behavior). Set either bound alone or both.
    YearFromRaw = fnConfigValue(ConfigTable, "YearFrom", null),
    YearFrom = if YearFromRaw = null then null else Number.From(YearFromRaw),
    YearToRaw = fnConfigValue(ConfigTable, "YearTo", null),
    YearTo = if YearToRaw = null then null else Number.From(YearToRaw),
    // ARCHIVE/CURRENT SPLIT MODE (optional — the incremental-refresh
    // pattern: closed years never change, so stop re-parsing them).
    // Set "CurrentYear" in ConfigTable to activate:
    //   - this query then only opens files whose year >= CurrentYear (plus
    //     files with no year in the name, which always stay live),
    //   - and prepends the static "Daily_Update_Archive" Excel table (the
    //     one-time output of the Daily_Update_Archive query) for history.
    // Not set = original behavior: every file in [YearFrom, YearTo].
    CurrentYearRaw = fnConfigValue(ConfigTable, "CurrentYear", null),
    CurrentYear = if CurrentYearRaw = null then null else Number.From(CurrentYearRaw),
    SplitMode = CurrentYear <> null,

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

    // Year filter — runs on file NAMES only; filtered-out files are never
    // opened. A file whose name doesn't end in a 4-digit year (ExtractYear
    // = null) is KEPT in every mode: it's off-convention, but silently
    // dropping it risks losing data — it comes through with ปี = null.
    //
    // Split mode: only CURRENT files (year >= CurrentYear, YearTo still an
    // upper bound) — closed years come from the Daily_Update_Archive table
    // instead (see ArchiveTable below), so their .xlsm are never parsed.
    // Normal mode: optional [YearFrom, YearTo] range as before.
    YearFilteredFiles = if SplitMode then
        Table.SelectRows(XlsmFiles, each
            let
                y = ExtractYear([Name])
            in
                y = null
                or (y >= CurrentYear and (YearTo = null or y <= YearTo))
        )
    else if YearFrom = null and YearTo = null then
        XlsmFiles
    else
        Table.SelectRows(XlsmFiles, each
            let
                y = ExtractYear([Name])
            in
                y = null
                or ((YearFrom = null or y >= YearFrom) and (YearTo = null or y <= YearTo))
        ),

    AddedSheetData = Table.AddColumn(YearFilteredFiles, "SheetData", each
        let
            FileName = [Name],
            // Binary.Buffer reads each file's bytes into memory once, so
            // any downstream re-evaluation hits the buffer instead of
            // re-reading the .xlsm from disk.
            Workbook = Excel.Workbook(Binary.Buffer([Content]), null, true),
            DailyUpdateRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
            RawTable = if Table.RowCount(DailyUpdateRow) > 0 then DailyUpdateRow{0}[Data] else null,
            // Column-count safety cap: real sheets here run ~200-something
            // columns. An Excel "used range" bloated out to thousands of
            // ghost columns gets dragged through FillDown/header-combine/
            // type-detection for nothing — 300 gives headroom over the real
            // column count without leaving the door open to thousands. Rows
            // are NOT capped here (unlike Daily_Update_HeaderList) because
            // this query needs every real data row, not just column names.
            ColumnCappedTable = if RawTable <> null then
                Table.SelectColumns(RawTable, List.FirstN(Table.ColumnNames(RawTable), 300))
            else
                null,
            // SelectedColumns is passed INTO the cleaner so it can trim to
            // just those columns right after building header names — before
            // its null-row removal and per-value type detection ever touch
            // the ~300-column width. (Previously the trim happened out here,
            // after all that work was already done on every column.)
            Cleaned = if ColumnCappedTable <> null then fnCleanMergedHeadersSkipJunk(ColumnCappedTable, JunkRows, HeaderRows, Separator, SelectedColumns) else null,
            WithSource = if Cleaned <> null then Table.AddColumn(Cleaned, "SourceFile", each FileName) else null,
            FileYear = ExtractYear(FileName),
            WithYear = if WithSource <> null then Table.AddColumn(WithSource, "ปี", each FileYear, type nullable number) else null
        in
            WithYear
    ),

    ValidOnly = Table.SelectRows(AddedSheetData, each [SheetData] <> null),

    // Split mode: read the static history back from the workbook's own
    // "Daily_Update_Archive" table — no file parsing, effectively free.
    // Reading Excel.CurrentWorkbook here (instead of referencing the
    // Daily_Update_Archive QUERY) keeps the Formula.Firewall happy: this
    // query already mixes Excel.CurrentWorkbook + Folder.Files in-query,
    // which is fine — referencing another query would not be.
    // Missing table is a hard error, NOT silently skipped: returning
    // current-year-only data while looking complete would be worse than
    // failing loudly.
    ArchiveTable = if not SplitMode then
        null
    else
        try Excel.CurrentWorkbook(){[Name = "Daily_Update_Archive"]}[Content]
        otherwise error [
            Reason = "MissingArchive",
            Message = "ตั้ง CurrentYear ไว้แต่ยังไม่มีตาราง Daily_Update_Archive ใน workbook — โหลด query Daily_Update_Archive ลง sheet ครั้งแรกก่อน (แล้วปิด Refresh on Refresh All ของมัน) หรือลบแถว CurrentYear ออกจาก ConfigTable เพื่อกลับโหมดเดิม"
        ],

    FreshTables = ValidOnly[SheetData],
    AllTables = (if ArchiveTable <> null then {ArchiveTable} else {}) & FreshTables,
    // Fail with a readable message instead of Table.Combine's cryptic error
    // when there is nothing at all to combine (wrong FolderPath, or no file
    // has a "Daily Update" sheet). In split mode the archive alone counts —
    // e.g. at new-year rollover before the new year's file exists.
    Combined = if List.Count(AllTables) = 0 then
        error [
            Reason = "NoData",
            Message = "ไม่พบไฟล์ " & FileExtension & " ที่มี sheet """ & SheetName & """ ในโฟลเดอร์ — ตรวจสอบ ConfigTable: FolderPath = " & FolderPath
        ]
    else
        Table.Combine(AllTables)
in
    Combined
