// ============================================================================
// Daily_Update_HeaderList.pq
// ============================================================================
// Reference table showing every cleaned header name from the "Daily Update"
// sheet alongside its original Excel column letter (A, B, C, ...). Reads
// only ONE .xlsm file — the one whose name sorts LAST (relies on file names
// ending in a 4-digit year, e.g. Daily_Update_2569.xlsm, so "sorts last"
// means "latest year") — meant as a lookup so you know what to type into
// the "select column" / SelectColumnTable sheet used by
// Daily_Update_Combined.pq. It intentionally does NOT apply the
// SelectColumnTable filter, so it always shows the FULL list of available
// headers, not just the ones currently selected.
//
// PREREQUISITE: Load fnCleanMergedHeaders.pq (as "fnCleanMergedHeaders"),
// fnConfigValue.pq (as "fnConfigValue"), and fnColumnLetterToIndex.pq (as
// "fnColumnLetterToIndex") as queries first. Does NOT need
// fnCleanMergedHeadersSkipJunk — the junk-row skip happens inline below
// (diagnostic test for a slowdown isolated to the "Cleaned" step; see
// SkippedJunkRows).
//
// All settings (folder path, sheet name, junk/header row counts, separator,
// file extension) come from the same "ConfigTable" that Daily_Update_Combined
// uses — one config sheet drives both queries. Config values are resolved
// LOCALLY here (via fnConfigValue) rather than through the "Config" query,
// because this query also calls Folder.Files/Excel.Workbook — mixing a
// data-source call with a value from another query trips Power Query's
// Formula.Firewall ("references other queries... may not directly access a
// data source"). See fnConfigValue.pq for details.
//
// Output columns:
//   ชื่อ header      — cleaned header name after fnCleanMergedHeaders
//   ตำแหน่ง column   — TRUE Excel column letter in the source file (see below)
//
// WHY "ตำแหน่ง column" NEEDS A CONFIGURABLE OFFSET:
// Excel.Workbook()[Data] returns a sheet's USED RANGE, which doesn't
// necessarily start at column A — if column A (and beyond) never had any
// content or formatting, Excel drops it from the used range entirely, so
// this query's first raw column is actually column B (or further) in the
// real file, with no way to ask Power Query where the used range truly
// starts. Set ConfigTable's "FirstColumnLetter" to whatever column the
// sheet's data ACTUALLY starts at (open the source file and check) — default
// is "A" (no offset) if left blank.
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
    // How many columns short of true column A the used range starts at —
    // see the header comment above for why this exists.
    FirstColumnLetter = Text.From(fnConfigValue(ConfigTable, "FirstColumnLetter", "A")),
    FirstColumnOffset = fnColumnLetterToIndex(FirstColumnLetter) - 1,

    Source = Folder.Files(FolderPath),
    // Exclude "~$" Excel lock files (left while a workbook is open) — same
    // extension but not real workbooks.
    XlsmFiles = Table.SelectRows(
        Source,
        each Text.Lower([Extension]) = FileExtension and not Text.StartsWith([Name], "~$")
    ),
    // Sort by file name descending so "first file" reliably means "latest
    // year" (file names end in a 4-digit year, e.g. Daily_Update_2569.xlsm)
    // rather than whatever order the OS/Folder.Files happens to return.
    SortedFiles = Table.Sort(XlsmFiles, {{"Name", Order.Descending}}),
    FirstFile = if Table.RowCount(SortedFiles) = 0 then
        error [
            Reason = "NoData",
            Message = "ไม่พบไฟล์ " & FileExtension & " ในโฟลเดอร์ — ตรวจสอบ ConfigTable: FolderPath = " & FolderPath
        ]
    else
        SortedFiles{0},

    Workbook = Excel.Workbook(FirstFile[Content], null, true),
    DailyUpdateRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
    RawTable = if Table.RowCount(DailyUpdateRow) = 0 then
        error [
            Reason = "NoData",
            Message = "ไฟล์ " & FirstFile[Name] & " ไม่มี sheet ชื่อ """ & SheetName & """"
        ]
    else
        DailyUpdateRow{0}[Data],
    // Column-count safety cap: real sheets here run ~200-something columns.
    // fnCleanMergedHeaders loops over every column several times (FillDown,
    // header combine, rename, ...), so an Excel "used range" that's bloated
    // out to thousands of ghost columns makes every one of those loops slow
    // even with the row cap below in place. 300 gives headroom over the
    // real column count without leaving the door open to thousands.
    RawColNames = Table.ColumnNames(RawTable),
    ColumnCappedTable = Table.SelectColumns(RawTable, List.FirstN(RawColNames, 300)),
    // This query only reads column NAMES (see HeaderNames below) — it never
    // uses a single data row — so cut down to just the junk+header rows
    // here. Otherwise fnCleanMergedHeadersSkipJunk still runs FillDown,
    // null-row removal, and type detection over every real data row for a
    // result that gets thrown away immediately after.
    // Capped at 50 as a safety net: JunkRows/HeaderRows come from
    // ConfigTable, so a typo there (e.g. HeaderRows = 5000) shouldn't be
    // able to pull thousands of rows — 50 is already generous headroom
    // over the normal 6 (3 junk + 3 header).
    // Table.Buffer forces this (now tiny — ≤50 rows × ≤300 cols) table to be
    // fully computed once and held in memory, instead of potentially being
    // re-evaluated back through Excel.Workbook() each time
    // fnCleanMergedHeaders calls Table.Column on it in a per-column loop.
    HeaderRowsOnly = Table.Buffer(Table.FirstN(ColumnCappedTable, List.Min({JunkRows + HeaderRows, 50}))),
    // Calling fnCleanMergedHeaders directly (skipping junk rows here
    // instead of going through the fnCleanMergedHeadersSkipJunk wrapper) —
    // diagnostic test to see whether the wrapper itself was contributing to
    // the slowdown at the "Cleaned" step, or whether it's purely
    // fnCleanMergedHeaders's own per-column work.
    SkippedJunkRows = Table.Skip(HeaderRowsOnly, JunkRows),
    Cleaned = fnCleanMergedHeaders(SkippedJunkRows, HeaderRows, Separator),

    // 1-based index → Excel column letter (1→A, 2→B, ..., 26→Z, 27→AA)
    ColIndexToLetter = (idx as number) as text =>
        let
            result = if idx <= 26 then
                Text.From(Character.FromNumber(64 + idx))
            else
                @ColIndexToLetter(Number.IntegerDivide(idx - 1, 26))
                & Text.From(Character.FromNumber(65 + Number.Mod(idx - 1, 26)))
        in
            result,

    HeaderNames = Table.ColumnNames(Cleaned),
    HeaderTable = Table.FromList(HeaderNames, Splitter.SplitByNothing(), {"ชื่อ header"}),
    WithIndex = Table.AddIndexColumn(HeaderTable, "ตำแหน่ง column", 1, 1),
    // + FirstColumnOffset shifts the position to match the TRUE Excel
    // column letter in the source file — see FirstColumnLetter above.
    WithLetter = Table.TransformColumns(WithIndex, {"ตำแหน่ง column", each ColIndexToLetter(_ + FirstColumnOffset)})
in
    WithLetter
