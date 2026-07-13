// ============================================================================
// Daily_Update_HeaderList.pq  (v2)
// ============================================================================
// Reference table: every cleaned header name from the "Daily Update" sheet
// + its true Excel column letter — a lookup for filling SelectColumnTable,
// so you never have to guess the combined ("Q1 Sales_Revenue_Jan"-style)
// spelling. It deliberately IGNORES SelectColumnTable, always showing the
// full list.
//
// Reads ONE file only — the one whose name sorts last (= latest year under
// the "..._<year>" naming convention).
//
// PERFORMANCE WARNING: even though only the top junk+header rows are used,
// Excel.Workbook still has to parse the whole sheet of that one file first —
// on an 80-100 MB .xlsm that's minutes, and no M trick can avoid it.
// So: right-click this query > Properties > UNTICK "Refresh this connection
// on Refresh All", and refresh it manually only when the source header
// layout actually changes.
//
// "ตำแหน่ง column" offset: Excel.Workbook returns the sheet's USED RANGE,
// which may not start at column A. Set ConfigTable's "FirstColumnLetter" to
// the column the data really starts at (default "A").
//
// PREREQUISITE: fnConfigValue, fnDailyUpdateConfig, fnCleanMergedHeaders,
// fnColumnLetterToIndex loaded as queries. See SETUP.md.
// ============================================================================

let
    ConfigTable = Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content],
    Cfg = fnDailyUpdateConfig(ConfigTable),
    // HeaderList-only setting, so it lives here rather than in the shared record.
    FirstColumnLetter = Text.From(fnConfigValue(ConfigTable, "FirstColumnLetter", "A")),
    FirstColumnOffset = fnColumnLetterToIndex(FirstColumnLetter) - 1,

    AllFiles = Folder.Files(Cfg[FolderPath]),
    Candidates = Table.SelectRows(AllFiles, each
        Text.Lower([Extension]) = Cfg[FileExtension] and not Text.StartsWith([Name], "~$")),
    // Name sorts descending -> first row = latest year.
    Sorted = Table.Sort(Candidates, {{"Name", Order.Descending}}),
    FirstFile = if Table.RowCount(Sorted) = 0 then
        error [
            Reason = "NoData",
            Message = "ไม่พบไฟล์ " & Cfg[FileExtension] & " ในโฟลเดอร์ — ตรวจสอบ ConfigTable: FolderPath = " & Cfg[FolderPath]
        ]
    else
        Sorted{0},

    // Binary.Buffer: read the file's bytes once, so editor preview /
    // re-evaluation doesn't re-read 100 MB from disk.
    Workbook = Excel.Workbook(Binary.Buffer(FirstFile[Content]), null, true),
    SheetRow = Table.SelectRows(Workbook, each [Item] = Cfg[SheetName] and [Kind] = "Sheet"),
    RawTable = if Table.RowCount(SheetRow) = 0 then
        error [
            Reason = "NoData",
            Message = "ไฟล์ " & FirstFile[Name] & " ไม่มี sheet ชื่อ """ & Cfg[SheetName] & """"
        ]
    else
        SheetRow{0}[Data],

    // 300-column ghost-range cap, then keep only the junk+header rows (50 =
    // typo-proof ceiling over the normal 6) and Table.Buffer that tiny table
    // so fnCleanMergedHeaders' per-column loops never re-evaluate back
    // through Excel.Workbook.
    Capped = Table.SelectColumns(RawTable, List.FirstN(Table.ColumnNames(RawTable), 300)),
    HeaderRowsOnly = Table.Buffer(Table.FirstN(Capped, List.Min({Cfg[JunkRows] + Cfg[HeaderRows], 50}))),
    Cleaned = fnCleanMergedHeaders(Table.Skip(HeaderRowsOnly, Cfg[JunkRows]), Cfg[HeaderRows], Cfg[Separator]),

    // 1-based index -> Excel column letter (1->A, 26->Z, 27->AA)
    ColIndexToLetter = (idx as number) as text =>
        let
            result = if idx <= 26 then
                Text.From(Character.FromNumber(64 + idx))
            else
                @ColIndexToLetter(Number.IntegerDivide(idx - 1, 26))
                & Text.From(Character.FromNumber(65 + Number.Mod(idx - 1, 26)))
        in
            result,

    HeaderTable = Table.FromList(Table.ColumnNames(Cleaned), Splitter.SplitByNothing(), {"ชื่อ header"}),
    WithIndex = Table.AddIndexColumn(HeaderTable, "ตำแหน่ง column", 1, 1),
    WithLetter = Table.TransformColumns(WithIndex, {"ตำแหน่ง column", each ColIndexToLetter(_ + FirstColumnOffset)})
in
    WithLetter
