// ============================================================================
// fnDailyUpdateCombine.pq
// ============================================================================
// The ONE combine engine behind Daily_Update_Combined and
// Daily_Update_Archive: lists the folder, filters files by NAME (extension,
// ~$ lock files, year range) so excluded files are never opened, then per
// file: Binary.Buffer -> Excel.Workbook -> find sheet -> 300-column cap ->
// fnCleanMergedHeadersSkipJunk (junk-skip + merged-header clean + trim to
// SelectedColumns) -> add SourceFile + ปี -> Table.Combine.
//
// PREREQUISITE: fnCleanMergedHeaders, fnCleanMergedHeadersSkipJunk loaded
// as queries.
//
// Takes ONE record argument (Options), so callers read like configuration:
//   Required fields — FolderPath, SheetName, FileExtension, JunkRows,
//     HeaderRows, Separator (pass the fnDailyUpdateConfig record and merge
//     the fields below onto it with &)
//   SelectedColumns : list of combined header names to keep (SelectColumnTable)
//   YearMin         : nullable — keep files whose name-year >= this
//   YearMax         : nullable — keep files whose name-year <= this
//   IncludeNullYear : logical — keep files with no 4-digit year in the name?
//                     (true for the live side: never silently drop a file;
//                      false for the archive: a changing off-convention file
//                      must not freeze inside a static snapshot)
//   AllowEmpty      : logical — when no usable file matches: true returns an
//                     empty table (split-mode rollover: archive alone is
//                     fine), false raises a readable Thai error
//
// Year = last 4 characters of the file's base name ("Daily_Update_2566.xlsm"
// -> 2566); names not ending in 4 digits -> null, ปี column = null.
// ============================================================================

let
    ExtractYear = (fileName as text) as nullable number =>
        let
            baseName = Text.BeforeDelimiter(fileName, ".", {0, RelativePosition.FromEnd}),
            last4 = Text.End(baseName, 4)
        in
            try Number.From(last4) otherwise null,

    Source = (Options as record) as table =>
        let
            FolderPath = Options[FolderPath],
            SheetName = Options[SheetName],
            FileExtension = Options[FileExtension],
            JunkRows = Options[JunkRows],
            HeaderRows = Options[HeaderRows],
            Separator = Options[Separator],
            SelectedColumns = Options[SelectedColumns],
            YearMin = Options[YearMin],
            YearMax = Options[YearMax],
            IncludeNullYear = Options[IncludeNullYear],
            AllowEmpty = Options[AllowEmpty],

            AllFiles = Folder.Files(FolderPath),
            // Name-level filtering — nothing here opens a file's content.
            // "~$..." are Excel lock files left while a workbook is open.
            Candidates = Table.SelectRows(AllFiles, each
                Text.Lower([Extension]) = FileExtension
                and not Text.StartsWith([Name], "~$")
                and (
                    let y = ExtractYear([Name])
                    in if y = null then
                        IncludeNullYear
                    else
                        (YearMin = null or y >= YearMin) and (YearMax = null or y <= YearMax)
                )
            ),

            AddedSheetData = Table.AddColumn(Candidates, "SheetData", each
                let
                    FileName = [Name],
                    // Binary.Buffer: read the file's bytes once; downstream
                    // re-evaluation hits memory, not disk.
                    Workbook = Excel.Workbook(Binary.Buffer([Content]), null, true),
                    SheetRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
                    RawTable = if Table.RowCount(SheetRow) > 0 then SheetRow{0}[Data] else null,
                    // 300-column safety cap: a used range bloated to
                    // thousands of ghost columns must not reach the
                    // per-column cleaning loops. Real sheets run ~200.
                    Capped = if RawTable <> null then
                        Table.SelectColumns(RawTable, List.FirstN(Table.ColumnNames(RawTable), 300))
                    else
                        null,
                    // SelectedColumns goes INTO the cleaner: unwanted columns
                    // are dropped right after header names are built, before
                    // the expensive null-row / type-detection scans.
                    Cleaned = if Capped <> null then
                        fnCleanMergedHeadersSkipJunk(Capped, JunkRows, HeaderRows, Separator, SelectedColumns)
                    else
                        null,
                    WithSource = if Cleaned <> null then Table.AddColumn(Cleaned, "SourceFile", each FileName) else null,
                    FileYear = ExtractYear(FileName),
                    WithYear = if WithSource <> null then Table.AddColumn(WithSource, "ปี", each FileYear, type nullable number) else null
                in
                    WithYear
            ),

            // Files without the target sheet are skipped, not errors.
            Usable = Table.SelectRows(AddedSheetData, each [SheetData] <> null)[SheetData],
            Result = if List.Count(Usable) > 0 then
                Table.Combine(Usable)
            else if AllowEmpty then
                #table({}, {})
            else
                error [
                    Reason = "NoData",
                    Message = "ไม่พบไฟล์ " & FileExtension & " ที่มี sheet """ & SheetName & """ ในช่วงปีที่กำหนด — ตรวจสอบ ConfigTable: FolderPath = " & FolderPath
                ]
        in
            Result
in
    Source
