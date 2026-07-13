// ============================================================================
// Daily_Update_FileCheck.pq
// ============================================================================
// Diagnostic: shows how the CURRENT ConfigTable settings will treat every
// file in the folder — which query will open it, or why it is skipped —
// WITHOUT opening a single file (name-level checks only), so this refreshes
// in under a second even against 100 MB sources.
//
// Use it to answer "why is refresh still slow?" (which files does
// Daily_Update_Combined actually open?) and "did CurrentYear/YearFrom do
// what I meant?" before paying for a real refresh.
//
// Output columns:
//   ไฟล์        — file name
//   ปี          — year parsed from the name (null if none)
//   ขนาด (MB)   — file size (big numbers here = expensive to parse)
//   สถานะ       — one of:
//     "Combined (ปีปัจจุบัน)"  opened by Daily_Update_Combined (split mode)
//     "Combined"               opened by Daily_Update_Combined (normal mode)
//     "Archive (ปีเก่า)"       belongs to the static Daily_Update_Archive
//     "Combined (ไม่มีปีในชื่อ)" no year in name — always live
//     "ข้าม: นอกช่วงปี"         excluded by YearFrom/YearTo
//     "ข้าม: lock file"         ~$ temp file while a workbook is open
//     "ข้าม: นามสกุลไม่ตรง"     different extension
//
// PREREQUISITE: fnConfigValue, fnDailyUpdateConfig loaded as queries.
// Refreshing this with Refresh All is harmless (it's name-only), but it's
// fine to untick it too.
// ============================================================================

let
    Cfg = fnDailyUpdateConfig(Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content]),
    SplitMode = Cfg[CurrentYear] <> null,

    ExtractYear = (fileName as text) as nullable number =>
        let
            baseName = Text.BeforeDelimiter(fileName, ".", {0, RelativePosition.FromEnd}),
            last4 = Text.End(baseName, 4)
        in
            try Number.From(last4) otherwise null,

    Classify = (name as text, extension as text) as text =>
        let
            y = ExtractYear(name),
            YearFrom = Cfg[YearFrom],
            YearTo = Cfg[YearTo],
            CurrentYear = Cfg[CurrentYear]
        in
            if Text.StartsWith(name, "~$") then
                "ข้าม: lock file"
            else if Text.Lower(extension) <> Cfg[FileExtension] then
                "ข้าม: นามสกุลไม่ตรง"
            else if y = null then
                "Combined (ไม่มีปีในชื่อ)"
            else if SplitMode then
                if y >= CurrentYear and (YearTo = null or y <= YearTo) then
                    "Combined (ปีปัจจุบัน)"
                else if y < CurrentYear and (YearFrom = null or y >= YearFrom) then
                    "Archive (ปีเก่า)"
                else
                    "ข้าม: นอกช่วงปี"
            else if (YearFrom = null or y >= YearFrom) and (YearTo = null or y <= YearTo) then
                "Combined"
            else
                "ข้าม: นอกช่วงปี",

    AllFiles = Folder.Files(Cfg[FolderPath]),
    Report = Table.SelectColumns(
        Table.AddColumn(
            Table.AddColumn(
                Table.AddColumn(AllFiles, "ปี", each ExtractYear([Name]), type nullable number),
                "ขนาด (MB)", each Number.Round([Attributes][Size] / 1048576, 1), type number),
            "สถานะ", each Classify([Name], [Extension]), type text),
        {"Name", "ปี", "ขนาด (MB)", "สถานะ"}
    ),
    Renamed = Table.RenameColumns(Report, {{"Name", "ไฟล์"}}),
    Sorted = Table.Sort(Renamed, {{"สถานะ", Order.Ascending}, {"ปี", Order.Descending}})
in
    Sorted
