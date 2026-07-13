// ============================================================================
// Daily_Update_Combined.pq  (v2 — thin caller over fnDailyUpdateCombine)
// ============================================================================
// The main output: every "Daily Update" sheet combined into one table.
//
// Two modes, chosen by ConfigTable's "CurrentYear" row:
//
//   NORMAL (CurrentYear not set): combine every file in [YearFrom, YearTo].
//
//   SPLIT (CurrentYear set — the incremental-refresh pattern): closed years
//   never change, so their 80-100 MB .xlsm are parsed ONCE by the
//   Daily_Update_Archive query into a static worksheet table; this query
//   then only opens files of CurrentYear onward (plus files with no year in
//   the name — those always stay live) and prepends that archive table.
//   Per-refresh cost drops to parsing the current-year file(s) only.
//
// The archive is read from the WORKBOOK TABLE "Daily_Update_Archive", not
// from the Daily_Update_Archive query — referencing another data query
// alongside Folder.Files would trip the Formula.Firewall; reading a local
// table (like ConfigTable) does not. Missing table in split mode is a hard
// error: silently returning current-year-only data would look complete and
// be wrong.
//
// LOAD ORDER (queries this depends on): fnConfigValue, fnCleanMergedHeaders,
// fnCleanMergedHeadersSkipJunk, fnDailyUpdateConfig, fnDailyUpdateCombine.
// Workbook tables required: ConfigTable, SelectColumnTable
// (+ Daily_Update_Archive in split mode). See SETUP.md.
// ============================================================================

let
    Cfg = fnDailyUpdateConfig(Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content]),
    SelectedColumns = List.Buffer(List.RemoveNulls(Table.Column(
        Excel.CurrentWorkbook(){[Name = "SelectColumnTable"]}[Content], "ColumnName"))),
    SplitMode = Cfg[CurrentYear] <> null,

    Fresh = fnDailyUpdateCombine(Cfg & [
        SelectedColumns = SelectedColumns,
        // Split mode: current year onward; normal mode: the plain YearFrom
        // lower bound. YearTo stays an upper bound in both modes.
        YearMin = if SplitMode then Cfg[CurrentYear] else Cfg[YearFrom],
        YearMax = Cfg[YearTo],
        // Off-convention files (no year in name) always stay on the live side.
        IncludeNullYear = true,
        // Split mode tolerates zero current files (new-year rollover before
        // the new file exists) — the archive alone is still a valid result.
        AllowEmpty = SplitMode
    ]),

    Archive = if not SplitMode then
        null
    else
        try Excel.CurrentWorkbook(){[Name = "Daily_Update_Archive"]}[Content]
        otherwise error [
            Reason = "MissingArchive",
            Message = "ตั้ง CurrentYear ไว้แต่ยังไม่มีตาราง Daily_Update_Archive ใน workbook — โหลด query Daily_Update_Archive ลง sheet ครั้งแรกก่อน (แล้วปิด Refresh on Refresh All ของมัน) หรือลบแถว CurrentYear ออกจาก ConfigTable เพื่อกลับโหมดเดิม"
        ],

    Result = if Archive = null then Fresh else Table.Combine({Archive, Fresh})
in
    Result
