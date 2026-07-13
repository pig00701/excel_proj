// ============================================================================
// Daily_Update_Archive.pq  (v2 — thin caller over fnDailyUpdateCombine)
// ============================================================================
// The STATIC half of the archive/current split: combines every CLOSED-year
// file (name-year strictly below ConfigTable's "CurrentYear", respecting
// YearFrom as the lower bound) with the exact same engine and cleaning as
// Daily_Update_Combined. Files with no year in the name are excluded here
// on purpose — they stay on the live side, so a changing off-convention
// file can never go silently stale inside a frozen snapshot.
//
// SETUP (one time):
//   1. Add "CurrentYear" to ConfigTable (e.g. 2569).
//   2. Load this query TO A TABLE on its own worksheet; the Excel table
//      must be named "Daily_Update_Archive" (Excel names it after the
//      query — verify in Table Design > Table Name).
//   3. Right-click this query > Properties > UNTICK "Refresh this
//      connection on Refresh All". This first load is the LAST slow one.
//
// NEW-YEAR ROLLOVER (once a year): bump CurrentYear in ConfigTable, then
// manually refresh THIS query once (it swallows the year that just closed).
//
// IF SelectColumnTable CHANGES: manually refresh this query once, or the
// archived history keeps the old column set.
//
// LOAD ORDER: same prerequisites as Daily_Update_Combined — see SETUP.md.
// ============================================================================

let
    Cfg = fnDailyUpdateConfig(Excel.CurrentWorkbook(){[Name = "ConfigTable"]}[Content]),
    SelectedColumns = List.Buffer(List.RemoveNulls(Table.Column(
        Excel.CurrentWorkbook(){[Name = "SelectColumnTable"]}[Content], "ColumnName"))),

    CurrentYear = if Cfg[CurrentYear] = null then
        error [
            Reason = "MissingConfig",
            Message = "ยังไม่ได้ตั้งค่า CurrentYear ใน ConfigTable — เพิ่มแถว Setting = CurrentYear, Value = ปีปัจจุบันตามชื่อไฟล์ (เช่น 2569) แล้วปีที่ต่ำกว่านี้จะถูกเก็บเข้า archive"
        ]
    else
        Cfg[CurrentYear],

    Result = fnDailyUpdateCombine(Cfg & [
        SelectedColumns = SelectedColumns,
        YearMin = Cfg[YearFrom],
        YearMax = CurrentYear - 1,
        IncludeNullYear = false,
        AllowEmpty = false
    ])
in
    Result
