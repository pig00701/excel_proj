// ============================================================================
// fnDailyUpdateConfig.pq
// ============================================================================
// Parses the "ConfigTable" Excel Table into one settings record, shared by
// every Daily_Update_* query — the single place where setting names,
// defaults, and validation live.
//
// This is a PURE function: the caller reads ConfigTable itself (via
// Excel.CurrentWorkbook) and passes it in. That keeps the Formula.Firewall
// happy — each query touches its own data sources directly and only calls
// functions, never other data queries. See fnConfigValue.pq for background.
//
// PREREQUISITE: fnConfigValue loaded as a query.
//
// Returned record fields:
//   FolderPath    (required — errors with a clear message when missing)
//   SheetName     default "Daily Update"
//   FileExtension default ".xlsm" (leading dot added if missing)
//   JunkRows      default 3
//   HeaderRows    default 3
//   Separator     default "_"
//   YearFrom      default null  (no lower bound)
//   YearTo        default null  (no upper bound)
//   CurrentYear   default null  (null = archive/current split OFF)
// ============================================================================

let
    Source = (ConfigTable as table) as record =>
        let
            FolderPathRaw = fnConfigValue(ConfigTable, "FolderPath", null),
            FolderPath = if FolderPathRaw = null then
                error [
                    Reason = "MissingConfig",
                    Message = "ยังไม่ได้ตั้งค่า FolderPath ใน ConfigTable — เพิ่มแถว Setting = FolderPath, Value = โฟลเดอร์ที่เก็บไฟล์ต้นทาง"
                ]
            else
                Text.Trim(Text.From(FolderPathRaw)),
            ExtensionRaw = Text.Lower(Text.Trim(Text.From(fnConfigValue(ConfigTable, "FileExtension", ".xlsm")))),
            OptionalNumber = (name as text) as nullable number =>
                let raw = fnConfigValue(ConfigTable, name, null)
                in if raw = null then null else Number.From(raw)
        in
            [
                FolderPath = FolderPath,
                SheetName = Text.From(fnConfigValue(ConfigTable, "SheetName", "Daily Update")),
                FileExtension = if Text.StartsWith(ExtensionRaw, ".") then ExtensionRaw else "." & ExtensionRaw,
                JunkRows = Number.From(fnConfigValue(ConfigTable, "JunkRows", 3)),
                HeaderRows = Number.From(fnConfigValue(ConfigTable, "HeaderRows", 3)),
                Separator = Text.From(fnConfigValue(ConfigTable, "Separator", "_")),
                YearFrom = OptionalNumber("YearFrom"),
                YearTo = OptionalNumber("YearTo"),
                CurrentYear = OptionalNumber("CurrentYear")
            ]
in
    Source
