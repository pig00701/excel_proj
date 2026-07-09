// ============================================================================
// fnConfigValue.pq
// ============================================================================
// Looks up one setting's Value from a ConfigTable-shaped table (columns
// "Setting" and "Value") by name, case/space-insensitive on the Setting
// column, falling back to `default` when the row is missing or blank.
//
// Load this as a query named "fnConfigValue".
//
// WHY THIS EXISTS (read before reusing the old "Config" query pattern):
// Any query that needs a config value which then feeds a data-source call
// in the SAME query — Folder.Files(FolderPath), Web.Contents(url), etc. —
// must read its OWN copy of ConfigTable locally (Excel.CurrentWorkbook) and
// call this function directly. Do NOT resolve that value via a separate
// "Config" query and reference it from the data-source-calling query: Power
// Query's firewall treats that as combining two data sources (the workbook,
// via Config, and the folder/web/etc., via this query) through a value that
// crossed a query boundary, and blocks it with:
//   "Formula.Firewall: Query '...' (step '...') references other queries or
//   steps, so it may not directly access a data source. Please rebuild this
//   data combination."
// A plain function (this one) is exempt — it doesn't access any data
// source itself, only the table you hand it when you call it.
// ============================================================================

(configTable as table, name as text, default as any) as any =>
    let
        matches = Table.SelectRows(
            configTable,
            each [Setting] <> null
                and Text.Lower(Text.Trim(Text.From([Setting]))) = Text.Lower(name)
        ),
        raw = if Table.RowCount(matches) > 0 then matches{0}[Value] else null,
        value = if raw = null or Text.Trim(Text.From(raw)) = "" then default else raw
    in
        value
