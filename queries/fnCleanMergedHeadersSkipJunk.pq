// ============================================================================
// fnCleanMergedHeadersSkipJunk.pq
// ============================================================================
// Wraps fnCleanMergedHeaders for sheets that have junk/title rows sitting
// ABOVE the real merged header — e.g. "Daily Update", where rows 1-3 are a
// report title / generated-on / department line, rows 4-6 are the real
// 3-row merged header, and data starts at row 7.
//
// It does exactly one thing: Table.Skip the junk rows, then hand off to
// fnCleanMergedHeaders unchanged. Kept as its own query so sheets WITHOUT
// junk rows (Sales_Report, Employee_Data) keep calling fnCleanMergedHeaders
// directly — this wrapper only exists for the sheets that need it.
//
// PREREQUISITE: Load fnCleanMergedHeaders.pq as a query named
// "fnCleanMergedHeaders" first.
//
// Parameters:
//   SourceTable  : The raw table from Excel (headers NOT promoted)
//   JunkRows     : Number of junk/title rows above the real header (e.g. 3)
//   HeaderRows   : Number of real merged-header rows (e.g. 3)
//   Separator    : Character to join header levels (default "_")
// ============================================================================

let
    Source = (SourceTable as table, JunkRows as number, HeaderRows as number, optional Separator as text) as table =>
        let
            SkippedJunkRows = Table.Skip(SourceTable, JunkRows),
            Cleaned = fnCleanMergedHeaders(SkippedJunkRows, HeaderRows, Separator)
        in
            Cleaned
in
    Source
