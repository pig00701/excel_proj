// ============================================================================
// Sales_Report_Clean.pq
// ============================================================================
// Reads the Sales_Report sheet from sample_data.xlsx
// Cleans 3-row merged headers and removes fully-null rows
//
// PREREQUISITE: Load fnCleanMergedHeaders.pq as a query named "fnCleanMergedHeaders" first
// ============================================================================

let
    Source = Excel.Workbook(
        File.Contents("C:\Users\User\Downloads\new_sb\second-brain\01-Projects\excel_proj\sample_data.xlsx"),
        null,
        true
    ),
    SalesSheet = Source{[Item="Sales_Report", Kind="Sheet"]}[Data],
    Cleaned = fnCleanMergedHeaders(SalesSheet, 3, "_")
in
    Cleaned