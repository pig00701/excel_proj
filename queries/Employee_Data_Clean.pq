// ============================================================================
// Employee_Data_Clean.pq
// ============================================================================
// Reads the Employee_Data sheet from sample_data.xlsx
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
    EmpSheet = Source{[Item="Employee_Data", Kind="Sheet"]}[Data],
    Cleaned = fnCleanMergedHeaders(EmpSheet, 3, "_")
in
    Cleaned