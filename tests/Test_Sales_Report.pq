// ============================================================================
// Test_Sales_Report.pq
// ============================================================================
// Test query: Validates that Sales_Report_Clean produces correct output.
// Run this after Sales_Report_Clean to verify:
//   1. No merged header artifacts remain
//   2. No fully-null rows exist
//   3. Correct number of data rows
//   4. Correct column names
// ============================================================================

let
    Source = Sales_Report_Clean,

    // --- Test 1: Column count (I renamed to Column_I, J kept as junk) ---
    ExpectedColCount = 10,
    ActualColCount = Table.ColumnCount(Source),
    Test1_ColCount = if ActualColCount = ExpectedColCount then "PASS" else "FAIL: Expected " & Text.From(ExpectedColCount) & " got " & Text.From(ActualColCount),

    // --- Test 2: Row count (10 data rows, 3 fully-null removed, row 6 survives) ---
    ExpectedRowCount = 7,
    ActualRowCount = Table.RowCount(Source),
    Test2_RowCount = if ActualRowCount = ExpectedRowCount then "PASS" else "FAIL: Expected " & Text.From(ExpectedRowCount) & " got " & Text.From(ActualRowCount),

    // --- Test 3: No null rows ---
    AllColNames = Table.ColumnNames(Source),
    NullRowCount = Table.RowCount(
        Table.SelectRows(
            Source,
            each List.AllTrue(
                List.Transform(AllColNames, (col) => Record.Field(_, col) = null)
            )
        )
    ),
    Test3_NoNullRows = if NullRowCount = 0 then "PASS" else "FAIL: Found " & Text.From(NullRowCount) & " null rows",

    // --- Test 4: Expected column names ---
    ExpectedCols = {"Product Info_ID", "Product Info_Name", "Q1 Sales_Revenue_Jan", "Q1 Sales_Revenue_Feb", "Q1 Sales_Units_Mar", "Q2 Sales_Revenue_Apr", "Q2 Sales_Revenue_May", "Q2 Sales_Units_Jun"},
    ActualCols = Table.ColumnNames(Source),
    Test4_ColNames = if ActualCols = ExpectedCols then "PASS" else "FAIL: Column mismatch",

    // --- Test 5: First row data check (must NOT be eaten by promote) ---
    FirstRow = Source{0},
    Test5_FirstRow = if Record.Field(FirstRow, "Product Info_ID") = "P001" then "PASS" else "FAIL",

    // --- Test 6: Last row data check ---
    LastRow = Source{Table.RowCount(Source) - 1},
    Test6_LastRow = if Record.Field(LastRow, "Product Info_ID") = "P006" then "PASS" else "FAIL",

    // --- Test 7: All expected IDs present ---
    ActualIDs = Table.Column(Source, "Product Info_ID"),
    ExpectedIDs = {"P001", "P002", "P003", "P004", "P005", "P006"},
    Test7_AllIDs = if ActualIDs = ExpectedIDs then "PASS" else "FAIL: Missing rows",

    // --- Build results table ---
    Results = #table(
        {"Test", "Result"},
        {
            {"1. Column Count", Test1_ColCount},
            {"2. Row Count", Test2_RowCount},
            {"3. No Null Rows", Test3_NoNullRows},
            {"4. Column Names", Test4_ColNames},
            {"5. First Row Data", Test5_FirstRow},
            {"6. Last Row Data", Test6_LastRow},
            {"7. All IDs Present", Test7_AllIDs}
        }
    )
in
    Results