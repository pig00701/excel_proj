// ============================================================================
// Test_Employee_Data.pq
// ============================================================================
// Test query: Validates that Employee_Data_Clean produces correct output.
// ============================================================================

let
    Source = Employee_Data_Clean,

    // --- Test 1: Column count ---
    ExpectedColCount = 6,
    ActualColCount = Table.ColumnCount(Source),
    Test1_ColCount = if ActualColCount = ExpectedColCount then "PASS" else "FAIL: Expected " & Text.From(ExpectedColCount) & " got " & Text.From(ActualColCount),

    // --- Test 2: Row count (4 data rows, 2 null rows removed = 4) ---
    ExpectedRowCount = 4,
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
    ExpectedCols = {"Emp ID", "Name_First", "Name_Last", "Contact_Email", "Contact_Phone", "Contact_Address"},
    ActualCols = Table.ColumnNames(Source),
    Test4_ColNames = if ActualCols = ExpectedCols then "PASS" else "FAIL: Column mismatch. Got: " & Text.Combine(ActualCols, ", "),

    // --- Test 5: First row data check (must NOT be eaten by promote) ---
    FirstRow = Source{0},
    Test5_FirstRow = if Record.Field(FirstRow, "Emp ID") = "E001" then "PASS" else "FAIL",

    // --- Test 6: All expected IDs present ---
    ActualIDs = Table.Column(Source, "Emp ID"),
    ExpectedIDs = {"E001", "E002", "E003", "E004"},
    Test6_AllIDs = if ActualIDs = ExpectedIDs then "PASS" else "FAIL: Missing rows",

    // --- Build results table ---
    Results = #table(
        {"Test", "Result"},
        {
            {"1. Column Count", Test1_ColCount},
            {"2. Row Count", Test2_RowCount},
            {"3. No Null Rows", Test3_NoNullRows},
            {"4. Column Names", Test4_ColNames},
            {"5. First Row Data", Test5_FirstRow},
            {"6. All IDs Present", Test6_AllIDs}
        }
    )
in
    Results