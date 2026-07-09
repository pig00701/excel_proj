// ============================================================================
// fnCleanMergedHeaders.pq
// ============================================================================
// Purpose: Generic function to clean Excel data with merged multi-row headers
//          and remove fully-null rows.
//
// Parameters:
//   SourceTable  : The raw table from Excel (headers NOT promoted)
//   HeaderRows   : Number of header rows (e.g., 3 for 3-row merged headers)
//   Separator    : Character to join header levels (default "_")
//
// Strategy:
//   1. Extract the top N rows as the header section
//   2. For each header row, fill nulls LEFT-TO-RIGHT (merged cells)
//   3. Fill DOWN within header section (merged cells across rows)
//   4. Combine header rows into a single text row
//   5. Columns whose combined header is still empty/null (blank across ALL
//      header rows) get a fallback name "Column_A", "Column_B", ... (Excel
//      column letter, based on the column's TRUE original position) —
//      duplicate combined names get suffix "_2", "_3", ...
//   6. Skip header rows, rename data columns with combined header
//   7. Remove rows where ALL columns are null
//   8. Auto-detect and set column types
//
// IMPORTANT: fallback names are assigned AFTER FillRight/FillDown (steps
// 2-3), not before. Assigning them first and writing the placeholder text
// into the header cells would give FillRight something to leak into a
// genuinely-empty cell of a NEIGHBORING real column, corrupting that
// column's combined name with an unrelated "_Column_X" suffix.
// ============================================================================

let
    Source = (SourceTable as table, HeaderRows as number, optional Separator as text) as table =>
    let
        Sep = if Separator = null then "_" else Separator,

        // --- Helper: Fill a list left-to-right (for merged cells in a row) ---
        // blankFlags marks columns that are blank across EVERY header row
        // (not part of any merge, just genuinely empty) — those positions
        // always stay null: they don't receive a fill from their left
        // neighbor, and they don't pass anything on to their right
        // neighbor either. Without this, a genuinely blank column sitting
        // right next to a real one gets silently swallowed into that real
        // column's name (columns treated as one merge when they're not),
        // and — the opposite failure — a blank column would otherwise
        // never get any text written into it, so it can't leak stray text
        // into its right neighbor either.
        FillRight = (lst as list, blankFlags as list) as list =>
            List.Accumulate(
                List.Zip({lst, blankFlags}),
                {{}, null}, // {results, lastNonNullValue}
                (state, current) =>
                    let
                        val = current{0},
                        isBlankCol = current{1},
                        results = state{0}
                    in
                        if isBlankCol then
                            {results & {null}, null}
                        else if val <> null then
                            {results & {val}, val}
                        else if state{1} <> null then
                            {results & {state{1}}, state{1}}
                        else
                            {results & {null}, null}
            ){0},

        // Helper: 1-based index → Excel column letter (1→A, 2→B, ..., 26→Z, 27→AA)
        ColIndexToLetter = (idx as number) as text =>
            let
                result = if idx <= 26 then
                    Text.From(Character.FromNumber(64 + idx))
                else
                    @ColIndexToLetter(Number.IntegerDivide(idx - 1, 26))
                    & Text.From(Character.FromNumber(65 + Number.Mod(idx - 1, 26)))
            in
                result,

        // --- Step 1: Extract header section (top N rows) ---
        HeaderSection = Table.FirstN(SourceTable, HeaderRows),
        HeaderColNames = Table.ColumnNames(HeaderSection),

        // Identify columns that are blank across ALL header rows (true gaps,
        // not part of any merge) BEFORE any filling happens, so FillRight
        // can protect them instead of guessing from already-modified data.
        IsBlankColumn = List.Transform(
            HeaderColNames,
            (colName) =>
                let
                    colValues = Table.Column(HeaderSection, colName)
                in
                    List.AllTrue(List.Transform(colValues, each _ = null or Text.Trim(Text.From(_)) = ""))
        ),

        // --- Step 2: Fill RIGHT within each header row (merged cells) ---
        // Operates on the RAW header values — nulls here are either a real
        // gap or a merge continuation, nothing else, so FillRight can't
        // mistake a placeholder for real merged content.
        HeaderRows_AsLists = Table.ToRows(HeaderSection),
        HeaderRows_FilledRight = List.Transform(HeaderRows_AsLists, each FillRight(_, IsBlankColumn)),
        HeaderFilledRight = Table.FromRows(HeaderRows_FilledRight, HeaderColNames),

        // --- Step 3: Fill DOWN within header section (merged cells across rows) ---
        HeaderFilledDown = Table.FillDown(
            HeaderFilledRight,
            Table.ColumnNames(HeaderFilledRight)
        ),

        // --- Step 4: Combine header rows into one, deduplicating consecutive identical parts ---
        ColNames = Table.ColumnNames(HeaderFilledDown),
        RawCombined = List.Transform(
            ColNames,
            each
                let
                    parts = List.Transform(
                        Table.Column(HeaderFilledDown, _),
                        each if _ = null then "" else Text.From(_)
                    ),
                    // Remove consecutive duplicates (e.g., "ID_ID" → "ID")
                    deduped = List.Accumulate(
                        parts,
                        {},
                        (state, current) =>
                            if List.Count(state) > 0 and current = List.Last(state) then state
                            else state & {current}
                    ),
                    // Remove empty parts
                    nonEmpty = List.Select(deduped, each _ <> ""),
                    result = Text.Combine(nonEmpty, Sep)
                in
                    result
        ),

        // --- Step 5: Empty/null combined headers → fallback "Column_A", "Column_B", ... ---
        // Only columns genuinely blank across ALL header rows land here —
        // computed from each column's TRUE original position (idx), so the
        // letter always matches its real Excel column regardless of
        // whether neighboring columns got a fallback name too.
        WithFallback = List.Transform(
            List.Zip({RawCombined, List.Positions(RawCombined)}),
            each
                let
                    name = _{0},
                    idx = _{1}
                in
                    if name = null or name = "" then "Column_" & ColIndexToLetter(idx + 1) else name
        ),
        // Deduplicate: if a name appears more than once, suffix with _2, _3, ...
        // Record.AddField errors if the field already exists, which a
        // repeated name always does by the second occurrence — remove it
        // first (MissingField.Ignore so the FIRST occurrence, where it
        // doesn't exist yet, doesn't error either) so this can overwrite
        // the running count instead of crashing on real duplicate headers.
        CombinedHeaderRow = List.Accumulate(
            List.Zip({WithFallback, {1..List.Count(WithFallback)}}),
            {[], {}},  // {seenCount record, result list}
            (state, current) =>
                let
                    name = current{0},
                    seen = state{0},
                    count = Record.FieldOrDefault(seen, name, 0),
                    newName = if count = 0 then name else name & "_" & Text.From(count + 1),
                    newSeen = Record.AddField(Record.RemoveFields(seen, {name}, MissingField.Ignore), name, count + 1)
                in
                    {newSeen, state{1} & {newName}}
        ){1},

        // --- Step 6: Get data section (skip header rows), rename with combined header ---
        DataSection = Table.Skip(SourceTable, HeaderRows),
        RenamedData = Table.RenameColumns(
            DataSection,
            List.Zip({HeaderColNames, CombinedHeaderRow})
        ),

        // --- Step 7: Remove rows where ALL columns are null ---
        AllColNames = Table.ColumnNames(RenamedData),
        CleanedData = Table.SelectRows(
            RenamedData,
            each not List.AllTrue(
                List.Transform(
                    AllColNames,
                    (col) => Record.Field(_, col) = null
                )
            )
        ),

        // --- Step 8: Auto-detect types (text → number where possible) ---
        // Checks ALL non-null values in the column, not just the first one.
        // A single sample would misclassify mixed columns (e.g. numbers with
        // an occasional "N/A") as numeric, which then errors out every cell
        // that fails Number.From during the actual type conversion below.
        TypedData = Table.TransformColumnTypes(
            CleanedData,
            List.Transform(
                AllColNames,
                each
                    let
                        colName = _,
                        allValues = List.RemoveNulls(Table.Column(CleanedData, colName)),
                        isNumeric = List.Count(allValues) > 0 and List.AllTrue(
                            List.Transform(
                                allValues,
                                each try Value.Is(Number.From(_), type number) otherwise false
                            )
                        ),
                        targetType = if isNumeric then type number else type text
                    in
                        {colName, targetType}
            )
        )
    in
        TypedData
in
    Source
