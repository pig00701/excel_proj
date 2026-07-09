// ============================================================================
// fnColumnLetterToIndex.pq
// ============================================================================
// Converts an Excel column letter to its 1-based index (reverse of the
// ColIndexToLetter helper used elsewhere): "A" → 1, "B" → 2, ..., "Z" → 26,
// "AA" → 27, ...
//
// Load this as a query named "fnColumnLetterToIndex". Used by
// Daily_Update_HeaderList.pq to turn a configured "true starting column"
// letter into a numeric offset — see that file for why the offset exists:
// Excel.Workbook()[Data] returns a sheet's USED RANGE, not always starting
// at column A, and Power Query exposes no way to ask it where the used
// range actually starts. If your source sheet's real data starts at column
// B (column A never had any content), set ConfigTable's FirstColumnLetter
// to "B" so the position column reports true Excel letters again.
// ============================================================================

(letter as text) as number =>
    let
        Alphabet = {"A".."Z"},
        Letters = Text.ToList(Text.Upper(Text.Trim(letter))),
        Result = List.Accumulate(
            Letters,
            0,
            (state, ch) => state * 26 + (List.PositionOf(Alphabet, ch) + 1)
        )
    in
        Result
