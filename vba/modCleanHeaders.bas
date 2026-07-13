Attribute VB_Name = "modCleanHeaders"
' ============================================================================
' modCleanHeaders.bas
' ============================================================================
' VBA port of fnCleanMergedHeaders / fnCleanMergedHeadersSkipJunk (Power
' Query) — cleans a sheet grid that has junk/title rows on top and a
' multi-row merged header, working entirely on in-memory arrays (one bulk
' read, no per-cell worksheet calls).
'
' CleanMergedHeaders(rawGrid, junkRows, headerRows, separator, columnsToKeep)
'   rawGrid       : 1-based 2D Variant array of the sheet from A1
'                   (use modUtils.ReadSheetGrid)
'   junkRows      : title/garbage rows above the real header (skipped)
'   headerRows    : merged-header rows to combine into one name
'   separator     : joins header levels ("Order Info" + "ID" -> Order Info_ID)
'   columnsToKeep : Collection of combined names to keep, in output order
'                   (missing names are skipped, like MissingField.Ignore);
'                   pass Nothing to keep every column
'
' Returns a 1-based 2D Variant array: row 1 = combined header names,
' rows 2.. = data rows. Rows blank in EVERY kept column are dropped.
'
' Same rules as the Power Query version:
'   1. Columns blank across ALL header rows are protected: they never
'      receive a fill from the left and never pass text to the right —
'      otherwise a genuinely blank column next to a real one would be
'      swallowed into that column's name, or leak text into its neighbor.
'   2. Fill RIGHT within each header row (merged cells across columns),
'      then fill DOWN (merged cells across rows).
'   3. Combine levels top-to-bottom, dropping consecutive duplicates
'      ("ID"+"ID" -> "ID") and empty parts.
'   4. Still-empty names get a fallback "Column_A", "Column_B", ... from the
'      column's TRUE original position; fallbacks are assigned AFTER the
'      fills so placeholder text can never leak into a real column's name.
'   5. Duplicate combined names get suffix _2, _3, ...
' ============================================================================
Option Explicit

Public Function CleanMergedHeaders(ByVal rawGrid As Variant, _
                                   ByVal junkRows As Long, _
                                   ByVal headerRows As Long, _
                                   ByVal separator As String, _
                                   ByVal columnsToKeep As Collection) As Variant
    Dim nRows As Long
    Dim nCols As Long
    nRows = UBound(rawGrid, 1)
    nCols = UBound(rawGrid, 2)

    Dim headerTop As Long
    headerTop = junkRows + 1
    If junkRows + headerRows > nRows Then
        Err.Raise vbObjectError + 520, "CleanMergedHeaders", _
            "ข้อมูลมีแค่ " & nRows & " แถว น้อยกว่า JunkRows (" & junkRows & _
            ") + HeaderRows (" & headerRows & ") — เช็คค่าใน ConfigTable"
    End If

    ' --- Step 1: columns blank across ALL header rows (true gaps) ---------
    Dim isBlankCol() As Boolean
    ReDim isBlankCol(1 To nCols)
    Dim c As Long
    Dim r As Long
    For c = 1 To nCols
        isBlankCol(c) = True
        For r = headerTop To junkRows + headerRows
            If Not IsBlankValue(rawGrid(r, c)) Then
                isBlankCol(c) = False
                Exit For
            End If
        Next r
    Next c

    ' --- Step 2+3: fill RIGHT per header row, then fill DOWN --------------
    ' hdr(i, c) = header level i (1..headerRows) of column c after filling.
    ' Blank cells are held as vbNullString.
    Dim hdr() As String
    ReDim hdr(1 To headerRows, 1 To nCols)
    Dim lastVal As String
    Dim i As Long
    For i = 1 To headerRows
        lastVal = vbNullString
        For c = 1 To nCols
            If isBlankCol(c) Then
                hdr(i, c) = vbNullString
                lastVal = vbNullString            ' blank column breaks the merge run
            ElseIf Not IsBlankValue(rawGrid(junkRows + i, c)) Then
                hdr(i, c) = CStr(rawGrid(junkRows + i, c))
                lastVal = hdr(i, c)
            ElseIf lastVal <> vbNullString Then
                hdr(i, c) = lastVal               ' merged cell continues to the right
            Else
                hdr(i, c) = vbNullString
            End If
        Next c
    Next i
    ' Fill down within each column (merged cells across header rows).
    For c = 1 To nCols
        For i = 2 To headerRows
            If hdr(i, c) = vbNullString Then hdr(i, c) = hdr(i - 1, c)
        Next i
    Next c

    ' --- Step 4: combine levels; consecutive duplicates collapse ----------
    Dim combined() As String
    ReDim combined(1 To nCols)
    Dim part As String
    Dim lastPart As String
    Dim nameParts As String
    For c = 1 To nCols
        nameParts = vbNullString
        lastPart = vbNullString
        For i = 1 To headerRows
            part = hdr(i, c)
            If part <> lastPart Then              ' skip consecutive duplicate
                If part <> vbNullString Then      ' skip empty part
                    If nameParts = vbNullString Then
                        nameParts = part
                    Else
                        nameParts = nameParts & separator & part
                    End If
                End If
                lastPart = part
            End If
        Next i
        combined(c) = nameParts
    Next c

    ' --- Step 5: fallback names from TRUE column position, then dedupe ----
    For c = 1 To nCols
        If combined(c) = vbNullString Then
            combined(c) = "Column_" & ColIndexToLetter(c)
        End If
    Next c
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 0                          ' case-sensitive, like M
    Dim countSoFar As Long
    Dim baseName As String
    For c = 1 To nCols
        baseName = combined(c)
        If seen.Exists(baseName) Then
            countSoFar = seen(baseName)
            combined(c) = baseName & "_" & (countSoFar + 1)
            seen(baseName) = countSoFar + 1
        Else
            seen.Add baseName, 1
        End If
    Next c

    ' --- Step 6: resolve output columns (keep order of columnsToKeep) -----
    Dim outIdx() As Long
    Dim nOut As Long
    If columnsToKeep Is Nothing Then
        nOut = nCols
        ReDim outIdx(1 To nOut)
        For c = 1 To nCols
            outIdx(c) = c
        Next c
    Else
        ' Map each combined name to its column index (first occurrence).
        Dim nameToIdx As Object
        Set nameToIdx = CreateObject("Scripting.Dictionary")
        nameToIdx.CompareMode = 0
        For c = 1 To nCols
            If Not nameToIdx.Exists(combined(c)) Then nameToIdx.Add combined(c), c
        Next c
        ReDim outIdx(1 To columnsToKeep.Count)
        nOut = 0
        Dim wanted As Variant
        For Each wanted In columnsToKeep
            If nameToIdx.Exists(CStr(wanted)) Then   ' missing name -> skipped
                nOut = nOut + 1
                outIdx(nOut) = nameToIdx(CStr(wanted))
            End If
        Next wanted
        If nOut = 0 Then
            ' No requested column exists in this grid: return header-only
            ' result with zero columns is useless — return empty (0 data
            ' rows, 0 cols marker) as a 1x1 flag the caller checks via
            ' CleanedColumnCount. Simpler: raise, the caller shows which
            ' file failed with clearer context.
            Err.Raise vbObjectError + 521, "CleanMergedHeaders", _
                "ไม่มีชื่อคอลัมน์ไหนใน SelectColumnTable ตรงกับ header ของไฟล์นี้เลย " & _
                "— เช็คชื่อด้วย RunDailyUpdateHeaderList"
        End If
        ReDim Preserve outIdx(1 To nOut)
    End If

    ' --- Step 7: collect data rows, dropping rows blank in all kept cols --
    Dim dataStart As Long
    dataStart = junkRows + headerRows + 1
    Dim keepRows() As Long
    Dim nKeep As Long
    ReDim keepRows(1 To Application.WorksheetFunction.Max(1, nRows - dataStart + 1))
    nKeep = 0
    Dim k As Long
    Dim allBlank As Boolean
    For r = dataStart To nRows
        allBlank = True
        For k = 1 To nOut
            If Not IsBlankValue(rawGrid(r, outIdx(k))) Then
                allBlank = False
                Exit For
            End If
        Next k
        If Not allBlank Then
            nKeep = nKeep + 1
            keepRows(nKeep) = r
        End If
    Next r

    ' --- Step 8: build result: row 1 = header names, rows 2.. = data ------
    Dim result() As Variant
    ReDim result(1 To nKeep + 1, 1 To nOut)
    For k = 1 To nOut
        result(1, k) = combined(outIdx(k))
    Next k
    Dim outR As Long
    For outR = 1 To nKeep
        For k = 1 To nOut
            If IsError(rawGrid(keepRows(outR), outIdx(k))) Then
                result(outR + 1, k) = Empty       ' #N/A etc. -> blank, not a poisoned cell
            Else
                result(outR + 1, k) = rawGrid(keepRows(outR), outIdx(k))
            End If
        Next k
    Next outR

    CleanMergedHeaders = result
End Function
