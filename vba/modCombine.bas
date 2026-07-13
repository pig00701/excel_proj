Attribute VB_Name = "modCombine"
' ============================================================================
' modCombine.bas
' ============================================================================
' Main entrypoint: RunDailyUpdateCombine — VBA port of Daily_Update_Combined
' (Power Query). One click does the whole pipeline:
'
'   1. Read settings from ConfigTable + column list from SelectColumnTable
'   2. Scan FolderPath for *FileExtension files (skip ~$ lock files)
'   3. Per file: read sheet SheetName from A1, skip JunkRows, clean the
'      merged header (modCleanHeaders), keep only selected columns
'   4. Add SourceFile (file name) + ปี (last 4 chars of base name)
'   5. Combine everything into one table on sheet "Daily_Update_Combined"
'
' Order matters and mirrors the PQ design: columns are trimmed PER FILE
' before combining, so the union step never carries hundreds of unused
' columns. Files without the target sheet are skipped (counted as warnings);
' if NO file is usable the run fails with a message that names FolderPath —
' that means the path is wrong or no file has the sheet, not that the code
' is broken.
'
' Wire this Sub to a button on the report workbook. Re-running replaces the
' previous output completely (same as Refresh in the PQ version).
' ============================================================================
Option Explicit

Private Const OUTPUT_SHEET As String = "Daily_Update_Combined"

Public Sub RunDailyUpdateCombine()
    Dim prevScreenUpdating As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalculation As XlCalculation
    Dim prevDisplayAlerts As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Dim cfg As DailyUpdateConfig
    cfg = ReadDailyUpdateConfig()

    Dim selectedCols As Collection
    Set selectedCols = ReadSelectedColumns()

    Dim files As Collection
    Set files = ListSourceFiles(cfg.FolderPath, cfg.FileExtension)

    ' Per-file cleaned results (each item: Array(cleanedGrid, fileName, year))
    Dim cleanedParts As New Collection
    Dim warnings As New Collection
    Dim totalDataRows As Long
    totalDataRows = 0

    Dim fileName As Variant
    For Each fileName In files
        Dim part As Variant
        part = CleanOneFile(cfg, selectedCols, CStr(fileName), warnings)
        If Not IsEmpty(part) Then
            cleanedParts.Add Array(part, CStr(fileName), YearFromFileName(CStr(fileName)))
            totalDataRows = totalDataRows + (UBound(part, 1) - 1)
        End If
    Next fileName

    If cleanedParts.Count = 0 Then
        Err.Raise vbObjectError + 530, "RunDailyUpdateCombine", _
            "ไม่พบไฟล์ที่ใช้ได้เลยใน FolderPath = " & cfg.FolderPath & vbCrLf & _
            "(ไฟล์ " & cfg.FileExtension & " ที่มี sheet """ & cfg.SheetName & """) " & _
            "— เช็คว่า path ถูกต้อง และไฟล์มี sheet ชื่อนี้จริง ไม่ใช่โค้ดพัง"
    End If

    ' --- Union into the final grid: selected columns + SourceFile + ปี ----
    ' Column order = SelectColumnTable order; a column missing from some
    ' file stays blank for that file's rows (same as Table.Combine).
    Dim nOutCols As Long
    nOutCols = selectedCols.Count + 2
    Dim output() As Variant
    ReDim output(1 To totalDataRows + 1, 1 To nOutCols)

    Dim k As Long
    Dim colName As Variant
    k = 0
    For Each colName In selectedCols
        k = k + 1
        output(1, k) = CStr(colName)
    Next colName
    output(1, nOutCols - 1) = "SourceFile"
    output(1, nOutCols) = ChrW(3611) & ChrW(3637)   ' "ปี" — ChrW keeps the module ANSI-safe

    Dim outRow As Long
    outRow = 1
    Dim item As Variant
    For Each item In cleanedParts
        Dim grid As Variant
        Dim srcName As String
        Dim srcYear As Variant
        grid = item(0)
        srcName = item(1)
        srcYear = item(2)

        ' Map this file's cleaned columns to the output positions.
        Dim mapIdx() As Long
        ReDim mapIdx(1 To selectedCols.Count)
        Dim c As Long
        Dim w As Long
        w = 0
        For Each colName In selectedCols
            w = w + 1
            mapIdx(w) = 0
            For c = 1 To UBound(grid, 2)
                If CStr(grid(1, c)) = CStr(colName) Then
                    mapIdx(w) = c
                    Exit For
                End If
            Next c
        Next colName

        Dim r As Long
        For r = 2 To UBound(grid, 1)
            outRow = outRow + 1
            For w = 1 To selectedCols.Count
                If mapIdx(w) > 0 Then output(outRow, w) = grid(r, mapIdx(w))
            Next w
            output(outRow, nOutCols - 1) = srcName
            output(outRow, nOutCols) = srcYear
        Next r
    Next item

    WriteOutputTable output, OUTPUT_SHEET, OUTPUT_SHEET

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts

    Dim msg As String
    msg = "รวมข้อมูลเสร็จ: " & cleanedParts.Count & " ไฟล์, " & totalDataRows & _
          " แถว -> sheet """ & OUTPUT_SHEET & """"
    If warnings.Count > 0 Then
        Dim wMsg As Variant
        msg = msg & vbCrLf & vbCrLf & "ข้ามไป " & warnings.Count & " ไฟล์:"
        For Each wMsg In warnings
            msg = msg & vbCrLf & "  - " & wMsg
        Next wMsg
    End If
    MsgBox msg, vbInformation, "Daily Update Pipeline"
    Exit Sub

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts
    ReportError "RunDailyUpdateCombine", errNum, errDesc
End Sub

' Open one source workbook read-only, clean its target sheet, close it.
' Returns Empty (and appends a warning) when the file has no target sheet
' or none of the selected columns — those files are skipped, not fatal.
Private Function CleanOneFile(ByRef cfg As DailyUpdateConfig, _
                              ByVal selectedCols As Collection, _
                              ByVal fileName As String, _
                              ByVal warnings As Collection) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim raw As Variant

    CleanOneFile = Empty
    Set wb = Workbooks.Open( _
        fileName:=cfg.FolderPath & Application.PathSeparator & fileName, _
        ReadOnly:=True, UpdateLinks:=0)

    On Error GoTo CleanFail
    Set ws = FindSheet(wb, cfg.SheetName)
    If ws Is Nothing Then
        warnings.Add fileName & " (ไม่มี sheet """ & cfg.SheetName & """)"
    Else
        raw = ReadSheetGrid(ws)
        CleanOneFile = CleanMergedHeaders(raw, cfg.JunkRows, cfg.HeaderRows, _
                                          cfg.Separator, selectedCols)
    End If
    wb.Close SaveChanges:=False
    Exit Function

CleanFail:
    ' Per-file problems (too few rows, no matching column, ...) skip the
    ' file with a visible warning instead of killing the whole run.
    warnings.Add fileName & " (" & Err.Description & ")"
    CleanOneFile = Empty
    wb.Close SaveChanges:=False
End Function

' Write a header+data grid to a cleared sheet and wrap it in a ListObject.
Public Sub WriteOutputTable(ByRef output As Variant, _
                            ByVal sheetName As String, _
                            ByVal tableName As String)
    Dim ws As Worksheet
    Set ws = GetOutputSheet(sheetName)

    Dim target As Range
    Set target = ws.Range("A1").Resize(UBound(output, 1), UBound(output, 2))
    target.Value = output

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = tableName
    ws.Columns.AutoFit
End Sub
