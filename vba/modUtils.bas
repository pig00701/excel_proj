Attribute VB_Name = "modUtils"
' ============================================================================
' modUtils.bas
' ============================================================================
' Shared helpers used by every module in the Daily Update VBA pipeline.
' No workbook-specific logic lives here — safe to reuse in other projects.
' ============================================================================
Option Explicit

' A cell counts as blank when it is Empty, Null, an error value, or
' whitespace-only text — mirrors "null" cells in the Power Query version.
Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        IsBlankValue = True
    ElseIf IsError(v) Then
        IsBlankValue = True
    Else
        IsBlankValue = (Trim$(CStr(v)) = vbNullString)
    End If
End Function

' 1-based column index -> Excel column letter (1->A, 26->Z, 27->AA)
Public Function ColIndexToLetter(ByVal idx As Long) As String
    Dim result As String
    Dim n As Long
    n = idx
    Do While n > 0
        result = Chr$(65 + ((n - 1) Mod 26)) & result
        n = (n - 1) \ 26
    Loop
    ColIndexToLetter = result
End Function

' Find a ListObject (worksheet table) by name anywhere in this workbook.
' Returns Nothing when not found — caller decides whether that is fatal.
Public Function FindListObject(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObject = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindListObject = Nothing
End Function

' Find a worksheet by name in any workbook (Excel sheet names are
' case-insensitive, so compare with vbTextCompare).
Public Function FindSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set FindSheet = ws
            Exit Function
        End If
    Next ws
    Set FindSheet = Nothing
End Function

' Get a fresh worksheet in ThisWorkbook for output. An existing sheet is
' DELETED and recreated rather than cleared: Cells.Clear leaves PivotTables
' and Power Query result ranges behind, and creating the output ListObject
' on top of those raises error 1004 ("a table cannot overlap...").
' The output sheet is fully machine-generated, so dropping it is safe.
Public Function GetOutputSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim tabIndex As Long
    Dim prevAlerts As Boolean

    tabIndex = 0
    Set ws = FindSheet(ThisWorkbook, sheetName)
    If Not ws Is Nothing Then
        tabIndex = ws.Index
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = prevAlerts
    End If

    If tabIndex > 0 And tabIndex <= ThisWorkbook.Worksheets.Count Then
        ' Recreate at the same tab position the old sheet occupied.
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(tabIndex))
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = sheetName
    Set GetOutputSheet = ws
End Function

' Year ("ปี") from a file name: the LAST 4 characters of the base name
' before the extension — a direct positional take, NOT a pattern search
' (same rule as the Power Query version). Non-numeric -> Empty (blank cell).
Public Function YearFromFileName(ByVal fileName As String) As Variant
    Dim baseName As String
    Dim dotPos As Long
    Dim lastFour As String

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
    Else
        baseName = fileName
    End If

    If Len(baseName) < 4 Then
        YearFromFileName = Empty
        Exit Function
    End If

    lastFour = Right$(baseName, 4)
    If lastFour Like "####" Then
        YearFromFileName = CLng(lastFour)
    Else
        YearFromFileName = Empty
    End If
End Function

' List file names in folderPath matching fileExtension, skipping Excel
' lock files ("~$..." — they appear whenever someone has a file open).
' Returns a Collection of file names (not full paths), in Dir() order.
Public Function ListSourceFiles(ByVal folderPath As String, ByVal fileExtension As String) As Collection
    Dim files As New Collection
    Dim f As String
    Dim ext As String

    ext = LCase$(fileExtension)
    f = Dir(folderPath & Application.PathSeparator & "*" & fileExtension)
    Do While Len(f) > 0
        ' Dir's pattern match is loose (e.g. *.xls also matches .xlsx),
        ' so re-check the extension exactly, and drop ~$ lock files.
        If Left$(f, 2) <> "~$" And LCase$(Right$(f, Len(ext))) = ext Then
            files.Add f
        End If
        f = Dir()
    Loop
    Set ListSourceFiles = files
End Function

' Read a sheet's full grid starting from A1 (like Power Query does) into a
' 1-based 2D Variant array. Always returns a 2D array, even for one cell.
Public Function ReadSheetGrid(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim raw As Variant

    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With

    If lastRow = 1 And lastCol = 1 Then
        Dim single2D(1 To 1, 1 To 1) As Variant
        single2D(1, 1) = ws.Cells(1, 1).Value
        ReadSheetGrid = single2D
    Else
        raw = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
        ReadSheetGrid = raw
    End If
End Function

' Standard error reporter: show procedure name + error info, used by every
' entrypoint's error handler after application state has been restored.
Public Sub ReportError(ByVal procName As String, ByVal errNumber As Long, ByVal errDescription As String)
    MsgBox "เกิดข้อผิดพลาดใน " & procName & vbCrLf & vbCrLf & _
           "Error " & errNumber & ": " & errDescription, _
           vbCritical, "Daily Update Pipeline"
End Sub
