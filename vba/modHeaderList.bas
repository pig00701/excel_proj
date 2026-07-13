Attribute VB_Name = "modHeaderList"
' ============================================================================
' modHeaderList.bas
' ============================================================================
' Helper entrypoint: RunDailyUpdateHeaderList — VBA port of
' Daily_Update_HeaderList (Power Query).
'
' Reads the FIRST usable file in FolderPath (NOT filtered through
' SelectColumnTable), cleans its merged header exactly the same way the
' main pipeline does, then lists every combined header name with its Excel
' column position (A, B, C, ...) on sheet "Daily_Update_HeaderList".
'
' Workflow: run this first, look at the list, then type the names you want
' into SelectColumnTable[ColumnName] and run RunDailyUpdateCombine.
' ============================================================================
Option Explicit

Private Const OUTPUT_SHEET As String = "Daily_Update_HeaderList"

Public Sub RunDailyUpdateHeaderList()
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

    Dim files As Collection
    Set files = ListSourceFiles(cfg.FolderPath, cfg.FileExtension)
    If files.Count = 0 Then
        Err.Raise vbObjectError + 540, "RunDailyUpdateHeaderList", _
            "ไม่พบไฟล์ " & cfg.FileExtension & " ใน FolderPath = " & cfg.FolderPath
    End If

    ' First file that actually has the target sheet.
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim usedFile As String
    Dim fileName As Variant
    Dim cleaned As Variant
    cleaned = Empty
    For Each fileName In files
        Set wb = Workbooks.Open( _
            fileName:=cfg.FolderPath & Application.PathSeparator & CStr(fileName), _
            ReadOnly:=True, UpdateLinks:=0)
        Set ws = FindSheet(wb, cfg.SheetName)
        If Not ws Is Nothing Then
            ' Keep ALL columns (columnsToKeep = Nothing) — the whole point
            ' is to see every available header name.
            cleaned = CleanMergedHeaders(ReadSheetGrid(ws), cfg.JunkRows, _
                                         cfg.HeaderRows, cfg.Separator, Nothing)
            usedFile = CStr(fileName)
        End If
        wb.Close SaveChanges:=False
        If Not IsEmpty(cleaned) Then Exit For
    Next fileName

    If IsEmpty(cleaned) Then
        Err.Raise vbObjectError + 541, "RunDailyUpdateHeaderList", _
            "ไม่มีไฟล์ไหนใน FolderPath = " & cfg.FolderPath & _
            " ที่มี sheet """ & cfg.SheetName & """"
    End If

    ' Build the reference list: header name + original column letter.
    ' Cleaned columns preserve original order (no selection), so column i
    ' of the cleaned grid IS original column i.
    Dim nCols As Long
    nCols = UBound(cleaned, 2)
    Dim output() As Variant
    ReDim output(1 To nCols + 1, 1 To 2)
    output(1, 1) = "HeaderName"
    output(1, 2) = "ColumnPosition"
    Dim c As Long
    For c = 1 To nCols
        output(c + 1, 1) = cleaned(1, c)
        output(c + 1, 2) = ColIndexToLetter(c)
    Next c

    WriteOutputTable output, OUTPUT_SHEET, OUTPUT_SHEET

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts

    MsgBox "อ่าน header จากไฟล์ """ & usedFile & """ ได้ " & nCols & _
           " คอลัมน์ -> sheet """ & OUTPUT_SHEET & """" & vbCrLf & vbCrLf & _
           "เลือกชื่อที่ต้องการแล้วพิมพ์ลง SelectColumnTable ก่อนรัน RunDailyUpdateCombine", _
           vbInformation, "Daily Update Pipeline"
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
    ReportError "RunDailyUpdateHeaderList", errNum, errDesc
End Sub
