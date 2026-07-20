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
    Dim prevAutomationSecurity As MsoAutomationSecurity

    prevScreenUpdating = Application.ScreenUpdating
    prevEnableEvents = Application.EnableEvents
    prevCalculation = Application.Calculation
    prevDisplayAlerts = Application.DisplayAlerts
    prevAutomationSecurity = Application.AutomationSecurity

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Dim cfg As DailyUpdateConfig
    cfg = ReadDailyUpdateConfig()

    ' Candidate files as FULL paths. CurrentFilePath (when set) goes first —
    ' it is the live file, so its layout is the most current reference for
    ' picking column names; the FolderPath files follow as fallback.
    Dim files As New Collection
    If cfg.CurrentFilePath <> vbNullString Then
        If FileExists(cfg.CurrentFilePath) Then files.Add cfg.CurrentFilePath
    End If
    Dim folderFile As Variant
    For Each folderFile In ListSourceFiles(cfg.FolderPath, cfg.FileExtension)
        files.Add cfg.FolderPath & Application.PathSeparator & CStr(folderFile)
    Next folderFile
    If files.Count = 0 Then
        Err.Raise vbObjectError + 540, "RunDailyUpdateHeaderList", _
            "ไม่พบไฟล์ " & cfg.FileExtension & " ใน FolderPath = " & cfg.FolderPath
    End If

    ' First file that actually has the target sheet. Only the header block
    ' is read (merge-aware, no heuristic guessing) — same name resolution
    ' the combine pipeline uses, and no full-sheet read at all.
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim usedFile As String
    Dim fileName As Variant
    Dim names() As String
    Dim nCols As Long
    nCols = 0
    For Each fileName In files
        Set wb = Workbooks.Open( _
            fileName:=CStr(fileName), ReadOnly:=True, UpdateLinks:=0)
        Set ws = FindSheet(wb, cfg.SheetName)
        If Not ws Is Nothing Then
            Dim lastRow As Long
            Dim lastCol As Long
            With ws.UsedRange
                lastRow = .Row + .Rows.Count - 1
                lastCol = .Column + .Columns.Count - 1
            End With
            If lastRow >= cfg.JunkRows + cfg.HeaderRows Then
                names = BuildCombinedHeaders( _
                    ReadHeaderBlockExpanded(ws, cfg.JunkRows + cfg.HeaderRows, lastCol), _
                    cfg.JunkRows, cfg.HeaderRows, cfg.Separator, useHeuristicFill:=False)
                nCols = lastCol
                usedFile = FileNameFromPath(CStr(fileName))
            End If
        End If
        wb.Close SaveChanges:=False
        If nCols > 0 Then Exit For
    Next fileName

    If nCols = 0 Then
        Err.Raise vbObjectError + 541, "RunDailyUpdateHeaderList", _
            "ไม่มีไฟล์ไหนใน FolderPath = " & cfg.FolderPath & _
            " ที่มี sheet """ & cfg.SheetName & """ พร้อมข้อมูลครบถึงแถว header"
    End If

    ' Build the reference list: header name + original column letter.
    Dim output() As Variant
    ReDim output(1 To nCols + 1, 1 To 2)
    output(1, 1) = "HeaderName"
    output(1, 2) = "ColumnPosition"
    Dim c As Long
    For c = 1 To nCols
        output(c + 1, 1) = names(c)
        output(c + 1, 2) = ColIndexToLetter(c)
    Next c

    WriteOutputTable output, OUTPUT_SHEET, OUTPUT_SHEET

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts
    Application.AutomationSecurity = prevAutomationSecurity

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
    Application.AutomationSecurity = prevAutomationSecurity
    ReportError "RunDailyUpdateHeaderList", errNum, errDesc
End Sub
