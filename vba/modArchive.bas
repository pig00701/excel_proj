Attribute VB_Name = "modArchive"
' ============================================================================
' modArchive.bas
' ============================================================================
' Archive entrypoint: RunDailyUpdateArchive — VBA port of the incremental-
' refresh idea from Daily_Update_Archive.pq (v2), adapted to the "current
' file lives somewhere else" setup:
'
'   - FolderPath      = closed old-year files (they never change)
'   - CurrentFilePath = the single live file, any name, any location
'
' This macro does the expensive work ONCE: opens every file in FolderPath,
' cleans it exactly like the main pipeline, and freezes the combined rows
' onto a hidden sheet/table "Daily_Update_Archive" in this workbook.
' RunDailyUpdateCombine then only opens CurrentFilePath per run and
' prepends these frozen rows — daily cost drops to one Workbooks.Open.
'
' WHEN TO RE-RUN THIS MACRO (the archive is a snapshot, not live):
'   1. SelectColumnTable changed — combine detects this and refuses to run
'      until the archive is rebuilt with the new column set
'   2. someone edited old-year data — nothing can detect this; re-run
'   3. new year rollover — the just-closed file moves into FolderPath
'
' The frozen table stores the exact column layout (selected columns +
' SourceFile), which is what lets combine verify staleness cheaply.
' ============================================================================
Option Explicit

Private Const ARCHIVE_SHEET As String = "Daily_Update_Archive"

Public Sub RunDailyUpdateArchive()
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

    Dim selectedCols As Collection
    Set selectedCols = ReadSelectedColumns()

    Dim files As Collection
    Set files = ListSourceFiles(cfg.FolderPath, cfg.FileExtension)

    Dim cleanedParts As New Collection
    Dim warnings As New Collection
    Dim totalDataRows As Long
    totalDataRows = 0

    Dim fileName As Variant
    For Each fileName In files
        Dim part As Variant
        part = CleanOneFile(cfg, selectedCols, _
                            cfg.FolderPath & Application.PathSeparator & CStr(fileName), _
                            CStr(fileName), warnings)
        If Not IsEmpty(part) Then
            cleanedParts.Add Array(part, CStr(fileName))
            totalDataRows = totalDataRows + (UBound(part, 1) - 1)
        End If
    Next fileName

    If cleanedParts.Count = 0 Then
        Err.Raise vbObjectError + 535, "RunDailyUpdateArchive", _
            "ไม่พบไฟล์ปีเก่าที่ใช้ได้เลยใน FolderPath = " & cfg.FolderPath & vbCrLf & _
            "(ไฟล์ " & cfg.FileExtension & " ที่มี sheet """ & cfg.SheetName & """)"
    End If

    Dim output() As Variant
    output = NewOutputGrid(selectedCols, totalDataRows)
    Dim outRow As Long
    outRow = 1
    FillParts output, outRow, cleanedParts, selectedCols

    WriteOutputTable output, ARCHIVE_SHEET, ARCHIVE_TABLE

    ' Hide the frozen sheet — it is plumbing, not a report page. (Excel
    ' refuses to hide the only visible sheet; the report workbook always
    ' has config/select column sheets, so this is safe here.)
    FindSheet(ThisWorkbook, ARCHIVE_SHEET).Visible = xlSheetHidden

    ' Refresh the YearFrom/YearTo dropdowns in ConfigTable so the year-range
    ' macro can only ever be pointed at years that actually exist in this
    ' archive (rows are created if missing).
    ApplyYearDropdowns cleanedParts

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts
    Application.AutomationSecurity = prevAutomationSecurity

    Dim msg As String
    msg = "สร้าง archive เสร็จ: " & cleanedParts.Count & " ไฟล์, " & totalDataRows & _
          " แถว แช่ไว้ใน sheet ซ่อน """ & ARCHIVE_SHEET & """" & vbCrLf & _
          "ต่อจากนี้ RunDailyUpdateCombine จะเปิดแค่ไฟล์ CurrentFilePath"
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
    Application.AutomationSecurity = prevAutomationSecurity
    ReportError "RunDailyUpdateArchive", errNum, errDesc
End Sub

' Attach a Data Validation dropdown listing the distinct years found in the
' archived files to the Value cells of ConfigTable's YearFrom/YearTo rows.
' Missing rows are appended (with blank values = no bound). Files without
' a 4-digit year suffix contribute nothing to the list.
Private Sub ApplyYearDropdowns(ByVal cleanedParts As Collection)
    ' Collect distinct years from the archived file names.
    Dim years As Object
    Set years = CreateObject("Scripting.Dictionary")
    Dim item As Variant
    Dim y As Variant
    For Each item In cleanedParts
        y = YearFromFileName(CStr(item(1)))
        If Not IsEmpty(y) Then
            If Not years.Exists(CLng(y)) Then years.Add CLng(y), True
        End If
    Next item
    If years.Count = 0 Then Exit Sub

    ' Sort ascending (tiny list — simple insertion into an array).
    Dim arr() As Long
    ReDim arr(1 To years.Count)
    Dim n As Long
    Dim key As Variant
    n = 0
    For Each key In years.Keys
        n = n + 1
        Dim j As Long
        j = n
        Do While j > 1
            If arr(j - 1) > CLng(key) Then
                arr(j) = arr(j - 1)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j) = CLng(key)
    Next key
    Dim listText As String
    For n = 1 To UBound(arr)
        listText = listText & IIf(n > 1, ",", "") & CStr(arr(n))
    Next n

    Dim lo As ListObject
    Set lo = FindListObject("ConfigTable")
    If lo Is Nothing Then Exit Sub
    ApplyDropdownToSetting lo, "YearFrom", listText
    ApplyDropdownToSetting lo, "YearTo", listText
End Sub

Private Sub ApplyDropdownToSetting(ByVal lo As ListObject, _
                                   ByVal settingName As String, _
                                   ByVal listText As String)
    Dim settingCol As Long
    Dim valueCol As Long
    settingCol = lo.ListColumns("Setting").Index
    valueCol = lo.ListColumns("Value").Index

    Dim valueCell As Range
    Set valueCell = Nothing
    Dim r As Long
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            If LCase$(Trim$(CStr(lo.DataBodyRange.Cells(r, settingCol).Value & vbNullString))) _
               = LCase$(settingName) Then
                Set valueCell = lo.DataBodyRange.Cells(r, valueCol)
                Exit For
            End If
        Next r
    End If
    If valueCell Is Nothing Then
        Dim newRow As ListRow
        Set newRow = lo.ListRows.Add
        newRow.Range.Cells(1, settingCol).Value = settingName
        Set valueCell = newRow.Range.Cells(1, valueCol)
    End If

    With valueCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=listText
        .InCellDropdown = True
        .IgnoreBlank = True             ' blank stays allowed = "no bound"
        .InputTitle = settingName
        .InputMessage = "เลือกปีจาก archive (เว้นว่าง = ไม่จำกัด)"
    End With
End Sub