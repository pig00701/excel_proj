Attribute VB_Name = "modYearRange"
' ============================================================================
' modYearRange.bas
' ============================================================================
' Ad-hoc year-range extraction: RunDailyUpdateYearRange.
'
' Pulls rows for a limited year range (e.g. 2564-2566) STRAIGHT FROM THE
' FROZEN ARCHIVE — zero Workbooks.Open, so it finishes in a blink no
' matter how big the source files are. The year of each row is derived
' from its SourceFile value (last 4 chars of the base name), the same rule
' the whole pipeline uses.
'
' Output goes to its own sheet "Daily_Update_YearRange" — completely
' separate from Daily_Update_Combined (the daily full output) and from the
' archive itself, so range queries never disturb the daily flow.
'
' ConfigTable rows used (both optional):
'   YearFrom | 2564    blank = no lower bound
'   YearTo   | 2566    blank = no upper bound
' RunDailyUpdateArchive attaches a dropdown (Data Validation) to these two
' Value cells listing exactly the years present in the archive.
'
' Prerequisites: archive built (RunDailyUpdateArchive) and matching the
' current SelectColumnTable — validated by the same check combine uses.
' Rows whose SourceFile has no 4-digit year suffix can't be classified
' and are skipped (reported in the summary). The live CurrentFilePath
' file is NOT included here by design: its name carries no year — current
' data lives on Daily_Update_Combined.
' ============================================================================
Option Explicit

Private Const OUTPUT_SHEET As String = "Daily_Update_YearRange"

Public Sub RunDailyUpdateYearRange()
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

    ' Same archive read + column-set validation as archive-mode combine.
    Dim archiveRows As Variant
    archiveRows = ReadArchiveRows(selectedCols)
    If IsEmpty(archiveRows) Then
        Err.Raise vbObjectError + 550, "RunDailyUpdateYearRange", _
            "ตาราง archive ว่างเปล่า — รัน RunDailyUpdateArchive ก่อน"
    End If

    Dim nCols As Long
    nCols = UBound(archiveRows, 2)          ' selected cols + SourceFile
    Dim srcCol As Long
    srcCol = nCols

    ' --- Filter rows by the year derived from SourceFile ------------------
    Dim total As Long
    total = UBound(archiveRows, 1)
    Dim keepRows() As Long
    ReDim keepRows(1 To total)
    Dim nKeep As Long
    Dim nNoYear As Long
    nKeep = 0
    nNoYear = 0
    Dim r As Long
    Dim y As Variant
    For r = 1 To total
        y = YearFromFileName(CStr(archiveRows(r, srcCol) & vbNullString))
        If IsEmpty(y) Then
            nNoYear = nNoYear + 1               ' no year in name -> can't classify
        ElseIf (cfg.YearFrom = 0 Or y >= cfg.YearFrom) And _
               (cfg.YearTo = 0 Or y <= cfg.YearTo) Then
            nKeep = nKeep + 1
            keepRows(nKeep) = r
        End If
    Next r

    ' --- Build output: same layout as the archive (header + kept rows) ----
    Dim output() As Variant
    output = NewOutputGrid(selectedCols, nKeep)
    Dim k As Long
    Dim outRow As Long
    outRow = 1
    For r = 1 To nKeep
        outRow = outRow + 1
        For k = 1 To nCols
            output(outRow, k) = archiveRows(keepRows(r), k)
        Next k
    Next r

    WriteOutputTable output, OUTPUT_SHEET, OUTPUT_SHEET

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts

    Dim rangeText As String
    If cfg.YearFrom = 0 And cfg.YearTo = 0 Then
        rangeText = "ทุกปีใน archive"
    Else
        rangeText = IIf(cfg.YearFrom = 0, "...", CStr(cfg.YearFrom)) & _
                    " - " & IIf(cfg.YearTo = 0, "...", CStr(cfg.YearTo))
    End If
    Dim msg As String
    msg = "ดึงช่วงปี " & rangeText & " จาก archive ได้ " & nKeep & _
          " แถว -> sheet """ & OUTPUT_SHEET & """"
    If nNoYear > 0 Then
        msg = msg & vbCrLf & "(ข้าม " & nNoYear & _
              " แถวที่ชื่อไฟล์ต้นทางไม่มีเลขปีท้ายชื่อ ระบุปีไม่ได้)"
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
    ReportError "RunDailyUpdateYearRange", errNum, errDesc
End Sub