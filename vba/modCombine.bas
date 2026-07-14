Attribute VB_Name = "modCombine"
' ============================================================================
' modCombine.bas
' ============================================================================
' Main entrypoint: RunDailyUpdateCombine — VBA port of Daily_Update_Combined
' (Power Query). One click does the whole pipeline:
'
'   1. Read settings from ConfigTable + column list from SelectColumnTable
'   2. Collect source data:
'        NORMAL MODE (no CurrentFilePath): scan FolderPath for
'        *FileExtension files (skip ~$ lock files) and clean every file.
'        ARCHIVE MODE (CurrentFilePath set): read the frozen old-year rows
'        from the hidden Daily_Update_Archive table (built once by
'        RunDailyUpdateArchive in modArchive), then open ONLY the single
'        current file — per-run cost drops to one Workbooks.Open.
'   3. Per opened file: read just the header block to resolve the selected
'      columns' positions BY NAME in that file, then pull only those
'      columns (targeted read — see CleanOneFile)
'   4. Add SourceFile (file name)
'   5. Combine everything into one table on sheet "Daily_Update_Combined"
'
' Archive-mode safety: the archive table's header row is compared against
' the CURRENT SelectColumnTable — any mismatch is a hard error telling the
' user to rebuild the archive, because silently combining would leave the
' new column blank for every old-year row and look complete while wrong.
'
' Files without the target sheet are skipped (counted as warnings); if NO
' source is usable the run fails with a message that names the path — that
' means the path is wrong or no file has the sheet, not that the code is
' broken.
'
' Wire this Sub to a button on the report workbook. Re-running replaces the
' previous output completely (same as Refresh in the PQ version).
' ============================================================================
Option Explicit

Private Const OUTPUT_SHEET As String = "Daily_Update_Combined"
Public Const ARCHIVE_TABLE As String = "Daily_Update_Archive"

Public Sub RunDailyUpdateCombine()
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
    ' Source .xlsm files may carry their own macros/Workbook_Open code —
    ' never let them run while we bulk-open files.
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Dim cfg As DailyUpdateConfig
    cfg = ReadDailyUpdateConfig()

    Dim selectedCols As Collection
    Set selectedCols = ReadSelectedColumns()

    Dim archiveMode As Boolean
    archiveMode = (cfg.CurrentFilePath <> vbNullString)

    ' --- Collect fresh (opened-file) parts --------------------------------
    Dim cleanedParts As New Collection
    Dim warnings As New Collection
    Dim totalDataRows As Long
    totalDataRows = 0

    If archiveMode Then
        If Not FileExists(cfg.CurrentFilePath) Then
            Err.Raise vbObjectError + 532, "RunDailyUpdateCombine", _
                "ไม่พบไฟล์ CurrentFilePath = " & cfg.CurrentFilePath & _
                " — เช็ค path ใน ConfigTable"
        End If
        Dim curPart As Variant
        curPart = CleanOneFile(cfg, selectedCols, cfg.CurrentFilePath, _
                               FileNameFromPath(cfg.CurrentFilePath), warnings)
        If Not IsEmpty(curPart) Then
            cleanedParts.Add Array(curPart, FileNameFromPath(cfg.CurrentFilePath))
            totalDataRows = UBound(curPart, 1) - 1
        End If
    Else
        Dim files As Collection
        Set files = ListSourceFiles(cfg.FolderPath, cfg.FileExtension)
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
    End If

    ' --- Archive rows (frozen old years) ----------------------------------
    Dim archiveRows As Variant
    Dim archiveCount As Long
    archiveRows = Empty
    archiveCount = 0
    If archiveMode Then
        archiveRows = ReadArchiveRows(selectedCols)   ' validates column match
        If Not IsEmpty(archiveRows) Then archiveCount = UBound(archiveRows, 1)
    End If

    If cleanedParts.Count = 0 And archiveCount = 0 Then
        If archiveMode Then
            Err.Raise vbObjectError + 530, "RunDailyUpdateCombine", _
                "ไม่มีข้อมูลเลย: archive ว่าง และไฟล์ CurrentFilePath ใช้ไม่ได้ " & _
                "— รัน RunDailyUpdateArchive ก่อน และเช็คไฟล์ปัจจุบัน"
        Else
            Err.Raise vbObjectError + 530, "RunDailyUpdateCombine", _
                "ไม่พบไฟล์ที่ใช้ได้เลยใน FolderPath = " & cfg.FolderPath & vbCrLf & _
                "(ไฟล์ " & cfg.FileExtension & " ที่มี sheet """ & cfg.SheetName & """) " & _
                "— เช็คว่า path ถูกต้อง และไฟล์มี sheet ชื่อนี้จริง ไม่ใช่โค้ดพัง"
        End If
    End If

    ' --- Assemble: header row + archive rows first, then fresh rows -------
    Dim output() As Variant
    output = NewOutputGrid(selectedCols, archiveCount + totalDataRows)

    Dim nOutCols As Long
    nOutCols = UBound(output, 2)
    Dim outRow As Long
    outRow = 1
    Dim r As Long
    Dim k As Long
    If archiveCount > 0 Then
        ' Archive columns are already in exactly this order (validated by
        ' ReadArchiveRows), so this is a straight block copy.
        For r = 1 To archiveCount
            outRow = outRow + 1
            For k = 1 To nOutCols
                output(outRow, k) = archiveRows(r, k)
            Next k
        Next r
    End If
    FillParts output, outRow, cleanedParts, selectedCols

    WriteOutputTable output, OUTPUT_SHEET, OUTPUT_SHEET

    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalculation
    Application.DisplayAlerts = prevDisplayAlerts
    Application.AutomationSecurity = prevAutomationSecurity

    Dim msg As String
    If archiveMode Then
        msg = "รวมข้อมูลเสร็จ (archive mode): archive " & archiveCount & _
              " แถว + ไฟล์ปัจจุบัน " & totalDataRows & " แถว -> sheet """ & OUTPUT_SHEET & """"
    Else
        msg = "รวมข้อมูลเสร็จ: " & cleanedParts.Count & " ไฟล์, " & totalDataRows & _
              " แถว -> sheet """ & OUTPUT_SHEET & """"
    End If
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
    ReportError "RunDailyUpdateCombine", errNum, errDesc
End Sub

' Allocate the output grid (header row included): selected columns in
' SelectColumnTable order + SourceFile. Shared with modArchive so both
' outputs always have identical column layout.
Public Function NewOutputGrid(ByVal selectedCols As Collection, _
                              ByVal totalDataRows As Long) As Variant()
    Dim nOutCols As Long
    nOutCols = selectedCols.Count + 1
    Dim output() As Variant
    ReDim output(1 To totalDataRows + 1, 1 To nOutCols)

    Dim k As Long
    Dim colName As Variant
    k = 0
    For Each colName In selectedCols
        k = k + 1
        output(1, k) = CStr(colName)
    Next colName
    output(1, nOutCols) = "SourceFile"
    NewOutputGrid = output
End Function

' Copy every cleaned part into output starting after row `outRow`,
' mapping each part's columns to the output positions by name. A column
' missing from some file stays blank for that file's rows (same as
' Table.Combine). Shared with modArchive.
Public Sub FillParts(ByRef output() As Variant, ByRef outRow As Long, _
                     ByVal cleanedParts As Collection, _
                     ByVal selectedCols As Collection)
    Dim nOutCols As Long
    nOutCols = UBound(output, 2)

    Dim item As Variant
    Dim colName As Variant
    For Each item In cleanedParts
        Dim grid As Variant
        Dim srcName As String
        grid = item(0)
        srcName = item(1)

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
            output(outRow, nOutCols) = srcName
        Next r
    Next item
End Sub

' Read the frozen archive rows, first validating that the archive was
' built with EXACTLY the current SelectColumnTable (same names, same
' order). Returns the data rows as a 2D array, or Empty for a zero-row
' archive. Any mismatch is fatal by design: combining a stale archive
' would leave new columns blank on old-year rows and look complete.
Private Function ReadArchiveRows(ByVal selectedCols As Collection) As Variant
    Dim lo As ListObject
    Set lo = FindListObject(ARCHIVE_TABLE)
    If lo Is Nothing Then
        Err.Raise vbObjectError + 533, "ReadArchiveRows", _
            "ตั้ง CurrentFilePath ไว้ (archive mode) แต่ยังไม่มีตาราง " & ARCHIVE_TABLE & _
            " — รัน RunDailyUpdateArchive ครั้งแรกก่อน " & _
            "หรือลบแถว CurrentFilePath ออกจาก ConfigTable เพื่อกลับโหมดเดิม"
    End If

    ' Header must be: selected columns in order + SourceFile.
    Dim hdr As Variant
    hdr = lo.HeaderRowRange.Value
    Dim expectedCount As Long
    expectedCount = selectedCols.Count + 1
    Dim mismatch As Boolean
    mismatch = (UBound(hdr, 2) <> expectedCount)
    If Not mismatch Then
        Dim k As Long
        Dim colName As Variant
        k = 0
        For Each colName In selectedCols
            k = k + 1
            If CStr(hdr(1, k)) <> CStr(colName) Then mismatch = True
        Next colName
        If CStr(hdr(1, expectedCount)) <> "SourceFile" Then mismatch = True
    End If
    If mismatch Then
        Err.Raise vbObjectError + 534, "ReadArchiveRows", _
            "รายชื่อคอลัมน์ใน SelectColumnTable ไม่ตรงกับตอนที่สร้าง archive " & _
            "— รัน RunDailyUpdateArchive ใหม่ก่อน เพื่อให้ข้อมูลปีเก่ามีคอลัมน์ครบชุดเดียวกัน"
    End If

    If lo.DataBodyRange Is Nothing Then
        ReadArchiveRows = Empty
    Else
        ReadArchiveRows = lo.DataBodyRange.Value   ' >= 2 columns, always 2D
    End If
End Function

' Open one source workbook read-only, clean its target sheet, close it.
' Returns Empty (and appends a warning) when the file has no target sheet
' or none of the selected columns — those files are skipped, not fatal.
'
' Targeted read: instead of pulling the whole sheet, this reads only the
' small header block (JunkRows+HeaderRows rows) to resolve each selected
' column's position IN THIS FILE by name, then pulls just those columns'
' data. A 300-column sheet trimmed to a few dozen names transfers ~10x
' less through the COM boundary. Positions are re-resolved per file — we
' never trust the HeaderList sheet's positions, because files from
' different years may lay their columns out differently.
Public Function CleanOneFile(ByRef cfg As DailyUpdateConfig, _
                             ByVal selectedCols As Collection, _
                             ByVal fullPath As String, _
                             ByVal displayName As String, _
                             ByVal warnings As Collection) As Variant
    Dim wb As Workbook
    Dim ws As Worksheet

    CleanOneFile = Empty
    Set wb = Workbooks.Open(fileName:=fullPath, ReadOnly:=True, UpdateLinks:=0)

    On Error GoTo CleanFail
    Set ws = FindSheet(wb, cfg.SheetName)
    If ws Is Nothing Then
        warnings.Add displayName & " (ไม่มี sheet """ & cfg.SheetName & """)"
    Else
        Dim lastRow As Long
        Dim lastCol As Long
        With ws.UsedRange
            lastRow = .Row + .Rows.Count - 1
            lastCol = .Column + .Columns.Count - 1
        End With

        Dim headerBottom As Long
        headerBottom = cfg.JunkRows + cfg.HeaderRows
        If lastRow < headerBottom Then
            Err.Raise vbObjectError + 520, "CleanOneFile", _
                "ข้อมูลมีแค่ " & lastRow & " แถว น้อยกว่า JunkRows+HeaderRows (" & _
                headerBottom & ") — เช็คค่าใน ConfigTable"
        End If

        ' 1) Small read: header block only, all columns.
        Dim headerGrid As Variant
        headerGrid = Read2D(ws, 1, 1, headerBottom, lastCol)

        Dim names() As String
        names = BuildCombinedHeaders(headerGrid, cfg.JunkRows, cfg.HeaderRows, cfg.Separator)

        ' 2) Resolve selected names -> column positions in THIS file.
        Dim nameToIdx As Object
        Set nameToIdx = CreateObject("Scripting.Dictionary")
        nameToIdx.CompareMode = 0
        Dim c As Long
        For c = 1 To lastCol
            If Not nameToIdx.Exists(names(c)) Then nameToIdx.Add names(c), c
        Next c

        Dim outIdx() As Long
        Dim outNames() As String
        ReDim outIdx(1 To selectedCols.Count)
        ReDim outNames(1 To selectedCols.Count)
        Dim nOut As Long
        nOut = 0
        Dim wanted As Variant
        For Each wanted In selectedCols
            If nameToIdx.Exists(CStr(wanted)) Then
                nOut = nOut + 1
                outIdx(nOut) = nameToIdx(CStr(wanted))
                outNames(nOut) = CStr(wanted)
            End If
        Next wanted
        If nOut = 0 Then
            Err.Raise vbObjectError + 521, "CleanOneFile", _
                "ไม่มีชื่อคอลัมน์ไหนใน SelectColumnTable ตรงกับ header ของไฟล์นี้เลย " & _
                "— เช็คชื่อด้วย RunDailyUpdateHeaderList"
        End If

        ' 3) Targeted reads: only the selected columns' data rows.
        Dim dataStart As Long
        dataStart = headerBottom + 1
        Dim nData As Long
        nData = lastRow - dataStart + 1
        If nData < 0 Then nData = 0

        Dim cols() As Variant                      ' cols(k) = 2D array (nData x 1)
        ReDim cols(1 To nOut)
        Dim k As Long
        If nData > 0 Then
            For k = 1 To nOut
                cols(k) = Read2D(ws, dataStart, outIdx(k), lastRow, outIdx(k))
            Next k
        End If

        ' 4) Assemble: header row + data rows, dropping all-blank rows.
        Dim keepRows() As Long
        Dim nKeep As Long
        If nData > 0 Then ReDim keepRows(1 To nData)
        nKeep = 0
        Dim r As Long
        Dim allBlank As Boolean
        For r = 1 To nData
            allBlank = True
            For k = 1 To nOut
                If Not IsBlankValue(cols(k)(r, 1)) Then
                    allBlank = False
                    Exit For
                End If
            Next k
            If Not allBlank Then
                nKeep = nKeep + 1
                keepRows(nKeep) = r
            End If
        Next r

        Dim result() As Variant
        ReDim result(1 To nKeep + 1, 1 To nOut)
        For k = 1 To nOut
            result(1, k) = outNames(k)
        Next k
        Dim outR As Long
        For outR = 1 To nKeep
            For k = 1 To nOut
                If IsError(cols(k)(keepRows(outR), 1)) Then
                    result(outR + 1, k) = Empty    ' #N/A etc. -> blank
                Else
                    result(outR + 1, k) = cols(k)(keepRows(outR), 1)
                End If
            Next k
        Next outR

        CleanOneFile = result
    End If
    wb.Close SaveChanges:=False
    Exit Function

CleanFail:
    ' Per-file problems (too few rows, no matching column, ...) skip the
    ' file with a visible warning instead of killing the whole run.
    warnings.Add displayName & " (" & Err.Description & ")"
    CleanOneFile = Empty
    wb.Close SaveChanges:=False
End Function

' Write a header+data grid to a fresh sheet and wrap it in a ListObject.
Public Sub WriteOutputTable(ByRef output As Variant, _
                            ByVal sheetName As String, _
                            ByVal tableName As String)
    Dim ws As Worksheet
    Set ws = GetOutputSheet(sheetName)

    ' A leftover table with the same name anywhere else in the workbook
    ' (e.g. the old Power Query output loaded on another sheet) would make
    ' the lo.Name assignment below fail — fail early with a fix instead.
    Dim existing As ListObject
    Set existing = FindListObject(tableName)
    If Not existing Is Nothing Then
        Err.Raise vbObjectError + 531, "WriteOutputTable", _
            "มีตารางชื่อ """ & tableName & """ อยู่แล้วบน sheet """ & _
            existing.Parent.Name & """ (น่าจะเป็นผลลัพธ์ Power Query เดิม) " & _
            "— ลบตาราง/query เก่านั้นออกก่อน แล้วรันใหม่"
    End If

    Dim target As Range
    Set target = ws.Range("A1").Resize(UBound(output, 1), UBound(output, 2))
    target.Value = output

    Dim lo As ListObject
    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = tableName
    ws.Columns.AutoFit
End Sub