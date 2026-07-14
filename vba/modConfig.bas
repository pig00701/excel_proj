Attribute VB_Name = "modConfig"
' ============================================================================
' modConfig.bas
' ============================================================================
' Parameter-sheet pattern: all settings live in worksheet tables so users
' change behavior from the sheet and re-run — never by editing VBA.
'
' ConfigTable (sheet "config", Ctrl+T, table name "ConfigTable"):
'   Setting        | Value                 | หมายเหตุ (query ไม่ใช้)
'   FolderPath     | C:\...\DailyFiles     | โฟลเดอร์ไฟล์ต้นทาง — บังคับต้องมี
'   SheetName      | Daily Update          | ชื่อ sheet ที่ดึง
'   JunkRows       | 3                     | แถวขยะบนสุดที่ข้าม
'   HeaderRows     | 3                     | แถว merged header
'   Separator      | _                     | ตัวคั่นชื่อ header
'   FileExtension  | .xlsm                 | นามสกุลไฟล์ที่สแกน
'   CurrentFilePath| E:\...\ยอดวันนี้.xlsm  | (optional) เปิด archive mode
'
' Rules (same as the Power Query version):
'   - Setting names match case-insensitively, ignoring surrounding spaces.
'   - Only FolderPath is required; everything else has a default, so rows
'     other than FolderPath can be deleted entirely.
'   - CurrentFilePath set = ARCHIVE MODE: FolderPath becomes the closed
'     old-year files that RunDailyUpdateArchive freezes onto a hidden
'     sheet; RunDailyUpdateCombine then opens ONLY CurrentFilePath (full
'     path to a single file — any name, any location) and appends it to
'     the frozen archive. Delete the row to return to normal scan-all mode.
'
' SelectColumnTable (sheet "select column", table name "SelectColumnTable"):
'   one column "ColumnName" listing the combined header names to keep.
'   Extra columns (e.g. หมายเหตุ) are ignored.
' ============================================================================
Option Explicit

Public Type DailyUpdateConfig
    FolderPath As String
    SheetName As String
    JunkRows As Long
    HeaderRows As Long
    Separator As String
    FileExtension As String
    CurrentFilePath As String   ' non-empty = archive mode
End Type

Public Function ReadDailyUpdateConfig() As DailyUpdateConfig
    Dim cfg As DailyUpdateConfig
    Dim lo As ListObject
    Dim settingCol As Long
    Dim valueCol As Long
    Dim data As Variant
    Dim r As Long
    Dim settingName As String
    Dim settingValue As Variant

    ' Defaults — every setting except FolderPath can be omitted.
    cfg.FolderPath = vbNullString
    cfg.SheetName = "Daily Update"
    cfg.JunkRows = 3
    cfg.HeaderRows = 3
    cfg.Separator = "_"
    cfg.FileExtension = ".xlsm"

    Set lo = FindListObject("ConfigTable")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 513, "ReadDailyUpdateConfig", _
            "ไม่พบตาราง ConfigTable ใน workbook นี้ — สร้าง sheet ""config"" " & _
            "แล้วทำตารางคอลัมน์ Setting/Value (Ctrl+T ตั้งชื่อ ConfigTable)"
    End If

    settingCol = lo.ListColumns("Setting").Index
    valueCol = lo.ListColumns("Value").Index

    If Not lo.DataBodyRange Is Nothing Then
        data = lo.DataBodyRange.Value
        ' Single data row comes back as a scalar-shaped array only when the
        ' table has one column; with Setting+Value it is always 2D.
        For r = 1 To UBound(data, 1)
            settingName = LCase$(Trim$(CStr(data(r, settingCol) & vbNullString)))
            settingValue = data(r, valueCol)
            If Not IsBlankValue(settingValue) Then
                Select Case settingName
                    Case "folderpath":    cfg.FolderPath = Trim$(CStr(settingValue))
                    Case "sheetname":     cfg.SheetName = Trim$(CStr(settingValue))
                    Case "junkrows":      cfg.JunkRows = CLng(settingValue)
                    Case "headerrows":    cfg.HeaderRows = CLng(settingValue)
                    Case "separator":     cfg.Separator = CStr(settingValue)
                    Case "fileextension": cfg.FileExtension = Trim$(CStr(settingValue))
                    Case "currentfilepath": cfg.CurrentFilePath = Trim$(CStr(settingValue))
                End Select
            End If
        Next r
    End If

    If cfg.FolderPath = vbNullString Then
        Err.Raise vbObjectError + 514, "ReadDailyUpdateConfig", _
            "ConfigTable ยังไม่ได้ตั้งค่า FolderPath — ใส่ path โฟลเดอร์ไฟล์ต้นทางก่อน"
    End If

    ' Normalize: no trailing separator on the folder, extension starts with "."
    Do While Right$(cfg.FolderPath, 1) = "\" Or Right$(cfg.FolderPath, 1) = "/"
        cfg.FolderPath = Left$(cfg.FolderPath, Len(cfg.FolderPath) - 1)
    Loop
    If Left$(cfg.FileExtension, 1) <> "." Then cfg.FileExtension = "." & cfg.FileExtension

    ReadDailyUpdateConfig = cfg
End Function

' Column names to keep, in table order, from SelectColumnTable[ColumnName].
' Blank rows are skipped. Missing table or empty list is a hard error —
' silently keeping nothing (or everything) would look complete and be wrong.
Public Function ReadSelectedColumns() As Collection
    Dim result As New Collection
    Dim lo As ListObject
    Dim colIdx As Long
    Dim data As Variant
    Dim r As Long

    Set lo = FindListObject("SelectColumnTable")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 515, "ReadSelectedColumns", _
            "ไม่พบตาราง SelectColumnTable — สร้าง sheet ""select column"" " & _
            "ทำตารางที่มีคอลัมน์ ColumnName (Ctrl+T ตั้งชื่อ SelectColumnTable)"
    End If

    colIdx = lo.ListColumns("ColumnName").Index

    If Not lo.DataBodyRange Is Nothing Then
        If lo.ListColumns.Count = 1 And lo.ListRows.Count = 1 Then
            ' One-cell table: .Value is a scalar, not an array.
            If Not IsBlankValue(lo.DataBodyRange.Value) Then
                result.Add Trim$(CStr(lo.DataBodyRange.Value))
            End If
        Else
            data = lo.DataBodyRange.Value
            For r = 1 To UBound(data, 1)
                If Not IsBlankValue(data(r, colIdx)) Then
                    result.Add Trim$(CStr(data(r, colIdx)))
                End If
            Next r
        End If
    End If

    If result.Count = 0 Then
        Err.Raise vbObjectError + 516, "ReadSelectedColumns", _
            "SelectColumnTable ว่างเปล่า — พิมพ์ชื่อคอลัมน์ที่ต้องการเก็บอย่างน้อย 1 ชื่อ " & _
            "(ดูชื่อได้จากผลของ RunDailyUpdateHeaderList)"
    End If

    Set ReadSelectedColumns = result
End Function
