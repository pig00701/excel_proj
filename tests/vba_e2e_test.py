# -*- coding: utf-8 -*-
# End-to-end test of the VBA Daily Update pipeline via Excel COM.
# Scenarios:
#   T1 normal mode combine (2 files in one folder)
#   T2 header list
#   T3 archive mode: build archive from old folder, combine with a
#      differently-named current file in another folder
#   T4 column-set drift -> combine must fail with rebuild-archive error
import os, re, shutil, sys, tempfile, traceback
import win32com.client

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRATCH = os.path.join(tempfile.gettempdir(), "daily_update_vba_e2e")
VB_OBJ_ERR = -2147221504  # vbObjectError

def prep_dirs():
    if os.path.exists(SCRATCH):
        shutil.rmtree(SCRATCH)
    os.makedirs(os.path.join(SCRATCH, "all_files"))
    os.makedirs(os.path.join(SCRATCH, "old_files"))
    os.makedirs(os.path.join(SCRATCH, "current"))
    os.makedirs(os.path.join(SCRATCH, "mods"))
    for f in ("Daily_Update_2566.xlsx", "Daily_Update_2567.xlsx"):
        shutil.copy(os.path.join(REPO, f), os.path.join(SCRATCH, "all_files", f))
    shutil.copy(os.path.join(REPO, "Daily_Update_2566.xlsx"),
                os.path.join(SCRATCH, "old_files", "Daily_Update_2566.xlsx"))
    # a second "closed year" (same content, different year in the name)
    # so the year-range filter has something to include AND exclude
    shutil.copy(os.path.join(REPO, "Daily_Update_2566.xlsx"),
                os.path.join(SCRATCH, "old_files", "Daily_Update_2565.xlsx"))
    # current file: different name, different folder (the user's scenario)
    shutil.copy(os.path.join(REPO, "Daily_Update_2567.xlsx"),
                os.path.join(SCRATCH, "current", "TodaySales.xlsx"))

def patch_modules():
    """Patch MsgBox -> cell logging so macros run unattended; write as cp874
    (Thai ANSI) because VBE imports .bas as ANSI."""
    mods = []
    for name in ("modUtils", "modConfig", "modCleanHeaders", "modCombine",
                 "modArchive", "modHeaderList", "modYearRange"):
        src = open(os.path.join(REPO, "vba", name + ".bas"), encoding="utf-8").read()
        src = re.sub(
            r'MsgBox "[^"]*" & procName[^\n]*\n[^\n]*\n\s*vbCritical, "Daily Update Pipeline"',
            'ThisWorkbook.Worksheets("config").Range("Z2").Value = '
            '"ERR|" & errNumber & "|" & procName & "|" & errDescription',
            src)
        src = src.replace('MsgBox msg, vbInformation, "Daily Update Pipeline"',
                          'ThisWorkbook.Worksheets("config").Range("Z1").Value = msg')
        # multi-line MsgBox in modHeaderList
        src = re.sub(
            r'MsgBox "[^"]*" & usedFile(?:.|\n)*?vbInformation, "Daily Update Pipeline"',
            'ThisWorkbook.Worksheets("config").Range("Z1").Value = "HeaderList OK " & nCols & " cols from " & usedFile',
            src)
        out = os.path.join(SCRATCH, "mods", name + ".bas")
        open(out, "w", encoding="cp874", errors="replace", newline="\r\n").write(src)
        mods.append(out)
    return mods

def build_workbook(xl, mods):
    wb = xl.Workbooks.Add()
    while wb.Worksheets.Count > 1:
        wb.Worksheets(wb.Worksheets.Count).Delete()
    ws = wb.Worksheets(1); ws.Name = "config"
    ws.Range("A1").Value = "Setting"; ws.Range("B1").Value = "Value"
    ws.Range("C1").Value = "หมายเหตุ"   # extra column must be ignored
    rows = [("FolderPath", os.path.join(SCRATCH, "all_files")),
            ("SheetName", "Daily Update"),
            ("JunkRows", 3), ("HeaderRows", 3), ("Separator", "_"),
            ("FileExtension", ".xlsx"),
            ("CurrentFilePath", "")]
    for i, (k, v) in enumerate(rows, start=2):
        ws.Range(f"A{i}").Value = k; ws.Range(f"B{i}").Value = v
    lo = ws.ListObjects.Add(1, ws.Range(f"A1:C{len(rows)+1}"), None, 1)
    lo.Name = "ConfigTable"

    ws2 = wb.Worksheets.Add(After=ws); ws2.Name = "select column"
    ws2.Range("A1").Value = "ColumnName"
    for i, n in enumerate(["Order Info_ID", "Order Info_Date", "Metrics_Qty",
                           "Metrics_Total"], start=2):
        ws2.Range(f"A{i}").Value = n
    lo2 = ws2.ListObjects.Add(1, ws2.Range("A1:A5"), None, 1)
    lo2.Name = "SelectColumnTable"

    for m in mods:
        wb.VBProject.VBComponents.Import(m)
    return wb

def set_config(wb, setting, value):
    lo = wb.Worksheets("config").ListObjects("ConfigTable")
    for r in range(1, lo.ListRows.Count + 1):
        if str(lo.DataBodyRange.Cells(r, 1).Value).strip() == setting:
            lo.DataBodyRange.Cells(r, 2).Value = value
            return
    raise RuntimeError("setting not found: " + setting)

def clear_log(wb):
    ws = wb.Worksheets("config")
    ws.Range("Z1").Value = None; ws.Range("Z2").Value = None

def read_log(wb):
    ws = wb.Worksheets("config")
    return ws.Range("Z1").Value, ws.Range("Z2").Value

def table_dims(wb, name):
    for sh in wb.Worksheets:
        for lo in sh.ListObjects:
            if lo.Name == name:
                nrows = lo.DataBodyRange.Rows.Count if lo.DataBodyRange else 0
                headers = [lo.HeaderRowRange.Cells(1, c).Value
                           for c in range(1, lo.HeaderRowRange.Columns.Count + 1)]
                return sh.Name, headers, nrows, lo
    return None, None, None, None

results = []
def check(label, cond, detail=""):
    results.append((label, bool(cond), detail))
    print(("PASS " if cond else "FAIL ") + label + ("  | " + str(detail) if detail else ""))

def main():
    prep_dirs()
    mods = patch_modules()
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False; xl.DisplayAlerts = False
    try:
        try:
            wb = build_workbook(xl, mods)
        except Exception:
            print("VBProject import blocked -> enabling AccessVBOM and retrying")
            xl.Quit()
            import winreg
            ver = "16.0"
            key = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                rf"Software\Microsoft\Office\{ver}\Excel\Security")
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False; xl.DisplayAlerts = False
            wb = build_workbook(xl, mods)

        # ---- T1: normal mode ----
        clear_log(wb)
        xl.Run("RunDailyUpdateCombine")
        z1, z2 = read_log(wb)
        sh, headers, nrows, _ = table_dims(wb, "Daily_Update_Combined")
        check("T1 combine ran without error", z2 is None, z2)
        check("T1 output table exists", sh is not None, sh)
        check("T1 5 data rows (3+2)", nrows == 5, nrows)
        check("T1 headers = selected + SourceFile",
              headers == ["Order Info_ID", "Order Info_Date", "Metrics_Qty",
                          "Metrics_Total", "SourceFile"], headers)

        # ---- T2: header list ----
        clear_log(wb)
        xl.Run("RunDailyUpdateHeaderList")
        z1, z2 = read_log(wb)
        sh, headers, nrows, _ = table_dims(wb, "Daily_Update_HeaderList")
        check("T2 headerlist ran without error", z2 is None, z2)
        check("T2 5 header rows", nrows == 5, nrows)
        check("T2 headers", headers == ["HeaderName", "ColumnPosition"], headers)

        # ---- T3: archive mode ----
        set_config(wb, "FolderPath", os.path.join(SCRATCH, "old_files"))
        set_config(wb, "CurrentFilePath",
                   os.path.join(SCRATCH, "current", "TodaySales.xlsx"))
        clear_log(wb)
        xl.Run("RunDailyUpdateArchive")
        z1, z2 = read_log(wb)
        check("T3a archive ran without error", z2 is None, z2)
        sh, headers, nrows, _ = table_dims(wb, "Daily_Update_Archive")
        check("T3a archive table 6 rows (2565+2566)", nrows == 6, nrows)
        check("T3a archive sheet hidden",
              wb.Worksheets("Daily_Update_Archive").Visible != -1)

        clear_log(wb)
        xl.Run("RunDailyUpdateCombine")
        z1, z2 = read_log(wb)
        sh, headers, nrows, lo = table_dims(wb, "Daily_Update_Combined")
        check("T3b combine (archive mode) ran without error", z2 is None, z2)
        check("T3b 8 rows total (6 archive + 2 current)", nrows == 8, nrows)
        srcs = [lo.DataBodyRange.Cells(r, 5).Value for r in range(1, nrows + 1)]
        check("T3b SourceFile: archive rows then TodaySales.xlsx",
              srcs[:3] == ["Daily_Update_2565.xlsx"] * 3
              and srcs[3:6] == ["Daily_Update_2566.xlsx"] * 3
              and srcs[6:] == ["TodaySales.xlsx"] * 2, srcs)

        # ---- T4: column drift must hard-error ----
        lo2 = wb.Worksheets("select column").ListObjects("SelectColumnTable")
        lo2.ListRows.Add()
        lo2.DataBodyRange.Cells(lo2.ListRows.Count, 1).Value = "Metrics_Price"
        clear_log(wb)
        xl.Run("RunDailyUpdateCombine")
        z1, z2 = read_log(wb)
        expected_err = VB_OBJ_ERR + 534
        got = str(z2).split("|") if z2 else []
        check("T4 combine refused (stale archive)",
              len(got) >= 2 and got[0] == "ERR" and int(got[1]) == expected_err,
              z2)

        # T4b: rebuild archive then combine succeeds with new column
        clear_log(wb)
        xl.Run("RunDailyUpdateArchive")
        xl.Run("RunDailyUpdateCombine")
        z1, z2 = read_log(wb)
        sh, headers, nrows, _ = table_dims(wb, "Daily_Update_Combined")
        check("T4b rebuild then combine ok", z2 is None and nrows == 8,
              (z2, nrows))
        check("T4b new column present", headers and "Metrics_Price" in headers,
              headers)

        # ---- T5: year-range extraction from archive ----
        cfg_lo = wb.Worksheets("config").ListObjects("ConfigTable")
        cfg_settings = [str(cfg_lo.DataBodyRange.Cells(r, 1).Value)
                        for r in range(1, cfg_lo.ListRows.Count + 1)]
        check("T5 YearFrom/YearTo rows auto-added by archive",
              "YearFrom" in cfg_settings and "YearTo" in cfg_settings,
              cfg_settings)
        yf_row = cfg_settings.index("YearFrom") + 1
        formula = cfg_lo.DataBodyRange.Cells(yf_row, 2).Validation.Formula1
        check("T5 dropdown lists archive years", formula == "2565,2566", formula)

        set_config(wb, "YearFrom", 2566)
        set_config(wb, "YearTo", 2566)
        clear_log(wb)
        xl.Run("RunDailyUpdateYearRange")
        z1, z2 = read_log(wb)
        sh, headers, nrows, lo = table_dims(wb, "Daily_Update_YearRange")
        check("T5 yearrange ran without error", z2 is None, z2)
        check("T5 3 rows (2566 only, 2565 excluded)", nrows == 3, nrows)
        if nrows:
            srcs = set(lo.DataBodyRange.Cells(r, len(headers)).Value
                       for r in range(1, nrows + 1))
            check("T5 all rows from 2566 file",
                  srcs == {"Daily_Update_2566.xlsx"}, srcs)

        wb.Close(SaveChanges=False)
    finally:
        xl.Quit()

    fails = [r for r in results if not r[1]]
    print(f"\n===== {len(results)-len(fails)}/{len(results)} passed =====")
    sys.exit(1 if fails else 0)

if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        sys.exit(2)
