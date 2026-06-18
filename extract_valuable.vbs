Option Explicit
' Extract valuable PSR data into sheet "Valuable Data" inside the master .xlsb.
' Close the .xlsb in Excel/Cursor before running.

Dim excel, wb, dest, src, fso, root, xlsb, row, i, shName, label, blocks, block
Const DEST_SHEET = "Valuable Data"
Const XL_PASTE_VALUES = -4163

Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(WScript.ScriptFullName)
xlsb = root & "\Copy of Copy of PSR_RAWDATA - new.xlsb"

If Not fso.FileExists(xlsb) Then
  WScript.Echo "Not found: " & xlsb
  WScript.Quit 1
End If

Dim fullSheets(5, 1)
fullSheets(0, 0) = "PSR&AVAIL" : fullSheets(0, 1) = "PSR vs AVAIL (yearly trend)"
fullSheets(1, 0) = "BSC&RNC"   : fullSheets(1, 1) = "RNC 60-day chart"
fullSheets(2, 0) = "RNC"       : fullSheets(2, 1) = "RNC daily pivot"
fullSheets(3, 0) = "BSC"       : fullSheets(3, 1) = "BSC daily pivot"
fullSheets(4, 0) = "LAC BSC_RNC" : fullSheets(4, 1) = "LAC performance"
fullSheets(5, 0) = "Raw data for Avail Vs PSR" : fullSheets(5, 1) = "Yearly PSR source"

Dim rawBlocks(4, 1)
rawBlocks(0, 0) = "A1:N44"   : rawBlocks(0, 1) = "Raw paste RNC"
rawBlocks(1, 0) = "A45:P90"  : rawBlocks(1, 1) = "Raw paste BSC"
rawBlocks(2, 0) = "Q45:Z200" : rawBlocks(2, 1) = "Raw paste LAC"
rawBlocks(3, 0) = "T2:Y40"   : rawBlocks(3, 1) = "Raw paste Technology PSR"
rawBlocks(4, 0) = "AA2:AF40" : rawBlocks(4, 1) = "Raw paste Service PSR"

WScript.Echo "Opening (close file if locked): " & xlsb
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
excel.EnableEvents = False
excel.AskToUpdateLinks = False

On Error Resume Next
Set wb = excel.Workbooks.Open(xlsb, 0, False)
If Err.Number <> 0 Then
  WScript.Echo "ERROR opening workbook: " & Err.Description
  WScript.Echo "Close the .xlsb in Excel/Cursor and run again."
  excel.Quit
  WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Opened. Building sheet '" & DEST_SHEET & "'..."

On Error Resume Next
wb.Worksheets(DEST_SHEET).Delete
On Error GoTo 0

Set dest = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
dest.Name = DEST_SHEET
dest.Tab.Color = 5296274

row = 1
dest.Cells(row, 1).Value = "PSR valuable data export"
dest.Cells(row, 1).Font.Bold = True
row = row + 2

For i = 0 To 5
  shName = fullSheets(i, 0)
  label = fullSheets(i, 1)
  WScript.Echo "  " & shName
  On Error Resume Next
  Set src = wb.Worksheets(shName)
  If Err.Number <> 0 Then
    WScript.Echo "    skipped (not found)"
    Err.Clear
    dest.Cells(row, 1).Value = "[MISSING] " & label
    row = row + 2
  Else
    On Error GoTo 0
    dest.Cells(row, 1).Value = label
    dest.Cells(row, 1).Font.Bold = True
    row = row + 1
    src.UsedRange.Copy
    dest.Cells(row, 1).PasteSpecial XL_PASTE_VALUES
    excel.CutCopyMode = False
    row = row + src.UsedRange.Rows.Count + 1
  End If
  On Error GoTo 0
Next

On Error Resume Next
Set src = wb.Worksheets("Raw data - BSC_RNC")
If Err.Number = 0 Then
  dest.Cells(row, 1).Value = "Raw data - BSC_RNC (compact blocks)"
  dest.Cells(row, 1).Font.Bold = True
  row = row + 2
  For i = 0 To 4
    block = rawBlocks(i, 0)
    label = rawBlocks(i, 1)
    WScript.Echo "  Raw block " & block
    dest.Cells(row, 1).Value = label
    dest.Cells(row, 1).Font.Bold = True
    row = row + 1
    src.Range(block).Copy
    dest.Cells(row, 1).PasteSpecial XL_PASTE_VALUES
    excel.CutCopyMode = False
    row = row + src.Range(block).Rows.Count + 1
  Next
End If
On Error GoTo 0

dest.Columns.AutoFit
dest.Activate

WScript.Echo "Saving..."
wb.Save
wb.Close False
excel.Quit

WScript.Echo "Done. Open sheet '" & DEST_SHEET & "' in the workbook."
