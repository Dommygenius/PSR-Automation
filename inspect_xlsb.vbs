Option Explicit
Dim excel, wb, ws, fso, out
Set fso = CreateObject("Scripting.FileSystemObject")
out = fso.GetParentFolderName(WScript.ScriptFullName) & "\xlsb_inspect.txt"
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
On Error Resume Next
Set wb = excel.Workbooks.Open(fso.GetParentFolderName(WScript.ScriptFullName) & "\Copy of Copy of PSR_RAWDATA - new.xlsb", 0, True)
If Err.Number <> 0 Then
  Call WriteLine("ERROR open: " & Err.Description)
  WScript.Quit 1
End If
On Error GoTo 0
Call WriteLine("Workbook: " & wb.Name)
Dim sh
For Each sh In wb.Worksheets
  Call WriteLine("SHEET|" & sh.Name & "|visible=" & sh.Visible & "|used=" & sh.UsedRange.Address)
  Dim cht
  For Each cht In sh.ChartObjects()
    Call WriteLine("  CHART|" & cht.Name & "|" & cht.Chart.ChartTitle.Text)
  Next
Next
' named ranges with Avail
Dim nm
For Each nm In wb.Names
  If InStr(1, nm.Name, "Avail", vbTextCompare) > 0 Or InStr(1, nm.Name, "PSR", vbTextCompare) > 0 Or InStr(1, nm.Name, "Year", vbTextCompare) > 0 Then
    Call WriteLine("NAME|" & nm.Name & "|" & nm.RefersTo)
  End If
Next
wb.Close False
excel.Quit
WScript.Echo "Done: " & out

Sub WriteLine(s)
  Dim ts
  Set ts = fso.OpenTextFile(out, 8, True)
  ts.WriteLine s
  ts.Close
End Sub
