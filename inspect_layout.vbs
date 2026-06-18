Option Explicit
' Read-only probe: exports top-left grid from each sheet (no changes to workbook).
Dim excel, wb, ws, fso, out, sh, r
Set fso = CreateObject("Scripting.FileSystemObject")
out = fso.GetParentFolderName(WScript.ScriptFullName) & "\layout_probe.txt"
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
Set wb = excel.Workbooks.Open(fso.GetParentFolderName(WScript.ScriptFullName) & "\Copy of Copy of PSR_RAWDATA - new.xlsb", 0, True)
Set fso = CreateObject("Scripting.FileSystemObject")
Dim ts : Set ts = fso.CreateTextFile(out, True)
For Each sh In wb.Worksheets
  ts.WriteLine "===== SHEET: " & sh.Name & " | used=" & sh.UsedRange.Address & " ====="
  Set ws = sh
  On Error Resume Next
  Set r = ws.Range("A1:AE50")
  Dim arr : arr = r.Value
  Dim i, j, lines
  If IsArray(arr) Then
    For i = 1 To UBound(arr, 1)
      lines = ""
      For j = 1 To UBound(arr, 2)
        If Not IsEmpty(arr(i, j)) Then
          lines = lines & Chr(9) & arr(i, j)
        End If
      Next
      If Len(lines) > 0 Then ts.WriteLine "R" & i & lines
    Next
  End If
  On Error GoTo 0
  ts.WriteLine ""
Next
ts.Close
wb.Close False
excel.Quit
WScript.Echo "Wrote " & out
