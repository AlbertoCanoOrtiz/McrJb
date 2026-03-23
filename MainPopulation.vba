Sub Population()

Dim i As Long
Dim number As Long

Application.ScreenUpdating = False
ActiveSheet.DisplayPageBreaks = False
Application.CutCopyMode = False

ThisWorkbook.Sheets("Hoja2").Activate
number = Application.WorksheetFunction.CountA(Range("A:A"))

For i = 2 To number

Dim F As Formless
Set F = New Formless
F.SetReport("census") = ThisWorkbook.Sheets("Hoja2").Range("E" & i).Value
F.SetName = ThisWorkbook.Sheets("Hoja2").Range("E" & i).Value
F.SetCodEmpresa = ThisWorkbook.Sheets("Hoja2").Range("A" & i).Value
F.SetCodPymeColectivo = ThisWorkbook.Sheets("Hoja2").Range("B" & i).Value
F.SetPopulationQuery = "1.0Population"
F.SetConnetion = "CloudSettings"
F.SetRecordSet = F.Query()
F.RecordToSheet("Hoja1") = 1

F.CloseReport = "True"
Next i
ThisWorkbook.Sheets("Hoja1").Activate
End Sub
