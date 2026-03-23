Sub Main()

Dim i As Long
Dim number As Long

Application.ScreenUpdating = False
ActiveSheet.DisplayPageBreaks = False
Application.CutCopyMode = False

ThisWorkbook.Sheets("Hoja2").Activate
number = Application.WorksheetFunction.CountA(Range("E:E"))


For i = 2 To number

Dim F As Formless
Set F = New Formless
F.SetReport("disasters") = ThisWorkbook.Sheets("Hoja2").Range("E" & i).Value
F.SetName = ThisWorkbook.Sheets("Hoja2").Range("E" & i).Value
F.SetCodEmpresa = ThisWorkbook.Sheets("Hoja2").Range("A" & i).Value
F.SetCodPymeColectivo = ThisWorkbook.Sheets("Hoja2").Range("B" & i).Value
F.setCodAfiliado = ThisWorkbook.Sheets("Hoja2").Range("F" & i).Value
F.SetPolicy = ThisWorkbook.Sheets("Hoja2").Range("G" & i).Value
F.SetStartDate = ThisWorkbook.Sheets("Hoja2").Range("C" & i).Value
F.SetEndDate = ThisWorkbook.Sheets("Hoja2").Range("D" & i).Value
F.SetDisasterQuery = "DisasterFormless"
F.SetConnetion = "127Settings"
F.SetRecordSet = F.Query()
F.RecordToSheet("Hoja1") = 1
F.PivotTableOfReport("Hoja1") = "SINIESTRALIDAD"


F.PivotTableOfMembers = "Hoja1"
F.SummaryOfReport = "Hoja1"
F.SummaryOfReport = "Hoja5"


F.SetAuthorizationsQuery = "AuthorizationsFormless"
F.SetConnetion = "CloudSettings"
F.SetRecordSet = F.Query()
F.RecordToSheet("Hoja3") = 1
F.PivotTableOfReport("Hoja3") = "AUTORIZACIONES"

F.SetCallCenterQuery = "CallCenterFormless"
F.SetRecordSet = F.Query()
F.RecordToSheet("Hoja2") = 1

F.SetHeaders = "True"

Next i

ThisWorkbook.Sheets("Hoja1").Activate
End Sub


