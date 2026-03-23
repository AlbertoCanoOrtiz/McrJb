Option Explicit
Private wReport As Workbook
Private sName As String
Private sQuery As String
Private lCodEmpresa As Long
Private lCodPymeColectivo As Long
Private dStartDate As Date
Private dEndDate As Date
Private strPolicy As String
Private strCodAfiliado As String
Private cConnect As ADODB.Connection
Private rRecordSet As ADODB.RecordSet
Public Property Let SetName(Name As String)
sName = Name
End Property
Public Property Get Name() As String
Name = sName
End Property
Public Property Let SetStartDate(StartDate As Date)
dStartDate = StartDate
End Property
Public Property Get StartDate() As Date
StartDate = dStartDate
End Property
Public Property Let SetEndDate(EndDate As Date)
dEndDate = EndDate
End Property
Public Property Get EndDate() As Date
EndDate = dEndDate
End Property
Public Property Let SetCodEmpresa(CodEmpresa As Long)
lCodEmpresa = CodEmpresa
End Property
Public Property Get CodEmpresa()
CodEmpresa = lCodEmpresa
End Property
Public Property Let SetPolicy(Policy As String)
strPolicy = Policy
End Property
Public Property Get Policy()
Policy = strPolicy
End Property
Public Property Let setCodAfiliado(CodAfiliado As String)
strCodAfiliado = CodAfiliado
End Property
Public Property Get CodAfiliado()
CodAfiliado = strCodAfiliado
End Property
Public Property Let SetCodPymeColectivo(CodPymeColectivo As Long)
lCodPymeColectivo = CodPymeColectivo
End Property
Public Property Get CodPymeColectivo()
CodPymeColectivo = lCodPymeColectivo
End Property
Public Property Let SetDisasterQuery(NameQuery As String)
Dim myFile As String
Dim text As String
Dim textline As String
myFile = ThisWorkbook.Path & "\Querys\" & NameQuery & ".txt"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If textline <> "//" Then
        text = text & textline
    ElseIf textline = "//" Then
        text = text & "and  o.f_ocurrido between '" & Format(dStartDate, "yyyy/mm/dd") & "' and '" & Format(dEndDate, "yyyy/mm/dd") & "' "
        If lCodPymeColectivo <> 0 And strPolicy <> "" And lCodEmpresa <> 0 Then
            text = text & "and d.codpymecolectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and d.codempresa in ('" & lCodEmpresa & "') "
            text = text & "and d.cve_Derhab in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo <> 0 And strPolicy = "" And lCodEmpresa <> 0 Then
            text = text & "and d.codpymecolectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and d.codempresa in ('" & lCodEmpresa & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and d.Nomina in ('" & strCodAfiliado & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy <> "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and d.Nomina in ('" & strCodAfiliado & "') "
            text = text & "and d.cve_Derhab in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy <> "" And lCodEmpresa = 0 And strCodAfiliado = "" Then
            text = text & "and d.cve_Derhab in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa <> 0 And strCodAfiliado = "" Then
            text = text & "and d.codempresa in ('" & lCodEmpresa & "') "
        End If
        
    End If
Loop
Close #1
sQuery = text
End Property
Public Property Let SetCallCenterQuery(NameQuery As String)
Dim myFile As String
Dim text As String
Dim textline As String
myFile = ThisWorkbook.Path & "\Querys\" & NameQuery & ".txt"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If textline <> "//" Then
        text = text & textline
    ElseIf textline = "//" Then
        text = text & "and  b.FechaLlamada between '" & Format(dStartDate, "yyyy/mm/dd") & "' and '" & Format(dEndDate, "yyyy/mm/dd") & "' "
        If lCodPymeColectivo <> 0 And strPolicy <> "" And lCodEmpresa <> 0 Then
            text = text & "and a.codPymeColectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo <> 0 And strPolicy = "" And lCodEmpresa <> 0 Then
            text = text & "and a.codPymeColectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and a.codafiliado in ('" & strCodAfiliado & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and a.codafiliado in ('" & strCodAfiliado & "') "
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy <> "" And lCodEmpresa = 0 And strCodAfiliado = "" Then
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa <> 0 And strCodAfiliado = "" Then
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
        End If
     End If
Loop
Close #1
sQuery = text
End Property
Public Property Let SetAuthorizationsQuery(NameQuery As String)
Dim myFile As String
Dim text As String
Dim textline As String
myFile = ThisWorkbook.Path & "\Querys\" & NameQuery & ".txt"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If textline <> "//" Then
        text = text & textline
    ElseIf textline = "//" Then
        text = text & " gc.fecha between '" & Format(dStartDate, "yyyy/mm/dd") & "' and '" & Format(dEndDate, "yyyy/mm/dd") & "' "
        If lCodPymeColectivo <> 0 And strPolicy <> "" And lCodEmpresa <> 0 Then
            text = text & "and a.codPymeColectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo <> 0 And strPolicy = "" And lCodEmpresa <> 0 Then
            text = text & "and a.codPymeColectivo in ('" & lCodPymeColectivo & "') "
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and a.codafiliado in ('" & strCodAfiliado & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy <> "" And lCodEmpresa = 0 And strCodAfiliado <> "" Then
            text = text & "and a.codafiliado in ('" & strCodAfiliado & "') "
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy <> "" And lCodEmpresa = 0 And strCodAfiliado = "" Then
            text = text & "and a.poliza in ('" & strPolicy & "') "
        ElseIf lCodPymeColectivo = 0 And strPolicy = "" And lCodEmpresa <> 0 And strCodAfiliado = "" Then
            text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
        End If
    End If
Loop
Close #1
sQuery = text
End Property
Public Property Let SetPopulationQuery(NameQuery As String)
Dim myFile As String
Dim text As String
Dim textline As String
myFile = ThisWorkbook.Path & "\Querys\" & NameQuery & ".txt"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    If textline <> "//" Then
        text = text & textline
    ElseIf textline = "//" And lCodPymeColectivo <> 0 Then
        text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
        text = text & "and a.codpymecolectivo in ('" & lCodPymeColectivo & "') "
    Else
        text = text & "and a.codempresa in ('" & lCodEmpresa & "') "
    End If
Loop
Close #1
sQuery = text
End Property
Public Property Get Query() As String
Query = sQuery
End Property
Public Property Let SetReport(NameOfFile As String, NameOfReport As String)
Dim i As Long
Set wReport = Workbooks.Add
ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\" & NameOfFile & "\" & NameOfReport & ".xlsx"
If NameOfFile = "disasters" Then
    ActiveWorkbook.Sheets.Add
    ActiveWorkbook.Sheets.Add
    'ActiveWorkbook.Sheets.Add
    'ActiveWorkbook.Sheets.Add
    For i = 1 To 4
        ActiveWorkbook.Sheets("Hoja" & i).Activate
        Range("A1:xfd1048576").Interior.Color = RGB(255, 255, 255)
    Next i
ElseIf NameOfFile = "census" Then
    ActiveWorkbook.Sheets("Hoja1").Activate
    Range("A1:xfd1048576").Interior.Color = RGB(255, 255, 255)
End If
End Property
Public Property Get Report()
Set Report = wReport
End Property
Public Property Let SetConnetion(NameOfSettings As String)
Dim Settings As String
Settings = ReadTxt(NameOfSettings)
Set cConnect = New ADODB.Connection
cConnect.Open Settings
End Property
Public Property Let SetRecordSet(Query As String)
Set rRecordSet = New ADODB.RecordSet
Set rRecordSet = cConnect.Execute(Query)
End Property
Public Property Let RecordToSheet(NameOfSheet As String, NumberOfRow As Long)
wReport.Sheets(NameOfSheet).Activate
Dim j As Long
For j = 0 To (rRecordSet.Fields.Count - 1)
    On Error GoTo HANDLER
    Cells(NumberOfRow, (j + 1)).Value = rRecordSet.Fields(j).Name
    Cells(NumberOfRow, (j + 1)).Font.Name = "Arial"
    Cells(NumberOfRow, (j + 1)).Font.Bold = True
    Cells(NumberOfRow, (j + 1)).Font.Color = RGB(255, 255, 255)
    Cells(NumberOfRow, (j + 1)).Interior.Color = RGB(72, 61, 139)
    Cells(NumberOfRow, (j + 1)).Borders.LineStyle = xlContinuous
    Cells(NumberOfRow, (j + 1)).Borders.Weight = xlThin
Next j

wReport.Sheets(NameOfSheet).Range("A" & (NumberOfRow + 1)).CopyFromRecordset rRecordSet
    
    Exit Property
HANDLER:
    Cells(NumberOfRow, j).Value = "Unknow"
Resume Next
End Property
Public Property Let PivotTableOfReport(NameOfSheet As String, NameOfReport As String)
Dim pc As PivotCache
Dim pt As PivotTable
Dim Col As Long
wReport.Sheets(NameOfSheet).Activate
Col = Application.WorksheetFunction.CountA(Range("A1:A1048000"))

On Error GoTo ehandler

If NameOfReport = "AUTORIZACIONES" Then
    Set pc = wReport.PivotCaches.Create(xlDatabase, NameOfSheet & "!A1:N" & Col)
    Set pt = pc.CreatePivotTable(wReport.Sheets(NameOfSheet).Range("P1"))
    With pt
        With .PivotFields("TipoGasto")
            .Orientation = xlRowField
        End With
       With .PivotFields("Monto")
           .Orientation = xlDataField
           .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
           .Function = xlSum
       End With
    End With
ElseIf NameOfReport = "SINIESTRALIDAD" Then
    Set pc = wReport.PivotCaches.Create(xlDatabase, NameOfSheet & "!A1:S" & Col)
    Set pt = pc.CreatePivotTable(wReport.Sheets(NameOfSheet).Range("U1"))
    With pt
        With .PivotFields("TipoGasto")
            .Orientation = xlRowField
        End With
       With .PivotFields("Importe")
            .Orientation = xlDataField
            .NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            .Function = xlSum
       End With
    End With
End If
Continue:
Exit Property

ehandler:

Resume Continue:

End Property
Public Property Let PivotTableOfMembers(Name As String)
Dim Col As Long
Dim rRng As Range
Dim pc As PivotCache
Dim pt As PivotTable
wReport.Sheets(Name).Activate
Col = Application.WorksheetFunction.CountA(Range("G:G"))
Sheets(Name).Range("g1:k" & Col).Copy Destination:=Sheets("Hoja5").Range("A1")
wReport.Sheets("Hoja5").Activate
DeleteColumn 2, 3
Set rRng = Range("A1", Range("B1").End(xlDown))
rRng.RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
Col = Application.WorksheetFunction.CountA(Range("A:A"))

On Error GoTo ehandler

Set pc = wReport.PivotCaches.Create(xlDatabase, "Hoja5!A1:B" & Col)
Set pt = pc.CreatePivotTable(wReport.Sheets("Hoja5").Range("d1"))
With pt
        With .PivotFields("TipoGasto")
             .Orientation = xlRowField
        End With
        With .PivotFields("Paciente")
             .Orientation = xlDataField
             .NumberFormat = "General"
             .Function = xlCount
        End With
End With

Set rRng = Range("A1", Range("A1").End(xlDown))
rRng.RemoveDuplicates Columns:=Array(1), Header:=xlYes
Col = Application.WorksheetFunction.CountA(Range("A:A"))
wReport.Sheets("Hoja4").Activate
Range("I30").Value = Col

Continue:

Exit Property

ehandler:

Resume Continue:

End Property
Public Property Let SummaryOfReport(Name As String)
Dim i As Long
Dim rRng As Range
On Error GoTo ehandler
If Name = "Hoja1" Then
    For i = 2 To 10
        If Sheets(Name).Range("u" & i).Value <> "" Then
            Sheets(Name).Range("u" & i).Copy Destination:=Sheets("Hoja4").Range("D" & (i + 8))
            Sheets("Hoja4").Active
            ColorsHeaders ("D" & (i + 8))
            Range("E" & (i + 8)).Borders.LineStyle = xlContinuous
            Range("E" & (i + 8)).Borders.Weight = xlThin
        End If
    Next i
    Set rRng = wReport.Sheets(Name).Range("u2:v10")
    wReport.Sheets("Hoja4").Activate
    For i = 10 To 18
        Range("E" & i).Value = Application.WorksheetFunction.VLookup(Range("D" & i).Value, rRng, 2, 0)
        Range("E" & i).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Next i
ElseIf Name = "Hoja5" Then
    For i = 2 To 10
        If Sheets(Name).Range("d" & i) <> "" Then
            Sheets(Name).Range("d" & i).Copy Destination:=Sheets("Hoja4").Range("D" & (i + 20))
            Sheets("Hoja4").Active
            ColorsHeaders ("D" & (i + 20))
            Range("E" & (i + 20)).Borders.LineStyle = xlContinuous
            Range("E" & (i + 20)).Borders.Weight = xlThin
        End If
    Next i
    Set rRng = wReport.Sheets("Hoja5").Range("d2:e10")
    wReport.Sheets("Hoja4").Activate
    On Error GoTo ehandler
    For i = 22 To 30
        Range("e" & i).Value = Application.WorksheetFunction.VLookup(Range("d" & i).Value, rRng, 2, 0)
    Next i
    
    Application.DisplayAlerts = False
    wReport.Sheets("Hoja5").Delete
    Application.DisplayAlerts = True
    
End If
Exit Property
ehandler:
    Range("e" & i).Value = ""
Resume Next
End Property
Public Property Let SetHeaders(str As String)
Application.DisplayAlerts = False

wReport.Sheets("Hoja1").Activate

Range("r2:s1048576").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

wReport.Sheets("Hoja4").Activate
Range("F1").Value = "REPORTE DE UTILIZACIÓN DEL FONDO DE PROTECCION"
Range("F1").Font.Size = 14
Range("F1").Font.Bold = True
Range("F1:P1").Merge
Range("F1:P1").HorizontalAlignment = xlCenter

Range("F2").Value = "Y SERVICIOS SOCIALES DE LOS EMPLEADOS DE " & UCase(Name())
Range("F2").Font.Size = 14
Range("F2").Font.Bold = True
Range("F2:P2").Merge
Range("F2:P2").HorizontalAlignment = xlCenter

Range("d9").Value = "Tipo de proveedor"
ColorsHeaders ("d9")
Range("d7:d9").Merge
Range("d7:d9").Borders.LineStyle = xlContinuous
Range("d7:d9").Borders.Weight = xlThin

Range("e9").Value = "Total general"
ColorsHeaders ("e9")
Range("e7:e9").Merge
Range("e7:e9").Borders.LineStyle = xlContinuous
Range("e7:e9").Borders.Weight = xlThin

Range("d21").Value = "Tipo de gasto"
ColorsHeaders ("d21")
Range("e21").Value = "Pacientes"
ColorsHeaders ("e21")

Range("g28").Value = "TOTAL DE PACIENTES QUE HAN HECHO USO DEL SERVICIO"
ColorsHeaders ("g28")
Range("g28:h31").Merge
Range("g28:h31").Borders.LineStyle = xlContinuous
Range("g28:h31").Borders.Weight = xlThin
Range("i28:j31").Merge
Range("i28:j31").Borders.LineStyle = xlContinuous
Range("i28:j31").Borders.Weight = xlThin

wReport.Sheets("Hoja2").Activate
Range("i1").Value = "Llamadas realizadas al Call Center MACC"
Range("i1").Font.Size = 14
Range("i1").Font.Bold = True
Range("i1:p1").Merge
Range("i1:p1").HorizontalAlignment = xlCenter


Range("j4").Value = "RESUMEN UTILIZACION DEL CALL CENTER"
Range("j4").Font.Size = 14
Range("j4").Font.Bold = True
Range("j4:o4").Merge
Range("j4:o4").HorizontalAlignment = xlCenter

Range("k6").Value = "MOTIVO DE LLAMADAS"
Range("k6:l6").Merge
Range("k6:l6").HorizontalAlignment = xlCenter
Range("m6").Value = "*Reporte Acumulado"
Range("m6:n6").Merge
Range("m6:n6").HorizontalAlignment = xlCenter

Range("k7").Value = "Motivo"
ColorsHeaders ("k7")
Range("k7").HorizontalAlignment = xlCenter

Range("l7").Value = "Cantidad"
ColorsHeaders ("l7")
Range("l7").HorizontalAlignment = xlCenter

Range("m7").Value = "%"
ColorsHeaders ("m7")
Range("m7").HorizontalAlignment = xlCenter

Range("k8").Value = "Autorizaciones"
Range("k9").Value = "Cita para Servicio"
Range("k10").Value = "Información General"
Range("k11").Value = "Administrativa"
Range("k12").Value = "Otros"
Range("k7:m12").Borders.LineStyle = xlContinuous
Range("k7:m12").Borders.Weight = xlThin

SummaryCallCenter

wReport.Sheets("Hoja3").Activate

Range("s1").Value = "FONDO DE PROTECCION Y SERVICIOS SOCIALES"
Range("s1").Font.Size = 14
Range("s1").Font.Bold = True
Range("s1:z1").Merge
Range("s1:z1").HorizontalAlignment = xlCenter

Range("s2").Value = "DE LOS EMPLEADOS DE " & UCase(Name())
Range("s2").Font.Size = 14
Range("s2").Font.Bold = True
Range("s2:z2").Merge
Range("s2:z2").HorizontalAlignment = xlCenter

Range("s3").Value = "*Reporte Acumulado"
Range("s3:t3").Merge
Range("s3:t3").HorizontalAlignment = xlCenter

Range("l2:m1048576").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"


wReport.Save
wReport.Close
Application.DisplayAlerts = True
End Property
Private Function ReadTxt(NameQuery As String) As String
Dim myFile As String
Dim text As String
Dim textline As String
myFile = ThisWorkbook.Path & "\Querys\" & NameQuery & ".txt"
Open myFile For Input As #1
Do Until EOF(1)
    Line Input #1, textline
    text = text & textline
Loop
Close #1
ReadTxt = text
End Function
Private Function DeleteColumn(NumberCol As Long, NumberDelete As Long)
Dim i As Long
For i = 1 To NumberDelete
    Columns(NumberCol).EntireColumn.Delete
Next i
End Function
Private Function SummaryCallCenter()
Dim i As Long
For i = 8 To 12
Range("l" & i).Value = Application.WorksheetFunction.CountIf(Range("F:F"), Range("k" & i).Value)
Next i

On Error GoTo ehandler
For i = 8 To 12
    Range("M" & i).Value = Range("l" & i) / Application.WorksheetFunction.Sum(Range("l8").Value, Range("l9").Value, Range("l10").Value, Range("l11").Value, Range("l12").Value)
    Range("M" & i).NumberFormat = "0%"
Next i

ActiveSheet.Shapes.AddChart.Select
ActiveSheet.Shapes(1).Top = 10
ActiveSheet.Shapes(1).Left = 10
ActiveChart.ChartType = xl3DPie
ActiveChart.PlotArea.Select
ActiveChart.SetSourceData Source:=Range("k8:l12")
ActiveChart.SeriesCollection(1).Explosion = 10
ActiveChart.HasTitle = False

Exit Function
ehandler:
    Range("M" & i).Value = 0
    Range("M" & i).NumberFormat = "0%"
Resume Next
End Function
Private Function ColorsHeaders(rng As String)
Range(rng).Font.Name = "Arial"
Range(rng).Font.Bold = True
Range(rng).Font.Color = RGB(255, 255, 255)
Range(rng).Interior.Color = RGB(72, 61, 139)
Range(rng).Borders.LineStyle = xlContinuous
Range(rng).Borders.Weight = xlThin
End Function
Public Property Let CloseReport(str As String)
Application.DisplayAlerts = False
wReport.Sheets("Hoja1").Activate
wReport.Save
wReport.Close
Application.DisplayAlerts = True
End Property

