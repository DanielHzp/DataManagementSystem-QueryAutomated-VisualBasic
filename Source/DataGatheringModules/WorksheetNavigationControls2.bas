Attribute VB_Name = "Módulo1"
Sub Auto_Open()
'
' Auto_Open Macro
'
Sheets("AMBIENTAL").Visible = True

Sheets("NIVELES_POZOS").Visible = False
Sheets("RESIDUOS").Visible = False
Sheets("RESIDUOS_SISMICA").Visible = False
Sheets("RESIDUOS_PERFORACION").Visible = False
Sheets("RESIDUOS_WORKOVER").Visible = False



End Sub
Sub Auto_Abrir()
'
' Auto_Open Macro
'
Call Auto_Open
'
End Sub


'Execute additional dropdownlist behavior and control worksheets visibility logic
        
Sub INGENIERO()
Sheets("USUARIOS").Visible = False
Sheets("BD COORDINADOR").Visible = False
                
                
Sheets("AMBIENTAL").Visible = True
Sheets("AMBIENTAL").Select
Sheets("USUARIOS").Visible = False
End Sub


Sub INGENIERO_BOGOTA()
     
                Sheets("AMBIENTAL_BOGOTA").Visible = True
                Sheets("AMBIENTAL_BOGOTA").Select
               
 Range("A1").EntireRow.Hidden = True
                               
                ' OCULTA COLUMNAS
                Range("F:F").EntireColumn.Hidden = True
                
                Range("BI:RJ").EntireColumn.Hidden = False
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = False
                
    
               

Range("A1").EntireRow.Hidden = False
Sheets("USUARIOS").Visible = False
End Sub

Sub RES_BOGOTA_CLIK_EN()
    Sheets("RESIDUOS_BOGOTA").Visible = True
    Sheets("RESIDUOS_BOGOTA").Select
    
End Sub
Sub OCULTA_RESIDUOS_CLICK_en()
    
    Range("A79:A388").EntireRow.Hidden = True
                
End Sub



Sub SELECCION_MES()
If Range("A2").Value = "TODOS" Then

Range("C1:AF1").EntireColumn.Hidden = False
Range("AG1:BJ1").EntireColumn.Hidden = False
Range("BK1:CN1").EntireColumn.Hidden = False
Range("CO1:DR1").EntireColumn.Hidden = False
Range("DS1:EV1").EntireColumn.Hidden = False
Range("EW1:FZ1").EntireColumn.Hidden = False
Range("GA1:HD1").EntireColumn.Hidden = False
Range("HE1:IH1").EntireColumn.Hidden = False
Range("II1:LJ1").EntireColumn.Hidden = False
Range("JM1:KP1").EntireColumn.Hidden = False
Range("KQ1:LT1").EntireColumn.Hidden = False
Range("LU1:MX1").EntireColumn.Hidden = False
End If

If Range("A2").Value = "ENERO" Then
Range("C1:AF1").EntireColumn.Hidden = False
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True
Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True
Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "FEBRERO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = False
Range("BK1:CN1").EntireColumn.Hidden = True
Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True
Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "MARZO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = False

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True
Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "ABRIL" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = False
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True
Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "MAYO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = False
Range("EW1:FZ1").EntireColumn.Hidden = True
Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "JUNIO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = False

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "JULIO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = False
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "AGOSTO" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = False
Range("II1:LJ1").EntireColumn.Hidden = True
Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "SEPTIEMBRE" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = False

Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "OCTUBRE" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True

Range("JM1:KP1").EntireColumn.Hidden = False
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = True

End If

If Range("A2").Value = "NOVIEMBRE" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True

Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = False
Range("LU1:MX1").EntireColumn.Hidden = True
End If

If Range("A2").Value = "DICIEMBRE" Then
Range("C1:AF1").EntireColumn.Hidden = True
Range("AG1:BJ1").EntireColumn.Hidden = True
Range("BK1:CN1").EntireColumn.Hidden = True

Range("CO1:DR1").EntireColumn.Hidden = True
Range("DS1:EV1").EntireColumn.Hidden = True
Range("EW1:FZ1").EntireColumn.Hidden = True

Range("GA1:HD1").EntireColumn.Hidden = True
Range("HE1:IH1").EntireColumn.Hidden = True
Range("II1:LJ1").EntireColumn.Hidden = True

Range("JM1:KP1").EntireColumn.Hidden = True
Range("KQ1:LT1").EntireColumn.Hidden = True
Range("LU1:MX1").EntireColumn.Hidden = False
End If


End Sub



Sub RESIDUOS_POR_ETAPA()

If Range("A70").Value = "OPERACION" Then
Range("A79:A140").EntireRow.Hidden = False
Range("A141:A202").EntireRow.Hidden = True
Range("A203:A264").EntireRow.Hidden = True
Range("A265:A326").EntireRow.Hidden = True
Range("A327:A388").EntireRow.Hidden = True
End If
If Range("A70").Value = "SISMICA" Then
Range("A79:A140").EntireRow.Hidden = True
Range("A141:A202").EntireRow.Hidden = False
Range("A203:A264").EntireRow.Hidden = True
Range("A265:A326").EntireRow.Hidden = True
Range("A327:A388").EntireRow.Hidden = True
End If
If Range("A70").Value = "OBRA CIVIL" Then
Range("A79:A140").EntireRow.Hidden = True
Range("A141:A202").EntireRow.Hidden = True
Range("A203:A264").EntireRow.Hidden = False
Range("A265:A326").EntireRow.Hidden = True
Range("A327:A388").EntireRow.Hidden = True
End If
If Range("A70").Value = "PERFORACION" Then
Range("A79:A140").EntireRow.Hidden = True
Range("A141:A202").EntireRow.Hidden = True
Range("A203:A264").EntireRow.Hidden = True
Range("A265:A326").EntireRow.Hidden = False
Range("A327:A388").EntireRow.Hidden = True
End If
If Range("A70").Value = "WORKOVER" Then
Range("A79:A140").EntireRow.Hidden = True
Range("A141:A202").EntireRow.Hidden = True
Range("A203:A264").EntireRow.Hidden = True
Range("A265:A326").EntireRow.Hidden = True
Range("A327:A388").EntireRow.Hidden = False
End If


End Sub




Sub FORMATO_RESIDUOS_POR_ETAPA()
If Range("A70").Value = "OPERACION" Then
On Error GoTo hereOP:
Call RESIDUOS_CLICK_en

Exit Sub
hereOP:
    MsgBox ("Error al cargar la hoja RESIDUOS OPERACIÓN")
End If
If Range("A70").Value = "SISMICA" Then
Call RESIDUOS_SISMICA_CLICK_en
End If
If Range("A70").Value = "OBRA CIVIL" Then
Call RESIDUOS_OBRA_CIVIL_CLICK_en
End If
If Range("A70").Value = "PERFORACION" Then
Call RESIDUOS_PERFORACION_CLICK_en
End If
If Range("A70").Value = "WORKOVER" Then
Call RESIDUOS_WORKOVER_CLICK_en

End If

End Sub



Sub RESIDUOS_POR_TIPO()

If Range("A75").Value = "RECICLABLES" Then

Range("A79:A93").EntireRow.Hidden = False
Range("A94:A108").EntireRow.Hidden = True
Range("A109:A124").EntireRow.Hidden = True
Range("A125:A140").EntireRow.Hidden = True

Range("A141:A155").EntireRow.Hidden = False
Range("A156:A170").EntireRow.Hidden = True
Range("A171:A186").EntireRow.Hidden = True
Range("A187:A202").EntireRow.Hidden = True

Range("A203:A217").EntireRow.Hidden = False
Range("A218:A232").EntireRow.Hidden = True
Range("A233:A248").EntireRow.Hidden = True
Range("A249:A264").EntireRow.Hidden = True

Range("A265:A279").EntireRow.Hidden = False
Range("A280:A294").EntireRow.Hidden = True
Range("A295:A310").EntireRow.Hidden = True
Range("A311:A326").EntireRow.Hidden = True

Range("A327:A341").EntireRow.Hidden = False
Range("A342:A356").EntireRow.Hidden = True
Range("A357:A372").EntireRow.Hidden = True
Range("A373:A388").EntireRow.Hidden = True



End If
If Range("A75").Value = "NO RECICLABLES" Then

Range("A79:A93").EntireRow.Hidden = True
Range("A94:A108").EntireRow.Hidden = False
Range("A109:A124").EntireRow.Hidden = True
Range("A125:A140").EntireRow.Hidden = True

Range("A141:A155").EntireRow.Hidden = True
Range("A156:A170").EntireRow.Hidden = False
Range("A171:A186").EntireRow.Hidden = True
Range("A187:A202").EntireRow.Hidden = True

Range("A203:A217").EntireRow.Hidden = True
Range("A218:A232").EntireRow.Hidden = False
Range("A233:A248").EntireRow.Hidden = True
Range("A249:A264").EntireRow.Hidden = True

Range("A265:A279").EntireRow.Hidden = True
Range("A280:A294").EntireRow.Hidden = False
Range("A295:A310").EntireRow.Hidden = True
Range("A311:A326").EntireRow.Hidden = True

Range("A327:A341").EntireRow.Hidden = True
Range("A342:A356").EntireRow.Hidden = False
Range("A357:A372").EntireRow.Hidden = True
Range("A373:A388").EntireRow.Hidden = True


End If
If Range("A75").Value = "PELIGROSOS" Then

Range("A79:A93").EntireRow.Hidden = True
Range("A94:A108").EntireRow.Hidden = True
Range("A109:A124").EntireRow.Hidden = False
Range("A125:A140").EntireRow.Hidden = True

Range("A141:A155").EntireRow.Hidden = True
Range("A156:A170").EntireRow.Hidden = True
Range("A171:A186").EntireRow.Hidden = False
Range("A187:A202").EntireRow.Hidden = True

Range("A203:A217").EntireRow.Hidden = True
Range("A218:A232").EntireRow.Hidden = True
Range("A233:A248").EntireRow.Hidden = False
Range("A249:A264").EntireRow.Hidden = True

Range("A265:A279").EntireRow.Hidden = True
Range("A280:A294").EntireRow.Hidden = True
Range("A295:A310").EntireRow.Hidden = False
Range("A311:A326").EntireRow.Hidden = True

Range("A327:A341").EntireRow.Hidden = True
Range("A342:A356").EntireRow.Hidden = True
Range("A357:A372").EntireRow.Hidden = False
Range("A373:A388").EntireRow.Hidden = True

End If

If Range("A75").Value = "ESPECIALES" Then

Range("A79:A93").EntireRow.Hidden = True
Range("A94:A108").EntireRow.Hidden = True
Range("A109:A124").EntireRow.Hidden = True
Range("A125:A140").EntireRow.Hidden = False

Range("A141:A155").EntireRow.Hidden = True
Range("A156:A170").EntireRow.Hidden = True
Range("A171:A186").EntireRow.Hidden = True
Range("A187:A202").EntireRow.Hidden = False

Range("A203:A217").EntireRow.Hidden = True
Range("A218:A232").EntireRow.Hidden = True
Range("A233:A248").EntireRow.Hidden = True
Range("A249:A264").EntireRow.Hidden = False

Range("A265:A279").EntireRow.Hidden = True
Range("A280:A294").EntireRow.Hidden = True
Range("A295:A310").EntireRow.Hidden = True
Range("A311:A326").EntireRow.Hidden = False

Range("A327:A341").EntireRow.Hidden = True
Range("A342:A356").EntireRow.Hidden = True
Range("A357:A372").EntireRow.Hidden = True
Range("A373:A388").EntireRow.Hidden = False

End If

End Sub

Sub RESIDUOS_CLICK_en()
    Sheets("RESIDUOS").Visible = True
    Sheets("RESIDUOS").Select
    
End Sub
Sub RESIDUOS_2020_CLICK_en()
    Sheets("RESIDUOS_2020").Visible = True
    Sheets("RESIDUOS_2020").Select
    
End Sub
Sub RESIDUOS_SISMICA_CLICK_en()
    Sheets("RESIDUOS_SISMICA").Visible = True
    Sheets("RESIDUOS_SISMICA").Select
    
End Sub
Sub RESIDUOS_SISMICA_2020_CLICK_en()
    Sheets("RESIDUOS_SISMICA_2020").Visible = True
    Sheets("RESIDUOS_SISMICA_2020").Select
    
End Sub

Sub RESIDUOS_OBRA_CIVIL_CLICK_en()
    Sheets("RESIDUOS_OBRA_CIVIL").Visible = True
    Sheets("RESIDUOS_OBRA_CIVIL").Select
    
End Sub
Sub RESIDUOS_OBRA_CIVIL_2020_CLICK_en()
    Sheets("RESIDUOS_OBRA_CIVIL_2020").Visible = True
    Sheets("RESIDUOS_OBRA_CIVIL_2020").Select
    
End Sub
Sub RESIDUOS_PERFORACION_CLICK_en()
    Sheets("RESIDUOS_PERFORACION").Visible = True
    Sheets("RESIDUOS_PERFORACION").Select
    
End Sub
Sub RESIDUOS_PERFORACION_2020_CLICK_en()
    Sheets("RESIDUOS_PERFORACION_2020").Visible = True
    Sheets("RESIDUOS_PERFORACION_2020").Select
    
End Sub
Sub RESIDUOS_WORKOVER_CLICK_en()
    Sheets("RESIDUOS_WORKOVER").Visible = True
    Sheets("RESIDUOS_WORKOVER").Select
    
End Sub
Sub RESIDUOS_WORKOVER_2020_CLICK_en()
    Sheets("RESIDUOS_WORKOVER_2020").Visible = True
    Sheets("RESIDUOS_WORKOVER_2020").Select
    
End Sub

Sub VOLVER_BOGOTA_ING_CLICK_en()
    Sheets("AMBIENTAL_BOGOTA").Visible = True
    ActiveSheet.Visible = False
    Sheets("AMBIENTAL_BOGOTA").Select
End Sub

Sub VOLVER_ING_CLICK_en()
    Sheets("AMBIENTAL").Visible = True
    ActiveSheet.Visible = False
    Sheets("AMBIENTAL").Select
    
End Sub









'This macro adds rows to the table where new data has to be added
'It is executed when the user clicks a button in the worksheet 'AGUA_RESID_BLOQ 2021'
Sub AGREGAR_FILAS_AGUARESIDUAL()

On Error GoTo here
Dim cantFilas As Integer

'Obtain number of rows to add on the worksheet
cantFilas = Application.InputBox(Prompt:="Ingrese el numero de filas a agregar", Type:=1)

'Add the rows below active data input cell
ActiveCell.EntireRow.Resize(cantFilas).Insert Shift:=x1Down
Exit Sub
here:
        MsgBox ("Close the window")

End Sub





'Data logic validation in the worksheet 'AGUA_RESID_BLOQ 2021'
'Alerts the user when inconsistencies are identified when the entire worksheet is iterated
Sub ReportarAlertasAguaBloque()
Dim i As Integer
Dim j As Integer
j = 0

On Error GoTo here


Range("Y3").Select
For i = 1 To 900

If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S)" & i + 3 & "  DEL MES DE ENERO"
j = j + 1
End If
Next i

Range("AU3").Select

For i = 1 To 900
If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S):" & i + 3 & "  DEL MES DE FEBRERO"
j = j + 1
End If
Next i

Range("BQ3").Select
For i = 1 To 900
If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S):" & i + 3 & "  DEL MES DE MARZO"
j = j + 1
End If
Next i

Range("CM3").Select
For i = 1 To 900
If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S):" & i + 3 & "  DEL MES DE ABRIL"
j = j + 1
End If
Next i

Range("DI3").Select
For i = 1 To 900
If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S):" & i + 3 & "  DEL MES DE MAYO"
j = j + 1
End If
Next i

Range("EE3").Select
For i = 1 To 900
If ActiveCell.Offset(i, 0).Text = "Las cantidades reportadas NO coinciden con el total de consumo" Then
MsgBox "REVISAR LOS DATOS DE LA(S) FILA(S):" & i + 3 & "  DEL MES DE JUNIO"
j = j + 1
End If
Next i

If j = 0 Then
MsgBox ("OK TODAS LAS CANTIDADES M3 REPORTADAS, TODOS LOS VOLUMENES COINCIDEN")
End If



Exit Sub



here:
MsgBox ("Keep verifying the reported amounts are correct")


End Sub














