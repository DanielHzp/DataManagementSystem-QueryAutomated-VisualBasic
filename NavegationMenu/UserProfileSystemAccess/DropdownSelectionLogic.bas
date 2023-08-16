Attribute VB_Name = "Módulo5"
Sub OCULTA_RESIDUOS_CLICK_en()
    
    Range("A109:A341").EntireRow.Hidden = True
                
End Sub



'The following macros execute on-click automated actions and control worksheet behavior depending on dropdown lists user selection

Sub RESIDUOS_POR_ETAPA()
If Range("A100").Value = "OPERACION" Then
Range("A109:A154").EntireRow.Hidden = False
Range("A155:A201").EntireRow.Hidden = True
Range("A202:A247").EntireRow.Hidden = True
Range("A248:A294").EntireRow.Hidden = True
Range("A295:A341").EntireRow.Hidden = True
End If
If Range("A100").Value = "SISMICA" Then
Range("A109:A154").EntireRow.Hidden = True
Range("A155:A201").EntireRow.Hidden = False
Range("A202:A247").EntireRow.Hidden = True
Range("A248:A294").EntireRow.Hidden = True
Range("A295:A341").EntireRow.Hidden = True
End If
If Range("A100").Value = "OBRA CIVIL" Then
Range("A109:A154").EntireRow.Hidden = True
Range("A155:A201").EntireRow.Hidden = True
Range("A202:A247").EntireRow.Hidden = False
Range("A248:A294").EntireRow.Hidden = True
Range("A295:A341").EntireRow.Hidden = True
End If
If Range("A100").Value = "PERFORACION" Then
Range("A109:A154").EntireRow.Hidden = True
Range("A155:A201").EntireRow.Hidden = True
Range("A202:A247").EntireRow.Hidden = True
Range("A248:A294").EntireRow.Hidden = False
Range("A295:A341").EntireRow.Hidden = True
End If
If Range("A100").Value = "WORKOVER" Then
Range("A109:A154").EntireRow.Hidden = True
Range("A155:A201").EntireRow.Hidden = True
Range("A202:A247").EntireRow.Hidden = True
Range("A248:A294").EntireRow.Hidden = True
Range("A295:A341").EntireRow.Hidden = False
End If

End Sub
Sub FORMATO_RESIDUOS_POR_ETAPA()
If Range("A100").Value = "OPERACION" Then
Call RESIDUOS_CLICK_en
End If
If Range("A100").Value = "SISMICA" Then
Call RESIDUOS_SISMICA_CLICK_en
End If
If Range("A100").Value = "OBRA CIVIL" Then
Call RESIDUOS_OBRA_CIVIL_CLICK_en
End If
If Range("A100").Value = "PERFORACION" Then
Call RESIDUOS_PERFORACION_CLICK_en
End If
If Range("A100").Value = "WORKOVER" Then
Call RESIDUOS_WORKOVER_CLICK_en
End If

End Sub

Sub VISTRA_COMPENSA()
If Range("B2").Value = "TOTAL EMPRESA" Then
Range("D2:BZ2").EntireColumn.Hidden = True
Range("CA2:CC2").EntireColumn.Hidden = False
Range("CD2:CO2").EntireColumn.Hidden = True
Range("CP2:DA2").EntireColumn.Hidden = True
End If
If Range("B2").Value = "CUENCA" Then
Range("D2:BZ2").EntireColumn.Hidden = True
Range("CA2:CC2").EntireColumn.Hidden = True
Range("CD2:CO2").EntireColumn.Hidden = False
Range("CP2:DA2").EntireColumn.Hidden = True
End If
If Range("B2").Value = "CAMPOS" Then
Range("D2:BZ2").EntireColumn.Hidden = False
Range("CA2:CC2").EntireColumn.Hidden = True
Range("CD2:CO2").EntireColumn.Hidden = False
Range("CP2:DA2").EntireColumn.Hidden = True
End If
End Sub

Sub RESIDUOS_POR_ETAPA_COORD()

Range("A103:A116").EntireRow.Hidden = True
Range("A117:A130").EntireRow.Hidden = True
Range("A131:A144").EntireRow.Hidden = True

If Range("A93").Value = "OPERACION" Then
Range("A145:A186").EntireRow.Hidden = False
Range("A187:A228").EntireRow.Hidden = True
Range("A229:A270").EntireRow.Hidden = True
Range("A271:A312").EntireRow.Hidden = True
Range("A313:A354").EntireRow.Hidden = True
End If
If Range("A93").Value = "SISMICA" Then
Range("A145:A186").EntireRow.Hidden = True
Range("A187:A228").EntireRow.Hidden = False
Range("A229:A270").EntireRow.Hidden = True
Range("A271:A312").EntireRow.Hidden = True
Range("A313:A354").EntireRow.Hidden = True
End If
If Range("A93").Value = "OBRA CIVIL" Then
Range("A145:A186").EntireRow.Hidden = True
Range("A187:A228").EntireRow.Hidden = True
Range("A229:A270").EntireRow.Hidden = False
Range("A271:A312").EntireRow.Hidden = True
Range("A313:A354").EntireRow.Hidden = True
End If
If Range("A93").Value = "PERFORACION" Then
Range("A145:A186").EntireRow.Hidden = True
Range("A187:A228").EntireRow.Hidden = True
Range("A229:A270").EntireRow.Hidden = True
Range("A271:A312").EntireRow.Hidden = False
Range("A313:A354").EntireRow.Hidden = True
End If
If Range("A93").Value = "WORKOVER" Then
Range("A145:A186").EntireRow.Hidden = True
Range("A187:A228").EntireRow.Hidden = True
Range("A229:A270").EntireRow.Hidden = True
Range("A271:A312").EntireRow.Hidden = True
Range("A313:A354").EntireRow.Hidden = False
End If
If Range("A93").Value = "NINGUNA" Then
Range("A145:A186").EntireRow.Hidden = True
Range("A187:A228").EntireRow.Hidden = True
Range("A229:A270").EntireRow.Hidden = True
Range("A271:A312").EntireRow.Hidden = True
Range("A313:A354").EntireRow.Hidden = True
End If
End Sub

Sub RESIDUOS_POR_TIPO()
If Range("A105").Value = "RECICLABLES" Then
Range("A109:A123").EntireRow.Hidden = False
Range("A124:A138").EntireRow.Hidden = True
Range("A139:A154").EntireRow.Hidden = True

Range("A155:A170").EntireRow.Hidden = False
Range("A171:A185").EntireRow.Hidden = True
Range("A186:A201").EntireRow.Hidden = True

Range("A202:A216").EntireRow.Hidden = False
Range("A217:A231").EntireRow.Hidden = True
Range("A232:A247").EntireRow.Hidden = True

Range("A248:A263").EntireRow.Hidden = False
Range("A264:A278").EntireRow.Hidden = True
Range("A279:A294").EntireRow.Hidden = True

Range("A295:A310").EntireRow.Hidden = False
Range("A311:A325").EntireRow.Hidden = True
Range("A326:A341").EntireRow.Hidden = True

End If
If Range("A105").Value = "ORDINARIOS" Then

Range("A109:A123").EntireRow.Hidden = True
Range("A124:A138").EntireRow.Hidden = False
Range("A139:A154").EntireRow.Hidden = True

Range("A155:A170").EntireRow.Hidden = True
Range("A171:A185").EntireRow.Hidden = False
Range("A186:A201").EntireRow.Hidden = True

Range("A202:A216").EntireRow.Hidden = True
Range("A217:A231").EntireRow.Hidden = False
Range("A232:A247").EntireRow.Hidden = True

Range("A248:A263").EntireRow.Hidden = True
Range("A264:A278").EntireRow.Hidden = False
Range("A279:A294").EntireRow.Hidden = True

Range("A295:A310").EntireRow.Hidden = True
Range("A311:A325").EntireRow.Hidden = False
Range("A326:A341").EntireRow.Hidden = True

End If
If Range("A105").Value = "PELIGROSOS" Then

Range("A109:A123").EntireRow.Hidden = True
Range("A124:A138").EntireRow.Hidden = True
Range("A139:A154").EntireRow.Hidden = False

Range("A155:A170").EntireRow.Hidden = True
Range("A171:A185").EntireRow.Hidden = True
Range("A186:A201").EntireRow.Hidden = False

Range("A202:A216").EntireRow.Hidden = True
Range("A217:A231").EntireRow.Hidden = True
Range("A232:A247").EntireRow.Hidden = False

Range("A248:A263").EntireRow.Hidden = True
Range("A264:A278").EntireRow.Hidden = True
Range("A279:A294").EntireRow.Hidden = False

Range("A295:A310").EntireRow.Hidden = True
Range("A311:A325").EntireRow.Hidden = True
Range("A326:A341").EntireRow.Hidden = False

End If

End Sub
Sub RESIDUOS_POR_TIPO_COORDINADOR()
If Range("A98").Value = "RECICLABLES" Then

Range("A103:A116").EntireRow.Hidden = False
Range("A117:A130").EntireRow.Hidden = True
Range("A131:A144").EntireRow.Hidden = True

End If
If Range("A98").Value = "ORDINARIOS" Then

Range("A103:A116").EntireRow.Hidden = True
Range("A117:A130").EntireRow.Hidden = False
Range("A131:A144").EntireRow.Hidden = True


End If
If Range("A98").Value = "PELIGROSOS" Then

Range("A103:A116").EntireRow.Hidden = True
Range("A117:A130").EntireRow.Hidden = False
Range("A131:A144").EntireRow.Hidden = True


End If
If Range("A98").Value = "NINGUNO" Then

Range("A103:A116").EntireRow.Hidden = True
Range("A117:A130").EntireRow.Hidden = True
Range("A131:A144").EntireRow.Hidden = True


End If

End Sub

Sub RESIDUOS_CLICK_en()
    Sheets("RESIDUOS").Visible = True
    Sheets("RESIDUOS").Select
    
End Sub
Sub RESIDUOS_SISMICA_CLICK_en()
    Sheets("RESIDUOS_SISMICA").Visible = True
    Sheets("RESIDUOS_SISMICA").Select
    
End Sub
Sub INVERSION_CLICK_en()
    Sheets("INVERSION").Visible = True
    Sheets("INVERSION").Select
End Sub
Sub VEDA_CLICK_en()
    Sheets("VEDA").Visible = True
    Sheets("VEDA").Select
End Sub
Sub COMPENSACION_CLICK_en()
    Sheets("COMPENSACION").Visible = True
    Sheets("COMPENSACION").Select
End Sub

Sub RESIDUOS_OBRA_CIVIL_CLICK_en()
    Sheets("RESIDUOS_OBRA_CIVIL").Visible = True
    Sheets("RESIDUOS_OBRA_CIVIL").Select
    
End Sub
Sub RESIDUOS_PERFORACION_CLICK_en()
    Sheets("RESIDUOS_PERFORACION").Visible = True
    Sheets("RESIDUOS_PERFORACION").Select
    
End Sub
Sub RESIDUOS_WORKOVER_CLICK_en()
    Sheets("RESIDUOS_WORKOVER").Visible = True
    Sheets("RESIDUOS_WORKOVER").Select
    
End Sub
Sub VOLVER_ING_CLICK_en()
    Sheets("INGENIERO").Visible = True
    ActiveSheet.Visible = False
    Sheets("INGENIERO").Select
End Sub
Sub VOLVER_ING_BOGCLICK_en()
    Sheets("INGENIERO_BOGOTA").Visible = True
    ActiveSheet.Visible = False
    Sheets("INGENIERO_BOGOTA").Select
End Sub

Sub VOLVER_COORD_CLICK_en()
    Sheets("REPORTE").Visible = True
    ActiveSheet.Visible = False
    Sheets("REPORTE").Select
End Sub

