Attribute VB_Name = "Módulo1"

'Handles worksheets default visibility behavior when the file is opened
Sub Auto_Open()
'
' Auto_Open Macro
'
Sheets("INICIO").Visible = True
Sheets("USUARIOS").Visible = False
Sheets("LIDER").Visible = False
Sheets("REPORTE").Visible = False
Sheets("LIDER_CAMPOS").Visible = False
Sheets("BD LIDER").Visible = False
Sheets("BD LIDER DATOS IMPORTA").Visible = False
Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
Sheets("INGENIERO").Visible = False
Sheets("INGENIERO_BOGOTA").Visible = False
Sheets("AGUA_MEDIDOR").Visible = False
Sheets("AGUA_BLOQUE").Visible = False
Sheets("CONTROL_VERTI").Visible = False
Sheets("AGUA_RESID_BLOQ").Visible = False
Sheets("NIVELES_POZOS").Visible = False
Sheets("RESIDUOS").Visible = False
Sheets("RESIDUOS_SISMICA").Visible = False
Sheets("RESIDUOS_PERFORACION").Visible = False
Sheets("RESIDUOS_BOGOTA").Visible = False
Sheets("RESIDUOS_WORKOVER").Visible = False
Sheets("COORDINADOR").Visible = False
Sheets("COORDINADOR_COMPENSACIONES").Visible = False
Sheets("BD COORDINADOR").Visible = False
Sheets("IC_PAPEL").Visible = False
Sheets("IC_ENERGIA").Visible = False
Sheets("IC_REFORESTACION").Visible = False
Sheets("IC_INC_AMB").Visible = False
Sheets("IC_AGUA").Visible = False
Sheets("IC_DIESEL").Visible = False
Sheets("IC_GAS").Visible = False
Sheets("IC_TOTAL_RS").Visible = False
Sheets("IC_VERTIMIENTOS").Visible = False
Sheets("IC_RECICLAJE").Visible = False
Sheets("INVERSION").Visible = False
Sheets("VEDA").Visible = False
Sheets("INCIDENTES").Visible = False
Sheets("COMPENSACION").Visible = False
Sheets("LIDER_PUTUMAYO").Visible = False
Sheets("GRAFICAS_COMPENSACIONES").Visible = False
UserForm1.Show
'
End Sub



'Controls  visibility for data entry worksheets navigation
Sub LIDER()
               
                Sheets("USUARIOS").Visible = False
                
                Sheets("LIDER").Visible = True
                Sheets("REPORTE").Visible = True
                Sheets("LIDER_CAMPOS").Visible = True
                
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = False
                Sheets("AGUA_MEDIDOR").Visible = False
                
                Sheets("LIDER").Select
                Range("A1").Select
       
End Sub




'The code below loads data entry cell ranges depending on the user profile and access credentials input
'Controls worksheet behavior when dropdown list values are selected by the user


Sub IC_AGUA()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = False
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = False
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = False
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = False
                
                Range("A1").Select
        End If
        
        If Range("H1").Value = "COORDINADOR VALLE MM" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR LLANOS" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR BOGOTA" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = False
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
        If Range("H1").Value = "LIDER" Then
                
                Sheets("IC_AGUA").Visible = True
                Sheets("IC_AGUA").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A102:A199").EntireRow.Hidden = True
                 Range("A200:A230").EntireRow.Hidden = False

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If

Sheets("USUARIOS").Visible = False
End Sub



Sub IC_VERTIMIENTOS()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = False
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = False
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = False
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = False
                
                Range("A1").Select
        End If
        
        If Range("H1").Value = "COORDINADOR VALLE MM" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR LLANOS" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR BOGOTA" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = False
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
        If Range("H1").Value = "LIDER" Then
                
                Sheets("IC_VERTIMIENTOS").Visible = True
                Sheets("IC_VERTIMIENTOS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A102:A199").EntireRow.Hidden = True
                 Range("A200:A230").EntireRow.Hidden = False

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
Sheets("USUARIOS").Visible = False
End Sub



'This subroutines recursively execute the same algorithm and dynamically display cell ranges for data entry
Sub IC_RECICLAJE()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = False
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = False
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = False
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = False
                
                Range("A1").Select
        End If
        
        If Range("H1").Value = "COORDINADOR VALLE MM" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR LLANOS" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR BOGOTA" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = False
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
        If Range("H1").Value = "LIDER" Then
                
                Sheets("IC_RECICLAJE").Visible = True
                Sheets("IC_RECICLAJE").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A102:A199").EntireRow.Hidden = True
                 Range("A200:A230").EntireRow.Hidden = False

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
Sheets("USUARIOS").Visible = False
End Sub


'Additional user profiles validation and display behavior
Sub IC_TOTAL_RS()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = False
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = False
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = False
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                Range("A:T").EntireColumn.Hidden = True
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = False
                
                Range("A1").Select
        End If
        
        If Range("H1").Value = "COORDINADOR VALLE MM" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR LLANOS" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = False
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        End If
        If Range("H1").Value = "COORDINADOR BOGOTA" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = False
                 ' FORMULARIO GRAN TOTAL
                 Range("A200:A230").EntireRow.Hidden = True
                
                

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
        If Range("H1").Value = "LIDER" Then
                
                Sheets("IC_TOTAL_RS").Visible = True
                Sheets("IC_TOTAL_RS").Select
                ' FORMULARIO OBS CUENCA PUT
                 Range("A2:A26").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA VMM
                 Range("A27:A51").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA LLANOS
                 Range("A52:A76").EntireRow.Hidden = True
                ' FORMULARIO OBS CUENCA BOGOTA
                 Range("A77:A101").EntireRow.Hidden = True
                 ' FORMULARIO GRAN TOTAL
                 Range("A102:A199").EntireRow.Hidden = True
                 Range("A200:A230").EntireRow.Hidden = False

                Range("A:T").EntireColumn.Hidden = False
                Range("U:AN").EntireColumn.Hidden = True
                Range("AO:BH").EntireColumn.Hidden = True
                Range("BI:CB").EntireColumn.Hidden = True
                Range("CC:CV").EntireColumn.Hidden = True
                
                Range("A1").Select
        
        End If
Sheets("USUARIOS").Visible = False
End Sub


'Loads 'BOGOTA' report for this specific profile
Sub REPORTE()
Sheets("REPORTE").Visible = True
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        
If Range("H1").Value = "COORDINADOR BOGOTA" Then
                Sheets("REPORTE").Select
                Range("A2:A13").EntireRow.Hidden = False
                Range("A4:B4").EntireRow.Hidden = True
                Range("A6:B7").EntireRow.Hidden = True
                Range("A11:B12").EntireRow.Hidden = True
Else
Sheets("REPORTE").Select
Range("A2:A13").EntireRow.Hidden = False

End If
Sheets("USUARIOS").Visible = False

End Sub
Sub COORDINADOR_COMPENSACIONES()
    Sheets("INVERSION").Visible = False
    Sheets("VEDA").Visible = False
    Sheets("COMPENSACION").Visible = False
    Sheets("GRAFICAS_COMPENSACIONES").Visible = True
    Sheets("COORDINADOR_COMPENSACIONES").Visible = True
    Sheets("USUARIOS").Visible = False
    
End Sub


'Controls worksheet behavior for the following user profile
Sub INGENIERO()

Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
If Range("H1").Value = "ING_1" Or Range("H1").Value = "ING_2" Or Range("H1").Value = "ING_3" Or Range("H1").Value = "ING_4" Or Range("H1").Value = "ING_5" Or Range("H1").Value = "ING_6" Or Range("H1").Value = "ING_7" Or Range("H1").Value = "ING_8" Or Range("H1").Value = "ING_9" Or Range("H1").Value = "ING_10" Or Range("H1").Value = "ING_11" Or Range("H1").Value = "ING_12" Or Range("H1").Value = "ING_13" Or Range("H1").Value = "ING_14" Or Range("H1").Value = "ING_15" Then

Call INGENIERO_PUTUMAYO
End If
        If Range("H1").Value = "ING_VMM_1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                Sheets("INGENIERO").Visible = True
                
                 
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = False
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True
                
        End If
        If Range("H1").Value = "ING_VMM_2" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False

                
                Sheets("INGENIERO").Visible = True
                 
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = False
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True

                
        End If
        If Range("H1").Value = "ING_VMM_3" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = False
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False


               ' OCULTA COLUMNAS
                Range("BY:LL").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("LM:MU").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("MV:TP").EntireColumn.Hidden = True


                
        End If
        If Range("H1").Value = "ING_VMM_4" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
  Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = False
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False
               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


                
        End If
        If Range("H1").Value = "ING_VMM_5" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = False
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


                
        End If
        
        If Range("H1").Value = "ING_VMM_6" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = False
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True




        End If
        If Range("H1").Value = "ING_VMM_7" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = False
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_VMM_8" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = False
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



        End If
        If Range("H1").Value = "ING_LLANOS_1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = False
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


        End If
        If Range("H1").Value = "ING_BGT_1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = False
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


        End If
        If Range("H1").Value = "ING_SINU_1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = False
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False
               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


        End If

Range("A109:A341").EntireRow.Hidden = True
Sheets("USUARIOS").Visible = False
End Sub




Sub INGENIERO_PUTUMAYO()
If Range("H1").Value = "ING_1" Then
                         
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
               
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = False
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True
                
                
        End If
        If Range("H1").Value = "ING_2" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
             
               
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = False
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


                
        End If
        If Range("H1").Value = "ING_3" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = False
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


                
        End If
        If Range("H1").Value = "ING_4" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
               
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = False
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_5" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                
                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = False
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



               
        End If
        If Range("H1").Value = "ING_6" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
  Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = False
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True




                
        End If
        If Range("H1").Value = "ING_7" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = False
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_8" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = False
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_9" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = False
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
         If Range("H1").Value = "ING_10" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = False
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_11" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = False
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_12" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = False
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_13" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
        If Range("H1").Value = "ING_14" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
         If Range("H1").Value = "ING_15" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = False
                Sheets("COORDINADOR").Visible = False
                Sheets("BD COORDINADOR").Visible = False
                Sheets("BD COORDINADOR DATOS IMPORTA").Visible = False
                
                Sheets("LIDER").Visible = False
                Sheets("REPORTE").Visible = False
                Sheets("LIDER_CAMPOS").Visible = False
                Sheets("BD LIDER").Visible = False
                Sheets("BD LIDER DATOS IMPORTA").Visible = False
                

                
                Sheets("INGENIERO").Visible = True
                
                
                Sheets("INGENIERO").Select
                Range("A1").EntireRow.Hidden = True
                Range("G:H").EntireColumn.Hidden = True
                Range("I:J").EntireColumn.Hidden = True
                Range("K:L").EntireColumn.Hidden = True
                Range("M:N").EntireColumn.Hidden = True
                Range("O:P").EntireColumn.Hidden = True
                Range("Q:R").EntireColumn.Hidden = True
                Range("S:T").EntireColumn.Hidden = True
                Range("U:V").EntireColumn.Hidden = True
                Range("W:X").EntireColumn.Hidden = True
                Range("Y:Z").EntireColumn.Hidden = True
                Range("AA:AB").EntireColumn.Hidden = True
                Range("AC:AD").EntireColumn.Hidden = True
                Range("AE:AF").EntireColumn.Hidden = True
                Range("AG:AH").EntireColumn.Hidden = True
                Range("AI:AJ").EntireColumn.Hidden = True
                Range("AK:AL").EntireColumn.Hidden = True
                Range("AM:AN").EntireColumn.Hidden = True
                Range("AO:AP").EntireColumn.Hidden = True
                Range("AQ:AR").EntireColumn.Hidden = True
                Range("AS:AT").EntireColumn.Hidden = True
                Range("AU:AV").EntireColumn.Hidden = True
                Range("AW:AX").EntireColumn.Hidden = True
                Range("AY:AZ").EntireColumn.Hidden = True
                Range("BA:BB").EntireColumn.Hidden = True
                Range("BC:BD").EntireColumn.Hidden = True
                Range("BE:BF").EntireColumn.Hidden = True

                
               ' graficas
                Range("BG:BO").EntireColumn.Hidden = False
                Range("BP:BX").EntireColumn.Hidden = False
                
    
                Range("A1").EntireRow.Hidden = False

               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True


               ' OCULTA COLUMNAS
                Range("BY:RJ").EntireColumn.Hidden = True
                ' MUESTRA COLUMNAS
                Range("RK:SS").EntireColumn.Hidden = False
                ' OCULTA COLUMNAS
                Range("ST:TP").EntireColumn.Hidden = True



                
        End If
       
End Sub

Sub INGENIERO_BOGOTA()
     
                Sheets("INGENIERO_BOGOTA").Visible = True
                Sheets("INGENIERO_BOGOTA").Select
               
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


Sub COORDINADOR()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo -  moqueta
                Range("G:H").EntireColumn.Hidden = False
                'cuencas putumayo - costayaco
                Range("I:J").EntireColumn.Hidden = True
                'cuencas putumayo -  santana
                Range("K:L").EntireColumn.Hidden = False
                'cuencas putumayo - cumplidor
                Range("M:N").EntireColumn.Hidden = True
                'cuencas putumayo - siriri
                Range("O:P").EntireColumn.Hidden = True
                'cuencas putumayo - pomorroso
                Range("Q:R").EntireColumn.Hidden = True
                'cuencas putumayo - nancy
                Range("S:T").EntireColumn.Hidden = True
                'cuencas putumayo - colibri
                Range("U:V").EntireColumn.Hidden = True
                'cuencas putumayo - vonu
                Range("W:X").EntireColumn.Hidden = True
                'cuencas putumayo - canelo nogal
                Range("y:z").EntireColumn.Hidden = False
                'cuencas putumayo - burdine
                Range("aa:ab").EntireColumn.Hidden = True
                'cuencas putumayo - alea 1848
                Range("ac:ad").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 1
                Range("ae:af").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 2
                Range("ag:ah").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 3
                Range("ai:aj").EntireColumn.Hidden = True
                
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = False
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo -  moqueta
                Range("G:H").EntireColumn.Hidden = True
                'cuencas putumayo - costayaco
                Range("I:J").EntireColumn.Hidden = False
                'cuencas putumayo -  santana
                Range("K:L").EntireColumn.Hidden = True
                'cuencas putumayo - cumplidor
                Range("M:N").EntireColumn.Hidden = True
                'cuencas putumayo - siriri
                Range("O:P").EntireColumn.Hidden = True
                'cuencas putumayo - pomorroso
                Range("Q:R").EntireColumn.Hidden = True
                'cuencas putumayo - nancy
                Range("S:T").EntireColumn.Hidden = True
                'cuencas putumayo - colibri
                Range("U:V").EntireColumn.Hidden = True
                'cuencas putumayo - vonu
                Range("W:X").EntireColumn.Hidden = False
                'cuencas putumayo - canelo nogal
                Range("y:z").EntireColumn.Hidden = True
                'cuencas putumayo - burdine
                Range("aa:ab").EntireColumn.Hidden = True
                'cuencas putumayo - alea 1848
                Range("ac:ad").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 1
                Range("ae:af").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 2
                Range("ag:ah").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 3
                Range("ai:aj").EntireColumn.Hidden = True
                
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = False
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo -  moqueta
                Range("G:H").EntireColumn.Hidden = True
                'cuencas putumayo - costayaco
                Range("I:J").EntireColumn.Hidden = True
                'cuencas putumayo -  santana
                Range("K:L").EntireColumn.Hidden = True
                'cuencas putumayo - cumplidor
                Range("M:N").EntireColumn.Hidden = False
                'cuencas putumayo - siriri
                Range("O:P").EntireColumn.Hidden = True
                'cuencas putumayo - pomorroso
                Range("Q:R").EntireColumn.Hidden = True
                'cuencas putumayo - nancy
                Range("S:T").EntireColumn.Hidden = False
                'cuencas putumayo - colibri
                Range("U:V").EntireColumn.Hidden = True
                'cuencas putumayo - vonu
                Range("W:X").EntireColumn.Hidden = True
                'cuencas putumayo - canelo nogal
                Range("y:z").EntireColumn.Hidden = True
                'cuencas putumayo - burdine
                Range("aa:ab").EntireColumn.Hidden = False
                'cuencas putumayo - alea 1848
                Range("ac:ad").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 1
                Range("ae:af").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 2
                Range("ag:ah").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 3
                Range("ai:aj").EntireColumn.Hidden = True
                
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = False
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo -  moqueta
                Range("G:H").EntireColumn.Hidden = True
                'cuencas putumayo - costayaco
                Range("I:J").EntireColumn.Hidden = True
                'cuencas putumayo -  santana
                Range("K:L").EntireColumn.Hidden = True
                'cuencas putumayo - cumplidor
                Range("M:N").EntireColumn.Hidden = True
                'cuencas putumayo - siriri
                Range("O:P").EntireColumn.Hidden = False
                'cuencas putumayo - pomorroso
                Range("Q:R").EntireColumn.Hidden = False
                'cuencas putumayo - nancy
                Range("S:T").EntireColumn.Hidden = True
                'cuencas putumayo - colibri
                Range("U:V").EntireColumn.Hidden = False
                'cuencas putumayo - vonu
                Range("W:X").EntireColumn.Hidden = True
                'cuencas putumayo - canelo nogal
                Range("y:z").EntireColumn.Hidden = True
                'cuencas putumayo - burdine
                Range("aa:ab").EntireColumn.Hidden = True
                'cuencas putumayo - alea 1848
                Range("ac:ad").EntireColumn.Hidden = False
                'cuencas putumayo - vacio 1
                Range("ae:af").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 2
                Range("ag:ah").EntireColumn.Hidden = True
                'cuencas putumayo - vacio 3
                Range("ai:aj").EntireColumn.Hidden = True
                
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = False
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        
        If Range("H1").Value = "COORDINADOR VALLE MM" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo
                Range("G:AJ").EntireColumn.Hidden = True
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = False
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = False
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        If Range("H1").Value = "COORDINADOR LLANOS" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo
                Range("G:AJ").EntireColumn.Hidden = True
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = False
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = True
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = False
                'bogota
                Range("BO:BP").EntireColumn.Hidden = True
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
        If Range("H1").Value = "COORDINADOR BOGOTA" Then
                
                Sheets("USUARIOS").Visible = False
                
                Sheets("REPORTE").Visible = True
                Sheets("COORDINADOR").Visible = True
                Sheets("COORDINADOR").Select
                
                Range("A1").EntireRow.Hidden = True
                
                Range("A:E").EntireColumn.Hidden = False
                Range("F:F").EntireColumn.Hidden = True
                
                'cuencas putumayo
                Range("G:AJ").EntireColumn.Hidden = True
                'cuencas vmm
                Range("AK:AZ").EntireColumn.Hidden = True
                'cuencas llanos
                Range("BA:BB").EntireColumn.Hidden = True
                'cuencas bogota
                Range("BC:BD").EntireColumn.Hidden = False
                'cuencas SINU
                Range("BE:BF").EntireColumn.Hidden = True
                
                'total general
                Range("BG:BH").EntireColumn.Hidden = True
                
                'totales por cuenca
                'putumayo
                Range("BI:BJ").EntireColumn.Hidden = True
                'vmm
                Range("BK:BL").EntireColumn.Hidden = True
                'llanos
                Range("BM:BN").EntireColumn.Hidden = True
                'bogota
                Range("BO:BP").EntireColumn.Hidden = False
                'SINU
                Range("BQ:BR").EntireColumn.Hidden = True
                'putumayo NORTE 1
                Range("BS:BT").EntireColumn.Hidden = True
                'putumayo NORTE 2
                Range("BU:BV").EntireColumn.Hidden = True
                'putumayo SUR 1
                Range("BW:BX").EntireColumn.Hidden = True
                'putumayo SUR 2
                Range("BY:BZ").EntireColumn.Hidden = True
                
                Range("A6:A102").EntireRow.Hidden = False
                Range("A103:A354").EntireRow.Hidden = True
                Range("A355:A400").EntireRow.Hidden = False
                Range("A1").EntireRow.Hidden = False
                
        End If
    
Range("A6:A102").EntireRow.Hidden = False
Range("A103:A354").EntireRow.Hidden = True
Range("A355:A400").EntireRow.Hidden = False
Range("A1").EntireRow.Hidden = False
Sheets("USUARIOS").Visible = False

End Sub



Sub LIDER_PUTUMAYO()
Sheets("REPORTE").Visible = True
Sheets("LIDER_PUTUMAYO").Visible = True
Sheets("USUARIOS").Visible = False
Sheets("LIDER_PUTUMAYO").Select
End Sub



'AT THIS POINT ALL USER/SYSTEM PROFILES HAVE BEEN VALIDATED AND FILE DISPLAY ACTIONS HAVE BEEN EXECUTED ACCORDINGLY


