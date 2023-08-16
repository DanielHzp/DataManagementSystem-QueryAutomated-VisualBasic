Attribute VB_Name = "Módulo6"



'The following macros control the visibility of data input reports on the current user/system profile

Sub IC_ENERGIA()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
                
                Sheets("IC_ENERGIA").Visible = True
                Sheets("IC_ENERGIA").Select
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
Sub IC_GAS()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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
                
                Sheets("IC_GAS").Visible = True
                Sheets("IC_GAS").Select
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

Sub IC_DIESEL()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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
                
                Sheets("IC_DIESEL").Visible = True
                Sheets("IC_DIESEL").Select
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

Sub IC_PAPEL()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
                
                Sheets("IC_PAPEL").Visible = True
                Sheets("IC_PAPEL").Select
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
Sub IC_INC_AMB()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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
                
                Sheets("IC_INC_AMB").Visible = True
                Sheets("IC_INC_AMB").Select
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

Sub IC_REFORESTACION()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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
                
                Sheets("IC_REFORESTACION").Visible = True
                Sheets("IC_REFORESTACION").Select
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

