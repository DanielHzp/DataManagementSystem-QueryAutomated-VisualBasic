Attribute VB_Name = "Module1"

'This macro controls worksheet navigation behavior


Sub Macro1_Estadisticas()
'
' Macro1_EST Macro
'

'Sheets("Menu").Select
    Sheets("Estadisticas").Visible = True
    Sheets("Estadisticas").Select
    Range("A1").Select
End Sub


Sub Button2_Click()

End Sub
Sub Macro1_Estadisticasvolver()
'
' Macro1_ESTvolver Macro
'

'Sheets("Estadisticas").Select
    Sheets("Estadisticas").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub

Sub OpenProd()

UserForm1.Show

End Sub
