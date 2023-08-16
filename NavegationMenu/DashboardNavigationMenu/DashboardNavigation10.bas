Attribute VB_Name = "Module10"


'This macro controls worksheet navigation behavior

Sub Macro10_MonitoreosGTE()
'
' Macro10_MonitoreosGTE Macro
'

'Sheets("Menu").Select
    Sheets("Monitoreos 2020").Visible = True
    Sheets("Monitoreos 2020").Select
    Range("A1").Select
End Sub

Sub Macro10_MonitoreosGTEvolver()
'
' Macro10_MonitoreosGTEvolver Macro
'

'Sheets("Monitoreos 2020").Select
    Sheets("Monitoreos 2020").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
