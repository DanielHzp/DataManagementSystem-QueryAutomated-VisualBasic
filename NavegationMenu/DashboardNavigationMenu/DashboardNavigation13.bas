Attribute VB_Name = "Module13"




'This macro controls worksheet navigation behavior

Sub Macro13_AGUA()
'
' Macro13_AGUA Macro
'

'Sheets("Menu").Select
    Sheets("Historico Monitoreos AGUA").Visible = True
    Sheets("Historico Monitoreos AGUA").Select
    Range("A1").Select
End Sub

Sub Macro13_AGUAvolver()
'
' Macro13_AGUAvolver Macro
'

'Sheets("Historico Monitoreos AGUA").Select
    Sheets("Historico Monitoreos AGUA").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
