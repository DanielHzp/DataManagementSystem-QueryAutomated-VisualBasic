Attribute VB_Name = "Module11"



'This macro controls worksheet navigation behavior

Sub Macro11_IncidentesGTE()
'
' Macro11_IncidenteGTE Macro
'

'Sheets("Menu").Select
    Sheets("Incidentes GTE").Visible = True
    Sheets("Incidentes GTE").Select
    Range("A1").Select
End Sub

Sub Macro11_IncidentesGTEvolver()
'
' Macro11_IncidentesGTEvolver Macro
'

'Sheets("Incidentes GTE").Select
    Sheets("Incidentes GTE").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub

