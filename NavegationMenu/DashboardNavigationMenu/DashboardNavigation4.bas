Attribute VB_Name = "Module4"


'This macro controls worksheet navigation behavior
Sub Macro4_Incidentes()
'
' Macro4_Incidentes Macro
'

'Sheets("Menu").Select
    Sheets("Incidentes_PUT").Visible = True
    Sheets("Incidentes_PUT").Select
    Range("A1").Select
End Sub

Sub Macro4_Incidentesvolver()
'
' Macro4_Incidentesvolver Macro
'

'Sheets("Incidentes_PUT").Select
    Sheets("Incidentes_PUT").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub


