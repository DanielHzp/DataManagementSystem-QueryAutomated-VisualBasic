Attribute VB_Name = "Module7"


'This macro controls worksheet navigation behavior

Sub Macro7_IncidentesVMM()
'
' Macro7_IncidentesVMM Macro
'

'Sheets("Menu").Select
    Sheets("Incidentes_VMM").Visible = True
    Sheets("Incidentes_VMM").Select
    Range("A1").Select
End Sub

Sub Macro7_IncidentesVMMvolver()
'
' Macro7_Incidentesvolver Macro
'

'Sheets("Incidentes_VMM").Select
    Sheets("Incidentes_VMM").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
