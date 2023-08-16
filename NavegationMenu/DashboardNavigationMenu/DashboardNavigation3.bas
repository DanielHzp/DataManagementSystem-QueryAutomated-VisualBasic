Attribute VB_Name = "Module3"


'This macro controls worksheet navigation behavior
Sub Macro3_Monitoreos()
'
' Macro2_Monitoreos Macro
'

'Sheets("Menu").Select
    Sheets("Monitoreos PUT").Visible = True
    Sheets("Monitoreos PUT").Select
    Range("A1").Select
End Sub

Sub Macro2_Monitoreosvolver()
'
' Macro2_Monitoreosvolver Macro
'

'Sheets("Monitoreos PUT").Select
    Sheets("Monitoreos PUT").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub

