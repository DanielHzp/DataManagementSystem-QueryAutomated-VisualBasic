Attribute VB_Name = "Module12"




'This macro controls worksheet navigation behavior

Sub Macro12_BBL()
'
' Macro12_BBL Macro
'

'Sheets("Menu").Select
    Sheets("BBL").Visible = True
    Sheets("BBL").Select
    Range("A1").Select
End Sub

Sub Macro12_BBLvolver()
'
' Macro12_IncidentesGTEvolver Macro
'

'Sheets("BBL").Select
    Sheets("BBL").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
