Attribute VB_Name = "Module2"

'This macro controls worksheet navigation behavior
Sub Macro2_Presupuesto()
'
' Macro2_PRESUPUESTO Macro
'

'Sheets("Menu").Select
    Sheets("Presupuesto PUT").Visible = True
    Sheets("Presupuesto PUT").Select
    Range("A1").Select
End Sub

Sub Macro2_Presupuestovolver()
'
' Macro2_PRESUPUESTOvolver Macro
'

'Sheets("Presupuesto PUT").Select
    Sheets("Presupuesto PUT").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub



