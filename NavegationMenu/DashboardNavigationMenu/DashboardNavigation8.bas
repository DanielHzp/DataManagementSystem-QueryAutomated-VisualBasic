Attribute VB_Name = "Module8"



'This macro controls worksheet navigation behavior

Sub Macro8_PresupuestoVMM()
'
' Macro8_PRESUPUESTOVMM Macro
'

'Sheets("Menu").Select
    Sheets("Presupuesto VMM").Visible = True
    Sheets("Presupuesto VMM").Select
    Range("A1").Select
End Sub

Sub Macro8_PresupuestoVMMvolver()
'
' Macro8_PRESUPUESTOVMMvolver Macro
'

'Sheets("Presupuesto VMM").Select
    Sheets("Presupuesto VMM").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub

