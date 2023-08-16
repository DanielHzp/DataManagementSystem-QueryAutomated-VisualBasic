Attribute VB_Name = "Module9"


'This macro controls worksheet navigation behavior

Sub Macro9_MonitoreosVMM()
'
' Macro9_MonitoreosVMM Macro
'

'Sheets("Menu").Select
    Sheets("Monitoreos VMM").Visible = True
    Sheets("Monitoreos VMM").Select
    Range("A1").Select
End Sub

Sub Macro9_MonitoreosVMMvolver()
'
' Macro9_MonitoreosVMMvolver Macro
'

'Sheets("Monitoreos VMM").Select
    Sheets("Monitoreos VMM").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
