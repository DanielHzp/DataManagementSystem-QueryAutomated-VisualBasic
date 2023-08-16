Attribute VB_Name = "Module6"


'This macro controls worksheet navigation behavior

Sub Macro6_InfoVMM()
'
' Macro6_InfoVMM Macro
'

'Sheets("Menu").Select
    Sheets("Info VMM").Visible = True
    Sheets("Info VMM").Select
    Range("A1").Select
End Sub

Sub Macro6_InfoVMMvolver()
'
' Macro6_InfoVMMvolver Macro
'

'Sheets("Info VMM").Select
    Sheets("Info VMM").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
