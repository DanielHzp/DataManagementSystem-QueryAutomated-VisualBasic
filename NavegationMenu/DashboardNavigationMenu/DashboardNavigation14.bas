Attribute VB_Name = "Module14"



'This macro controls worksheet navigation behavior

Sub Macro14_ICAS()
'
' Macro14_ICAS Macro
'

'Sheets("Menu").Select
    Sheets("ICAS").Visible = True
    Sheets("ICAS").Select
    Range("A1").Select
End Sub


Sub Macro14_ICASvolver()
'
' Macro14_ICASvolver Macro
'

'Sheets("ICAS").Select
    Sheets("ICAS").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub
