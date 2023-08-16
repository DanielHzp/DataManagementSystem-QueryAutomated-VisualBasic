Attribute VB_Name = "Module5"

'This macro controls worksheet navigation behavior

Sub Macro5_InfoPUT()
'
' Macro5_InfoPUT Macro
'

'Sheets("Menu").Select
    Sheets("Info PUT").Visible = True
    Sheets("Info PUT").Select
    Range("A1").Select
End Sub

Sub Macro5_InfoPUTvolver()
'
' Macro5_InfoPUTvolver Macro
'

'Sheets("Info PUT").Select
    Sheets("Info PUT").Visible = False
    Sheets("Menu").Select
    Range("A1").Select
End Sub

Sub MostrarTodasFilas()


Rows.EntireRow.Hidden = False


End Sub
