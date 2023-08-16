Attribute VB_Name = "Module2"
Option Explicit



'Loads the form panel that allows the user to load the data from the input files
'Executes the user form that creates and updates the database repository
Sub IniciarPanelActualizacion()
UserForm1.Show
End Sub

Sub ImportarDatosBD()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")

FileName = Application.GetOpenFilename(Title:="Seleccione la ING con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False

'BD AGUA
'BD VERTIMIENTOS
'BD COORDINADOR
End Sub

