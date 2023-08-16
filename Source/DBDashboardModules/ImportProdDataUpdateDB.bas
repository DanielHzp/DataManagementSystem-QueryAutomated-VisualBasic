Attribute VB_Name = "Module3"
Option Explicit





'Creates the database repository of production KPI reports
'Imports the dataset from .xlsm a external file loaded by the user
Sub ImportProduccion()


'Declare variables and aux
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Dim totalputN As Double, totalputS As Double, acrmch As Double, aym As Double, mar As Double, miniors As Double


'Save opened workbook reference
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here


'Open dialog box to load .xlsm file
FileName = Application.GetOpenFilename(Title:="Seleccione el reporte del último Q con datos de producción más recientes", MultiSelect:=True)
Application.ScreenUpdating = False

Workbooks.Open FileName(1)

'Clean dataset outputs to avoid overwritten data
LibroDestino.Sheets("COORDINADOR PUT").Range("BG5:BH16").Clear
LibroDestino.Sheets("COORDINADOR VMM").Range("BI5:BI16").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
   
   'Start data fetch from the imported .xlsm file (selected by the user)
    LibroDestino.Worksheets("COORDINADOR PUT").Range("BG5:BH16").Value = aWB.Sheets("BBL").Range("AI6:AJ17").Value
    LibroDestino.Worksheets("COORDINADOR VMM").Range("BI5:BI16").Value = aWB.Sheets("BBL").Range("AK6:AK17").Value
    aWB.Close savechanges:=False
    
    'Confirm that new data was added to the repository
    MsgBox ("Se actualizaron los datos de producción PUT y VMM")
    Application.ScreenUpdating = True
Exit Sub
here:
    MsgBox ("Se cerro la ventana")
    Application.ScreenUpdating = True
End Sub








