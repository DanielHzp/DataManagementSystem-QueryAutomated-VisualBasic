VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Load Input Files Data to Main Database"
   ClientHeight    =   10272
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10080
   OleObjectBlob   =   "InputDataImportForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



'The following macros automate dataset imports when a user clicks any of the form buttons and selects a .xlsm file
'The import algorithm has the same structure for each data file input that will be added to the DB


'Input file 1 .xlsm
Private Sub CommandButton1_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook

Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)



On Error GoTo here

'Open window to import .xlsm file
FileName = Application.GetOpenFilename(Title:="Seleccione la ING .xlsm con los datos a cargar", MultiSelect:=True)

'Avoid screen flickering while loading data
Application.ScreenUpdating = False

'Open the workbook selected to begin the data fetch
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A7994:H8881").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A5186:H5761").Clear
LibroDestino.Sheets("BD_PUT").Range("A1:I564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    
    'Save the datasets needed from the imported file in the data repository target cells
    'The populated tables are later synchronized with an external relational database
    
    LibroDestino.Worksheets("BD_AGUA").Range("A7994:H8881").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A5186")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A5186:H5761").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J17666:P19873").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    
    LibroDestino.Worksheets("BD_PUT").Range("A1:I564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    
    'Confirm that new data has been added to the repository
    MsgBox ("Se actualizaron los datos de la ING COSTAYACO")
    
    'Save date and time of the update
    Range("Z10").Value = Date & " " & Time
    Application.ScreenUpdating = True
    
Exit Sub
here:
    MsgBox ("Se cerro la ventana")
    Application.ScreenUpdating = True
End Sub

'The import algorithm has the same structure for each data file input that will be added to the DB
'For each button this macro dynamically stores new data in new cells


'---------------------------------------------VMM CARGUE INGs-----------------------------------------------------------------
'ACORDIONERO
Private Sub CommandButton10_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING ACORDIONERO con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A2:H889").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A2:H577").Clear
LibroDestino.Sheets("BD_VMM").Range("A1:I564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A2")
    LibroDestino.Worksheets("BD_AGUA").Range("A2:H889").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A1:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A1")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A2:H577").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
     LibroDestino.Worksheets("BD_RESIDUOS").Range("J2:P2209").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("A1:I564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING ACORDIONERO")
    Range("Z1").Value = Date & " " & Time
     Application.ScreenUpdating = True
Exit Sub
here:
    MsgBox ("Se cerro la ventana")
    Application.ScreenUpdating = True
End Sub


'CHUIRA
Private Sub CommandButton11_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING CHUIRA con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A890:H1777").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A578:H1153").Clear
LibroDestino.Sheets("BD_VMM").Range("M1:U564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A890")
    LibroDestino.Worksheets("BD_AGUA").Range("A890:H1777").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A578")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A578:H1153").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J2210:P4417").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("M1:U564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING CHUIRA")
     Range("Z2").Value = Date & " " & Time
      Application.ScreenUpdating = True
Exit Sub
here:
    MsgBox ("Se cerro la ventana")
    Application.ScreenUpdating = True
End Sub


'COLON
Private Sub CommandButton12_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING COLON con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A1778:H2665").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A1154:H1729").Clear
LibroDestino.Sheets("BD_VMM").Range("Y1:AG564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A1778")
    LibroDestino.Worksheets("BD_AGUA").Range("A1778:H2665").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A1154")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A1154:H1729").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J4418:P6625").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("Y1:AG564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING COLON")
     Range("Z3").Value = Date & " " & Time
     Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'JUGLAR
Private Sub CommandButton13_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING JUGLAR con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A2666:H3553").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A1730:H2305").Clear
LibroDestino.Sheets("BD_VMM").Range("AK1:AS564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A2666")
    LibroDestino.Worksheets("BD_AGUA").Range("A2666:H3553").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A1730")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A1730:H2305").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J6626:P8833").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("AK1:AS564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING JUGLAR")
     Range("Z4").Value = Date & " " & Time
      Application.ScreenUpdating = True
Exit Sub
here:
 Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'LOS ANGELES
Private Sub CommandButton14_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING LOS ANGELES con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A3554:H4441").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A2306:H2881").Clear
LibroDestino.Sheets("BD_VMM").Range("AW1:BE564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A3554")
    LibroDestino.Worksheets("BD_AGUA").Range("A3554:H4441").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A2306")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A2306:H2881").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
     LibroDestino.Worksheets("BD_RESIDUOS").Range("J8834:P11041").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("AW1:BE564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING LOS ANGELES")
     Range("Z5").Value = Date & " " & Time
      Application.ScreenUpdating = True
Exit Sub
here:
 Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'MONOARAÑA
Private Sub CommandButton15_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING MONOARAÑA con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A4442:H5329").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A2882:H3457").Clear
LibroDestino.Sheets("BD_VMM").Range("BI1:BQ564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
   ' aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A4442")
   LibroDestino.Worksheets("BD_AGUA").Range("A4442:H5329").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A2882")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A2882:H3457").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J11042:P13249").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("BI1:BQ564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING MONOARAÑA")
    Range("Z6").Value = Date & " " & Time
     Application.ScreenUpdating = True
Exit Sub
here:
 Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub

'SANTA LUCIA
Private Sub CommandButton16_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING SANTA LUCIA con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A5330:H6217").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A3458:H4033").Clear
LibroDestino.Sheets("BD_VMM").Range("BU1:CC564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A5330")
    LibroDestino.Worksheets("BD_AGUA").Range("A5330:H6217").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A3458")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A3458:H4033").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J13250:P15457").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("BU1:CC564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING SANTA LUCIA")
    Range("Z7").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'SAN ALBERTO
Private Sub CommandButton17_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING SAN ALBERTO con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A6218:H7105").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A4034:H4609").Clear
LibroDestino.Sheets("BD_VMM").Range("CG1:CO576").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A6218")
    LibroDestino.Worksheets("BD_AGUA").Range("A6218:H7105").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A4034")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A4034:H4609").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J15458:P17665").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("CG1:CO576").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J576").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING SAN ALBERTO")
    Range("Z8").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'SISMICA VMM
Private Sub CommandButton18_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING SISMICA VMM con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A7106:H7993").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A4610:H5185").Clear
LibroDestino.Sheets("BD_VMM").Range("CS1:DA564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A7106")
    LibroDestino.Worksheets("BD_AGUA").Range("A7106:H7993").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A4610")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A4610:H5185").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_VMM").Range("CS1:DA564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING SAN ALBERTO")
    Range("Z9").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub
'cARGUE ing BOGOTA
Private Sub CommandButton19_Click()
MsgBox "ING Bogota consultar externamente"
End Sub


'The import algorithm has the same structure for each data file input that will be added to the DB
'For each button this macro dynamically stores new data in new cells

'--------------------------------------------------------------------------------------------------------------------------------------------------
'-------------------------------------------PUT CARGUE INGs - COSTAYACO AL PRINCIPIO CÓDIGO ------------------------------------------------------
'CUMPLIDOR
Private Sub CommandButton2_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING CUMPLIDOR con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A8883:H9770").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A5762:H6337").Clear
LibroDestino.Sheets("BD_PUT").Range("M1:U564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A8883")
    LibroDestino.Worksheets("BD_AGUA").Range("A8883:H9770").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A5762")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A5762:H6337").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J19874:P22081").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("M1:U564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING CUMPLIDOR")
    Range("Z11").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'Displays date and time info. of the latest database update
'It is executed when the user clicks the button 'View latest update'
Private Sub CommandButton20_Click()
TextBox10.Value = Range("Z1").Value
TextBox11.Value = Range("Z2").Value
TextBox12.Value = Range("Z3").Value
TextBox13.Value = Range("Z4").Value
TextBox14.Value = Range("Z5").Value
TextBox15.Value = Range("Z6").Value
TextBox16.Value = Range("Z7").Value
TextBox17.Value = Range("Z8").Value
TextBox18.Value = Range("Z9").Value
TextBox1.Value = Range("Z10").Value
TextBox2.Value = Range("Z11").Value
TextBox3.Value = Range("Z12").Value
TextBox4.Value = Range("Z13").Value
TextBox5.Value = Range("Z14").Value
TextBox6.Value = Range("Z15").Value
TextBox7.Value = Range("Z16").Value
TextBox8.Value = Range("Z17").Value
TextBox9.Value = Range("Z18").Value

End Sub

'MOQUETA
Private Sub CommandButton3_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING MOQUETA con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A9771:H10658").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A6338:H6913").Clear
LibroDestino.Sheets("BD_PUT").Range("Y1:AG564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A9771")
    LibroDestino.Worksheets("BD_AGUA").Range("A9771:H10658").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A6338")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A6338:H6913").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J22082:P24289").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("Y1:AG564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING MOQUETA")
    Range("Z12").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub


'MARY
Private Sub CommandButton4_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING MARY con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A10659:H11546").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A6914:H7489").Clear
LibroDestino.Sheets("BD_PUT").Range("AK1:AS564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("A9771")
    LibroDestino.Worksheets("BD_AGUA").Range("A10659:H11546").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A6338")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A6914:H7489").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
     LibroDestino.Worksheets("BD_RESIDUOS").Range("J24290:P26497").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("AK1:AS564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING MARY")
    Range("Z13").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub



'TOROYACO
Private Sub CommandButton5_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING TOROYACO con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A11547:H12434").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A7490:H8065").Clear
LibroDestino.Sheets("BD_PUT").Range("AW1:BE564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("")
    LibroDestino.Worksheets("BD_AGUA").Range("A11547:H12434").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A...")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A7490:H8065").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J26498:P28705").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("AW1:BE564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING TOROYACO")
    Range("Z14").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub



'MIRAFLOR
Private Sub CommandButton6_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING MIRAFLOR con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A12435:H13322").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A8066:H8641").Clear
LibroDestino.Sheets("BD_PUT").Range("BI1:BQ564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("")
    LibroDestino.Worksheets("BD_AGUA").Range("A12435:H13322").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A...")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A8066:H8641").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("BI1:BQ564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING MIRAFLOR")
     Range("Z15").Value = Date & " " & Time
     Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub



'NANCY
Private Sub CommandButton7_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING NANCY con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A13323:H14210").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A8642:H9217").Clear
LibroDestino.Sheets("BD_PUT").Range("BU1:CC564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("")
    LibroDestino.Worksheets("BD_AGUA").Range("A13323:H14210").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A...")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A8642:H9217").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    LibroDestino.Worksheets("BD_RESIDUOS").Range("J28706:P30913").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("BU1:CC564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING NANCY")
    Range("Z16").Value = Date & " " & Time
     Application.ScreenUpdating = True
Exit Sub
here:
     Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub



'SURORIENTE
Private Sub CommandButton8_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING SURORIENTE con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A14211:H15098").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A9218:H9793").Clear
LibroDestino.Sheets("BD_PUT").Range("CG1:CO564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("")
    LibroDestino.Worksheets("BD_AGUA").Range("A14211:H15098").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A...")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A9218:H9793").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
     LibroDestino.Worksheets("BD_RESIDUOS").Range("J30914:P33121").Value = aWB.Sheets("BD VERTIMIENTOS").Range("Q1:W2208").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("CG1:CO564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING SURORIENTE")
    Range("Z17").Value = Date & " " & Time
         Application.ScreenUpdating = True
Exit Sub
here:
     Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub



'SISMICA PUT
Private Sub CommandButton9_Click()
Dim FileName As Variant, aWB As Workbook, tWB As Workbook, gWB As Workbook, LibroDestino As Workbook
Set tWB = ThisWorkbook
Set LibroDestino = Workbooks(tWB.Name)
'LibroDestino.Sheets("BD_AGUA")
On Error GoTo here
FileName = Application.GetOpenFilename(Title:="Seleccione la ING SISMICA PUT con los datos a cargar", MultiSelect:=True)
Application.ScreenUpdating = False
Workbooks.Open FileName(1)
LibroDestino.Sheets("BD_AGUA").Range("A15099:H15986").Clear
LibroDestino.Sheets("BD_VERTIMIENTOS").Range("A9794:H10369").Clear
LibroDestino.Sheets("BD_PUT").Range("CS1:DA564").Clear
    Set aWB = ActiveWorkbook
    aWB.Activate
    'aWB.Sheets("BD AGUA").Range("A2:H889").Copy Destination:=LibroDestino.Worksheets("BD_AGUA").Range("")
    LibroDestino.Worksheets("BD_AGUA").Range("A15099:H15986").Value = aWB.Sheets("BD AGUA").Range("A2:H889").Value
    'aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Copy Destination:=LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A...")
    LibroDestino.Worksheets("BD_VERTIMIENTOS").Range("A9794:H10369").Value = aWB.Sheets("BD VERTIMIENTOS").Range("A2:H577").Value
    'aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    LibroDestino.Worksheets("BD_PUT").Range("CS1:DA564").Value = aWB.Sheets("BD COORDINADOR").Range("B1:J564").Value
    aWB.Close savechanges:=False
    MsgBox ("Se actualizaron los datos de la ING SISMICA PUT")
    Range("Z18").Value = Date & " " & Time
    Application.ScreenUpdating = True
Exit Sub
here:
    Application.ScreenUpdating = True
    MsgBox ("Se cerro la ventana")
End Sub

'--------------------------------------------------------------------------------------------------------------------------------------------
