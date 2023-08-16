VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Data Viewer Summary 3 V1.0"
   ClientHeight    =   8484.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13164
   OleObjectBlob   =   "DataViewerSummaryForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 1
Private Sub CommandButton1_Click()
If UserForm5.ComboBox5.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
If UserForm5.ComboBox5.ListIndex = 0 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!E85:I97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
If UserForm5.ComboBox5.ListIndex = 1 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!K85:O97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
'MARY
If UserForm5.ComboBox5.ListIndex = 2 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!w85:AA97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
'MOQUETA
If UserForm5.ComboBox5.ListIndex = 3 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!Q85:U97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
'NANCY
If UserForm5.ComboBox5.ListIndex = 4 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!AO85:AS97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
'SURORIENTE
If UserForm5.ComboBox5.ListIndex = 5 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!AU85:AY97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
'TOROYACO
If UserForm5.ComboBox5.ListIndex = 6 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR PUT'!AC85:AG97"
TextBox4.Value = UserForm5.ComboBox5.Value
End If
End Sub



'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 2
'Show INGs data VMM in list form
Private Sub CommandButton2_Click()
If UserForm5.ComboBox6.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
'acordionero
If UserForm5.ComboBox6.ListIndex = 0 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!E85:I97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'chuira
If UserForm5.ComboBox6.ListIndex = 1 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!K85:O97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'COLON
If UserForm5.ComboBox6.ListIndex = 2 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!Q85:U97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'JUGLAR
If UserForm5.ComboBox6.ListIndex = 3 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!W85:AA97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'LOS ANGELES
If UserForm5.ComboBox6.ListIndex = 4 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!AC85:AG97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'MONOARAÑA
If UserForm5.ComboBox6.ListIndex = 5 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!AI85:AM97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'SAN ALBERTO
If UserForm5.ComboBox6.ListIndex = 6 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!AU85:AY97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
'SANTA LUCIA
If UserForm5.ComboBox6.ListIndex = 7 Then
UserForm5.ListBox6.RowSource = "'COORDINADOR VMM'!AO85:AS97"
TextBox4.Value = UserForm5.ComboBox6.Value
End If
End Sub



'DYNAMICALLY INSERTS REPORT COMMENTS ON THE 'REPORTS SUMMARY FORM'
Private Sub CommandButton3_Click()
'PUT
If UserForm5.ComboBox5.ListIndex = 0 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("j85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("j86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("j87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("j88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("j89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("j90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("j91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("j92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("j93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("j94").Value

End If
If UserForm5.ComboBox5.ListIndex = 1 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("p85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("p86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("p87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("p88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("p89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("p90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("p91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("p92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("p93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("p94").Value
End If
If UserForm5.ComboBox5.ListIndex = 2 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AB85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AB86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AB87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AB88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AB89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AB90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AB91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AB92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AB93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AB94").Value
End If
If UserForm5.ComboBox5.ListIndex = 3 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("V85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("V86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("V87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("V88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("V89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("V90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("V91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("V92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("V93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("V94").Value
End If
If UserForm5.ComboBox5.ListIndex = 4 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AT85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AT86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AT87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AT88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AT89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AT90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AT91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AT92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AT93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AT94").Value
End If
If UserForm5.ComboBox5.ListIndex = 5 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AZ85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AZ86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AZ87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AZ88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AZ89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AZ90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AZ91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AZ92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AZ93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AZ94").Value
End If
If UserForm5.ComboBox5.ListIndex = 6 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AH85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AH86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AH87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AH88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AH89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AH90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AH91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AH92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AH93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AH94").Value
End If

'vmm
If UserForm5.TextBox4.Value = "ACORDIONERO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("j85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("j86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("j87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("j88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("j89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("j90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("j91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("j92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("j93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("j94").Value

End If
If UserForm5.TextBox4.Value = "CHUIRA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("p85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("p86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("p87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("p88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("p89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("p90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("p91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("p92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("p93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("p94").Value
End If
If UserForm5.TextBox4.Value = "COLON" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("V85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("V86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("V87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("V88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("v89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("v90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("v91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("v92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("v93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("v94").Value
End If
If UserForm5.TextBox4.Value = "JUGLAR" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AB85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AB86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AB87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AB88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AB89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AB90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AB91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AB92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AB93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AB94").Value
End If
If UserForm5.TextBox4.Value = "LOS ANGELES" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AH85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AH86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AH87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AH88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AH89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AH90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AH91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AH92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AH93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AH94").Value
End If
If UserForm5.TextBox4.Value = "MONOARAÑA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AN85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AN86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AN87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AN88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AN89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AN90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AN91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AN92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AN93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AN94").Value
End If
If UserForm5.TextBox4.Value = "SAN ALBERTO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AZ85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AZ86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AZ87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AZ88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AZ89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AZ90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AZ91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AZ92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AZ93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AZ94").Value
End If
If UserForm5.TextBox4.Value = "SANTA LUCIA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AT85").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AT86").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AT87").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AT88").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AT89").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AT90").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AT91").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AT92").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AT93").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AT94").Value
End If


'OPEN FORM WHEN THE BUTTON 'VIEW DATA DETAILS' IS CLICKED
UserForm4.Show

End Sub

Private Sub Label2_Click()

End Sub
