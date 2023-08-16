VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Data Viewer Summary 2 V1.0"
   ClientHeight    =   8664.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13224
   OleObjectBlob   =   "DataViewerSummaryForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 1
Private Sub CommandButton3_Click()

If UserForm3.ComboBox3.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
If UserForm3.ComboBox3.ListIndex = 0 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!E59:I71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
If UserForm3.ComboBox3.ListIndex = 1 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!K59:O71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
If UserForm3.ComboBox3.ListIndex = 2 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!w59:AA71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
If UserForm3.ComboBox3.ListIndex = 3 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!Q59:U71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
'NANCY
If UserForm3.ComboBox3.ListIndex = 4 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!AO59:AS71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
'SURORIENTE
If UserForm3.ComboBox3.ListIndex = 5 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!AU59:AY71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
'TOROYACO
If UserForm3.ComboBox3.ListIndex = 6 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR PUT'!AC59:AG71"
TextBox3.Value = UserForm3.ComboBox3.Value
End If
End Sub


'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 2
Private Sub CommandButton4_Click()
If UserForm3.ComboBox4.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
'acordionero
If UserForm3.ComboBox4.ListIndex = 0 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!E59:I71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'chuira
If UserForm3.ComboBox4.ListIndex = 1 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!K59:O71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'COLON
If UserForm3.ComboBox4.ListIndex = 2 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!Q59:U71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'JUGLAR
If UserForm3.ComboBox4.ListIndex = 3 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!W59:AA71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'LOS ANGELES
If UserForm3.ComboBox4.ListIndex = 4 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!AC59:AG71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'MONOARAÑA
If UserForm3.ComboBox4.ListIndex = 5 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!AI59:AM71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'SAN ALBERTO
If UserForm3.ComboBox4.ListIndex = 6 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!AU59:AY71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If
'SANTA LUCIA
If UserForm3.ComboBox4.ListIndex = 7 Then
UserForm3.ListBox4.RowSource = "'COORDINADOR VMM'!AO59:AS71"
TextBox3.Value = UserForm3.ComboBox4.Value
End If

End Sub


'Dynamically inserts the data time frame from the imported datasets to the form list
Private Sub CommandButton5_Click()
'PUT
If UserForm3.ComboBox3.ListIndex = 0 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("j59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("j60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("j61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("j62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("j63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("j64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("j65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("j66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("j67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("j68").Value
End If
If UserForm3.ComboBox3.ListIndex = 1 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("p59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("p60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("p61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("p62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("p63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("p64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("p65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("p66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("p67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("p68").Value
End If
If UserForm3.ComboBox3.ListIndex = 2 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AB59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AB60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AB61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AB62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AB63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AB64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AB65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AB66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AB67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AB68").Value
End If
If UserForm3.ComboBox3.ListIndex = 3 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("V59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("V60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("V61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("V62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("V63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("V64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("V65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("V66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("V67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("V68").Value
End If
If UserForm3.ComboBox3.ListIndex = 4 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AT59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AT60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AT61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AT62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AT63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AT64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AT65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AT66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AT67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AT68").Value
End If
If UserForm3.ComboBox3.ListIndex = 5 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AZ59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AZ60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AZ61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AZ62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AZ63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AZ64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AZ65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AZ66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AZ67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AZ68").Value
End If
If UserForm3.ComboBox3.ListIndex = 6 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AH59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AH60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AH61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AH62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AH63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AH64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AH65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AH66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AH67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AH68").Value
End If

'vmm
If UserForm3.TextBox3.Value = "ACORDIONERO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("j59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("j60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("j61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("j62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("j63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("j64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("j65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("j66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("j67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("j68").Value


End If
If UserForm3.TextBox3.Value = "CHUIRA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("p59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("p60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("p61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("p62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("p63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("p64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("p65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("p66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("p67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("p68").Value
End If
If UserForm3.TextBox3.Value = "COLON" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("V59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("V60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("V61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("V62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("V63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("V64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("V65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("V66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("V67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("V68").Value
End If
If UserForm3.TextBox3.Value = "JUGLAR" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AB59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AB60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AB61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AB62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AB63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AB64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AB65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AB66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AB67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AB68").Value

End If
If UserForm3.TextBox3.Value = "LOS ANGELES" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AH59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AH60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AH61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AH62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AH63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AH64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AH65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AH66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AH67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AH68").Value
End If
If UserForm3.TextBox3.Value = "MONOARAÑA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AN59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AN60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AN61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AN62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AN63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AN64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AN65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AN66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AN67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AN68").Value

End If
If UserForm3.TextBox3.Value = "SAN ALBERTO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AZ59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AZ60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AZ61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AZ62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AZ63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AZ64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AZ65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AZ66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AZ67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AZ68").Value
End If
If UserForm3.TextBox3.Value = "SANTA LUCIA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AT59").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AT60").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AT61").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AT62").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AT63").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AT64").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AT65").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AT66").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AT67").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AT68").Value
End If



UserForm4.Show
End Sub
