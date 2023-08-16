VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "DB Data Summary 1"
   ClientHeight    =   8472.001
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13332
   OleObjectBlob   =   "DBDataSummaryForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 1

'CONSUMO AGUA SAMPLE DATASET
Private Sub CommandButton1_Click()
If UserForm2.ComboBox1.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
If UserForm2.ComboBox1.ListIndex = 0 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!E33:I45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
If UserForm2.ComboBox1.ListIndex = 1 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!K33:O45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
If UserForm2.ComboBox1.ListIndex = 2 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!w33:AA45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
If UserForm2.ComboBox1.ListIndex = 3 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!Q33:U45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
'NANCY
If UserForm2.ComboBox1.ListIndex = 4 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!AO33:AS45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
'SURORIENTE
If UserForm2.ComboBox1.ListIndex = 5 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!AU33:AY45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If
'TOROYACO
If UserForm2.ComboBox1.ListIndex = 6 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR PUT'!AC33:AG45"
TextBox1.Value = UserForm2.ComboBox1.Value
End If


End Sub


'Dynamically inserts data in the collection of the form depending on the comboBox selected value - repository 2

'SAMPLE DATASET

Private Sub CommandButton2_Click()
If UserForm2.ComboBox2.ListIndex = -1 Then
MsgBox "Seleccione una ING válida"
End If
'acordionero
If UserForm2.ComboBox2.ListIndex = 0 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!E33:I45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'chuira
If UserForm2.ComboBox2.ListIndex = 1 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!K33:O45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'COLON
If UserForm2.ComboBox2.ListIndex = 2 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!Q33:U45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'JUGLAR
If UserForm2.ComboBox2.ListIndex = 3 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!W33:AA45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'LOS ANGELES
If UserForm2.ComboBox2.ListIndex = 4 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!AC33:AG45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'MONOARAÑA
If UserForm2.ComboBox2.ListIndex = 5 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!AI33:AM45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'SAN ALBERTO
If UserForm2.ComboBox2.ListIndex = 6 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!AU33:AY45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If
'SANTA LUCIA
If UserForm2.ComboBox2.ListIndex = 7 Then
UserForm2.ListBox1.RowSource = "'COORDINADOR VMM'!AO33:AS45"
TextBox1.Value = UserForm2.ComboBox2.Value
End If

End Sub




'Dynamically display the datasets name in a caption text box
'The selected value of the combobox must be displayed in this field
Private Sub CommandButton3_Click()
'PUT
If UserForm2.ComboBox1.ListIndex = 0 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("j33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("j34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("j35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("j36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("j37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("j38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("j39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("j40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("j41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("j42").Value
End If
If UserForm2.ComboBox1.ListIndex = 1 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("p33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("p34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("p35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("p36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("p37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("p38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("p39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("p40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("p41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("p42").Value

End If
If UserForm2.ComboBox1.ListIndex = 2 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AB33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AB34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AB35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AB36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AB37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AB38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AB39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AB40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AB41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AB42").Value
End If
If UserForm2.ComboBox1.ListIndex = 3 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("V33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("V34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("V35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("V36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("V37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("V38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("V39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("V40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("V41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("V42").Value

End If
If UserForm2.ComboBox1.ListIndex = 4 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AT33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AT34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AT35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AT36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AT37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AT38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AT39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AT40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AT41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AT42").Value

End If
If UserForm2.ComboBox1.ListIndex = 5 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AZ33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AZ34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AZ35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AZ36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AZ37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AZ38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AZ39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AZ40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AZ41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AZ42").Value

End If
If UserForm2.ComboBox1.ListIndex = 6 Then
UserForm4.Label1.Caption = Sheets("COORDINADOR PUT").Range("AH33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR PUT").Range("AH34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR PUT").Range("AH35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR PUT").Range("AH36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR PUT").Range("AH37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR PUT").Range("AH38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR PUT").Range("AH39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR PUT").Range("AH40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR PUT").Range("AH41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR PUT").Range("AH42").Value
End If

'vmm
If UserForm2.TextBox1.Value = "ACORDIONERO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("j33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("j34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("j35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("j36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("j37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("j38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("j39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("j40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("j41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("j42").Value

End If
If UserForm2.TextBox1.Value = "CHUIRA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("p33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("p34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("p35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("p36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("p37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("p38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("p39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("p40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("p41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("p42").Value

End If
If UserForm2.TextBox1.Value = "COLON" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("V33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("V34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("V35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("V36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("V37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("V38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("V39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("V40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("V41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("V42").Value
End If
If UserForm2.TextBox1.Value = "JUGLAR" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AB33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AB34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AB35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AB36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AB37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AB38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AB39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AB40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AB41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AB42").Value

End If
If UserForm2.TextBox1.Value = "LOS ANGELES" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AH33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AH34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AH35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AH36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AH37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AH38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AH39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AH40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AH41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AH42").Value

End If
If UserForm2.TextBox1.Value = "MONOARAÑA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AN33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AN34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AN35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AN36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AN37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AN38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AN39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AN40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AN41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AN42").Value

End If
If UserForm2.TextBox1.Value = "SAN ALBERTO" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AZ33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AZ34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AZ35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AZ36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AZ37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AZ38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AZ39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AZ40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AZ41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AZ42").Value

End If
If UserForm2.TextBox1.Value = "SANTA LUCIA" Then
UserForm4.Label1.Caption = Sheets("COORDINADOR VMM").Range("AT33").Value
UserForm4.Label2.Caption = Sheets("COORDINADOR VMM").Range("AT34").Value
UserForm4.Label3.Caption = Sheets("COORDINADOR VMM").Range("AT35").Value
UserForm4.Label4.Caption = Sheets("COORDINADOR VMM").Range("AT36").Value
UserForm4.Label6.Caption = Sheets("COORDINADOR VMM").Range("AT37").Value
UserForm4.Label7.Caption = Sheets("COORDINADOR VMM").Range("AT38").Value
UserForm4.Label8.Caption = Sheets("COORDINADOR VMM").Range("AT39").Value
UserForm4.Label9.Caption = Sheets("COORDINADOR VMM").Range("AT40").Value
UserForm4.Label10.Caption = Sheets("COORDINADOR VMM").Range("AT41").Value
UserForm4.Label11.Caption = Sheets("COORDINADOR VMM").Range("AT42").Value


End If





'Load 'View Report Details' userform

UserForm4.Show

End Sub

Private Sub Label5_Click()

End Sub
