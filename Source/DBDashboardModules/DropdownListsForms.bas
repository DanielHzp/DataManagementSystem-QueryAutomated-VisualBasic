Attribute VB_Name = "Module4"
Option Explicit



'This macros populate the dropdown lists of the data summary forms


Sub IniciarResumenConsumoAgua()
'llenar PUT
UserForm2.ComboBox1.AddItem "COSTAYACO"
UserForm2.ComboBox1.AddItem "CUMPLIDOR"
UserForm2.ComboBox1.AddItem "MARY"
UserForm2.ComboBox1.AddItem "MOQUETA"
UserForm2.ComboBox1.AddItem "NANCY"
UserForm2.ComboBox1.AddItem "SURORIENTE"
UserForm2.ComboBox1.AddItem "TOROYACO"
'llenar VMM
UserForm2.ComboBox2.AddItem "ACORDIONERO"
UserForm2.ComboBox2.AddItem "CHUIRA"
UserForm2.ComboBox2.AddItem "COLON"
UserForm2.ComboBox2.AddItem "JUGLAR"
UserForm2.ComboBox2.AddItem "LOS ANGELES"
UserForm2.ComboBox2.AddItem "MONOARAÑA"
UserForm2.ComboBox2.AddItem "SAN ALBERTO"
UserForm2.ComboBox2.AddItem "SANTA LUCIA"
UserForm2.ListBox2.RowSource = "'COORDINADOR PUT'!D33:D45"
UserForm2.Show

End Sub

Sub IniciarResumenAguaResidual()
'llenar PUT
UserForm3.ComboBox3.AddItem "COSTAYACO"
UserForm3.ComboBox3.AddItem "CUMPLIDOR"
UserForm3.ComboBox3.AddItem "MARY"
UserForm3.ComboBox3.AddItem "MOQUETA"
UserForm3.ComboBox3.AddItem "NANCY"
UserForm3.ComboBox3.AddItem "SURORIENTE"
UserForm3.ComboBox3.AddItem "TOROYACO"
'llenar VMM
UserForm3.ComboBox4.AddItem "ACORDIONERO"
UserForm3.ComboBox4.AddItem "CHUIRA"
UserForm3.ComboBox4.AddItem "COLON"
UserForm3.ComboBox4.AddItem "JUGLAR"
UserForm3.ComboBox4.AddItem "LOS ANGELES"
UserForm3.ComboBox4.AddItem "MONOARAÑA"
UserForm3.ComboBox4.AddItem "SAN ALBERTO"
UserForm3.ComboBox4.AddItem "SANTA LUCIA"
UserForm3.ListBox3.RowSource = "'COORDINADOR PUT'!D33:D45"
UserForm3.Show

End Sub

Sub IniciarResumenResiduos()
'llenar PUT
UserForm5.ComboBox5.AddItem "COSTAYACO"
UserForm5.ComboBox5.AddItem "CUMPLIDOR"
UserForm5.ComboBox5.AddItem "MARY"
UserForm5.ComboBox5.AddItem "MOQUETA"
UserForm5.ComboBox5.AddItem "NANCY"
UserForm5.ComboBox5.AddItem "SURORIENTE"
UserForm5.ComboBox5.AddItem "TOROYACO"
'llenar VMM
UserForm5.ComboBox6.AddItem "ACORDIONERO"
UserForm5.ComboBox6.AddItem "CHUIRA"
UserForm5.ComboBox6.AddItem "COLON"
UserForm5.ComboBox6.AddItem "JUGLAR"
UserForm5.ComboBox6.AddItem "LOS ANGELES"
UserForm5.ComboBox6.AddItem "MONOARAÑA"
UserForm5.ComboBox6.AddItem "SAN ALBERTO"
UserForm5.ComboBox6.AddItem "SANTA LUCIA"
UserForm5.ListBox5.RowSource = "'COORDINADOR PUT'!D33:D45"
UserForm5.Show




End Sub




















'Sub desocultarfilas()
'Sheets("REPORTE").Range("A221:A268").Rows.EntireRow.Hidden = False
'End Sub
Sub desocultarcols()


ActiveSheet.Columns.EntireColumn.Hidden = False

End Sub
