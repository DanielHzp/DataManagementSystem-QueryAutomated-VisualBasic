Attribute VB_Name = "Module5"




'These macros control the worksheet automated behavior
'These commands are executed from worksheet click buttons

Sub reset()
ActiveSheet.UsedRange.SpecialCells (xlCellTypeLastCell)

End Sub

Sub desocultarCol()
ActiveSheet.Columns.EntireColumn.Hidden = False
ActiveSheet.Rows.EntireRow.Hidden = False

End Sub

Sub restablecerRango()
Range(Cells(1, 24), Cells(1, Columns.Count)).EntireColumn.Hidden = True


End Sub
Sub crearrangofilas()
 Range(Cells(200, 1), Cells(Rows.Count, 1)).EntireRow.Hidden = True
 
 End Sub

Sub ocultarTablaAnalisisResiduos()
Range(Cells(39, 1), Cells(87, 1)).EntireRow.Hidden = True
End Sub

Sub mostrarTablaAnalisisResiduos()
Range(Cells(39, 1), Cells(87, 1)).EntireRow.Hidden = False
End Sub

Sub mostrarTablaAnalisisConsumoAgua()
Range(Cells(42, 1), Cells(90, 1)).EntireRow.Hidden = False
End Sub
Sub ocultarTablaAnalisisConsumoAgua()
Range(Cells(42, 1), Cells(90, 1)).EntireRow.Hidden = True
End Sub
Sub ocultarceldascompensaciones()
Range(Cells(41, 1), Cells(Rows.Count, 1)).EntireRow.Hidden = True

End Sub
Sub ocultarcolumnascompensaciones()
Range(Cells(1, 22), Cells(1, Columns.Count)).EntireColumn.Hidden = True
End Sub

Sub mostrarTablaAguaResidual()
Range(Cells(42, 1), Cells(90, 1)).EntireRow.Hidden = False
End Sub
Sub ocultarTablaAguaResidual()
Range(Cells(42, 1), Cells(90, 1)).EntireRow.Hidden = True
End Sub
