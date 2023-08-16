Attribute VB_Name = "Module4"
Option Explicit




'The following code populates the waste/operational output datasets using the monthly data inputs of the user
Sub creartablaresiduos()

Application.ScreenUpdating = False


'try behavior
On Error GoTo here


Dim ultimacelda As Integer, a As Integer, c As Integer
c = 1

'Loop over the worksheet wastes table and copy the values
'Then transpose the table and paste it in a new dataset that will be manipulated by the database repository
'This is part of the data cleaning process required for consistency
For a = 1531 To 1608 Step 7
Sheets("RESIDUOS").Range("I" & a, "BB" & a).Copy

Sheets("BD VERTIMIENTOS").Range("W" & c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True
c = c + 46

Next a

'Recursively call different parts of the algorithm
Call ActualizarResiduosWo
Call ActualizarResiduosOC
Call ActualizarResiduosPer

Exit Sub
Application.ScreenUpdating = True

'Catch behavior
here:
    MsgBox "Runtime error, make sure the tables have the correct format"

End Sub





'The following macros repeat the same algorithm, which transforms table inputs into dataset outputs ready to be manipulated by the database repository

Sub ActualizarResiduosWo()
Dim a As Integer, c As Integer
c = 553
For a = 1531 To 1608 Step 7
Sheets("RESIDUOS_WORKOVER").Range("I" & a, "BB" & a).Copy
Sheets("BD VERTIMIENTOS").Range("W" & c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True
c = c + 46
Next a
End Sub

Sub ActualizarResiduosOC()
Dim a As Integer, c As Integer
c = 1105
For a = 1531 To 1608 Step 7
Sheets("RESIDUOS_OBRA_CIVIL").Range("I" & a, "BB" & a).Copy
Sheets("BD VERTIMIENTOS").Range("W" & c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True
c = c + 46
Next a
End Sub

Sub ActualizarResiduosPer()
Dim a As Integer, c As Integer
c = 1657
For a = 1531 To 1608 Step 7
Sheets("RESIDUOS_PERFORACION").Range("I" & a, "BB" & a).Copy
Sheets("BD VERTIMIENTOS").Range("W" & c).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, Transpose:=True
c = c + 46
Next a

End Sub
