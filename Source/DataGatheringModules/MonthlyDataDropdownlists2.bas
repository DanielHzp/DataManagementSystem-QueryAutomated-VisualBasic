Attribute VB_Name = "Módulo4"



'This macro controls dropdown lists worksheet behavior
Sub MES_AGUA_MEDIDOR()


If Range("B3").Value = "ENERO" Then

Range("E:X").EntireColumn.Hidden = False
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True


End If
If Range("B3").Value = "FEBRERO" Then

Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = False
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True


End If
If Range("B3").Value = "MARZO" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = False
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "ABRIL" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = False
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "MAYO" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = False
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "JUNIO" Then

Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = False
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "JULIO" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = False
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "AGOSTO" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = False
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "SEPTIEMBRE" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = False
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "OCTUBRE" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = False
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "NOVIEMBRE" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = False
Range("HQ:IJ").EntireColumn.Hidden = True

End If
If Range("B3").Value = "DICIEMBRE" Then
Range("E:X").EntireColumn.Hidden = True
Range("Y:AR").EntireColumn.Hidden = True
Range("AS:BL").EntireColumn.Hidden = True
Range("BM:CF").EntireColumn.Hidden = True
Range("CG:CZ").EntireColumn.Hidden = True
Range("DA:DT").EntireColumn.Hidden = True
Range("DU:EN").EntireColumn.Hidden = True
Range("EO:FH").EntireColumn.Hidden = True
Range("FI:GB").EntireColumn.Hidden = True
Range("GC:GV").EntireColumn.Hidden = True
Range("GW:HP").EntireColumn.Hidden = True
Range("HQ:IJ").EntireColumn.Hidden = False

End If


End Sub

