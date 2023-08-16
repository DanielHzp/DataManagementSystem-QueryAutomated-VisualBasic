Attribute VB_Name = "Module2"




'This macro controls dropdownlists display options and on click behavior
'Specifically when the user navigates between different monthly data sets


Sub MES_AGUA_RESBLOQUE()


If Range("B2").Value = "ENERO" Then

Range("D:S").EntireColumn.Hidden = False
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True


End If
If Range("B2").Value = "FEBRERO" Then

Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = False
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True



End If
If Range("B2").Value = "MARZO" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = False
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True

End If
If Range("B2").Value = "ABRIL" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = False
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True

End If
If Range("B2").Value = "MAYO" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = False
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True
End If
If Range("B2").Value = "JUNIO" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = False
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True

End If
If Range("B2").Value = "JULIO" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = False
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True
End If
If Range("B2").Value = "AGOSTO" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = False
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True
End If
If Range("B2").Value = "SEPTIEMBRE" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = False
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True
End If
If Range("B2").Value = "OCTUBRE" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = False
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = True

End If
If Range("B2").Value = "NOVIEMBRE" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = False
Range("FX:GM").EntireColumn.Hidden = True

End If
If Range("B2").Value = "DICIEMBRE" Then
Range("D:S").EntireColumn.Hidden = True
Range("T:AI").EntireColumn.Hidden = True
Range("AJ:AY").EntireColumn.Hidden = True
Range("AZ:BO").EntireColumn.Hidden = True
Range("BP:CE").EntireColumn.Hidden = True
Range("CF:CU").EntireColumn.Hidden = True
Range("CV:DK").EntireColumn.Hidden = True
Range("DL:EA").EntireColumn.Hidden = True
Range("EB:EQ").EntireColumn.Hidden = True
Range("ER:FG").EntireColumn.Hidden = True
Range("FH:FW").EntireColumn.Hidden = True
Range("FX:GM").EntireColumn.Hidden = False
End If


End Sub
Sub MES_AGUA_BLOQUE()


If Range("B2").Value = "ENERO" Then

Range("D:Y").EntireColumn.Hidden = False
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True


End If
If Range("B2").Value = "FEBRERO" Then

Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = False
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True


End If
If Range("B2").Value = "MARZO" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = False
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "ABRIL" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = False
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "MAYO" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = False
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "JUNIO" Then

Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = False
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "JULIO" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = False
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "AGOSTO" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = False
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "SEPTIEMBRE" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = False
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "OCTUBRE" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = False
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "NOVIEMBRE" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = False
Range("IL:JG").EntireColumn.Hidden = True

End If
If Range("B2").Value = "DICIEMBRE" Then
Range("D:Y").EntireColumn.Hidden = True
Range("Z:AU").EntireColumn.Hidden = True
Range("AV:BQ").EntireColumn.Hidden = True
Range("BR:CM").EntireColumn.Hidden = True
Range("CN:DI").EntireColumn.Hidden = True
Range("DJ:EE").EntireColumn.Hidden = True
Range("EF:FA").EntireColumn.Hidden = True
Range("FB:FW").EntireColumn.Hidden = True
Range("FX:GS").EntireColumn.Hidden = True
Range("GT:HO").EntireColumn.Hidden = True
Range("HP:IK").EntireColumn.Hidden = True
Range("IL:JG").EntireColumn.Hidden = False

End If


End Sub


Sub MES_VERTIMIENTOS()


If Range("B2").Value = "ENERO" Then

Range("F:V").EntireColumn.Hidden = False
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True


End If
If Range("B2").Value = "FEBRERO" Then

Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = False
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True


End If
If Range("B2").Value = "MARZO" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = False
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "ABRIL" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = False
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "MAYO" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = False
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "JUNIO" Then

Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = False
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "JULIO" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = False
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "AGOSTO" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = False
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "SEPTIEMBRE" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = False
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "OCTUBRE" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = False
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "NOVIEMBRE" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = False
Range("GK:HA").EntireColumn.Hidden = True

End If
If Range("B2").Value = "DICIEMBRE" Then
Range("F:V").EntireColumn.Hidden = True
Range("W:AM").EntireColumn.Hidden = True
Range("AN:BD").EntireColumn.Hidden = True
Range("BE:BU").EntireColumn.Hidden = True
Range("BV:CL").EntireColumn.Hidden = True
Range("CM:DC").EntireColumn.Hidden = True
Range("DD:DT").EntireColumn.Hidden = True
Range("DU:EK").EntireColumn.Hidden = True
Range("EL:FB").EntireColumn.Hidden = True
Range("FC:FS").EntireColumn.Hidden = True
Range("FT:GJ").EntireColumn.Hidden = True
Range("GK:HA").EntireColumn.Hidden = False

End If


End Sub




