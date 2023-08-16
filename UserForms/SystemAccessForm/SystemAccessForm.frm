VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   3456
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   6108
   OleObjectBlob   =   "SystemAccessForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
End Sub




'THE FOLLOWING MACRO HANDLES USER PROFILE INPUTS AND CONTROLS USERFORM BEHAVIOR BASED ON CREDENTIALS INFO.
Private Sub CommandButton2_Click()
Sheets("USUARIOS").Visible = True
Sheets("USUARIOS").Select
Dim usuario As String
Dim password As Variant
Dim DatoEncontrado
Blog = "SAMPLE DATA REPOSITORY NAME"

UsuarioExistente = Application.WorksheetFunction.CountIf(Sheets("USUARIOS").Range("D2:D50"), _
    Me.txtUsuario.Value)
Set Rango = Sheets("USUARIOS").Range("D2:D50")
If Me.txtUsuario.Value = "" Or Me.txtPassword.Value = "" Then
    
    Sheets("USUARIOS").Visible = False
    Sheets("INICIO").Visible = True
    Sheets("INICIO").Select
    
    
    MsgBox "Por favor introduce usuario y contraseña", vbExclamation, Blog
    Me.txtUsuario.SetFocus
    
    


ElseIf UsuarioExistente = 0 Then
    
    Sheets("USUARIOS").Visible = False
    Sheets("INICIO").Visible = True
    Sheets("INICIO").Select
    
    
    MsgBox "El usuario '" & Me.txtUsuario & "' no existe", vbExclamation, Blog
    
        
ElseIf UsuarioExistente = 1 Then

    'SEARCH INPUT CREDENTIALS IN WORKSHEET DATASETS
    DatoEncontrado = Rango.Find(What:=Me.txtUsuario.Value, MatchCase:=True).Address
    Contrasenia = Range(DatoEncontrado).Offset(0, 1).Value
    
    
    If Range(DatoEncontrado).Value = Me.txtUsuario.Value And Contrasenia = _
    Me.txtPassword.Value Then
        Range("H1").Value = Range(DatoEncontrado).Offset(0, -1).Value

        
        'Below is the code that loads and controls the workbook behavior based on user profile input
        Unload Me
        
        
        
        'Verify user profile METADATA
        'INVOKE NECCESARY METHODS/SUBROUTINES CODED IN SOURCE MODULES
        If Range("H1").Value = "LIDER" Then
            Call LIDER
            Call REPORTE
        End If
        If Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 1" Or Range("H1").Value = "COORDINADOR PUTUMAYO NORTE 2" Or Range("H1").Value = "COORDINADOR PUTUMAYO SUR 1" Or Range("H1").Value = "COORDINADOR PUTUMAYO SUR 2" Or Range("H1").Value = "COORDINADOR VALLE MM" Or Range("H1").Value = "COORDINADOR LLANOS" Or Range("H1").Value = "COORDINADOR BOGOTA" Then
            Call COORDINADOR
            Call REPORTE
        End If
        
        If Range("H1").Value = "ING_1" Or Range("H1").Value = "ING_2" Or Range("H1").Value = "ING_3" Or Range("H1").Value = "ING_4" Or Range("H1").Value = "ING_5" Or Range("H1").Value = "ING_6" Or Range("H1").Value = "ING_7" Or Range("H1").Value = "ING_8" Or Range("H1").Value = "ING_9" Or Range("H1").Value = "ING_10" Or Range("H1").Value = "ING_11" Or Range("H1").Value = "ING_12" Or Range("H1").Value = "ING_13" Or Range("H1").Value = "ING_14" Or Range("H1").Value = "ING_15" Then
            Call INGENIERO
        End If
        If Range("H1").Value = "ING_VMM_1" Or Range("H1").Value = "ING_VMM_2" Or Range("H1").Value = "ING_VMM_3" Or Range("H1").Value = "ING_VMM_4" Or Range("H1").Value = "ING_VMM_5" Or Range("H1").Value = "ING_VMM_6" Or Range("H1").Value = "ING_VMM_7" Or Range("H1").Value = "ING_VMM_8" Then
            Call INGENIERO
        End If
        If Range("H1").Value = "ING_LLANOS_1" Then
            Call INGENIERO
        End If
        If Range("H1").Value = "ING_BGT_1" Then
            Call INGENIERO_BOGOTA
        End If
        If Range("H1").Value = "ING_SINU_1" Then
            Call INGENIERO
        End If
        
        
        If Range("H1").Value = "COORDINADOR_COMPENSACIONES" Then
            Call COORDINADOR_COMPENSACIONES
        End If
        If Range("H1").Value = "LIDER_PUTUMAYO" Then
            Call LIDER_PUTUMAYO
            Call REPORTE
        End If
        
        
        
        End If
End If

End Sub


