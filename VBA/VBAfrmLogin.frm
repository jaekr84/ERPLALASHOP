VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "UserForm7"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   OleObjectBlob   =   "VBAfrmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnIngresar_Click()
    Dim usuario As String
    Dim clave As String

    usuario = LCase(Trim(txtUsuario.Value))
    clave = Trim(txtClave.Value)

    ' Usuario y clave válidos
    If usuario = "admin" And clave = "admin" Then
        Application.Visible = True ' <- Esta línea es clave para volver a mostrar Excel
        Call AlternarModoDesarrollador
        Unload Me
    Else
        MsgBox "Usuario o contraseña incorrectos.", vbCritical
        txtClave.Value = ""
        txtClave.SetFocus
    End If
End Sub



