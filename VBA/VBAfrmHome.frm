VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHome 
   Caption         =   "Home"
   ClientHeight    =   12360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17130
   OleObjectBlob   =   "VBAfrmHome.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmHome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnAltaArticulo_Click()
    frmAltaArticulo.Show
End Sub

Private Sub btnModoDesarrollador_Click()
    Call AlternarModoDesarrollador
End Sub

Private Sub CommandButton1_Click()
    If Not EsCajaAbiertaHoy() Then
        MsgBox "No hay caja abierta. Primero abrí caja para poder vender.", vbExclamation
        Exit Sub
    End If
    frmNuevaVenta.Show
End Sub


Private Sub CommandButton11_Click()
    frmCaja.Show
End Sub

Private Sub CommandButton12_Click()
    frmEtiquetas.Show
End Sub

Private Sub CommandButton13_Click()
    frmNuevoCliente.Show
End Sub

Private Sub CommandButton14_Click()
    frmNuevoProveedor.Show
End Sub

Private Sub CommandButton15_Click()
    frmCompra.Show
End Sub

Private Sub CommandButton16_Click()
    frmListadoCompras.Show
End Sub

Private Sub CommandButton17_Click()
    Call HacerBackup
End Sub

Private Sub CommandButton18_Click()
    frmAjusteStock.Show
End Sub

Private Sub CommandButton4_Click()
    frmConsulta.Show
End Sub

Private Sub CommandButton5_Click()
    frmEditarArticulo.Show
End Sub

Private Sub CommandButton6_Click()
       Dim respuesta As VbMsgBoxResult

    respuesta = MsgBox("¿Estás seguro de que querés cerrar el sistema?", vbQuestion + vbYesNo, "Confirmar salida")

    If respuesta = vbYes Then
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub

Private Sub CommandButton8_Click()
    FrmListadoVentas.Show
End Sub

Private Sub CommandButton9_Click()
    frmDashboard.Show
End Sub
Private Sub UserForm_Initialize()
    Call ActualizarTotalDiaEn(Me.lblTotalDia)
End Sub
