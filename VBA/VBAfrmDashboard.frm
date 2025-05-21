VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDashboard 
   Caption         =   "UserForm7"
   ClientHeight    =   13395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18720
   OleObjectBlob   =   "VBAfrmDashboard.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnActualizarDashboard_Click()
    Call CargarTopProductos("DIA", Me.lstArtDia)
    Call CargarTopProductos("SEM", Me.lstArtSem)
    Call CargarTopProductos("MES", Me.lstArtMes)
    Call CargarTopProductos("ANIO", Me.lstArtAnio)
    CargarTopCategorias Me
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub btnActualizarRotacion_Click()
    Call AnalizarRotacionAlta(Me)
End Sub

Private Sub lstDia_Click()

End Sub

Private Sub lstTalleAnio_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim hoy As Date
    Dim primerDia As Date
    Dim ultimoDia As Date

    hoy = Date
    primerDia = DateSerial(Year(hoy), Month(hoy), 1)
    ultimoDia = DateSerial(Year(hoy), Month(hoy) + 1, 0)

    txtDesde.Value = Format(primerDia, "dd/mm/yyyy")
    txtHasta.Value = Format(ultimoDia, "dd/mm/yyyy")
    
    Call ActualizarTotalesVentasDashboard(Me)

    Call CargarTopTalles(Me)
    CargarTopCategorias Me
    Call CargarTopProductos("DIA", Me.lstArtDia)
    Call CargarTopProductos("SEM", Me.lstArtSem)
    Call CargarTopProductos("MES", Me.lstArtMes)
    Call CargarTopProductos("ANIO", Me.lstArtAnio)

End Sub

Private Sub lstArtDia_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AbrirDetalleProductoDesdeLista Me.lstArtDia
End Sub

Private Sub lstArtSem_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AbrirDetalleProductoDesdeLista Me.lstArtSem
End Sub

Private Sub lstArtMes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AbrirDetalleProductoDesdeLista Me.lstArtMes
End Sub

Private Sub lstArtAnio_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    AbrirDetalleProductoDesdeLista Me.lstArtAnio
End Sub

