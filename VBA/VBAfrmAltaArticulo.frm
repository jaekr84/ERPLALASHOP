VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAltaArticulo 
   Caption         =   "UserForm2"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8505.001
   OleObjectBlob   =   "VBAfrmAltaArticulo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAltaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbProveedor_Change()

End Sub

Private Sub CommandButton2_Click()
    frmNuevoProveedor.Show vbModal
End Sub
Private Sub NuevoCategoria_Click()
    frmNuevaCategoria.Show vbModal
End Sub

Private Sub UserForm_Initialize()
    Dim hojaContadores As Worksheet
    Dim tblContador   As ListObject
    Dim proximoCodigo As Long

    ' ————————————— Generar y mostrar próximo código —————————————
    Set hojaContadores = ThisWorkbook.Sheets("Contadores")
    Set tblContador = hojaContadores.ListObjects("Contador")
    
    ' Calcular siguiente código y volcarlo a TxtCodigo
    proximoCodigo = tblContador.DataBodyRange(1, 1).Value + 1
    Me.txtCodigo.Value = proximoCodigo
    
    ' ————————————— Cargar datos en comboboxes —————————————
    ' Asume que tienes un procedimiento CargarProveedores que llena cbProveedor
    Call CargarProveedores

    ' Asume que tienes un procedimiento CargarCategorias que llena cbCategoria
    Call CargarCategorias
End Sub
Private Sub btnConfirmar_Click()
    Dim hojaStock      As Worksheet
    Dim hojaContadores As Worksheet
    Dim tblStock       As ListObject
    Dim tblContador    As ListObject
    Dim nuevoCodigo    As Long
    Dim ultimoCodBarra As Long
    Dim talles()       As String
    Dim colores()      As String
    Dim i As Long, j As Long
    Dim fila As ListRow
    Dim codBarraFinal  As String

    ' ————— Validaciones básicas —————
    If Trim(txtDescripcion.Value) = "" Or _
       Trim(cbProveedor.Value) = "" Or _
       Trim(cbCategoria.Value) = "" Or _
       Trim(txtCosto.Value) = "" Or _
       Trim(txtPrecio.Value) = "" Or _
       Trim(txtTalles.Value) = "" Or _
       Trim(txtColores.Value) = "" Then
       
        MsgBox "Completá todos los campos antes de confirmar.", vbExclamation
        Exit Sub
    End If

    ' ————— Referencias a hojas y tablas —————
    Set hojaStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = hojaStock.ListObjects("Stock")
    Set hojaContadores = ThisWorkbook.Sheets("Contadores")
    Set tblContador = hojaContadores.ListObjects("Contador")

    ' ————— Leer contadores (columna 1 = ÚltimoCódigo, col 2 = ÚltimoCodBarra) —————
    nuevoCodigo = tblContador.DataBodyRange(1, 1).Value + 1
    ultimoCodBarra = tblContador.DataBodyRange(1, 2).Value

    ' Mostrar el código base en el TextBox
    Me.txtCodigo.Value = nuevoCodigo

    ' ————— Dividir talles y colores —————
    talles = Split(txtTalles.Value, ",")
    colores = Split(txtColores.Value, ",")

    ' ————— Validar valores vacíos en arrays —————
    For i = 0 To UBound(talles)
        If Trim(talles(i)) = "" Then
            MsgBox "Revisá los talles: hay un valor vacío o una coma extra.", vbExclamation
            Exit Sub
        End If
    Next i

    For j = 0 To UBound(colores)
        If Trim(colores(j)) = "" Then
            MsgBox "Revisá los colores: hay un valor vacío o una coma extra.", vbExclamation
            Exit Sub
        End If
    Next j

    ' ————— Generar variantes y agregar a la tabla Stock —————
    For i = 0 To UBound(talles)
        For j = 0 To UBound(colores)
            ' Generar nuevo código de barras
            ultimoCodBarra = ultimoCodBarra + 1
            codBarraFinal = nuevoCodigo & Format(ultimoCodBarra, "00000000")

            ' Agregar fila
            Set fila = tblStock.ListRows.Add
            With fila.Range
                .Cells(1, 1).Value = nuevoCodigo                  ' Código base
                .Cells(1, 2).Value = txtDescripcion.Value         ' Descripción
                .Cells(1, 3).Value = Val(txtCosto.Value)          ' Costo
                .Cells(1, 4).Value = cbProveedor.Value            ' Proveedor
                .Cells(1, 5).Value = Val(txtPrecio.Value)         ' Precio venta
                .Cells(1, 6).Value = 0                            ' Stock inicial
                .Cells(1, 7).NumberFormat = "@"                   ' Forzar texto
                .Cells(1, 7).Value = codBarraFinal                ' Código de barras
                .Cells(1, 8).Value = cbCategoria.Value            ' Categoría
                .Cells(1, 9).Value = Trim(talles(i))              ' Talle
                .Cells(1, 10).Value = Trim(colores(j))            ' Color
                .Cells(1, 11).Value = Date                        ' Fecha alta
            End With
        Next j
    Next i

    ' ————— Actualizar contadores en la tabla Contador —————
    With tblContador.DataBodyRange
        .Cells(1, 1).Value = nuevoCodigo    ' Guarda el nuevo ÚltimoCódigo
        .Cells(1, 2).Value = ultimoCodBarra ' Guarda el nuevo ÚltimoCodBarra
    End With

    ' ————— Confirmación y refresco de frmCompra —————
    MsgBox "Producto generado correctamente con " & _
           (UBound(talles) + 1) * (UBound(colores) + 1) & " variantes.", vbInformation

    If VBA.UserForms.Count > 0 Then
        On Error Resume Next
        frmCompra.RefreshArticulos
        On Error GoTo 0
    End If

    Unload Me
End Sub





Private Sub btnNuevoProveedor_Click()
    frmNuevoProveedor.Show vbModal
End Sub


Private Sub CommandButton1_Click()
    Unload Me
End Sub

Public Sub RecibirNuevoProveedor(fichaProveedor As String)
    Me.CargarProveedores ' recarga desde la tabla
    cbProveedor.Value = fichaProveedor ' selecciona el nuevo
End Sub
Public Sub CargarProveedores()
    Dim hojaProv As Worksheet
    Dim tblProv As ListObject
    Dim i As Long
    Dim nombre As String

    Set hojaProv = ThisWorkbook.Sheets("Proveedores")
    Set tblProv = hojaProv.ListObjects("tblProveedores")

    cbProveedor.Clear

    For i = 1 To tblProv.ListRows.Count
        nombre = tblProv.DataBodyRange(i, 2).Value ' Columna B = Nombre
        cbProveedor.AddItem nombre
    Next i
End Sub




Public Sub RecibirNuevaCategoria(fichaCategoria As String)
    Me.CargarCategorias
    cbCategoria.Value = fichaCategoria
End Sub

Public Sub CargarCategorias()
    Dim hojaCat As Worksheet
    Dim tblCat As ListObject
    Dim i As Long
    Dim nombre As String

    Set hojaCat = ThisWorkbook.Sheets("Categorias")
    Set tblCat = hojaCat.ListObjects("tblCategorias")

    cbCategoria.Clear

    For i = 1 To tblCat.ListRows.Count
        nombre = tblCat.DataBodyRange(i, 2).Value ' Columna B = Nombre
        cbCategoria.AddItem nombre
    Next i
End Sub


