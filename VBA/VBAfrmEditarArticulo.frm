VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditarArticulo 
   Caption         =   "UserForm5"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150.001
   OleObjectBlob   =   "VBAfrmEditarArticulo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEditarArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnAgregarVariante_Click()
    Dim hojaStock As Worksheet, hojaContadores As Worksheet
    Dim tblStock As ListObject, tblContador As ListObject
    Dim nuevoCodBarra As String
    Dim codBarraNum As Long
    Dim fila As ListRow
    Dim codigoBase As String

    ' Validar campos obligatorios
    If Trim(txtNuevoTalle.Value) = "" Or Trim(txtNuevoColor.Value) = "" Then
        MsgBox "Completá talle y color para agregar una variante.", vbExclamation
        Exit Sub
    End If

    codigoBase = Trim(txtCodigo.Value)
    If codigoBase = "" Then
        MsgBox "Primero buscá un producto para poder agregar variantes.", vbExclamation
        Exit Sub
    End If

    ' Referencias
    Set hojaStock = ThisWorkbook.Sheets("Stock")
    Set hojaContadores = ThisWorkbook.Sheets("Contadores")
    Set tblStock = hojaStock.ListObjects("Stock")
    Set tblContador = hojaContadores.ListObjects("Contador")

    ' Generar nuevo código de barra
    codBarraNum = tblContador.DataBodyRange(1, 2).Value + 1
    nuevoCodBarra = codigoBase & Trim(txtNuevoTalle.Value) & Trim(txtNuevoColor.Value) & Format(codBarraNum, "00000")

    ' Agregar fila a Stock
    Set fila = tblStock.ListRows.Add
    With fila.Range
        .Cells(1, 1).Value = codigoBase                         ' Código
        .Cells(1, 2).Value = txtDescripcion.Value              ' Descripción
        .Cells(1, 3).Value = Val(txtCosto.Value)               ' Costo
        .Cells(1, 4).Value = txtProveedor.Value                ' Proveedor
        .Cells(1, 5).Value = Val(txtPrecio.Value)              ' Precio Venta
        .Cells(1, 6).Value = 0                                 ' Stock inicial
        .Cells(1, 7).Value = nuevoCodBarra                     ' Código de barra
        .Cells(1, 8).Value = txtCategoria.Value                ' Categoría
        .Cells(1, 9).Value = Trim(txtNuevoTalle.Value)         ' Talle
        .Cells(1, 10).Value = Trim(txtNuevoColor.Value)        ' Color
        .Cells(1, 11).Value = Date                             ' Fecha
    End With

    ' Actualizar contador
    tblContador.DataBodyRange(1, 2).Value = codBarraNum

    ' Agregar al ListBox
    ListBox1.AddItem txtNuevoTalle.Value
    ListBox1.List(ListBox1.ListCount - 1, 1) = txtNuevoColor.Value
    ListBox1.List(ListBox1.ListCount - 1, 2) = 0
    ListBox1.List(ListBox1.ListCount - 1, 3) = nuevoCodBarra

    ' Limpiar campos
    txtNuevoTalle.Value = ""
    txtNuevoColor.Value = ""

    MsgBox "Variante agregada correctamente.", vbInformation
End Sub


Private Sub btnBuscar_Click()
    Dim hojaStock As Worksheet
    Dim tblStock As ListObject
    Dim codigoBuscado As String
    Dim i As Long
    Dim primeraCargada As Boolean

    Set hojaStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = hojaStock.ListObjects("Stock")

    codigoBuscado = Trim(txtCodigoBuscar.Value)
    If codigoBuscado = "" Then
        MsgBox "Por favor ingresá un código para buscar.", vbExclamation
        Exit Sub
    End If

    ' Inicializar
    primeraCargada = False
    ListBox1.Clear

    ' Recorrer toda la tabla stock
    For i = 1 To tblStock.ListRows.Count
        With tblStock.ListRows(i).Range
            If Trim(.Cells(1, 1).Value) = codigoBuscado Then
                ' Cargar datos padre desde la primera coincidencia
                If Not primeraCargada Then
                    txtDescripcion.Value = .Cells(1, 2).Value
                    txtCosto.Value = .Cells(1, 3).Value
                    txtProveedor.Value = .Cells(1, 4).Value
                    txtPrecio.Value = .Cells(1, 5).Value
                    txtCategoria.Value = .Cells(1, 8).Value
                    txtCodigo.Value = .Cells(1, 1).Value
                    primeraCargada = True
                End If

                ' Agregar variante al ListBox
                ListBox1.AddItem .Cells(1, 9).Value ' Talle
                ListBox1.List(ListBox1.ListCount - 1, 1) = .Cells(1, 10).Value ' Color
                ListBox1.List(ListBox1.ListCount - 1, 2) = .Cells(1, 6).Value ' Stock
                ListBox1.List(ListBox1.ListCount - 1, 3) = .Cells(1, 7).Value ' Código de barra
                End If
            End With
    Next i

    If Not primeraCargada Then
        MsgBox "No se encontraron variantes con ese código.", vbExclamation
    End If
End Sub

Private Sub btnCerrar_Click()
 Unload Me
End Sub
Private Sub btnGuardarCambios_Click()
    Dim hojaStock As Worksheet
    Dim tblStock As ListObject
    Dim i As Long
    Dim codigoBase As String
    Dim contador As Long

    codigoBase = Trim(txtCodigo.Value)
    If codigoBase = "" Then
        MsgBox "No hay producto cargado para actualizar.", vbExclamation
        Exit Sub
    End If

    ' VALIDACIÓN de campos obligatorios
    If Trim(txtDescripcion.Value) = "" Or _
       Trim(txtProveedor.Value) = "" Or _
       Trim(txtCategoria.Value) = "" Or _
       Trim(txtPrecio.Value) = "" Or _
       Trim(txtCosto.Value) = "" Then

        MsgBox "Todos los campos deben estar completos antes de guardar.", vbExclamation
        Exit Sub
    End If

    Set hojaStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = hojaStock.ListObjects("Stock")
    contador = 0

    ' Actualizar cada fila con el mismo código base
    For i = 1 To tblStock.ListRows.Count
        With tblStock.ListRows(i).Range
            If Trim(.Cells(1, 1).Value) = codigoBase Then
                .Cells(1, 2).Value = txtDescripcion.Value
                .Cells(1, 3).Value = Val(txtCosto.Value)
                .Cells(1, 4).Value = txtProveedor.Value
                .Cells(1, 5).Value = Val(txtPrecio.Value)
                .Cells(1, 8).Value = txtCategoria.Value
                contador = contador + 1
            End If
        End With
    Next i

    MsgBox "Se actualizaron " & contador & " variantes del producto correctamente.", vbInformation

    ' Limpiar formulario
    txtCodigoBuscar.Value = ""
    txtCodigo.Value = ""
    txtDescripcion.Value = ""
    txtCosto.Value = ""
    txtProveedor.Value = ""
    txtPrecio.Value = ""
    txtCategoria.Value = ""
    txtNuevoTalle.Value = ""
    txtNuevoColor.Value = ""
    ListBox1.Clear
    txtCodigoBuscar.SetFocus
End Sub


