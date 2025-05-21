VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNuevoProveedor 
   Caption         =   "UserForm7"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   OleObjectBlob   =   "VBAfrmNuevoProveedor.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNuevoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancelar_Click()
    Unload Me
End Sub
Private Sub btnConfirmar_Click()
    Dim hojaProv As Worksheet
    Dim tblProv As ListObject
    Dim nuevaFila As ListRow
    Dim nuevoCodigo As String, nuevoNombre As String, nuevaDireccion As String
    Dim existe As Boolean
    Dim i As Long

    ' 1) Validar campos
    If Trim(Me.txtNombre.Value) = "" Or Trim(Me.txtDireccion.Value) = "" Then
        MsgBox "Completá el nombre y la dirección del proveedor.", vbExclamation
        Exit Sub
    End If

    ' 2) Preparar referencias
    Set hojaProv = ThisWorkbook.Sheets("Proveedores")
    Set tblProv = hojaProv.ListObjects("tblProveedores")

    ' 3) Tomar valores de los TextBoxes
    nuevoCodigo = Me.txtCodigo.Value
    nuevoNombre = Trim(Me.txtNombre.Value)
    nuevaDireccion = Trim(Me.txtDireccion.Value)

    ' 4) Verificar duplicado por nombre (columna 2)
    existe = False
    For i = 1 To tblProv.ListRows.Count
        If LCase(Trim(tblProv.ListRows(i).Range.Cells(1, 2).Value)) = LCase(nuevoNombre) Then
            existe = True
            Exit For
        End If
    Next i

    If existe Then
        MsgBox "Ya existe un proveedor con ese nombre.", vbExclamation
        Exit Sub
    End If

    ' 5) Agregar nueva fila a la tabla
    Set nuevaFila = tblProv.ListRows.Add
    With nuevaFila.Range
        .Cells(1, 1).Value = nuevoCodigo
        .Cells(1, 2).Value = nuevoNombre
        .Cells(1, 3).Value = nuevaDireccion
    End With

    ' 6) Actualizar cmbProveedor en frmCompra
    With frmCompra.cmbProveedor
        .AddItem nuevoCodigo                           ' Columna 0 = ID
        .List(.ListCount - 1, 1) = nuevoNombre         ' Columna 1 = Nombre
        .Value = nuevoCodigo                           ' Selecciono automáticamente
    End With

    ' 7) Informar al usuario y cerrar
    MsgBox "Proveedor agregado correctamente.", vbInformation
    Unload Me
End Sub




Private Sub UserForm_Initialize()
    Dim hojaProv As Worksheet
    Dim tblProv As ListObject
    Dim ultimaFila As Long
    Dim nuevoCodigo As Long

    Set hojaProv = ThisWorkbook.Sheets("Proveedores")
    Set tblProv = hojaProv.ListObjects("tblProveedores")

    If tblProv.ListRows.Count = 0 Then
        nuevoCodigo = 100
    Else
        ' Buscar el mayor ID actual
        ultimaFila = tblProv.DataBodyRange.Rows.Count
        nuevoCodigo = Application.WorksheetFunction.Max(tblProv.ListColumns(1).DataBodyRange) + 1
    End If

    txtCodigo.Value = nuevoCodigo
    txtCodigo.Enabled = False ' para que no se pueda modificar
    

End Sub


Public Sub CargarProveedores()
    Dim hojaProv As Worksheet
    Dim tblProv As ListObject
    Dim i As Long
    Dim id As String, nombre As String, direccion As String
    Dim fichaProveedor As String

    Set hojaProv = ThisWorkbook.Sheets("Proveedores")
    Set tblProv = hojaProv.ListObjects("tblProveedores")

    cbProveedor.Clear

    For i = 1 To tblProv.ListRows.Count
        With tblProv.ListRows(i).Range
            id = .Cells(1, 1).Value
            nombre = .Cells(1, 2).Value
            direccion = .Cells(1, 3).Value
        End With

        fichaProveedor = id & " - " & nombre & " | " & direccion
        cbProveedor.AddItem fichaProveedor
    Next i
End Sub

