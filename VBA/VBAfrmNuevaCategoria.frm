VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNuevaCategoria 
   Caption         =   "UserForm7"
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5955
   OleObjectBlob   =   "VBAfrmNuevaCategoria.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNuevaCategoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnConfirmar_Click()
    Dim hojaCat As Worksheet
    Dim tblCat As ListObject
    Dim nuevaFila As ListRow
    Dim nuevoId As Long
    Dim nuevoNombre As String
    Dim fichaCategoria As String
    Dim existe As Boolean
    Dim i As Long

    ' Validar campo
    If Trim(txtNombre.Value) = "" Then
        MsgBox "Completá el nombre de la categoría.", vbExclamation
        Exit Sub
    End If

    nuevoNombre = Trim(txtNombre.Value)

    Set hojaCat = ThisWorkbook.Sheets("Categorias")
    Set tblCat = hojaCat.ListObjects("tblCategorias")

    ' Verificar duplicado
    existe = False
    For i = 1 To tblCat.ListRows.Count
        If LCase(Trim(tblCat.DataBodyRange(i, 2).Value)) = LCase(nuevoNombre) Then
            existe = True
            Exit For
        End If
    Next i

    If existe Then
        MsgBox "Ya existe una categoría con ese nombre.", vbExclamation
        Exit Sub
    End If

    ' Obtener nuevo ID
    If tblCat.ListRows.Count = 0 Then
        nuevoId = 1
    Else
        nuevoId = Application.WorksheetFunction.Max(tblCat.ListColumns(1).DataBodyRange) + 1
    End If

    ' Agregar nueva fila
    Set nuevaFila = tblCat.ListRows.Add
    With nuevaFila.Range
        .Cells(1, 1).Value = nuevoId
        .Cells(1, 2).Value = nuevoNombre
    End With

    fichaCategoria = nuevoNombre ' Solo el nombre, ya que no mostrás el ID

    MsgBox "Categoría agregada correctamente.", vbInformation

    ' Actualizar ComboBox en UserForm2
    If frmAltaArticulo.Visible Then
        frmAltaArticulo.RecibirNuevaCategoria fichaCategoria
    End If

    Unload Me
End Sub


Private Sub btnGuardar_Click()

End Sub

If UserForm2.Visible Then
    UserForm2.RecibirNuevaCategoria fichaCategoria
End If

