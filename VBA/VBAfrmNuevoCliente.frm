VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNuevoCliente 
   Caption         =   "UserForm7"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6945
   OleObjectBlob   =   "VBAfrmNuevoCliente.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNuevoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnGenerar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nuevaFila As ListRow
    Dim nuevoId As Long
    Dim idFormateado As String
    Dim fichaCliente As String

    ' Validación básica
    If Trim(txtNombre.Value) = "" Or Trim(txtApellido.Value) = "" Then
        MsgBox "Completá al menos nombre y apellido.", vbExclamation
        Exit Sub
    End If

    ' Referencias
    Set ws = ThisWorkbook.Sheets("Clientes")
    Set tbl = ws.ListObjects("tblClientes")

    ' Obtener nuevo ID
    If tbl.ListRows.Count = 0 Then
        nuevoId = 1
    Else
        nuevoId = Application.WorksheetFunction.Max(tbl.ListColumns(1).DataBodyRange) + 1
    End If
    idFormateado = Format(nuevoId, "00000000")

    ' Agregar fila
    Set nuevaFila = tbl.ListRows.Add
    With nuevaFila.Range
        .Cells(1, 1).Value = nuevoId                      ' ID
        .Cells(1, 2).Value = txtNombre.Value              ' Nombre
        .Cells(1, 3).Value = txtApellido.Value            ' Apellido
        .Cells(1, 4).Value = txtTelefono.Value            ' Teléfono
        .Cells(1, 5).Value = txtDni.Value                 ' DNI
        .Cells(1, 6).Value = txtFechaNac.Value            ' Fecha Nac
    End With

    ' Mostrar ID en el formulario
    lblIdCliente.Caption = idFormateado

    ' Armar el texto para seleccionar en el combo
    fichaCliente = idFormateado & " - " & txtNombre.Value & " " & txtApellido.Value & " | " & txtDni.Value


    ' Confirmar y cerrar
    MsgBox "Cliente registrado y seleccionado.", vbInformation
    Unload Me
End Sub





Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim nuevoId As Long

    Set ws = ThisWorkbook.Sheets("Clientes")
    Set tbl = ws.ListObjects("tblClientes")

    If tbl.ListRows.Count = 0 Then
        nuevoId = 1
    Else
        nuevoId = Application.WorksheetFunction.Max(tbl.ListColumns(1).DataBodyRange) + 1
    End If

    lblIdCliente.Caption = "Número de Cliente: " & Format(nuevoId, "00000000")
End Sub



