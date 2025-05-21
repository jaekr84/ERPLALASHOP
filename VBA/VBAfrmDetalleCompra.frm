VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDetalleCompra 
   Caption         =   "UserForm1"
   ClientHeight    =   9825.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   OleObjectBlob   =   "VBAfrmDetalleCompra.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDetalleCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub CargarComprobante(nroComprobante As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim fecha As Date
    Dim total As Double
    Dim nombreProv As String

    ' Limpiar controles
    lstDetalle.Clear
    lblProveedor.Caption = ""
    lblFecha.Caption = ""
    lblComprobante.Caption = ""
    lblDescuento.Caption = "0"
    lblTotal.Caption = ""

    Set ws = ThisWorkbook.Sheets("Compras")
    Set tbl = ws.ListObjects("tblCompras")

    For Each fila In tbl.ListRows
        If Trim(CStr(fila.Range.Cells(1, 10).Value)) = nroComprobante Then
            ' Obtener datos generales solo una vez
            If lblProveedor.Caption = "" Then
                nombreProv = BuscarNombreProveedor(fila.Range.Cells(1, 2).Value)
                lblProveedor.Caption = nombreProv
                fecha = fila.Range.Cells(1, 1).Value
                lblFecha.Caption = Format(fecha, "dd/mm/yyyy")
                lblComprobante.Caption = nroComprobante
            End If

            ' Agregar ítem al ListBox
            Dim cod, desc, talle, color, cant, costo, subtotal As Variant
            cod = fila.Range.Cells(1, 3).Value
            desc = fila.Range.Cells(1, 4).Value
            talle = fila.Range.Cells(1, 5).Value
            color = fila.Range.Cells(1, 6).Value
            cant = fila.Range.Cells(1, 7).Value
            costo = fila.Range.Cells(1, 8).Value
            subtotal = fila.Range.Cells(1, 9).Value

            lstDetalle.AddItem cod
            lstDetalle.List(lstDetalle.ListCount - 1, 1) = desc
            lstDetalle.List(lstDetalle.ListCount - 1, 2) = talle
            lstDetalle.List(lstDetalle.ListCount - 1, 3) = color
            lstDetalle.List(lstDetalle.ListCount - 1, 4) = cant
            lstDetalle.List(lstDetalle.ListCount - 1, 5) = Format(costo, "#,##0")
            lstDetalle.List(lstDetalle.ListCount - 1, 6) = Format(subtotal, "#,##0")

            total = total + subtotal
        End If
    Next fila

    lblTotal.Caption = "$" & Format(total, "#,##0")
End Sub

Private Function BuscarNombreProveedor(idProv As Variant) As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Proveedores")
    Set tbl = ws.ListObjects("tblProveedores")

    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = idProv Then
            BuscarNombreProveedor = tbl.DataBodyRange(i, 2).Value ' columna 2: nombre
            Exit Function
        End If
    Next i

    BuscarNombreProveedor = "Proveedor desconocido"
End Function

Private Sub CommandButton3_Click()
    Unload Me
End Sub
