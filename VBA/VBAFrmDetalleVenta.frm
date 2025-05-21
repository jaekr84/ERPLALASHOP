VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDetalleVenta 
   Caption         =   "UserForm7"
   ClientHeight    =   9225.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "VBAFrmDetalleVenta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "FrmDetalleVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub CargarComprobante(ByVal numeroComprobante As String)
    Dim ws As Worksheet
    Dim i As Long, ultimaFila As Long
    Dim subtotal As Double, descuento As Double, total As Double
    Dim filaInicial As Boolean
    Dim f As Date, medioPago As String
    Dim cliente As String

    Set ws = ThisWorkbook.Sheets("Ventas")
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ListBoxDetalle.Clear
    subtotal = 0: total = 0: descuento = 0
    filaInicial = False

    For i = 2 To ultimaFila
        If ws.Cells(i, 12).Value = numeroComprobante Then
            If Not filaInicial Then
                f = ws.Cells(i, 1).Value
                medioPago = ws.Cells(i, 8).Value
                descuento = Val(ws.Cells(i, 13).Value)
                cliente = ws.Cells(i, 14).Value
                filaInicial = True
            End If

            ListBoxDetalle.AddItem ws.Cells(i, 3).Value ' Descripción
            ListBoxDetalle.List(ListBoxDetalle.ListCount - 1, 1) = ws.Cells(i, 9).Value ' Talle
            ListBoxDetalle.List(ListBoxDetalle.ListCount - 1, 2) = ws.Cells(i, 10).Value ' Color
            ListBoxDetalle.List(ListBoxDetalle.ListCount - 1, 3) = ws.Cells(i, 4).Value ' Cantidad
            ListBoxDetalle.List(ListBoxDetalle.ListCount - 1, 4) = ws.Cells(i, 5).Value ' Precio unitario

            subtotal = subtotal + ws.Cells(i, 6).Value
            total = total + ws.Cells(i, 7).Value
        End If
    Next i

    lblFecha.Caption = "Fecha: " & Format(f, "dd/mm/yyyy")
    lblComprobante.Caption = "Comprobante: " & numeroComprobante
    lblMedioPago.Caption = "Pago con: " & medioPago
    lblSubtotal.Caption = "Subtotal: $" & CLng(subtotal)
    lblDescuento.Caption = "Descuento: $" & CLng(descuento)
    lblTotal.Caption = "Total: $" & CLng(total)
    lblCliente.Caption = "Cliente: " & cliente
End Sub

Private Sub btnImprimirCambio_Click()
    Dim numComprobante As String, fechaVenta As String

    numComprobante = Replace(lblComprobante.Caption, "Comprobante: ", "")
    fechaVenta = Replace(lblFecha.Caption, "Fecha: ", "")

    Call ImprimirTicketDeCambioConWord(numComprobante, fechaVenta)
End Sub


Private Sub btnImprimirTicket_Click()
    Dim numComprobante As String
    Dim detalles() As Variant
    Dim i As Long

    numComprobante = Replace(lblComprobante.Caption, "Comprobante: ", "")
    ReDim detalles(ListBoxDetalle.ListCount - 1, 4)

    For i = 0 To ListBoxDetalle.ListCount - 1
        detalles(i, 0) = ListBoxDetalle.List(i, 0) ' Descripción
        detalles(i, 1) = ListBoxDetalle.List(i, 1) ' Talle
        detalles(i, 2) = ListBoxDetalle.List(i, 2) ' Color
        detalles(i, 3) = CLng(ListBoxDetalle.List(i, 3)) ' Cantidad
        detalles(i, 4) = CDbl(ListBoxDetalle.List(i, 4)) ' Precio unitario
    Next i

    Call ImprimirTicketConWord( _
        numComprobante, _
        Replace(lblFecha.Caption, "Fecha: ", ""), _
        Replace(lblMedioPago.Caption, "Pago con: ", ""), _
        Val(Replace(lblSubtotal.Caption, "Subtotal: $", "")), _
        Val(Replace(lblDescuento.Caption, "Descuento: $", "")), _
        Val(Replace(lblTotal.Caption, "Total: $", "")), _
        detalles)
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

