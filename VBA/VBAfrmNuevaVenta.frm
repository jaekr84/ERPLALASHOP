VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNuevaVenta 
   Caption         =   "UserForm1"
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
   OleObjectBlob   =   "VBAfrmNuevaVenta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmNuevaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAgregar_Click()
    Dim datos() As String
    Dim codigoBuscado As String
    Dim codigoBarra As String
    Dim i As Long, j As Long
    Dim hojaStock As Worksheet
    Dim descripcion As String, talle As String, color As String
    Dim cantidad As Double, precio As Double, total As Double
    Dim yaExiste As Boolean

    ' Validación
    If Trim(cbCodigo.Value) = "" Then
        MsgBox "Seleccioná un producto para agregar.", vbExclamation
        Exit Sub
    End If

    datos = Split(cbCodigo.Value, "|")
    If UBound(datos) < 4 Then
        MsgBox "El formato del producto no es válido.", vbExclamation
        Exit Sub
    End If

    codigoBuscado = Trim(datos(0))
    codigoBarra = Trim(datos(4)) ' viene del ComboBox

    Set hojaStock = ThisWorkbook.Sheets("Stock")

    ' Buscar producto exacto por código + código de barra
    For i = 2 To hojaStock.Cells(hojaStock.Rows.Count, 1).End(xlUp).Row
        If hojaStock.Cells(i, 1).Value = codigoBuscado And Trim(hojaStock.Cells(i, 7).Value) = codigoBarra Then
            descripcion = hojaStock.Cells(i, 2).Value
            talle = hojaStock.Cells(i, 9).Value
            color = hojaStock.Cells(i, 10).Value
            precio = hojaStock.Cells(i, 5).Value
            cantidad = 1
            total = cantidad * precio
            yaExiste = False

            With lstDetalle
                For j = 0 To .ListCount - 1
                    If Trim(.List(j, 8)) = codigoBarra Then
                        .List(j, 4) = CDbl(.List(j, 4)) + cantidad
                        .List(j, 6) = CDbl(.List(j, 4)) * precio
                        .List(j, 7) = IIf(CDbl(.List(j, 4)) < 0, "CAMBIO", "VENTA")
                        yaExiste = True
                        Exit For
                    End If
                Next j

                If Not yaExiste Then
                    .AddItem codigoBuscado
                    .List(.ListCount - 1, 1) = descripcion
                    .List(.ListCount - 1, 2) = talle
                    .List(.ListCount - 1, 3) = color
                    .List(.ListCount - 1, 4) = cantidad
                    .List(.ListCount - 1, 5) = precio
                    .List(.ListCount - 1, 6) = total
                    .List(.ListCount - 1, 7) = IIf(cantidad < 0, "CAMBIO", "VENTA")
                    .List(.ListCount - 1, 8) = codigoBarra
                End If
            End With

            Call CalcularTotalDetalle
            cbCodigo.Value = ""
            Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
            Exit Sub
        End If
    Next i

    MsgBox "No se encontró el producto en el stock.", vbExclamation
    cbCodigo.Value = ""
    Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
End Sub


Private Sub lstDetalle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim cantidad As String
    Dim fila As Long
    Dim precio As Double

    fila = lstDetalle.ListIndex
    If fila < 0 Then Exit Sub

    cantidad = InputBox("Ingresá la cantidad (usa negativo si es devolución):", "Editar cantidad", lstDetalle.List(fila, 4))

    If IsNumeric(cantidad) And cantidad <> "" Then
        lstDetalle.List(fila, 4) = CDbl(cantidad)
        precio = CDbl(lstDetalle.List(fila, 5))
        lstDetalle.List(fila, 6) = CDbl(cantidad) * precio

        If CDbl(cantidad) < 0 Then
            lstDetalle.List(fila, 7) = "CAMBIO"
        Else
            lstDetalle.List(fila, 7) = ""
        End If

        Call CalcularTotalDetalle
    Else
        MsgBox "Ingresá un número válido.", vbExclamation
    End If
End Sub


Private Sub btnConfirmar_Click()
    Dim hojaVentas As Worksheet, hojaStock As Worksheet, hojaMov As Worksheet, hojaPagos As Worksheet
    Dim tblVentas As ListObject, tblStock As ListObject, tblPagos As ListObject
    Dim i As Long, j As Long, filaMov As Long, filaVenta As Long
    Dim cod As String, desc As String, talle As String, color As String
    Dim cant As Double, precio As Double, subtotal As Double, totalProporcional As Double
    Dim categoria As String, tipoOperacion As String
    Dim fecha As Date, comprobante As String, descuento As Double
    Dim idCliente As String
    Dim totalBruto As Double, totalNetoFinal As Double
    Dim detalles() As Variant
    Dim fila As Long
    Dim medioPago1 As String, medioPago2 As String
    Dim montoPago1 As Double, montoPago2 As Double
    Dim medioPagoTexto As String
    Dim newStock As Double
    Dim codBarra As String


    ' VALIDACIONES
    If lstDetalle.ListCount = 0 Then
        MsgBox "No hay productos en la venta.", vbExclamation
        Exit Sub
    End If

    If Trim(cbClientes.Value) = "" Then
        MsgBox "Seleccioná un cliente.", vbExclamation
        Exit Sub
    End If

    If Trim(cmbMedioPago.Value) = "" Then
        MsgBox "Seleccioná un medio de pago.", vbExclamation
        Exit Sub
    End If
        
   

    ' OBTENER MEDIOS DE PAGO
    medioPago1 = Trim(cmbMedioPago.Value)
    medioPago2 = Trim(cmbMedioPago1.Value)
    
    If Trim(txtMontoPago.Value) = "" Then
        montoPago1 = NumeroLimpio(txtTotalNeto.Value)
        txtMontoPago.Value = Format(montoPago1, "#,##0")
        montoPago2 = 0
        txtMontoPago1.Value = "0"
        cmbMedioPago1.Value = ""
        cmbMedioPago1.Enabled = False
    Else
        montoPago1 = NumeroLimpio(txtMontoPago.Value)
        montoPago2 = NumeroLimpio(txtMontoPago1.Value)
    End If

    ' VALIDAR TOTALES
    totalBruto = NumeroLimpio(txtTotalBruto.Value)
    totalNetoFinal = NumeroLimpio(txtTotalNeto.Value)

    If montoPago1 + montoPago2 <> totalNetoFinal Then
        MsgBox "La suma de los pagos no coincide con el total a cobrar." & vbNewLine & _
               "Total: $" & Format(totalNetoFinal, "#,##0") & vbNewLine & _
               "Ingresado: $" & Format(montoPago1 + montoPago2, "#,##0"), vbExclamation
        Exit Sub
    End If

    ' PREPARACIÓN
    Set hojaVentas = ThisWorkbook.Sheets("Ventas")
    Set hojaStock = ThisWorkbook.Sheets("Stock")
    Set hojaMov = ThisWorkbook.Sheets("MovimientosStock")
    Set hojaPagos = ThisWorkbook.Sheets("RegMediosPago")
    Set tblVentas = hojaVentas.ListObjects("Tabla1")
    Set tblStock = hojaStock.ListObjects("Stock")
    Set tblPagos = hojaPagos.ListObjects("tblRegMediosPago")

    fecha = Date
    comprobante = lblComprobante.Caption
    descuento = Val(txtDescuento.Value)
    idCliente = ObtenerIDClientePorNombre(cbClientes.Value)

    ' REGISTRAR VENTAS Y DEVOLUCIONES
    For i = 0 To lstDetalle.ListCount - 1
        cod = lstDetalle.List(i, 0)
        desc = lstDetalle.List(i, 1)
        talle = lstDetalle.List(i, 2)
        color = lstDetalle.List(i, 3)
        cant = CDbl(lstDetalle.List(i, 4))
        precio = CDbl(lstDetalle.List(i, 5))
        subtotal = Abs(cant) * precio
        categoria = ObtenerCategoria(cod, talle, color)
        codBarra = lstDetalle.List(i, 8)

        totalProporcional = subtotal / totalBruto * totalNetoFinal
        totalProporcional = Int(totalProporcional)
        tipoOperacion = IIf(cant < 0, "Devolución", "Venta")

        ' AGREGAR A LA TABLA DE VENTAS
        With tblVentas.ListRows.Add.Range
            .Cells(1, 1).Value = fecha
            .Cells(1, 2).Value = cod
            .Cells(1, 3).Value = desc
            .Cells(1, 4).Value = cant
            .Cells(1, 5).Value = precio
            .Cells(1, 6).Value = subtotal
            .Cells(1, 7).Value = IIf(cant < 0, -totalProporcional, totalProporcional)
            .Cells(1, 8).Value = medioPago1
            .Cells(1, 9).Value = talle
            .Cells(1, 10).Value = color
            .Cells(1, 11).Value = categoria
            .Cells(1, 12).Value = comprobante
            .Cells(1, 13).Value = descuento
            .Cells(1, 14).Value = idCliente
            .Cells(1, 15).Value = tipoOperacion
            .Cells(1, 16).Value = montoPago1
            .Cells(1, 17).Value = montoPago2
        End With

        ' ————————— ACTUALIZAR STOCK —————————
        Dim r As ListRow
        For Each r In tblStock.ListRows
            ' comparo el código de barras en la columna 7 de la tabla Stock
            If r.Range.Cells(1, 7).Value = codBarra Then
                ' calculo el nuevo stock y evito que baje de cero
                newStock = r.Range.Cells(1, 6).Value - cant
                If newStock < 0 Then newStock = 0
                r.Range.Cells(1, 6).Value = newStock
                Exit For
            End If
        Next r


        ' REGISTRAR EN MOVIMIENTOS
        filaMov = hojaMov.Cells(hojaMov.Rows.Count, 1).End(xlUp).Row + 1
        hojaMov.Cells(filaMov, 1).Value = fecha
        hojaMov.Cells(filaMov, 2).Value = cod
        hojaMov.Cells(filaMov, 3).Value = desc
        hojaMov.Cells(filaMov, 4).Value = talle
        hojaMov.Cells(filaMov, 5).Value = color
        hojaMov.Cells(filaMov, 6).Value = Abs(cant)
        hojaMov.Cells(filaMov, 7).Value = tipoOperacion
    Next i

    ' GUARDAR EN RegMediosPago
    With tblPagos.ListRows.Add.Range
        .Cells(1, 1).Value = fecha
        .Cells(1, 2).Value = comprobante
        .Cells(1, 3).Value = medioPago1
        .Cells(1, 4).Value = montoPago1
        .Cells(1, 5).Value = medioPago2
        .Cells(1, 6).Value = montoPago2
        .Cells(1, 7).Value = montoPago1 + montoPago2
    End With

    ' DETALLES PARA TICKET
    ReDim detalles(0 To lstDetalle.ListCount - 1, 0 To 4)
    For fila = 0 To lstDetalle.ListCount - 1
        detalles(fila, 0) = lstDetalle.List(fila, 1)
        detalles(fila, 1) = lstDetalle.List(fila, 2)
        detalles(fila, 2) = lstDetalle.List(fila, 3)
        detalles(fila, 3) = lstDetalle.List(fila, 4)
        detalles(fila, 4) = lstDetalle.List(fila, 5)
    Next fila

    ' TEXTO DE PAGO
    medioPagoTexto = medioPago1 & ": $" & Format(montoPago1, "#,##0")
    If medioPago2 <> "" And montoPago2 > 0 Then
        medioPagoTexto = medioPagoTexto & " / " & medioPago2 & ": $" & Format(montoPago2, "#,##0")
    End If

    MsgBox "Venta registrada correctamente.", vbInformation

    ' IMPRIMIR TICKET
    Call ImprimirTicketConWord(comprobante, Format(fecha, "dd/mm/yyyy"), medioPagoTexto, totalBruto, descuento, totalNetoFinal, detalles)
           
    Call ActualizarTotalDiaEn(frmHome.lblTotalDia)
    
    ThisWorkbook.Save
    Call HacerBackup

    Unload Me
End Sub

Private Sub btnEliminar_Click()
    Dim fila As Long

    fila = lstDetalle.ListIndex
    If fila < 0 Then
        MsgBox "Seleccioná un producto para eliminar.", vbExclamation
        Exit Sub
    End If

    lstDetalle.RemoveItem fila
    Call CalcularTotalDetalle

    ' Reenfocar el escáner después de eliminar
    Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
End Sub


Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub btnNuevoCliente_Click()
    origenLlamada = "home"
    frmNuevoCliente.Show
End Sub

Private Sub CommandButton4_Click()
    Dim numComprobante As String, fechaVenta As String

    numComprobante = Replace(lblComprobante.Caption, "Comprobante: ", "")
    fechaVenta = Replace(lblFecha.Caption, "Fecha: ", "")

    Call ImprimirTicketDeCambioConWord(numComprobante, fechaVenta)
End Sub

Private Sub txtCodigoBarra_AfterUpdate()
    Dim hojaStock As Worksheet
    Dim i As Long, j As Long
    Dim codigoBarra As String
    Dim codigo As String, descripcion As String
    Dim talle As String, color As String
    Dim cantidad As Double, precio As Double, total As Double
    Dim yaExiste As Boolean
    Dim stockActual As Double

    codigoBarra = Trim(txtCodigoBarra.Value)
    If codigoBarra = "" Then Exit Sub

    Set hojaStock = ThisWorkbook.Sheets("Stock")

    For i = 2 To hojaStock.Cells(hojaStock.Rows.Count, 7).End(xlUp).Row
        If Trim(hojaStock.Cells(i, 7).Value) = codigoBarra Then
            stockActual = hojaStock.Cells(i, 6).Value

            If stockActual <= 0 Then
                MsgBox "No hay stock disponible para este producto.", vbExclamation
                txtCodigoBarra.Value = ""
                Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
                Exit Sub
            End If

            codigo = hojaStock.Cells(i, 1).Value
            descripcion = hojaStock.Cells(i, 2).Value
            talle = hojaStock.Cells(i, 9).Value
            color = hojaStock.Cells(i, 10).Value
            precio = hojaStock.Cells(i, 5).Value
            cantidad = 1
            total = cantidad * precio
            yaExiste = False

            With lstDetalle
                For j = 0 To .ListCount - 1
                    If Trim(.List(j, 8)) = codigoBarra Then
                        .List(j, 4) = CDbl(.List(j, 4)) + cantidad
                        .List(j, 6) = CDbl(.List(j, 4)) * precio
                        .List(j, 7) = IIf(CDbl(.List(j, 4)) < 0, "CAMBIO", "VENTA")
                        yaExiste = True
                        Exit For
                    End If
                Next j

                If Not yaExiste Then
                    .AddItem codigo
                    .List(.ListCount - 1, 1) = descripcion
                    .List(.ListCount - 1, 2) = talle
                    .List(.ListCount - 1, 3) = color
                    .List(.ListCount - 1, 4) = cantidad
                    .List(.ListCount - 1, 5) = precio
                    .List(.ListCount - 1, 6) = total
                    .List(.ListCount - 1, 7) = IIf(cantidad < 0, "CAMBIO", "VENTA")
                    .List(.ListCount - 1, 8) = codigoBarra
                End If
            End With

            Call CalcularTotalDetalle
            txtCodigoBarra.Value = ""
            Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
            Exit Sub
        End If
    Next i

    ' Si no se encontró ningún producto:
    MsgBox "Código de barra no encontrado en el stock.", vbExclamation
    txtCodigoBarra.Value = ""
    Application.OnTime Now + TimeValue("00:00:01"), "ForzarFocoEnCodigoBarra"
End Sub



Private Sub txtMontoPago_Change()
    Dim total As Double
    Dim monto As Double

    If Not IsNumeric(txtMontoPago.Value) Then Exit Sub

    total = Val(Replace(txtTotalNeto.Value, ".", ""))
    monto = Val(Replace(txtMontoPago.Value, ".", "")) ' ? limpieza para cálculo

    txtMontoPago.Value = Format(monto, "#,##0") ' ? volver a escribirlo formateado
    txtMontoPago1.Value = Format(total - monto, "#,##0")

 
End Sub

Private Sub UserForm_Initialize()
    Dim wsCaja As Worksheet
    Dim tblCaja As ListObject
    Dim i As Long

    ' Verificar si hay caja abierta
    Set wsCaja = ThisWorkbook.Sheets("Caja")
    Set tblCaja = wsCaja.ListObjects("tblCaja")
    
    If tblCaja.ListRows.Count = 0 Then
        MsgBox "No hay caja abierta. Primero abrí caja para poder vender.", vbExclamation
        Unload Me
        Exit Sub
    End If

    ' Inicialización normal
    lblComprobante.Caption = ObtenerNuevoComprobanteVenta()
    lblFecha.Caption = Format(Now, "dd/mm/yyyy HH:nn")
    Call CargarMediosPago(cmbMedioPago)
    Call CargarMediosPago(cmbMedioPago1)
    Call CargarClientesEnCombo(cbClientes)
    Call CargarProductosEnCombo(cbCodigo)

    ' Seleccionar "Consumidor Final" como cliente por defecto
    For i = 0 To cbClientes.ListCount - 1
        If InStr(1, cbClientes.List(i), "Consumidor Final", vbTextCompare) > 0 Then
            cbClientes.ListIndex = i
            Exit For
        End If
    Next i

    btnConfirmar.Enabled = True
   
End Sub

Sub CargarClientesEnCombo(cmb As ComboBox)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim id As String, nombre As String, apellido As String, dni As String
    Dim fichaCliente As String

    Set ws = ThisWorkbook.Sheets("Clientes")
    Set tbl = ws.ListObjects("tblClientes")
    cmb.Clear

    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            id = .Cells(1, 1).Value
            nombre = .Cells(1, 2).Value
            apellido = .Cells(1, 3).Value
            dni = .Cells(1, 5).Value
        End With

        fichaCliente = nombre & " " & apellido & " | ID " & id & " | DNI " & dni
        cmb.AddItem fichaCliente
    Next i
End Sub

Sub CargarProductosEnCombo(cmb As ComboBox)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim codigo As String, descripcion As String, talle As String, color As String, codBarra As String
    Dim itemTexto As String

    Set ws = ThisWorkbook.Sheets("Stock")
    Set tbl = ws.ListObjects("Stock")
    cmb.Clear

    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            codigo = .Cells(1, 1).Value
            descripcion = .Cells(1, 2).Value
            codBarra = .Cells(1, 7).Value  ' Código de barra (col 7)
            talle = .Cells(1, 9).Value
            color = .Cells(1, 10).Value
        End With

        itemTexto = codigo & " | " & descripcion & " | Talle " & talle & " | " & color & " | " & codBarra
        cmb.AddItem itemTexto
    Next i
End Sub

Function ObtenerIDClientePorNombre(nombreCompleto As String) As String
    Dim partes() As String

    partes = Split(nombreCompleto, "|")
    If UBound(partes) >= 1 Then
        ObtenerIDClientePorNombre = Trim(Replace(partes(1), "ID", ""))
    Else
        ObtenerIDClientePorNombre = ""
    End If
End Function


Sub CalcularTotalDetalle()
    Dim i As Long
    Dim total As Double

    For i = 0 To lstDetalle.ListCount - 1
        total = total + Val(lstDetalle.List(i, 6)) ' Col 6 = total sin descuento
    Next i

    txtTotalBruto.Value = Format(total, "#,##0")

    ' Llamar siempre al cálculo de total neto
    Call CalcularTotalConDescuento
End Sub

Private Sub btnActualizarCantidad_Click()
    Dim fila As Long
    Dim nuevaCantidad As Double
    Dim precio As Double

    fila = lstDetalle.ListIndex
    If fila < 0 Then
        MsgBox "Seleccioná un producto de la lista para modificar la cantidad.", vbExclamation
        Exit Sub
    End If

    If Not IsNumeric(txtCantidad.Value) Then
        MsgBox "La cantidad debe ser un número válido.", vbExclamation
        Exit Sub
    End If

    nuevaCantidad = CDbl(txtCantidad.Value)
    precio = CDbl(lstDetalle.List(fila, 5)) ' Precio

    lstDetalle.List(fila, 4) = nuevaCantidad
    lstDetalle.List(fila, 6) = nuevaCantidad * precio

    Call CalcularTotalDetalle
End Sub

Sub CalcularTotalConDescuento()
    Dim totalBruto As Double
    Dim descuentoPorc As Double
    Dim totalNeto As Double
    Dim totalRedondeado As Double

    If IsNumeric(txtTotalBruto.Value) Then
        totalBruto = CDbl(NumeroLimpio(txtTotalBruto.Value))
    Else
        totalBruto = 0
    End If

    If IsNumeric(txtDescuento.Value) Then
        descuentoPorc = CDbl(txtDescuento.Value) / 100
    Else
        descuentoPorc = 0
    End If

    totalNeto = totalBruto * (1 - descuentoPorc)

    ' Solo redondear si hay descuento
    If descuentoPorc > 0 Then
        totalRedondeado = Application.WorksheetFunction.Floor(totalNeto, 1000)
    Else
        totalRedondeado = totalNeto
    End If

    txtTotalNeto.Value = Format(totalRedondeado, "#,##0")

    ' ?? Asignar directamente a txtMontoPago también
    txtMontoPago.Value = Format(totalRedondeado, "#,##0")

    ' ?? Limpiar o desactivar el segundo medio de pago
    txtMontoPago1.Value = "0"
End Sub

Private Sub txtDescuento_Change()
    Call CalcularTotalConDescuento
Exit Sub

End Sub

Function ObtenerCategoria(codigo As String, talle As String, color As String) As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Stock")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("Stock")
    Dim i As Long

    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            If .Cells(1, 1).Value = codigo And _
               .Cells(1, 9).Value = talle And _
               .Cells(1, 10).Value = color Then
                ObtenerCategoria = .Cells(1, 8).Value
                Exit Function
            End If
        End With
    Next i
    ObtenerCategoria = "Sin categoría"
End Function


