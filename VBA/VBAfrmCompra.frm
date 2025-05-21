VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCompra 
   Caption         =   "UserForm1"
   ClientHeight    =   9960.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12045
   OleObjectBlob   =   "VBAfrmCompra.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAgregar_Click()
    Dim modo As String
    Dim wsStock   As Worksheet
    Dim tblStock  As ListObject
    Dim r         As ListRow
    Dim codigo    As String, descripcion As String, talle As String
    Dim color     As String, codBarra As String
    Dim costo     As Double
    Dim lidx      As Long, foundIdx As Long
    Dim codigoBuscar As String
    Dim exists    As Boolean
    Dim currentQty As Long

    modo = Me.btnAgregar.Tag
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = wsStock.ListObjects("Stock")

    If modo = "single" Then
        ' ————————— Single —————————
        If Me.cmbArticulo.ListIndex = -1 Then
            MsgBox "Primero seleccioná un artículo.", vbExclamation
            Exit Sub
        End If

        ' Leo del combo
        codigo = Me.cmbArticulo.Value
        descripcion = Me.cmbArticulo.List(Me.cmbArticulo.ListIndex, 1)
        talle = Me.cmbArticulo.List(Me.cmbArticulo.ListIndex, 2)
        color = Me.cmbArticulo.List(Me.cmbArticulo.ListIndex, 3)
        codBarra = Me.cmbArticulo.List(Me.cmbArticulo.ListIndex, 4)
        ' Extraigo costo de la hoja
        For Each r In tblStock.ListRows
            With r.Range
                If .Cells(1, 1).Value = codigo And _
                   .Cells(1, 9).Value = talle And _
                   .Cells(1, 10).Value = color Then
                    costo = .Cells(1, 3).Value
                    Exit For
                End If
            End With
        Next r

        ' Agrego o acumulo en lstDetalle
        With Me.lstDetalle
            exists = False
            For lidx = 0 To .ListCount - 1
                If .List(lidx, 0) = codigo And _
                   .List(lidx, 2) = talle And _
                   .List(lidx, 3) = color Then
                    exists = True: foundIdx = lidx: Exit For
                End If
            Next lidx

            If exists Then
                currentQty = Val(.List(foundIdx, 5)) + 1
                .List(foundIdx, 5) = currentQty
                .List(foundIdx, 7) = currentQty * costo
            Else
                lidx = .ListCount
                .AddItem codigo
                .List(lidx, 1) = descripcion
                .List(lidx, 2) = talle
                .List(lidx, 3) = color
                .List(lidx, 4) = codBarra
                .List(lidx, 5) = 1             ' Cantidad inicial
                .List(lidx, 6) = costo
                .List(lidx, 7) = costo         ' Subtotal = 1 * costo
            End If
        End With

    ElseIf modo = "bulk" Then
        ' ————————— Bulk —————————
        codigoBuscar = Trim(Me.txtCargaTodas.Value)
        If codigoBuscar = "" Then
            MsgBox "Ingresá el código para carga masiva.", vbExclamation
            Exit Sub
        End If

        With Me.lstDetalle
            For Each r In tblStock.ListRows
                If CStr(r.Range.Cells(1, 1).Value) = codigoBuscar Then
                    descripcion = r.Range.Cells(1, 2).Value
                    talle = r.Range.Cells(1, 9).Value
                    color = r.Range.Cells(1, 10).Value
                    codBarra = r.Range.Cells(1, 7).Value
                    costo = r.Range.Cells(1, 3).Value

                    ' ¿Ya existe esa variante?
                    exists = False
                    For lidx = 0 To .ListCount - 1
                        If .List(lidx, 0) = codigoBuscar And _
                           .List(lidx, 2) = talle And _
                           .List(lidx, 3) = color Then
                            exists = True: foundIdx = lidx: Exit For
                        End If
                    Next lidx

                    If exists Then
                        currentQty = Val(.List(foundIdx, 5)) + 1
                        .List(foundIdx, 5) = currentQty
                        .List(foundIdx, 7) = currentQty * costo
                    Else
                        lidx = .ListCount
                        .AddItem codigoBuscar
                        .List(lidx, 1) = descripcion
                        .List(lidx, 2) = talle
                        .List(lidx, 3) = color
                        .List(lidx, 4) = codBarra
                        .List(lidx, 5) = 1             ' Cantidad inicial
                        .List(lidx, 6) = costo
                        .List(lidx, 7) = costo         ' Subtotal
                    End If
                End If
            Next r
            Me.txtCargaTodas.Value = ""  ' vuelve a single
        End With

    End If  ' fin If modo

    ' Recalcular totales
    CalcularTotalCantidad1   ' actualiza lblTotalCant
    UpdateTotals             ' actualiza lblSubTotal y lblTotal
End Sub





Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnNuevoProveedor_Click()
    Call frmNuevoProveedor.Show
End Sub




Private Sub CommandButton3_Click()
    frmAltaArticulo.Show
End Sub

Private Sub btnEliminar_Click()
    Dim idx As Long
    
    ' Obtener índice de la línea seleccionada
    idx = Me.lstDetalle.ListIndex
    If idx = -1 Then
        MsgBox "Seleccioná una línea para eliminar.", vbExclamation
        Exit Sub
    End If

    ' Eliminar la línea del ListBox
    Me.lstDetalle.RemoveItem idx

    ' Recalcular totales
     Call CalcularTotalCantidad1   ' actualiza lblTotalCant
     Call UpdateTotals             ' actualiza lblSubTotal y lblTotal
End Sub

Private Sub btnConfirmar_Click()
    Dim wsCompras   As Worksheet, tblCompras As ListObject
    Dim wsMov       As Worksheet, tblMov     As ListObject
    Dim wsCnt       As Worksheet, tblCnt     As ListObject
    Dim rowC        As ListRow, rowM         As ListRow
    
    Dim fechaCompra As Date
    Dim compID      As Long
    Dim compNumText As String
    Dim provID      As String
    
    ' Variables para detalle
    Dim codigo_i    As String
    Dim desc_i      As String
    Dim talle_i     As String
    Dim color_i     As String
    Dim qty_i       As Long
    Dim cost_i      As Double
    Dim sub_i       As Double
    
    Dim i           As Long
    
    ' ——— Datos de encabezado ———
    fechaCompra = CDate(Me.lblFecha.Caption)
    compNumText = Mid(Me.lblComprobante.Caption, 2)
    compID = CLng(compNumText)
    provID = Me.cmbProveedor.List(Me.cmbProveedor.ListIndex, 0)
    
    ' ——— Referencias a tablas ———
    Set wsCompras = ThisWorkbook.Sheets("Compras")
    Set tblCompras = wsCompras.ListObjects("tblCompras")
    
    Set wsMov = ThisWorkbook.Sheets("MovimientosStock")
    Set tblMov = wsMov.ListObjects("Movimientos")
    
    Set wsCnt = ThisWorkbook.Sheets("Contadores")
    Set tblCnt = wsCnt.ListObjects("tblCompCompra")
    
    ' ——— Volcar líneas de detalle ———
    With Me.lstDetalle
        For i = 0 To .ListCount - 1
            ' 1) Extraigo de lstDetalle a variables
            codigo_i = .List(i, 0)
            desc_i = .List(i, 1)
            talle_i = .List(i, 2)
            color_i = .List(i, 3)
            qty_i = CLng(.List(i, 5))
            cost_i = CDbl(.List(i, 6))
            sub_i = CDbl(.List(i, 7))
            
            ' 2) Registro en Compras
            Set rowC = tblCompras.ListRows.Add
            With rowC.Range
                .Cells(1, 1).Value = fechaCompra
                .Cells(1, 2).Value = provID
                .Cells(1, 3).Value = codigo_i
                .Cells(1, 4).Value = desc_i
                .Cells(1, 5).Value = talle_i
                .Cells(1, 6).Value = color_i
                .Cells(1, 7).Value = qty_i
                .Cells(1, 8).Value = cost_i
                .Cells(1, 9).Value = sub_i
                .Cells(1, 10).Value = Me.lblComprobante.Caption
            End With
            
            ' 3) Registro en MovimientosStock
            Set rowM = tblMov.ListRows.Add
            With rowM.Range
                .Cells(1, 1).Value = fechaCompra
                .Cells(1, 2).Value = codigo_i
                .Cells(1, 3).Value = desc_i
                .Cells(1, 4).Value = talle_i
                .Cells(1, 5).Value = color_i
                .Cells(1, 6).Value = qty_i
                .Cells(1, 7).Value = "Compra"
            End With
        Next i
    End With
    
    ' ——— Actualizo contador ———
    tblCnt.DataBodyRange(1, 2).Value = compID
    
      ' ————— ACTUALIZAR STOCK EN “Stock” —————
    Dim r As ListRow
    Dim codBarra As String
    
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = wsStock.ListObjects("Stock")

    With Me.lstDetalle
        For i = 0 To .ListCount - 1
            ' Extraigo el codBarra y la cantidad de la fila i
            codBarra = .List(i, 4)       ' columna 4 en tu lstDetalle
            qty_i = CLng(.List(i, 5))    ' columna 5 = cantidad

            ' Recorro la tabla Stock
            For Each r In tblStock.ListRows
                If r.Range.Cells(1, 7).Value = codBarra Then
                    ' Sumo la cantidad (compra) o resto si qty_i es negativo
                    r.Range.Cells(1, 6).Value = r.Range.Cells(1, 6).Value + qty_i
                    Exit For
                End If
            Next r
        Next i
    End With


    MsgBox "Compra " & Me.lblComprobante.Caption & " registrada correctamente.", vbInformation
    Call HacerBackup
    Unload Me
End Sub



Private Sub lblTotalCant_Click()

End Sub

Private Sub txtDescuento_Change()
    Call UpdateTotals
End Sub

Private Sub UserForm_Initialize()
    Dim wsCnt     As Worksheet
    Dim tblCnt    As ListObject
    Dim nextComp  As Long
    
    Dim wsProv    As Worksheet
    Dim tblProv   As ListObject
    
    Dim wsStock   As Worksheet
    Dim tblStock  As ListObject
    
    Dim r         As ListRow

    ' ————————————— Generar Comprobante —————————————
    Set wsCnt = ThisWorkbook.Sheets("Contadores")
    Set tblCnt = wsCnt.ListObjects("tblCompCompra")
    nextComp = tblCnt.DataBodyRange(1, 2).Value + 1
    Me.lblComprobante.Caption = "C" & Format(nextComp, "000000000")

    ' ————————————— Mostrar Fecha —————————————
    Me.lblFecha.Caption = Format(Date, "dd/mm/yyyy")

    ' ——— Carga de Proveedores ———
    Set wsProv = ThisWorkbook.Sheets("Proveedores")
    Set tblProv = wsProv.ListObjects("tblProveedores")
    With Me.cmbProveedor
        .Clear
        For Each r In tblProv.ListRows
            .AddItem r.Range.Cells(1, 1).Value        ' Col 0 = ID
            .List(.ListCount - 1, 1) = r.Range.Cells(1, 2).Value  ' Col 1 = Nombre
        Next r
    End With

    ' ——— Carga de Artículos ———
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = wsStock.ListObjects("Stock")
    With Me.cmbArticulo
        .Clear
        For Each r In tblStock.ListRows
            .AddItem r.Range.Cells(1, 1).Value        ' Col 0 = Código
            .List(.ListCount - 1, 1) = r.Range.Cells(1, 2).Value  ' Col 1 = Descripción
            .List(.ListCount - 1, 2) = r.Range.Cells(1, 9).Value  ' Col 2 = Talle
            .List(.ListCount - 1, 3) = r.Range.Cells(1, 10).Value ' Col 3 = Color
            .List(.ListCount - 1, 4) = r.Range.Cells(1, 7).Value  ' Col 4 = Cód. Barra (oculto)
            .List(.ListCount - 1, 5) = r.Range.Cells(1, 3).Value  ' Col 5 = Costo     (oculto)
        Next r
    End With

    ' ——— Inicializar modo del botón Agregar ———
    Me.btnAgregar.Tag = "single"

    ' ——— Limpiar controles de totales ———
    Me.txtCargaTodas.Value = ""
    Me.lblSubtotal.Caption = "$0"
    Me.lblTotal.Caption = "$0"
    Me.lblTotalCant.Caption = "0"
End Sub




Private Sub txtCargaTodas_Change()
    If Len(Me.txtCargaTodas.Text) > 0 Then
        Me.btnAgregar.Tag = "bulk"
    Else
        Me.btnAgregar.Tag = "single"
    End If
End Sub

Private Sub CalcularTotalCantidad()
    Dim i As Long, total As Long
    total = 0

    With Me.lstDetalle
        For i = 0 To .ListCount - 1
            total = total + Val(.List(i, 5))   ' Columna 5 = “Cantidad”
        Next i
    End With

    Me.lblTotalCant.Caption = total
End Sub

Public Sub RefreshArticulos()
    Dim wsStock  As Worksheet
    Dim tblStock As ListObject
    Dim r        As ListRow
    
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = wsStock.ListObjects("Stock")
    
    With Me.cmbArticulo
        .Clear
        For Each r In tblStock.ListRows
            .AddItem r.Range.Cells(1, 1).Value             ' Código
            .List(.ListCount - 1, 1) = r.Range.Cells(1, 2).Value   ' Descripción
            .List(.ListCount - 1, 2) = r.Range.Cells(1, 9).Value   ' Talle
            .List(.ListCount - 1, 3) = r.Range.Cells(1, 10).Value  ' Color
            .List(.ListCount - 1, 4) = r.Range.Cells(1, 7).Value   ' Cód. Barra (oculto)
            .List(.ListCount - 1, 5) = r.Range.Cells(1, 3).Value   ' Costo     (oculto)
        Next r
    End With
End Sub

' ————————————


' ————————————
' DblClick en lstDetalle para editar cantidad
' ————————————
Private Sub lstDetalle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim idx As Long
    Dim nuevaCant As Variant
    idx = Me.lstDetalle.ListIndex
    If idx = -1 Then Exit Sub

    nuevaCant = InputBox( _
        Prompt:="Ingrese la nueva cantidad:", _
        Title:="Editar cantidad", _
        Default:=Me.lstDetalle.List(idx, 5) _
    )
    If nuevaCant = vbNullString Then Exit Sub
    If Not IsNumeric(nuevaCant) Then
        MsgBox "Cantidad inválida.", vbExclamation: Exit Sub
    End If

    Me.lstDetalle.List(idx, 5) = Val(nuevaCant)
    Me.lstDetalle.List(idx, 7) = Val(nuevaCant) * Val(Me.lstDetalle.List(idx, 6))

    Call CalcularTotalCantidad1
    Call UpdateTotals
End Sub

' ——————————————
' Calcula la cantidad total de unidades
' ——————————————
Private Sub CalcularTotalCantidad1()
    Dim i As Long
    Dim totalCant As Long
    
    totalCant = 0
    With Me.lstDetalle
        For i = 0 To .ListCount - 1
            totalCant = totalCant + CLng(.List(i, 5))   ' Col 5 = Cantidad
        Next i
    End With
    
    Me.lblTotalCant.Caption = totalCant

End Sub

' ——————————————
' Recalcula Subtotal, aplica Descuento% y muestra Total
' ——————————————
Private Sub UpdateTotals()
    Dim i As Long
    Dim subtotal As Double
    Dim total    As Double
    Dim descPct  As Double
    
    ' 1) Sumar todos los Subtotales (col 7)
    subtotal = 0
    With Me.lstDetalle
        For i = 0 To .ListCount - 1
            subtotal = subtotal + CDbl(.List(i, 7))   ' Col 7 = Subtotal línea
        Next i
    End With
    
    ' Mostrar Subtotal formateado
    Me.lblSubtotal.Caption = " Subtotal " & "$" & Format(subtotal, "#,##0")
    
    ' 2) Leer Descuento%
    If Me.txtDescuento.Value <> "" And IsNumeric(Me.txtDescuento.Value) Then
        descPct = CDbl(Me.txtDescuento.Value) / 100
    Else
        descPct = 0
    End If
    
    ' 3) Calcular Total con descuento
    total = subtotal * (1 - descPct)
    Me.lblTotal.Caption = " Total " & "$" & Format(total, "#,##0")
End Sub

