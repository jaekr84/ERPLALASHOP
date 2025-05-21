Attribute VB_Name = "modRotacion"
Sub AnalizarRotacionAlta(frm As Object)
    Dim wsMov As Worksheet, wsVentas As Worksheet, wsStock As Worksheet
    Dim tblMov As ListObject, tblStock As ListObject
    Dim dictUltimoIngreso As Object, dictVentas As Object, dictDescripcion As Object
    Dim fila As ListRow, i As Long
    Dim clave As String, fechaIngreso As Date, fechaVenta As Date
    Dim diasRotacion As Long, cantidadVendida As Long

    Dim desde As Date, hasta As Date
    On Error GoTo Validacion
    desde = CDate(frm.txtDesde.Value)
    hasta = CDate(frm.txtHasta.Value)
    On Error GoTo 0

    Set wsMov = ThisWorkbook.Sheets("MovimientosStock")
    Set wsVentas = ThisWorkbook.Sheets("Ventas")
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblMov = wsMov.ListObjects("Movimientos")
    Set tblStock = wsStock.ListObjects("Stock")

    Set dictUltimoIngreso = CreateObject("Scripting.Dictionary")
    Set dictVentas = CreateObject("Scripting.Dictionary")
    Set dictDescripcion = CreateObject("Scripting.Dictionary")

    Dim resultado(), temp As Variant
    Dim idx As Long: idx = 0
    ReDim resultado(1 To 1000, 1 To 8)

    ' Buscar ingresos de stock en el rango de fechas
    For Each fila In tblMov.ListRows
        If Trim(fila.Range.Cells(1, 7).Value) = "Compra" Then
            fechaIngreso = fila.Range.Cells(1, 1).Value
            If fechaIngreso >= desde And fechaIngreso <= hasta Then
                clave = fila.Range.Cells(1, 2).Value & "|" & fila.Range.Cells(1, 4).Value & "|" & fila.Range.Cells(1, 5).Value
                If Not dictUltimoIngreso.exists(clave) Or fechaIngreso > dictUltimoIngreso(clave) Then
                    dictUltimoIngreso(clave) = fechaIngreso
                    dictDescripcion(clave) = fila.Range.Cells(1, 3).Value ' Descripción
                End If
            End If
        End If
    Next fila

    ' Buscar ventas posteriores al ingreso
    For i = 2 To wsVentas.Cells(wsVentas.Rows.Count, 1).End(xlUp).Row
        clave = wsVentas.Cells(i, 2).Value & "|" & wsVentas.Cells(i, 10).Value & "|" & wsVentas.Cells(i, 11).Value
        If dictUltimoIngreso.exists(clave) Then
            If wsVentas.Cells(i, 1).Value >= dictUltimoIngreso(clave) Then
                dictVentas(clave & "|cant") = dictVentas(clave & "|cant") + wsVentas.Cells(i, 4).Value
                dictVentas(clave & "|ult") = wsVentas.Cells(i, 1).Value
            End If
        End If
    Next i

    ' Procesar los productos sin stock actual
    For Each fila In tblStock.ListRows
        clave = fila.Range.Cells(1, 1).Value & "|" & fila.Range.Cells(1, 9).Value & "|" & fila.Range.Cells(1, 10).Value
        If dictUltimoIngreso.exists(clave) Then
            If fila.Range.Cells(1, 6).Value = 0 Then
                cantidadVendida = dictVentas(clave & "|cant")
                fechaIngreso = dictUltimoIngreso(clave)
                fechaVenta = dictVentas(clave & "|ult")
                diasRotacion = fechaVenta - fechaIngreso

                If cantidadVendida > 0 Then
                    idx = idx + 1
                    resultado(idx, 1) = Split(clave, "|")(0)
                    resultado(idx, 2) = dictDescripcion(clave)
                    resultado(idx, 3) = Split(clave, "|")(1)
                    resultado(idx, 4) = Split(clave, "|")(2)
                    resultado(idx, 5) = cantidadVendida
                    resultado(idx, 6) = diasRotacion
                    resultado(idx, 7) = Format(fechaVenta, "dd/mm/yyyy")
                    resultado(idx, 8) = "?"
                End If
            End If
        End If
    Next fila

    ' Ordenar por días de rotación (menor a mayor)
    Dim j As Long, k As Long
    For i = 1 To idx - 1
        For j = i + 1 To idx
            If resultado(i, 6) > resultado(j, 6) Then
                For k = 1 To 8
                    temp = resultado(i, k)
                    resultado(i, k) = resultado(j, k)
                    resultado(j, k) = temp
                Next k
            End If
        Next j
    Next i

    ' Mostrar en ListBox
    With frm.lstCodigos
        .ColumnCount = 8
        .ColumnWidths = "50;130;40;50;50;80;80;60"
        For i = 1 To idx
            .AddItem resultado(i, 1)
            For j = 2 To 8
                .List(.ListCount - 1, j - 1) = resultado(i, j)
            Next j
        Next i
    End With
    Exit Sub

Validacion:
    MsgBox "Ingresá fechas válidas en ambos campos.", vbExclamation
End Sub

