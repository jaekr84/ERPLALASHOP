Attribute VB_Name = "modReportes"
Sub CargarTopCategorias(frm As Object)
    Dim ws As Worksheet
    Dim ultFila As Long, i As Long
    Dim fechaVenta As Date, categoria As String, cantidad As Long
    Dim hoy As Date, inicioSemana As Date, inicioMes As Date, inicioAnio As Date
    Dim dictDia As Object, dictSemana As Object, dictMes As Object, dictAnio As Object

    Set dictDia = CreateObject("Scripting.Dictionary")
    Set dictSemana = CreateObject("Scripting.Dictionary")
    Set dictMes = CreateObject("Scripting.Dictionary")
    Set dictAnio = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Sheets("Ventas")
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hoy = Date
    inicioSemana = hoy - Weekday(hoy, vbMonday) + 1
    inicioMes = DateSerial(Year(hoy), Month(hoy), 1)
    inicioAnio = DateSerial(Year(hoy), 1, 1)

    For i = 2 To ultFila
        If IsDate(ws.Cells(i, 1).Value) Then
            fechaVenta = ws.Cells(i, 1).Value
            categoria = Trim(ws.Cells(i, 11).Value)
            cantidad = Val(ws.Cells(i, 4).Value)

            If categoria <> "" And cantidad > 0 Then
                If fechaVenta = hoy Then dictDia(categoria) = dictDia(categoria) + cantidad
                If fechaVenta >= inicioSemana Then dictSemana(categoria) = dictSemana(categoria) + cantidad
                If fechaVenta >= inicioMes Then dictMes(categoria) = dictMes(categoria) + cantidad
                If fechaVenta >= inicioAnio Then dictAnio(categoria) = dictAnio(categoria) + cantidad
            End If
        End If
    Next i

    ' Cargar y ordenar cada diccionario
    CargarListaOrdenada dictDia, frm.lstDia
    CargarListaOrdenada dictSemana, frm.lstSemana
    CargarListaOrdenada dictMes, frm.lstMes
    CargarListaOrdenada dictAnio, frm.lstAnio
End Sub

Sub CargarListaOrdenada(dict As Object, lst As MSForms.ListBox)
    Dim arr(), i As Long, j As Long, tempCat As String, tempCant As Long

    ReDim arr(1 To dict.Count, 1 To 2)
    i = 0
    For Each k In dict.Keys
        i = i + 1
        arr(i, 1) = k
        arr(i, 2) = dict(k)
    Next k

    ' Ordenar de mayor a menor por cantidad
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i, 2) < arr(j, 2) Then
                tempCat = arr(i, 1): tempCant = arr(i, 2)
                arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2)
                arr(j, 1) = tempCat: arr(j, 2) = tempCant
            End If
        Next j
    Next i

 
    For i = 1 To Application.Min(10, UBound(arr))
        lst.AddItem arr(i, 1)
        lst.List(lst.ListCount - 1, 1) = arr(i, 2)
    Next i
End Sub

Sub CargarTopProductos(rango As String, lst As MSForms.ListBox)
    Dim ws As Worksheet, tbl As ListObject
    Dim dict As Object
    Dim i As Long
    Dim clave As String, cantidad As Double
    Dim fechaVenta As Date
    Dim fechaDesde As Date
    Dim productosOrdenados As Variant
    Dim fila As Variant

    Set ws = ThisWorkbook.Sheets("ventas")
    Set tbl = ws.ListObjects("Tabla1")
    Set dict = CreateObject("Scripting.Dictionary")

    ' Definir fecha desde según el rango
    Select Case UCase(rango)
        Case "DIA": fechaDesde = Date
        Case "SEM": fechaDesde = Date - 6
        Case "MES": fechaDesde = DateSerial(Year(Date), Month(Date), 1)
        Case "ANIO": fechaDesde = DateSerial(Year(Date), 1, 1)
    End Select

    ' Agrupar por código de artículo (columna 2), sumar cantidades (columna 4)
    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            fechaVenta = .Cells(1, 1).Value
            If IsDate(fechaVenta) And fechaVenta >= fechaDesde Then
                clave = Trim(.Cells(1, 2).Value) ' Código del artículo
                cantidad = Val(.Cells(1, 4).Value) ' Cantidad
                If dict.exists(clave) Then
                    dict(clave) = dict(clave) + cantidad
                Else
                    dict.Add clave, cantidad
                End If
            End If
        End With
    Next i

    ' Ordenar por cantidad vendida descendente
    productosOrdenados = SortDictionaryByValue(dict)
    
    For Each fila In productosOrdenados
        If lst.ListCount < 30 Then
            lst.AddItem fila(0)
            lst.List(lst.ListCount - 1, 1) = fila(1)
        Else
            Exit For
        End If
    Next fila

End Sub

Function SortDictionaryByValue(dict As Object) As Variant
    Dim arr() As Variant, temp As Variant
    Dim i As Long, j As Long

    ReDim arr(0 To dict.Count - 1)
    i = 0
    For Each clave In dict.Keys
        arr(i) = Array(clave, dict(clave))
        i = i + 1
    Next

    For i = 0 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)(1) < arr(j)(1) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i

    SortDictionaryByValue = arr
End Function
Sub AbrirDetalleProductoDesdeLista(lst As MSForms.ListBox)
    If lst.ListIndex = -1 Then Exit Sub

    Dim codigo As String
    codigo = lst.List(lst.ListIndex, 0)

    On Error Resume Next ' Evita error si el form no tiene MostrarDetalle
    frmDetalleProducto.MostrarDetalle codigo
    On Error GoTo 0

    frmDetalleProducto.Show
End Sub

Sub CargarTopTalles(frm As Object)
    Dim ws As Worksheet
    Dim ultFila As Long, i As Long
    Dim fechaVenta As Date, talle As String, cantidad As Long
    Dim hoy As Date, inicioSemana As Date, inicioMes As Date, inicioAnio As Date
    Dim dictDia As Object, dictSemana As Object, dictMes As Object, dictAnio As Object

    Set dictDia = CreateObject("Scripting.Dictionary")
    Set dictSemana = CreateObject("Scripting.Dictionary")
    Set dictMes = CreateObject("Scripting.Dictionary")
    Set dictAnio = CreateObject("Scripting.Dictionary")

    Set ws = ThisWorkbook.Sheets("Ventas")
    ultFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hoy = Date
    inicioSemana = hoy - Weekday(hoy, vbMonday) + 1
    inicioMes = DateSerial(Year(hoy), Month(hoy), 1)
    inicioAnio = DateSerial(Year(hoy), 1, 1)

    For i = 2 To ultFila
        If IsDate(ws.Cells(i, 1).Value) Then
            fechaVenta = ws.Cells(i, 1).Value
            talle = Trim(ws.Cells(i, 9).Value) ' Columna J = Talle
            cantidad = Val(ws.Cells(i, 4).Value) ' Columna D = Cantidad

            If talle <> "" And cantidad > 0 Then
                If fechaVenta = hoy Then dictDia(talle) = dictDia(talle) + cantidad
                If fechaVenta >= inicioSemana Then dictSemana(talle) = dictSemana(talle) + cantidad
                If fechaVenta >= inicioMes Then dictMes(talle) = dictMes(talle) + cantidad
                If fechaVenta >= inicioAnio Then dictAnio(talle) = dictAnio(talle) + cantidad
            End If
        End If
    Next i

    ' Cargar y ordenar en los ListBox del form
    CargarListaOrdenada dictDia, frm.lstTalleDia
    CargarListaOrdenada dictSemana, frm.lstTalleSem
    CargarListaOrdenada dictMes, frm.lstTalleMes
    CargarListaOrdenada dictAnio, frm.lstTalleAnio
End Sub

Sub ActualizarTotalesVentasDashboard(frm As Object)
    Dim ws As Worksheet
    Dim i As Long
    Dim fechaVenta As Date
    Dim total As Double
    Dim hoy As Date, inicioSemana As Date, inicioMes As Date, inicioAnio As Date

    Set ws = ThisWorkbook.Sheets("Ventas")
    hoy = Date
    inicioSemana = hoy - Weekday(hoy, vbMonday) + 1
    inicioMes = DateSerial(Year(hoy), Month(hoy), 1)
    inicioAnio = DateSerial(Year(hoy), 1, 1)

    Dim totalDia As Double, totalSem As Double, totalMes As Double, totalAnio As Double

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(i, 1).Value) Then
            fechaVenta = ws.Cells(i, 1).Value
            total = Val(ws.Cells(i, 7).Value) ' Columna 7 = Total (verificá que sea la correcta)

            If fechaVenta = hoy Then totalDia = totalDia + total
            If fechaVenta >= inicioSemana Then totalSem = totalSem + total
            If fechaVenta >= inicioMes Then totalMes = totalMes + total
            If fechaVenta >= inicioAnio Then totalAnio = totalAnio + total
        End If
    Next i

    ' Mostrar en etiquetas del formulario
    frm.lblTotalDia.Caption = Format(totalDia, "$ #,##0")
    frm.lblTotalSem.Caption = Format(totalSem, "$ #,##0")
    frm.lblTotalMes.Caption = Format(totalMes, "$ #,##0")
    frm.lblTotalAnio.Caption = Format(totalAnio, "$ #,##0")
End Sub

