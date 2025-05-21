Attribute VB_Name = "Module1"
' --- Variables globales ---
Public modoDevActivo As Boolean
Public origenLlamada As String
Public clienteRecienCreado As String

' --- Alternar entre modo desarrollador y modo app ---
Sub AlternarModoDesarrollador()
    Dim ws As Worksheet

    If Not modoDevActivo Then
        Application.Visible = True
        For Each ws In ThisWorkbook.Worksheets
            ws.Visible = xlSheetVisible
        Next ws
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        ThisWorkbook.Windows(1).DisplayHeadings = True
        ThisWorkbook.Windows(1).DisplayGridlines = True
        MsgBox "Modo desarrollador activado.", vbInformation
        modoDevActivo = True
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        ThisWorkbook.Windows(1).DisplayHeadings = False
        ThisWorkbook.Windows(1).DisplayGridlines = False
        Application.Visible = False
        modoDevActivo = False
    End If
End Sub

' --- Funciones globales ---
Function ObtenerIDClientePorNombre(nombreCompleto As String) As String
    Dim partes() As String
    partes = Split(nombreCompleto, "|")
    If UBound(partes) >= 1 Then
        ObtenerIDClientePorNombre = Trim(Replace(partes(1), "ID", ""))
    Else
        ObtenerIDClientePorNombre = ""
    End If
End Function

Function NumeroLimpio(valorTexto As String) As Double
    NumeroLimpio = Val(Replace(valorTexto, ".", ""))
End Function

Function ObtenerNuevoComprobanteVenta() As String
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim nuevoNumero As Long, nuevoComprobante As String

    Set ws = ThisWorkbook.Sheets("Contadores")
    Set tbl = ws.ListObjects("tblCompVenta")

    For i = 1 To tbl.ListRows.Count
        If tbl.DataBodyRange(i, 1).Value = "UltimoComprobanteVenta" Then
            nuevoNumero = tbl.DataBodyRange(i, 2).Value + 1
            tbl.DataBodyRange(i, 2).Value = nuevoNumero
            ObtenerNuevoComprobanteVenta = "V" & Format(nuevoNumero, "0000000")
            Exit Function
        End If
    Next i

    MsgBox "No se encontró 'UltimoComprobanteVenta' en la tabla tblCompVenta.", vbCritical
    ObtenerNuevoComprobanteVenta = ""
End Function

Sub CargarClientesEnCombo(cmb As ComboBox)
    Dim ws As Worksheet, tbl As ListObject, i As Long
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
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim codigo As String, descripcion As String, talle As String, color As String
    Dim itemTexto As String

    Set ws = ThisWorkbook.Sheets("Stock")
    Set tbl = ws.ListObjects("Stock")
    cmb.Clear

    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            codigo = .Cells(1, 1).Value
            descripcion = .Cells(1, 2).Value
            talle = .Cells(1, 9).Value
            color = .Cells(1, 10).Value
        End With
        itemTexto = codigo & " | " & descripcion & " | Talle " & talle & " | " & color
        cmb.AddItem itemTexto
    Next i
End Sub

Sub CargarMediosPago(cmb As ComboBox)
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Set ws = ThisWorkbook.Sheets("MediosPago")
    Set tbl = ws.ListObjects("Tabla2")
    cmb.Clear
    For i = 1 To tbl.ListRows.Count
        cmb.AddItem tbl.DataBodyRange(i, 1).Value
    Next i
End Sub

Public Sub ImprimirTicketConWord(comprobante As String, _
    fecha As String, medioPago As String, _
    subtotal As Double, descuento As Double, total As Double, _
    detalles As Variant)

    Dim wdApp As Object, wdDoc As Object
    Dim i As Long
    Dim nombreNegocio As String, direccion As String, cuit As String
    Dim redes As String, leyendaCambio As String
    Dim iva As Double
    Dim descripcion As String, talle As String, color As String
    Dim cantidad As Long, precioUnitario As Currency

    ' Datos del negocio
    With ThisWorkbook.Sheets("DatosNegocio")
        nombreNegocio = .Range("B1").Value
        direccion = .Range("B2").Value
        cuit = .Range("B3").Value
        redes = .Range("B4").Value
        leyendaCambio = .Range("B5").Value
    End With

    iva = subtotal * 0.21

    ' Crear documento Word
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Add

    With wdDoc.PageSetup
        .TopMargin = wdApp.CentimetersToPoints(0.2)
        .BottomMargin = wdApp.CentimetersToPoints(0.2)
        .LeftMargin = wdApp.CentimetersToPoints(0.2)
        .RightMargin = wdApp.CentimetersToPoints(0.2)
        .PageWidth = wdApp.CentimetersToPoints(7)
        .PageHeight = wdApp.CentimetersToPoints(29.7)
    End With

    With wdDoc.Content
        .Font.Name = "Courier New"
        .Font.Size = 9
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.LineSpacingRule = 0
        .ParagraphFormat.Alignment = 1 ' Centrado

        .InsertAfter nombreNegocio & vbCrLf
        .InsertAfter direccion & vbCrLf
        .InsertAfter "CUIT: " & cuit & vbCrLf
        .InsertAfter redes & vbCrLf
        .InsertAfter "Fecha: " & fecha & vbCrLf
        .InsertAfter "Comprobante: " & comprobante & vbCrLf
        .InsertAfter String(30, "-") & vbCrLf
        .InsertAfter "Detalle de productos" & vbCrLf
        .InsertAfter String(30, "-") & vbCrLf
        .ParagraphFormat.Alignment = 0 ' Alineación a la izquierda
    End With

For i = 0 To UBound(detalles, 1)
    descripcion = detalles(i, 0)
    talle = detalles(i, 1)
    color = detalles(i, 2)
    cantidad = detalles(i, 3)
    precioUnitario = detalles(i, 4)

    With wdDoc.Paragraphs.last.Range
        .InsertAfter descripcion & " " & talle & " " & color
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
        End With
        .InsertParagraphAfter
    End With

    With wdDoc.Paragraphs.last.Range
        .InsertAfter "x " & cantidad & " a $" & Format(precioUnitario, "#,##0") & " = $" & Format(cantidad * precioUnitario, "#,##0")
        With .ParagraphFormat
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LineSpacingRule = wdLineSpaceSingle
        End With
        .InsertParagraphAfter
    End With
Next i


    With wdDoc.Content
        .InsertAfter String(30, "-") & vbCrLf
        .InsertAfter "Subtotal: $" & Format(subtotal, "#,##0") & vbCrLf
        .InsertAfter "Descuento: $" & Format(descuento, "#,##0") & vbCrLf
        .InsertAfter "IVA (21%): $" & Format(iva, "#,##0") & vbCrLf
        .InsertAfter "TOTAL: $" & Format(total, "#,##0") & vbCrLf
        .InsertAfter "Pago con: " & medioPago & vbCrLf
        .InsertAfter String(30, "-") & vbCrLf

        If Trim(leyendaCambio) <> "" Then
            .InsertAfter leyendaCambio & vbCrLf
        End If
    End With

    wdDoc.PrintOut Background:=False
    wdDoc.Close False
    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

Public Sub ImprimirTicketDeCambioConWord(numeroComprobante As String, fecha As String)
    Dim wdApp As Object, wdDoc As Object
    Dim nombreNegocio As String, direccion As String, cuit As String
    Dim redes As String, leyendaCambio As String

    ' Obtener datos del negocio desde la hoja DatosNegocio
    With ThisWorkbook.Sheets("DatosNegocio")
        nombreNegocio = .Range("B1").Value
        direccion = .Range("B2").Value
        cuit = .Range("B3").Value
        redes = .Range("B4").Value
        leyendaCambio = .Range("B5").Value
    End With

    ' Crear documento Word
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Add

    ' Configurar página para ticket térmico
    With wdDoc.PageSetup
        .TopMargin = wdApp.CentimetersToPoints(0.2)
        .BottomMargin = wdApp.CentimetersToPoints(0.2)
        .LeftMargin = wdApp.CentimetersToPoints(0.2)
        .RightMargin = wdApp.CentimetersToPoints(0.2)
        .PageWidth = wdApp.CentimetersToPoints(7)
        .PageHeight = wdApp.CentimetersToPoints(29.7)
    End With

    ' Escribir contenido del ticket
    With wdDoc.Content
        .Font.Name = "Courier New"
        .Font.Size = 9
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.LineSpacingRule = 0
        .ParagraphFormat.Alignment = 1 ' Centrado

        .InsertAfter nombreNegocio & vbCrLf
        .InsertAfter direccion & vbCrLf
        .InsertAfter "CUIT: " & cuit & vbCrLf
        .InsertAfter redes & vbCrLf
        .InsertAfter "Fecha: " & fecha & vbCrLf
        .InsertAfter "Comprobante: " & numeroComprobante & vbCrLf
        .InsertAfter String(30, "-") & vbCrLf
        .InsertAfter "TICKET DE CAMBIO" & vbCrLf
        .InsertAfter String(30, "-") & vbCrLf
        .InsertAfter vbCrLf

        If Trim(leyendaCambio) <> "" Then
            .InsertAfter leyendaCambio & vbCrLf
        Else
            .InsertAfter "Condiciones de cambio no definidas." & vbCrLf
        End If
    End With

    ' Imprimir directamente y cerrar
    wdDoc.PrintOut Background:=False
    wdDoc.Close False
    wdApp.Quit

    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

Public Sub ActualizarTotalDiaEn(control As MSForms.label)
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim totalHoy As Double
    Dim hoy As Date

    Set ws = ThisWorkbook.Sheets("RegMediosPago")
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    hoy = Date

    totalHoy = 0
    For i = 2 To ultimaFila
        If IsDate(ws.Cells(i, 1).Value) Then
            If ws.Cells(i, 1).Value = hoy Then
                totalHoy = totalHoy + Val(ws.Cells(i, 7).Value)
            End If
        End If
    Next i

    control.Caption = "$ " & Format(totalHoy, "#,##0")
End Sub

Public Sub ActualizarTotalSemanaEn(control As MSForms.label)
    Dim ws As Worksheet, hoy As Date, inicioSemana As Date
    Dim i As Long, total As Double

    Set ws = ThisWorkbook.Sheets("RegMediosPago")
    hoy = Date
    inicioSemana = hoy - Weekday(hoy, vbMonday) + 1 ' lunes

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(i, 1).Value) Then
            If ws.Cells(i, 1).Value >= inicioSemana And ws.Cells(i, 1).Value <= hoy Then
                total = total + Val(ws.Cells(i, 7).Value)
            End If
        End If
    Next i

    control.Caption = "$ " & Format(total, "#,##0")
End Sub

Public Sub ActualizarTotalMesEn(control As MSForms.label)
    Dim ws As Worksheet, hoy As Date, inicioMes As Date
    Dim i As Long, total As Double

    Set ws = ThisWorkbook.Sheets("RegMediosPago")
    hoy = Date
    inicioMes = DateSerial(Year(hoy), Month(hoy), 1)

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(i, 1).Value) Then
            If ws.Cells(i, 1).Value >= inicioMes And ws.Cells(i, 1).Value <= hoy Then
                total = total + Val(ws.Cells(i, 7).Value)
            End If
        End If
    Next i

    control.Caption = "$ " & Format(total, "#,##0")
End Sub

Public Sub ActualizarTotalAnioEn(control As MSForms.label)
    Dim ws As Worksheet, hoy As Date, inicioAnio As Date
    Dim i As Long, total As Double

    Set ws = ThisWorkbook.Sheets("RegMediosPago")
    hoy = Date
    inicioAnio = DateSerial(Year(hoy), 1, 1)

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsDate(ws.Cells(i, 1).Value) Then
            If ws.Cells(i, 1).Value >= inicioAnio And ws.Cells(i, 1).Value <= hoy Then
                total = total + Val(ws.Cells(i, 7).Value)
            End If
        End If
    Next i

    control.Caption = "$ " & Format(total, "#,##0")
End Sub

Public Sub ForzarFocoEnCodigoBarra()
    On Error Resume Next
    frmNuevaVenta.txtCodigoBarra.SetFocus
End Sub

