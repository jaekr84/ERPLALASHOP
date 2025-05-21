Attribute VB_Name = "modCaja"
' Módulo: Gestión de Caja

' ——————————————
' Comprueba si ya hay una apertura de caja sin cierre para hoy
' ——————————————
Public Function EsCajaAbiertaHoy() As Boolean
    Dim ws   As Worksheet
    Dim tbl  As ListObject
    Dim lr   As ListRow

    On Error GoTo NoExiste
    Set ws = ThisWorkbook.Sheets("Caja")
    Set tbl = ws.ListObjects("tblCaja")
    If tbl.ListRows.Count = 0 Then GoTo NoExiste

    For Each lr In tbl.ListRows
        ' Columna 1 = Fecha, Columna 5 = MontoCierre
        If lr.Range.Cells(1, 1).Value = Date _
           And Trim(lr.Range.Cells(1, 5).Value) = "" Then
            EsCajaAbiertaHoy = True
            Exit Function
        End If
    Next lr

NoExiste:
    EsCajaAbiertaHoy = False
End Function


' ——————————————
' Abre la caja: registra un movimiento de apertura para cada medio de pago
' Si se pasa montoEfectivo, lo usa; si no, lo pide con InputBox
' ——————————————
Public Sub AbrirCaja(Optional ByVal montoEfectivo As Variant)
    Dim wsCaja   As Worksheet
    Dim tblCaja  As ListObject
    Dim wsMed    As Worksheet
    Dim lastRow  As Long
    Dim lr       As ListRow
    Dim medio    As String
    Dim horaNow  As String
    Dim inicial  As Double
    Dim i        As Long

    ' Validar que no haya ya caja abierta hoy
    If EsCajaAbiertaHoy() Then
        MsgBox "Ya existe una caja abierta para hoy.", vbExclamation
        Exit Sub
    End If

    Set wsCaja = ThisWorkbook.Sheets("Caja")
    Set tblCaja = wsCaja.ListObjects("tblCaja")
    Set wsMed = ThisWorkbook.Sheets("MediosPago")

    horaNow = Format(Time, "hh:mm:ss")
    lastRow = wsMed.Cells(wsMed.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        medio = Trim(wsMed.Cells(i, "A").Value)
        If medio <> "" Then
            If UCase(medio) = "EFECTIVO" Then
                If IsMissing(montoEfectivo) Then
                    inicial = CDbl(InputBox("Ingrese el efectivo inicial para apertura de caja:", _
                                             "Apertura de Caja"))
                Else
                    inicial = CDbl(montoEfectivo)
                End If
            Else
                inicial = 0
            End If

            Set lr = tblCaja.ListRows.Add
            With lr.Range
                .Cells(1, 1).Value = Date                ' Fecha
                .Cells(1, 2).Value = horaNow             ' HoraApertura
                .Cells(1, 3).Value = medio               ' MedioPago
                .Cells(1, 4).Value = inicial             ' MontoInicial
                .Cells(1, 5).Value = ""                  ' MontoCierre
                .Cells(1, 6).Value = ""                  ' Diferencia
                .Cells(1, 7).Value = Environ("Username") ' Usuario
                .Cells(1, 8).Value = "Apertura"          ' Tipo de operacin
            End With
        End If
    Next i

    MsgBox "Caja abierta correctamente.", vbInformation
End Sub


' ——————————————
' Puente para abrir caja con monto proporcionado desde formulario
' ——————————————
Public Sub AbrirCajaConMontoEfectivo(ByVal montoEfectivo As Double)
    AbrirCaja montoEfectivo
End Sub


