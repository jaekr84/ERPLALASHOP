VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCaja 
   Caption         =   "UserForm7"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5640
   OleObjectBlob   =   "VBAfrmCaja.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub TextBoxEfectivoReal_Change()
    Dim wsCaja As Worksheet
    Dim tbl As ListObject
    Dim totalEfectivoVend As Double
    Dim efectivoReal As Double
    Dim efectivoInicial As Double
    Dim diferencia As Double
    Dim i As Long

    ' 1. Buscar total vendido en efectivo
    totalEfectivoVend = 0
    For i = 0 To ListBoxResumen.ListCount - 1
        If UCase(Trim(ListBoxResumen.List(i, 0))) = "EFECTIVO" Then
            If IsNumeric(ListBoxResumen.List(i, 1)) Then
                totalEfectivoVend = CDbl(ListBoxResumen.List(i, 1))
            End If
            Exit For
        End If
    Next i

    ' 2. Buscar efectivo inicial desde tabla
    efectivoInicial = 0
    Set wsCaja = ThisWorkbook.Sheets("Caja")
    Set tbl = wsCaja.ListObjects("tblCaja")

    For i = tbl.ListRows.Count To 1 Step -1
        With tbl.ListRows(i).Range
            If .Cells(1, 1).Value = Date And UCase(.Cells(1, 3).Value) = "EFECTIVO" Then
                efectivoInicial = .Cells(1, 4).Value
                Exit For
            End If
        End With
    Next i

    ' 3. Calcular diferencia si ingreso válido
    If IsNumeric(TextBoxEfectivoReal.Value) Then
        efectivoReal = CDbl(TextBoxEfectivoReal.Value)
        diferencia = efectivoReal - (totalEfectivoVend + efectivoInicial)
        LabelDiferencia.Caption = "Diferencia: " & Format(diferencia, "#,##0.00")
        
        ' Cambiar color
        If diferencia < 0 Then
            LabelDiferencia.ForeColor = RGB(192, 0, 0) ' rojo
        Else
            LabelDiferencia.ForeColor = RGB(0, 96, 0) ' verde oscuro
        End If
    Else
        LabelDiferencia.Caption = ""
        LabelDiferencia.ForeColor = RGB(0, 0, 0) ' negro por defecto
    End If
End Sub





Private Sub UserForm_Initialize()
    CargarMediosPago
    CargarResumenVentasDelDia
End Sub

Private Sub CargarMediosPago()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("MediosPago")
    
    ListBoxMediosPago.Clear
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ListBoxMediosPago.AddItem ws.Cells(i, 1).Value
    Next i
End Sub

Private Sub ListBoxMediosPago_Click()
    If UCase(ListBoxMediosPago.Value) = "EFECTIVO" Then
        TextBoxMontoInicial.Enabled = True
    Else
        TextBoxMontoInicial.Enabled = False
        TextBoxMontoInicial.Value = ""
    End If
End Sub

Private Sub btnAbrirCaja_Click()
    Dim monto As Double

    ' Validar que se ingresó un número
    If Not IsNumeric(Me.txtMontoApertura.Value) Or Me.txtMontoApertura.Value = "" Then
        MsgBox "Ingresá un monto de efectivo inicial válido.", vbExclamation
        Exit Sub
    End If

    monto = CDbl(Me.txtMontoApertura.Value)

    ' Llamar a la sub que pasa el monto a AbrirCaja
    Call AbrirCajaConMontoEfectivo(monto)

    ' Opcional: desactivar el botón o limpiar el campo
    Me.btnAbrirCaja.Enabled = False
    Me.txtMontoApertura.Enabled = False
    Unload Me
End Sub


Private Sub CargarResumenVentasDelDia()
    Dim wsVentas As Worksheet
    Dim dict As Object
    Dim i As Long, ultima As Long
    Dim fechaHoy As Date
    Dim medio As Variant
    Dim total As Double
    
    Set wsVentas = ThisWorkbook.Sheets("ventas")
    Set dict = CreateObject("Scripting.Dictionary")
    
    fechaHoy = Date
    ultima = wsVentas.Cells(wsVentas.Rows.Count, 1).End(xlUp).Row
    
    ' Recorrer ventas del día y acumular por medio de pago
    For i = 2 To ultima
        If wsVentas.Cells(i, 1).Value = fechaHoy Then
            medio = UCase(Trim(wsVentas.Cells(i, 8).Value)) ' Columna H: MedioPago
            total = wsVentas.Cells(i, 7).Value               ' Columna G: Total
            If dict.exists(medio) Then
                dict(medio) = dict(medio) + total
            Else
                dict.Add medio, total
            End If
        End If
    Next i
    
    ' Mostrar en ListBoxResumen sin formato
    ListBoxResumen.Clear
    Dim clave As Variant
    For Each clave In dict.Keys
        ListBoxResumen.AddItem clave
        ListBoxResumen.List(ListBoxResumen.ListCount - 1, 1) = dict(clave)
    Next clave
End Sub


Private Sub btnCerrarCaja_Click()
    Dim wsCaja As Worksheet
    Dim tbl As ListObject
    Dim i As Long, j As Long
    Dim medio As String
    Dim totalVend As Double
    Dim efectivoReal As Double

    If Not IsNumeric(TextBoxEfectivoReal.Value) Then
        MsgBox "Ingresá el efectivo real contado.", vbExclamation
        Exit Sub
    End If
    efectivoReal = CDbl(TextBoxEfectivoReal.Value)
    
    Dim totalEfectivoVend As Double
    Dim diferencia As Double
    
    ' Buscar el total vendido en EFECTIVO
    For j = 0 To ListBoxResumen.ListCount - 1
        If UCase(Trim(ListBoxResumen.List(j, 0))) = "EFECTIVO" Then
            If IsNumeric(ListBoxResumen.List(j, 1)) Then
                totalEfectivoVend = CDbl(ListBoxResumen.List(j, 1))
            End If
            Exit For
        End If
    Next j
    ' Buscar efectivo inicial desde tblCaja
    Dim efectivoInicial As Double
    efectivoInicial = 0
    




    Set wsCaja = ThisWorkbook.Sheets("Caja")
    Set tbl = wsCaja.ListObjects("tblCaja")
    
    For i = tbl.ListRows.Count To 1 Step -1
        With tbl.ListRows(i).Range
            If .Cells(1, 1).Value = Date And UCase(.Cells(1, 3).Value) = "EFECTIVO" Then
                efectivoInicial = .Cells(1, 4).Value ' MontoInicial
                Exit For
            End If
        End With
    Next i
    
    ' Calcular diferencia real
    diferencia = efectivoReal - (totalEfectivoVend + efectivoInicial)
    
    If Abs(diferencia) > 0.01 Then
        Dim r As VbMsgBoxResult
        r = MsgBox("Hay una diferencia de $" & Format(diferencia, "#,##0.00") & " en el cierre de caja." & vbCrLf & _
                   "¿Deseás continuar igualmente?", vbExclamation + vbYesNo, "Diferencia detectada")
        If r = vbNo Then Exit Sub
    End If



    Set wsCaja = ThisWorkbook.Sheets("Caja")
    Set tbl = wsCaja.ListObjects("tblCaja")

    ' Cerrar cada fila abierta hoy
    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            If .Cells(1, 1).Value = Date And .Cells(1, 5).Value = "" Then ' sin cierre
                medio = UCase(Trim(.Cells(1, 3).Value))
                totalVend = 0
                
                ' Buscar en el ListBoxResumen el total correspondiente a ese medio
                For j = 0 To ListBoxResumen.ListCount - 1
                    If UCase(Trim(ListBoxResumen.List(j, 0))) = medio Then
                        If IsNumeric(ListBoxResumen.List(j, 1)) Then
                            totalVend = CDbl(ListBoxResumen.List(j, 1))
                        End If
                        Exit For
                    End If
                Next j

                .Cells(1, 5).Value = totalVend ' MontoCierre

                If medio = "EFECTIVO" Then
                    .Cells(1, 6).Value = efectivoReal - totalVend ' Diferencia
                Else
                    .Cells(1, 6).Value = 0
                End If

                .Cells(1, 7).Value = Environ("Username") & " / " & Format(Time, "hh:mm:ss")
                .Cells(1, 8).Value = "Cierre"
            End If
        End With
    Next i

    MsgBox "Caja cerrada correctamente.", vbInformation
    Unload Me
End Sub

