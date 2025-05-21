VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListadoCompras 
   Caption         =   "UserForm1"
   ClientHeight    =   9360.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5730
   OleObjectBlob   =   "VBAfrmListadoCompras.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmListadoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton4_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    ' Asignar fecha actual como valor por defecto
    txtFechaDesde.Value = Format(Date, "dd/mm/yyyy")
    txtFechaHasta.Value = Format(Date, "dd/mm/yyyy")
End Sub
Private Sub txtFechaDesde_AfterUpdate()
    If IsDate(txtFechaDesde.Value) Then
        txtFechaDesde.Value = Format(CDate(txtFechaDesde.Value), "dd/mm/yyyy")
    Else
        MsgBox "La fecha ingresada no es válida.", vbExclamation
        txtFechaDesde.Value = Format(Date, "dd/mm/yyyy")
    End If
End Sub

Private Sub txtFechaHasta_AfterUpdate()
    If IsDate(txtFechaHasta.Value) Then
        txtFechaHasta.Value = Format(CDate(txtFechaHasta.Value), "dd/mm/yyyy")
    Else
        MsgBox "La fecha ingresada no es válida.", vbExclamation
        txtFechaHasta.Value = Format(Date, "dd/mm/yyyy")
    End If
End Sub

Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim dictTotales As Object
    Dim dictFechas As Object
    Dim fechaDesde As Date, fechaHasta As Date
    Dim fechaCompra As Date
    Dim nroComp As String
    Dim subtotal As Double

    ' Validar fechas
    If Not IsDate(txtFechaDesde.Value) Or Not IsDate(txtFechaHasta.Value) Then
        MsgBox "Las fechas ingresadas no son válidas.", vbExclamation
        Exit Sub
    End If

    fechaDesde = CDate(txtFechaDesde.Value)
    fechaHasta = CDate(txtFechaHasta.Value)

    Set ws = ThisWorkbook.Sheets("Compras")
    Set tbl = ws.ListObjects("tblCompras")
    Set dictTotales = CreateObject("Scripting.Dictionary")
    Set dictFechas = CreateObject("Scripting.Dictionary")

    For Each fila In tbl.ListRows
        fechaCompra = fila.Range.Cells(1, 1).Value          ' Col 1: Fecha
        nroComp = Trim(CStr(fila.Range.Cells(1, 10).Value)) ' Col 10: Nº comprobante
        subtotal = fila.Range.Cells(1, 9).Value              ' Col 9: Subtotal

        If IsDate(fechaCompra) Then
            If fechaCompra >= fechaDesde And fechaCompra <= fechaHasta Then
                ' Guardar o acumular el subtotal
                If dictTotales.exists(nroComp) Then
                    dictTotales(nroComp) = dictTotales(nroComp) + subtotal
                Else
                    dictTotales.Add nroComp, subtotal
                    dictFechas.Add nroComp, fechaCompra
                End If
            End If
        End If
    Next fila

    ' Limpiar el ListBox
    lstComprobantes.Clear

    ' Cargar en el ListBox
    Dim k As Variant
    For Each k In dictTotales.Keys
        lstComprobantes.AddItem Format(dictFechas(k), "dd/mm/yyyy")           ' Col 0: Fecha
        lstComprobantes.List(lstComprobantes.ListCount - 1, 1) = k            ' Col 1: Comprobante
        lstComprobantes.List(lstComprobantes.ListCount - 1, 2) = Format(dictTotales(k), "#,##0") ' Col 2: Total
    Next k
    
    ' Calcular y mostrar el total general
    Dim totalGeneral As Double
    For Each k In dictTotales.Keys
        totalGeneral = totalGeneral + dictTotales(k)
    Next k

    lblTotal.Caption = "$" & Format(totalGeneral, "#,##0")

End Sub

Private Sub lstComprobantes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim nroComprobante As String
    If lstComprobantes.ListIndex = -1 Then Exit Sub

    nroComprobante = lstComprobantes.List(lstComprobantes.ListIndex, 1)

    frmDetalleCompra.CargarComprobante nroComprobante
    frmDetalleCompra.Show
End Sub


