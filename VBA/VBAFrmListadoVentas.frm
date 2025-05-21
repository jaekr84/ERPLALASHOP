VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListadoVentas 
   Caption         =   "UserForm7"
   ClientHeight    =   9495.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5610
   OleObjectBlob   =   "VBAFrmListadoVentas.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "FrmListadoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim fila As ListRow
    Dim dictComprobantes As Object
    Dim fechaDesde As Date, fechaHasta As Date
    Dim fechaVenta As Date
    Dim nroComp As String, fechaTexto As String
    Dim clave As String
    
    If Not IsDate(txtFechaDesde.Value) Or Not IsDate(txtFechaHasta.Value) Then
        MsgBox "Las fechas ingresadas no son válidas.", vbExclamation
        Exit Sub
    End If
    
    fechaDesde = CDate(txtFechaDesde.Value)
    fechaHasta = CDate(txtFechaHasta.Value)

    Set ws = ThisWorkbook.Sheets("Ventas")
    Set tbl = ws.ListObjects("Tabla1")
    Set dictComprobantes = CreateObject("Scripting.Dictionary")

    ' Recorrer las filas de la tabla
    For Each fila In tbl.ListRows
        fechaVenta = fila.Range.Cells(1, 1).Value ' Columna A
        nroComp = Trim(fila.Range.Cells(1, 12).Value) ' Columna L

        If IsDate(fechaVenta) Then
            If fechaVenta >= fechaDesde And fechaVenta <= fechaHasta Then
                ' Agrupar por comprobante
                If Not dictComprobantes.exists(nroComp) Then
                    fechaTexto = Format(fechaVenta, "dd/mm/yyyy")
                    dictComprobantes.Add nroComp, fechaTexto
                End If
            End If
        End If
    Next fila

    ' Cargar en ListBox
    ListBoxComprobantes.Clear
    Dim k As Variant
    For Each k In dictComprobantes.Keys
        ListBoxComprobantes.AddItem dictComprobantes(k) ' ? ahora la fecha es la columna 0
        ListBoxComprobantes.List(ListBoxComprobantes.ListCount - 1, 1) = k ' ? comprobante es la columna 1

    Next k
End Sub




Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Asignar fecha actual como valor por defecto
    txtFechaDesde.Value = Format(Date, "dd/mm/yyyy")
    txtFechaHasta.Value = Format(Date, "dd/mm/yyyy")
    Call CargarComprobantes
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

Sub CargarComprobantes()
    Dim ws As Worksheet
    Dim ultimaFila As Long, i As Long
    Dim dictComprobantes As Object
    Dim fechaDesde As Date, fechaHasta As Date
    Dim f As Date
    Dim comprobante As Variant

    Set ws = ThisWorkbook.Sheets("Ventas")
    Set dictComprobantes = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    fechaDesde = CDate(txtFechaDesde.Value)
    fechaHasta = CDate(txtFechaHasta.Value)
    On Error GoTo 0

    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ListBoxComprobantes.Clear

    For i = 2 To ultimaFila
        If IsDate(ws.Cells(i, 1).Value) Then
            f = ws.Cells(i, 1).Value
            comprobante = ws.Cells(i, 12).Value ' Columna L

            If f >= fechaDesde And f <= fechaHasta Then
                If Not dictComprobantes.exists(comprobante) Then
                    dictComprobantes.Add comprobante, f
                End If
            End If
        End If
    Next i

    For Each comprobante In dictComprobantes.Keys
        ListBoxComprobantes.AddItem dictComprobantes(comprobante) ' Col 0: Fecha
        ListBoxComprobantes.List(ListBoxComprobantes.ListCount - 1, 1) = comprobante ' Col 1: Comprobante
    Next comprobante
End Sub

Private Sub ListBoxComprobantes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim comprobante As String
    If ListBoxComprobantes.ListIndex = -1 Then Exit Sub
    comprobante = ListBoxComprobantes.List(ListBoxComprobantes.ListIndex, 1)
    
    FrmDetalleVenta.CargarComprobante comprobante
    FrmDetalleVenta.Show
End Sub

