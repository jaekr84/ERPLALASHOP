VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAjusteStock 
   Caption         =   "UserForm1"
   ClientHeight    =   10695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8565.001
   OleObjectBlob   =   "VBAfrmAjusteStock.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAjusteStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscar_Click()
    Dim wsStock As Worksheet
    Dim tblStock As ListObject
    Dim i As Long
    Dim codigoBuscado As String
    Dim cod As String, desc As String, talle As String, color As String
    Dim stockActual As Double, codBarra As String
    Dim encontrado As Boolean

    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set tblStock = wsStock.ListObjects("Stock")

    codigoBuscado = Trim(txtCodigo.Value)
    If codigoBuscado = "" Then
        MsgBox "Ingresá un código de producto.", vbExclamation
        Exit Sub
    End If

    listDetalle.Clear
    encontrado = False

    For i = 1 To tblStock.ListRows.Count
        With tblStock.ListRows(i).Range
            cod = .Cells(1, 1).Value
            If cod = codigoBuscado Then
                desc = .Cells(1, 2).Value
                stockActual = .Cells(1, 6).Value
                codBarra = .Cells(1, 7).Value
                talle = .Cells(1, 9).Value
                color = .Cells(1, 10).Value

                listDetalle.AddItem cod
                listDetalle.List(listDetalle.ListCount - 1, 1) = desc
                listDetalle.List(listDetalle.ListCount - 1, 2) = talle
                listDetalle.List(listDetalle.ListCount - 1, 3) = color
                listDetalle.List(listDetalle.ListCount - 1, 4) = stockActual
                listDetalle.List(listDetalle.ListCount - 1, 5) = stockActual
                listDetalle.List(listDetalle.ListCount - 1, 6) = codBarra
                encontrado = True
            End If
        End With
    Next i

    If Not encontrado Then
        MsgBox "No se encontraron variantes para este código.", vbInformation
    End If
End Sub


Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub btnConfirmar_Click()
    Dim wsStock As Worksheet, wsMov As Worksheet
    Dim tblStock As ListObject, tblMov As ListObject
    Dim i As Long, j As Long
    Dim cod As String, codBarra As String
    Dim desc As String, talle As String, color As String
    Dim stockActual As Double, stockNuevo As Double, diferencia As Double
    Dim filaMov As ListRow
    Dim movCount As Long
    Dim filaStock As Range
    Dim hoy As Date
    Dim wsHist As Worksheet
    Dim tblHist As ListObject
    Dim filaHist As ListRow
    
    Set wsHist = ThisWorkbook.Sheets("HistorialAjustes")
    Set tblHist = wsHist.ListObjects("tblHistorialAjustes")
    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set wsMov = ThisWorkbook.Sheets("MovimientosStock")
    Set tblStock = wsStock.ListObjects("Stock")
    Set tblMov = wsMov.ListObjects("Movimientos")
    
    hoy = Date
    movCount = 0

    For i = 0 To listDetalle.ListCount - 1
        cod = listDetalle.List(i, 0)
        desc = listDetalle.List(i, 1)
        talle = listDetalle.List(i, 2)
        color = listDetalle.List(i, 3)
        stockActual = CDbl(listDetalle.List(i, 4))
        stockNuevo = CDbl(listDetalle.List(i, 5))
        codBarra = listDetalle.List(i, 6)

        If stockNuevo <> stockActual Then
            diferencia = stockNuevo - stockActual
            
            ' Buscar la fila correspondiente en la tabla Stock
            For j = 1 To tblStock.ListRows.Count
                With tblStock.ListRows(j).Range
                    If .Cells(1, 1).Value = cod And _
                       .Cells(1, 7).Value = codBarra And _
                       .Cells(1, 9).Value = talle And _
                       .Cells(1, 10).Value = color Then
                       
                        .Cells(1, 6).Value = stockNuevo ' ? actualizar stock
                        Exit For
                    End If
                End With
            Next j
            
            ' Registra en hoja de HistorialAjustes
            Set filaHist = tblHist.ListRows.Add
            With filaHist.Range
                .Cells(1, 1).Value = Date
                .Cells(1, 2).Value = cod
                .Cells(1, 3).Value = desc
                .Cells(1, 4).Value = talle
                .Cells(1, 5).Value = color
                .Cells(1, 6).Value = stockActual
                .Cells(1, 7).Value = stockNuevo
                .Cells(1, 8).Value = stockNuevo - stockActual
            End With

            
            ' Registrar en MovimientosStock
            Set filaMov = tblMov.ListRows.Add
            With filaMov.Range
                .Cells(1, 1).Value = hoy
                .Cells(1, 2).Value = cod
                .Cells(1, 3).Value = desc
                .Cells(1, 4).Value = talle
                .Cells(1, 5).Value = color
                .Cells(1, 6).Value = Abs(diferencia)
                .Cells(1, 7).Value = "Ajuste"
            End With
            
            movCount = movCount + 1
        End If
    Next i

    If movCount > 0 Then
        MsgBox "Se aplicaron " & movCount & " ajustes de stock correctamente.", vbInformation
    Else
        MsgBox "No hubo diferencias de stock que ajustar.", vbInformation
    End If
End Sub


Private Sub listDetalle_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim fila As Long
    Dim stockNuevo As Variant
    Dim valorActual As Double

    fila = listDetalle.ListIndex
    If fila < 0 Then Exit Sub

    valorActual = listDetalle.List(fila, 5) ' columna de Stock nuevo
    stockNuevo = InputBox("Ingresá el nuevo stock para esta variante:", "Editar stock", valorActual)

    If stockNuevo = "" Then Exit Sub ' canceló

    If IsNumeric(stockNuevo) Then
        If CDbl(stockNuevo) >= 0 Then
            listDetalle.List(fila, 5) = CDbl(stockNuevo)
        Else
            MsgBox "El valor debe ser mayor o igual a cero.", vbExclamation
        End If
    Else
        MsgBox "Ingresá un número válido.", vbExclamation
    End If
End Sub

