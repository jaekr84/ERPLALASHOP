VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDetalleProducto 
   Caption         =   "UserForm7"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "VBAfrmDetalleProducto.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmDetalleProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Public Sub MostrarDetalle(codigo As String)
    Dim wsStock As Worksheet, wsVentas As Worksheet, wsCompras As Worksheet
    Dim ultimaFila As Long, i As Long
    Dim vendidas As Double, ingresadas As Double
    Dim primeraFecha As Variant
    Dim descripcion As String, proveedor As String

    Set wsStock = ThisWorkbook.Sheets("Stock")
    Set wsVentas = ThisWorkbook.Sheets("Ventas")
    Set wsCompras = ThisWorkbook.Sheets("Compras") ' cambialo si usás otra hoja

    ' Buscar descripción y fecha desde la hoja Stock
    ultimaFila = wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaFila
        If Trim(wsStock.Cells(i, 1).Value) = codigo Then
            descripcion = wsStock.Cells(i, 2).Value
            primeraFecha = wsStock.Cells(i, 11).Value
            Exit For
        End If
    Next i

    ' Buscar proveedor en hoja Compras
    proveedor = "(no encontrado)"
    ultimaFila = wsCompras.Cells(wsCompras.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaFila
        If Trim(wsCompras.Cells(i, 3).Value) = codigo Then ' Col 3 = Código artículo
            proveedor = wsCompras.Cells(i, 2).Value          ' Col 2 = Proveedor
            Exit For
        End If
    Next i

    ' Contar VENTAS
    vendidas = 0
    ultimaFila = wsVentas.Cells(wsVentas.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaFila
        If Trim(wsVentas.Cells(i, 2).Value) = codigo Then
            vendidas = vendidas + Val(wsVentas.Cells(i, 4).Value)
        End If
    Next i

    ' Contar INGRESOS
    ingresadas = 0
    ultimaFila = wsCompras.Cells(wsCompras.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaFila
        If Trim(wsCompras.Cells(i, 3).Value) = codigo Then ' Col 3 = Código
            ingresadas = ingresadas + Val(wsCompras.Cells(i, 6).Value) ' Col 6 = Cantidad
        End If
    Next i

    ' Mostrar en etiquetas
    lblCodigo.Caption = codigo
    lblDescripcion.Caption = descripcion
    lblProveedor.Caption = proveedor
    lblFecha.Caption = Format(primeraFecha, "dd/mm/yyyy")
    lblVendidas.Caption = vendidas
    lblIngresadas.Caption = ingresadas
End Sub

