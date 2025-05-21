VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsulta 
   Caption         =   "UserForm4"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8220.001
   OleObjectBlob   =   "VBAfrmConsulta.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    txtCodBarra.SetFocus
End Sub

Private Sub btnBuscar_Click()
    Dim hojaStock As Worksheet
    Dim datos As Variant
    Dim i As Long
    Dim texto As String

    texto = LCase(Trim(txtBuscar.Value))
    Set hojaStock = ThisWorkbook.Sheets("Stock")

    ' Cargar los datos a memoria
    datos = hojaStock.Range("A2:K" & hojaStock.Cells(hojaStock.Rows.Count, 1).End(xlUp).Row).Value

    ListBox1.Clear

    For i = 1 To UBound(datos)
        If LCase(CStr(datos(i, 1))) Like "*" & texto & "*" Then
            ListBox1.AddItem datos(i, 1)              ' Código
            ListBox1.List(ListBox1.ListCount - 1, 1) = datos(i, 2) ' Producto
            ListBox1.List(ListBox1.ListCount - 1, 2) = datos(i, 9) ' Talle
            ListBox1.List(ListBox1.ListCount - 1, 3) = datos(i, 10) ' Color
            ListBox1.List(ListBox1.ListCount - 1, 4) = datos(i, 5) ' Precio
            ListBox1.List(ListBox1.ListCount - 1, 5) = datos(i, 6) ' Stock
        End If
    Next i

    ' Mostrar mensaje si no se encontraron coincidencias
    If ListBox1.ListCount = 0 Then
        MsgBox "No se encontró ningún producto con ese código.", vbExclamation
        txtBuscar.Value = ""
    End If
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub txtCodBarra_AfterUpdate()
    Dim hojaStock As Worksheet
    Dim datos As Variant
    Dim i As Long
    Dim textoBarra As String
    Dim encontrado As Boolean

    textoBarra = LCase(Trim(txtCodBarra.Value))
    Set hojaStock = ThisWorkbook.Sheets("Stock")
    datos = hojaStock.Range("A2:K" & hojaStock.Cells(hojaStock.Rows.Count, 1).End(xlUp).Row).Value

    ListBox1.Clear
    encontrado = False

    For i = 1 To UBound(datos)
        If LCase(CStr(datos(i, 7))) = textoBarra Then
            ListBox1.AddItem datos(i, 1)               ' Código
            ListBox1.List(ListBox1.ListCount - 1, 1) = datos(i, 2) ' Producto
            ListBox1.List(ListBox1.ListCount - 1, 2) = datos(i, 9) ' Talle
            ListBox1.List(ListBox1.ListCount - 1, 3) = datos(i, 10) ' Color
            ListBox1.List(ListBox1.ListCount - 1, 4) = datos(i, 5)  ' Precio
            ListBox1.List(ListBox1.ListCount - 1, 5) = datos(i, 6)  ' Stock
            encontrado = True
        End If
    Next i

        If Not encontrado Then
        MsgBox "Código de barra no encontrado.", vbExclamation
    End If

    txtCodBarra.Value = ""

    ' Forzar foco luego de que termine el evento
   Call ReasignarFocoSiCorresponde

End Sub

Private Sub ReasignarFocoSiCorresponde()
    ' Solo reasigna el foco si el campo activo NO cambió manualmente
    If Me.ActiveControl Is Nothing Then Exit Sub
    If Me.ActiveControl.Name = "txtCodBarra" Then Exit Sub
    If Me.ActiveControl.Name = "txtBuscar" Then Exit Sub ' el usuario está buscando, no interrumpir

    txtCodBarra.SetFocus
End Sub

