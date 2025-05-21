VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEtiquetas 
   Caption         =   "UserForm7"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "VBAfrmEtiquetas.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Dim codigoBase As String
    Dim fila As Range
    Dim cod As String, desc As String, talle As String, color As String

    Set ws = ThisWorkbook.Sheets("Stock")
    Set tbl = ws.ListObjects("Stock")

    codigoBase = Trim(txtCodigoBase.Value)
    If codigoBase = "" Then
        MsgBox "Ingresá un código base para buscar.", vbExclamation
        Exit Sub
    End If

    lstVariantes.Clear

    For i = 1 To tbl.ListRows.Count
        Set fila = tbl.ListRows(i).Range
        cod = fila.Cells(1, 1).Value

        If cod = codigoBase Then
            desc = fila.Cells(1, 2).Value
            talle = fila.Cells(1, 9).Value
            color = fila.Cells(1, 10).Value

            lstVariantes.AddItem cod
            lstVariantes.List(lstVariantes.ListCount - 1, 1) = desc
            lstVariantes.List(lstVariantes.ListCount - 1, 2) = talle
            lstVariantes.List(lstVariantes.ListCount - 1, 3) = color
            lstVariantes.List(lstVariantes.ListCount - 1, 4) = 1 ' Cantidad por defecto
        End If
    Next i

    If lstVariantes.ListCount = 0 Then
        MsgBox "No se encontraron variantes para el código ingresado.", vbInformation
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnExportarCSV_Click()
    Dim rutaCSV As String
    Dim rutaScript As String
    Dim i As Long
    Dim linea As String
    Dim archivo As Integer
    
    If lstVariantes.ListCount = 0 Then
        MsgBox "No hay variantes para exportar.", vbExclamation
        Exit Sub
    End If

    ' Ruta del archivo CSV (CORREGIDA)
    rutaCSV = "C:\Users\lalab\OneDrive\Documentos\ERP LALA SHOP\Etiquetas ERP\temp_etiquetas.csv"
    archivo = FreeFile

    ' Crear archivo CSV
    Open rutaCSV For Output As #archivo
    Print #archivo, "codigo,descripcion,talle,color,cod_barra,cantidad"

    ' Escribir cada línea del ListBox
    For i = 0 To lstVariantes.ListCount - 1
        linea = lstVariantes.List(i, 0) & "," & _
                lstVariantes.List(i, 1) & "," & _
                lstVariantes.List(i, 2) & "," & _
                lstVariantes.List(i, 3) & "," & _
                ObtenerCodigoBarra(lstVariantes.List(i, 0), lstVariantes.List(i, 2), lstVariantes.List(i, 3)) & "," & _
                lstVariantes.List(i, 4)
        Print #archivo, linea
    Next i
    Close #archivo

    ' Ruta completa a python.exe (ajustala si es necesario)
    Dim rutaPython As String
    rutaPython = "C:\Users\lalab\AppData\Local\Programs\Python\Python313\python.exe"

    ' Ruta al script
    rutaScript = "C:\Users\lalab\OneDrive\Documentos\ERP LALA SHOP\Etiquetas ERP\generar_etiquetas_alta_calidad.py"

    ' Ejecutar script de Python
    Shell """" & rutaPython & """ """ & rutaScript & """", vbNormalFocus

    MsgBox "Archivo CSV generado y script de etiquetas ejecutado correctamente.", vbInformation
End Sub

Function ObtenerCodigoBarra(cod As String, talle As String, color As String) As String
    Dim ws As Worksheet, tbl As ListObject
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Stock")
    Set tbl = ws.ListObjects("Stock")

    For i = 1 To tbl.ListRows.Count
        With tbl.ListRows(i).Range
            If .Cells(1, 1).Value = cod And _
               .Cells(1, 9).Value = talle And _
               .Cells(1, 10).Value = color Then
                ObtenerCodigoBarra = .Cells(1, 7).Value
                Exit Function
            End If
        End With
    Next i

    ObtenerCodigoBarra = ""
End Function

Private Sub lstVariantes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim cantidad As String
    Dim fila As Long

    fila = lstVariantes.ListIndex
    If fila < 0 Then Exit Sub

    cantidad = InputBox("Ingresá la cantidad de etiquetas para esta variante:", "Cantidad", lstVariantes.List(fila, 4))

    If IsNumeric(cantidad) And Val(cantidad) >= 0 Then
        lstVariantes.List(fila, 4) = Val(cantidad)
    Else
        MsgBox "Ingresá un número válido.", vbExclamation
    End If
End Sub

Private Sub UserForm_Click()
End Sub


