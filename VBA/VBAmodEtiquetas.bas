Attribute VB_Name = "modEtiquetas"
Sub GenerarEtiquetasCentradas(frm As Object)
    Dim wordApp As Object, doc As Object, section As Object
    Dim wsStock As Worksheet
    Dim rutaBase As String, nombreArchivo As String
    Dim docPath As String, pdfPath As String
    Dim codigoBase As String, descripcion As String, precio As Double, codBarra As String
    Dim talle As String, color As String, linea As String
    Dim i As Long, j As Long, cantidad As Long

    ' Ruta
    rutaBase = "C:\Users\rafa\Desktop\Etiquetas ERP\"
    If Dir(rutaBase, vbDirectory) = "" Then MkDir rutaBase

    ' Buscar código base
    codigoBase = Trim(frm.txtBuscarCodigo.Value)
    If codigoBase = "" Then
        MsgBox "Ingresá un código base.", vbExclamation
        Exit Sub
    End If

    ' Buscar en Stock
    Set wsStock = ThisWorkbook.Sheets("Stock")
    descripcion = ""
    For i = 2 To wsStock.Cells(wsStock.Rows.Count, 1).End(xlUp).Row
        If Trim(wsStock.Cells(i, 1).Value) = codigoBase Then
            descripcion = wsStock.Cells(i, 2).Value
            precio = wsStock.Cells(i, 5).Value
            codBarra = wsStock.Cells(i, 7).Value
            Exit For
        End If
    Next i

    If descripcion = "" Then
        MsgBox "No se encontró el producto en Stock.", vbExclamation
        Exit Sub
    End If

    ' Crear Word
    Set wordApp = CreateObject("Word.Application")
    Set doc = wordApp.Documents.Add

    ' Configurar primera sección
    With doc.Sections(1).PageSetup
.PageWidth = 141.75
.PageHeight = 70.875
.TopMargin = 5
.BottomMargin = 5
.LeftMargin = 5
.RightMargin = 5

    End With

    ' Configurar estilo
    With doc.Styles("Normal").Font
        .Name = "Arial"
        .Size = 6
    End With

    ' Agregar etiquetas
    Dim totalEtiquetas As Long: totalEtiquetas = 0
    For i = 0 To frm.lstVariantes.ListCount - 1
        If IsNumeric(frm.lstVariantes.List(i, 2)) Then
            cantidad = frm.lstVariantes.List(i, 2)
            If cantidad > 0 Then
                talle = frm.lstVariantes.List(i, 0)
                color = frm.lstVariantes.List(i, 1)
                For j = 1 To cantidad
                    If totalEtiquetas > 0 Then
                        Set section = doc.Sections.Add
                        With section.PageSetup
.PageWidth = 141.75
.PageHeight = 70.875
.TopMargin = 5
.BottomMargin = 5
.LeftMargin = 5
.RightMargin = 5
                            .RightMargin = 14.17   ' 0.5 cm



                        End With
                    End If

                    With doc.Paragraphs.last.Range
                        .ParagraphFormat.Alignment = 1 ' Centrado
                        .Text = vbCrLf & _
                                "Código: " & codigoBase & " | $" & precio & vbCrLf & _
                                descripcion & vbCrLf & _
                                "Talle: " & talle & " | Color: " & color & vbCrLf & _
                                "*" & codBarra & "*"
                    End With

                    totalEtiquetas = totalEtiquetas + 1
                Next j
            End If
        End If
    Next i

    ' Guardar
    nombreArchivo = "Etiquetas_" & Format(Now, "yyyymmdd_hhmmss")
    docPath = rutaBase & nombreArchivo & ".docx"
    pdfPath = rutaBase & nombreArchivo & ".pdf"

    doc.SaveAs2 docPath
    doc.SaveAs2 pdfPath, FileFormat:=17 ' PDF

    ' Abrir ambos
    Shell "explorer.exe """ & docPath & """", vbNormalFocus
    Shell "explorer.exe """ & pdfPath & """", vbNormalFocus

    doc.Close False
    wordApp.Quit
    Set doc = Nothing: Set wordApp = Nothing

    MsgBox "Se generaron " & totalEtiquetas & " etiquetas correctamente.", vbInformation
End Sub


