Attribute VB_Name = "ExportFrmMod"
Sub ExportarTodosComponentesVBA()
    Dim vbComp As VBIDE.VBComponent
    Dim rutaExport As String
    Dim ext As String
    
    ' 1) Ruta absoluta de exportacion (asegurate de que exista la carpeta completa)
    rutaExport = "C:\Users\rafa\Desktop\ERP LALA SHOP\VBA"
    
    ' Si no existe la carpeta, la creamos (MkDir falla si faltan niveles intermedios)
    If Dir(rutaExport, vbDirectory) = "" Then
        ' Creamos recursivamente cada nivel
        CrearCarpetaCompleta rutaExport
    End If
    
    ' Recorre cada componente del VBAProject
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule:  ext = ".bas"
            Case vbext_ct_ClassModule: ext = ".cls"
            Case vbext_ct_MSForm:      ext = ".frm"
            Case Else:                 ext = ""
        End Select
        
        If ext <> "" Then
            vbComp.Export rutaExport & vbComp.Name & ext
        End If
    Next vbComp
    
    MsgBox "Exportacion completada en:" & vbCrLf & rutaExport, vbInformation
End Sub

' Funcion auxiliar para crear toda la ruta aunque falten carpetas intermedias
Private Sub CrearCarpetaCompleta(ByVal sRuta As String)
    Dim arr(), sCamino As String
    Dim i As Long
    
    ' Quitar barra final
    If Right(sRuta, 1) = "\" Then sRuta = Left(sRuta, Len(sRuta) - 1)
    arr = Split(sRuta, "\")
    
    sCamino = arr(0) & "\"
    For i = 1 To UBound(arr)
        sCamino = sCamino & arr(i) & "\"
        If Dir(sCamino, vbDirectory) = "" Then MkDir sCamino
    Next i
End Sub

