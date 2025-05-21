Attribute VB_Name = "ImpFrmMod"
Option Explicit
' Requiere referencia a:
'   - Microsoft Visual Basic for Applications Extensibility 5.3
' Debe estar activo: Confiar en el acceso al modelo de objetos de VBA

Sub ImportarFormsYModulos()
    Dim rutaCarpeta As String
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim proyecto As VBIDE.VBProject
    Dim compExistente As VBIDE.VBComponent
    Dim nombreComponente As String
    Dim ext As String
    
    ' >>> Ajusta esta ruta a tu carpeta de importación <<<
    rutaCarpeta = "PEGAR ACA LA RUTA"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rutaCarpeta) Then
        MsgBox "No existe la carpeta: " & rutaCarpeta, vbExclamation
        Exit Sub
    End If
    
    Set carpeta = fso.GetFolder(rutaCarpeta)
    Set proyecto = ThisWorkbook.VBProject
    
    ' (Opcional) Elimina módulos/forms con el mismo nombre antes de importar
    For Each archivo In carpeta.Files
        ext = LCase(fso.GetExtensionName(archivo.Name))
        If ext = "frm" Or ext = "bas" Or ext = "cls" Then
            nombreComponente = fso.GetBaseName(archivo.Name)
            On Error Resume Next
            Set compExistente = proyecto.VBComponents(nombreComponente)
            If Not compExistente Is Nothing Then
                proyecto.VBComponents.Remove compExistente
            End If
            On Error GoTo 0
        End If
    Next archivo
    
    ' Importa todos los .frm, .bas y .cls de la carpeta
    For Each archivo In carpeta.Files
        ext = LCase(fso.GetExtensionName(archivo.Name))
        Select Case ext
            Case "frm", "bas", "cls"
                proyecto.VBComponents.Import archivo.Path
        End Select
    Next archivo
    
    MsgBox "Importación de UserForms y Módulos completada.", vbInformation
End Sub

