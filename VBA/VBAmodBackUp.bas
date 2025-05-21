Attribute VB_Name = "modBackUp"
Sub HacerBackup(Optional ByVal motivo As String = "general")
    Dim rutaTemp As String
    Dim rutaDestino As String
    Dim nombreArchivo As String
    Dim rutaBackup As String
    Dim fechaHoy As String

    On Error GoTo ErrHandler

    ' RUTA DE BACKUP DESTINO
    rutaBackup = "PEGAR ACA LA RUTA"
    If Right(rutaBackup, 1) <> "\" Then rutaBackup = rutaBackup & "\"

    If Dir(rutaBackup, vbDirectory) = "" Then
        MsgBox "La carpeta de backup no existe.", vbCritical
        Exit Sub
    End If

    ' NOMBRE DE ARCHIVO DE BACKUP
    fechaHoy = Format(Date, "yyyymmdd")
    nombreArchivo = "ERP_BACKUP_" & motivo & "_" & fechaHoy & ".xlsm"
    rutaDestino = rutaBackup & nombreArchivo

    ' GUARDAR UNA COPIA TEMPORAL PRIMERO
    rutaTemp = Environ("TEMP") & "\ERP_TEMP_BACKUP_" & motivo & ".xlsm"
    ThisWorkbook.SaveCopyAs rutaTemp

    ' BORRAR EL DESTINO SI YA EXISTE
    If Dir(rutaDestino) <> "" Then Kill rutaDestino

    ' COPIAR LA COPIA TEMPORAL AL DESTINO FINAL
    FileCopy rutaTemp, rutaDestino

    ' OPCIONAL: eliminar la copia temporal
    Kill rutaTemp

    ' MsgBox "Backup creado correctamente en: " & rutaDestino, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "ERROR " & Err.Number & ": " & Err.Description & vbCrLf & "Destino: " & rutaDestino, vbCritical
End Sub

