' This removes any existing uninst.exe on disk during an upgrade so that SensorID is preserved

On Error Resume Next

' Clear any existing Errors
Err.Clear

dim fsStr
dim unStr
dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

If Not Err Then
  If fso.FileExists(Session.Property("CustomActionData")) Then
    fso.DeleteFile(Session.Property("CustomActionData"))
  End If
End If

' Clear any Errors
Err.Clear