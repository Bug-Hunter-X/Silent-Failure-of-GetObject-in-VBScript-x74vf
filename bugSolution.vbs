Function GetObjectSafe(path)
  Dim obj, ErrObj
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    ErrObj = Err.Description
    Err.Clear
    Set obj = Nothing
    MsgBox "Error accessing file: " & ErrObj, vbCritical
  End If
  Set GetObjectSafe = obj
End Function

'Example Usage
Set myExcel = GetObjectSafe("C:\\path\\to\\your\\excel.xls")
if not myexcel is nothing then
  MsgBox "Excel object created successfully."
else
  MsgBox "Failed to create Excel object."
end if 