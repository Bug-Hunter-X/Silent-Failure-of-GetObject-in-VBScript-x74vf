Function GetObject(path)
  On Error Resume Next
  Set obj = GetObject(path)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = Nothing
  End If
  Set GetObject = obj
End Function

' Example usage:
Set myExcel = GetObject("C:\\path\\to\\your\\excel.xls")
if not myexcel is nothing then
  MsgBox "Excel object created successfully."
else
  MsgBox "Failed to create Excel object."
end if 