Function MyFunction(param1)
  On Error GoTo ErrorHandler
  ' Some code here that might cause an error
  If IsEmpty(param1) Then
    Err.Raise vbError, , "Parameter 'param1' cannot be empty"
  End If
  ' Rest of the function code
  Exit Function
ErrorHandler:
  MsgBox "Error: " & Err.Description & ". Please check your input.", vbCritical
  Err.Clear
End Function