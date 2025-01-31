Function MyFunction(param1)
  ' Some code here that might cause an error
  If IsEmpty(param1) Then
    Err.Raise 13, , "Type mismatch: param1 is empty"
  End If
  ' Rest of the function code
End Function