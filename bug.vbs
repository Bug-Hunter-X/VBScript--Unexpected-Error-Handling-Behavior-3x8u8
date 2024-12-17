Function MyFunction(param1)
  ' Some code here
  If param1 = "" Then
    Err.Raise 1001, , "Parameter cannot be empty"
  End If
End Function