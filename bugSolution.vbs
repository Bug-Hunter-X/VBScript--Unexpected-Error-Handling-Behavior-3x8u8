Function MyFunction(param1)
  On Error GoTo ErrorHandler
  ' Some code here
  If param1 = "" Then
    Err.Raise 1001, , "Parameter cannot be empty"
  End If
  Exit Function
ErrorHandler:
  MsgBox "Error Number: " & Err.Number & "
Description: " & Err.Description, vbCritical
  ' Handle the error appropriately. You can log it, retry, or take other actions.
End Function