Function MyFunction(param1, param2)
  If IsEmpty(param1) Or IsEmpty(param2) Then
    Err.Raise vbError, , "Parameters cannot be empty. Please provide values for both param1 and param2."
    Exit Function 'added for early exit upon error
  End If
  On Error Resume Next
    Dim result : result = CInt(param1) + CInt(param2) 'Explicit type conversion for robustness
  If Err.Number <> 0 Then
    Err.Raise vbError, , "Invalid parameter type. Parameters must be numeric."
    Exit Function 'added for early exit upon error
  End If
  On Error GoTo 0
  ' ...rest of the function
  MyFunction = result
End Function