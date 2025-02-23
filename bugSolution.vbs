Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 9999, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
End Function

Sub Main
  On Error GoTo ErrorHandler
  Dim result
  result = MyFunction("")
  Exit Sub

ErrorHandler:
  If Err.Number = 9999 Then
    MsgBox "Error: " & Err.Description, vbCritical
  Else
    MsgBox "An unexpected error occurred: " & Err.Number & " - " & Err.Description, vbCritical
  End If
  Err.Clear
End Sub