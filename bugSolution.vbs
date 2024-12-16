Function correctedAdd(a, b)
  'Explicit type checking prevents type mismatch errors
  If IsNumeric(a) And IsNumeric(b) Then
    Dim sum
    sum = CDbl(a) + CDbl(b) 'Use CDbl for safer numeric operations
    correctedAdd = sum
  Else
    correctedAdd = "Error: Inputs must be numeric"
  End If
End Function

'Early binding with object creation avoids late-binding errors
Dim objFSO
Set objFSO = CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists("test.txt") Then
  WScript.Echo "File exists!"
Else
  WScript.Echo "File does not exist."
End If

Set objFSO = Nothing 'Properly release the object