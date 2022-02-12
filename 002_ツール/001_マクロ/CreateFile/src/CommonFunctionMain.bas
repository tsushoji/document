Attribute VB_Name = "CommonFunctionMain"
Option Explicit
Function IsContainStrAry(ByVal ary As Variant, ByVal str As String) As Boolean
  IsContainStrAry = False
  
  Dim i As Long
  
  For i = 0 To UBound(ary)
    If StrComp(ary(i), str) = 0 Then
      IsContainStrAry = True
    End If
  Next i

End Function

Function IsRegMatch(ByVal str As String, ByVal pattern As String) As Boolean
  IsRegMatch = False
  
  Dim reg As Object
  Set reg = CreateObject("VBScript.RegExp")
  
  With reg
    .pattern = pattern
    .IgnoreCase = False '‘å•¶š‚Æ¬•¶š‚ğ‹æ•Ê‚·‚é
    .Global = True      '•¶š—ñ‘S‘Ì‚ğŒŸõ‚·‚é
    
    If .test(str) Then
      IsRegMatch = True
    End If
    
  End With
  
  Set reg = Nothing

End Function

Function IsExistFile(ByVal fso As FileSystemObject, ByVal filePath As String) As Boolean
  IsExistFile = False
  
  If fso Is Nothing Then
    Set fso = New FileSystemObject
  End If
  
  If fso.FileExists(filePath) Then
    IsExistFile = True
  End If
  
End Function
