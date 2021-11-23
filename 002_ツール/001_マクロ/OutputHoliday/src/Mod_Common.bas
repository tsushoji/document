Attribute VB_Name = "Mod_Common"
Option Explicit
Function isExistWorksheet(ByVal WB As Workbook, ByVal WBSheetName As String) As Boolean
    Dim WBSheet As Worksheet
    
    On Error GoTo ErrorHandler
    
    For Each WBSheet In WB.Worksheets
        If WBSheet.Name = WBSheetName Then
            isExistWorksheet = True
            Exit Function
        End If
    Next
    isExistWorksheet = False
    
    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Function
