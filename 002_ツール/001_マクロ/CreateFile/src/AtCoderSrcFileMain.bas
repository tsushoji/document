Attribute VB_Name = "AtCoderSrcFileMain"
Option Explicit

Type AtCoderSrcFileInfo
    manageNum As Long
    problemNum As String
    problemName As String
    problemURL As String
End Type

Sub CreateAtCoderSrcFile()
  Dim atCoderSrcFileWSheet As Worksheet
  Set atCoderSrcFileWSheet = Worksheets(CommonConstant.ATCODER_SRC_FILE_SHEET_NAME)
  
  Dim srcFileLangAry() As Variant
  srcFileLangAry = Array(CommonConstant.CREATE_SRC_FILE_LAUNG_CSHARP)
  
  Dim srcFileLang As String
  Dim targetPageURL As String
  
  srcFileLang = atCoderSrcFileWSheet.Range(CommonConstant.CREATE_SRC_FILE_LAUNG_RANGE).Value
  targetPageURL = atCoderSrcFileWSheet.Range(CommonConstant.TARGET_PAGE_URL_RANGE).Value
  
  If AtCoderSrcFileSub.ValidateCreateAtCoderSourceFileInput(srcFileLang, targetPageURL, srcFileLangAry) = False Then
    Exit Sub
  End If
  
  Dim atCoderSrcFile() As AtCoderSrcFileInfo
  
  Dim manageNumRowLastNum As Long
  manageNumRowLastNum = Cells(Rows.Count, CommonConstant.MANAGE_NUM_COL_NUM).End(xlUp).Row
  
  Dim manageNum As Long
  manageNum = manageNumRowLastNum + 1 - 2
  
  Dim atCoderSrcFileLastIndex As Long
  
  If AtCoderSrcFileSub.ScrapWeb_Chrome(targetPageURL, atCoderSrcFile, atCoderSrcFileLastIndex, manageNum) = False Then
    Exit Sub
  End If
  
  If atCoderSrcFileLastIndex < 0 Then
    Exit Sub
  End If
  
  Call AtCoderSrcFileSub.PostAtCoderProblemDataToExcel(atCoderSrcFileWSheet, manageNumRowLastNum + 1, atCoderSrcFile, atCoderSrcFileLastIndex)
  
  Dim nameSpace As String
  nameSpace = atCoderSrcFileWSheet.Range(CommonConstant.NAME_SPACE_RANGE).Value
  
  If srcFileLang = CommonConstant.CREATE_SRC_FILE_LAUNG_CSHARP Then
    Call CreateAtCoderSrcFile_CSharp(atCoderSrcFileWSheet, manageNumRowLastNum + 1, nameSpace)
  End If

End Sub



