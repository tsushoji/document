Attribute VB_Name = "AtCoderSrcFileSub"
Option Explicit
Function ValidateCreateAtCoderSourceFileInput(ByVal srcFileLang As String, ByVal targetPageURL As String, ByVal srcFileLangAry As Variant) As Boolean
  ValidateCreateAtCoderSourceFileInput = False
  
  If CommonFunctionMain.IsContainStrAry(srcFileLangAry, srcFileLang) = False Then
    Exit Function
  End If
  
  If Len(targetPageURL) = 0 Then
    Exit Function
  End If
  
  If CommonFunctionMain.IsRegMatch(targetPageURL, CommonConstant.TARGET_PAGE_URL_REG_PATTERN) = False Then
    Exit Function
  End If
  
  ValidateCreateAtCoderSourceFileInput = True

End Function

Function ScrapWeb_Chrome(ByVal targetPageURL As String, ByRef atCoderSrcFile() As AtCoderSrcFileMain.AtCoderSrcFileInfo, ByRef atCoderSrcFileLastIndex As Long, ByVal manageNum As Long) As Boolean
  ScrapWeb_Chrome = False

  Dim driver As New Selenium.WebDriver
  
  driver.Start "Chrome"
  driver.Get targetPageURL
  
  Dim targetElements As WebElements
  Set targetElements = driver.FindElementsByTag("table")
  
  Dim targetElement As WebElement
  Dim tbodyElement As WebElement
  Dim trElements As WebElements
  
  Dim i As Long
  
  Dim isExistElement As Boolean
  isExistElement = False
  
  Dim atCoderProblemDataCount As Long
  
  For i = 1 To targetElements.Count
    Set targetElement = targetElements.Item(i)
    If targetElement.Attribute("class") = CommonConstant.ATCODER_PROBLEM_DATA_TABLE_CLASS_NAME Then
      Set tbodyElement = targetElement.FindElementByTag("tbody")
      Set trElements = tbodyElement.FindElementsByTag("tr")
      atCoderProblemDataCount = trElements.Count
      isExistElement = True
    End If
  Next
  
  If isExistElement = False Or atCoderProblemDataCount < 1 Then
    Exit Function
  End If
  
  Dim trElement As WebElement
  Dim tdElements As WebElements
  
  ReDim atCoderSrcFile(atCoderProblemDataCount - 1)
  
  Dim atCoderSrcFileIndex As Long
  atCoderSrcFileIndex = -1
  
  Dim problemNum As String
  Dim problemName As String
  Dim problemURL As String
  
  Dim problemNumElement As WebElement
  Dim problemNameElement As WebElement
  
  For i = 1 To atCoderProblemDataCount
    Set trElement = trElements.Item(i)
    Set tdElements = trElement.FindElementsByTag("td")
    Set problemNumElement = tdElements.Item(1).FindElementByTag("a")
    Set problemNameElement = tdElements.Item(2).FindElementByTag("a")
    
    problemNum = problemNumElement.Text
    problemName = problemNameElement.Text
    problemURL = problemNameElement.Attribute("href")
    
    atCoderSrcFileIndex = atCoderSrcFileIndex + 1
    
    atCoderSrcFile(atCoderSrcFileIndex).manageNum = manageNum
    atCoderSrcFile(atCoderSrcFileIndex).problemNum = problemNum
    atCoderSrcFile(atCoderSrcFileIndex).problemName = problemName
    atCoderSrcFile(atCoderSrcFileIndex).problemURL = problemURL
    
    manageNum = manageNum + 1
  Next
  
  atCoderSrcFileLastIndex = atCoderSrcFileIndex
  
  ScrapWeb_Chrome = True
  
End Function

Sub PostAtCoderProblemDataToExcel(ByRef WSheet As Worksheet, ByVal startRowNum As Long, ByRef atCoderSrcFile() As AtCoderSrcFileMain.AtCoderSrcFileInfo, ByVal atCoderSrcFileLastIndex As Long)
  Dim i As Long
  
  For i = 0 To atCoderSrcFileLastIndex
    WSheet.Cells(startRowNum, 2).Value = atCoderSrcFile(i).manageNum
    WSheet.Cells(startRowNum, 3).Value = atCoderSrcFile(i).problemNum
    WSheet.Cells(startRowNum, 4).Value = atCoderSrcFile(i).problemName
    WSheet.Cells(startRowNum, 5).Value = atCoderSrcFile(i).problemURL
    
    startRowNum = startRowNum + 1
  Next
  
End Sub

Sub CreateAtCoderSrcFile_CSharp(ByRef WSheet As Worksheet, ByVal startRowNum As Long, ByVal nameSpace As String)
  Dim rowLastNum As Long
  rowLastNum = Cells(Rows.Count, CommonConstant.MANAGE_NUM_COL_NUM).End(xlUp).Row
  
  Dim i As Long
  
  Dim isCompleted As String
  
  Dim problemNum As String
  Dim problemName As String
  Dim problemURL As String
  
  Dim srcFilePath As String
  
  Dim fso As FileSystemObject
  Set fso = New FileSystemObject
  
  Dim ads As ADODB.Stream
  Set ads = New ADODB.Stream
  
  Dim srcFileTemplateStr As String
  
  ChDir ThisWorkbook.Path
  
  ads.Charset = "UTF-8"
  ads.Open
  ads.LoadFromFile CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_PATH
  srcFileTemplateStr = ads.ReadText(-1) ' Ç∑Ç◊Çƒì«Ç›çûÇ›
  ads.Close
  
  Dim srcFileStr As String
  
  Dim srcFileTemplateFilePathExtension As String
  srcFileTemplateFilePathExtension = fso.GetExtensionName(CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_PATH)
  
  For i = startRowNum To rowLastNum
    isCompleted = WSheet.Cells(i, 6).Value
    If isCompleted <> "ÅZ" Then
      WSheet.Cells(i, 6).Value = "Å~"
      
      problemNum = WSheet.Cells(i, 3).Value
      problemName = WSheet.Cells(i, 4).Value
      problemURL = WSheet.Cells(i, 5).Value
      
      srcFilePath = CommonConstant.ATCODER_SRC_FILE_FOLDER_PATH & "\" & problemNum & "." & srcFileTemplateFilePathExtension
      
      If CommonFunctionMain.IsExistFile(fso, srcFilePath) Then
        fso.DeleteFile (srcFilePath)
      End If
      
      srcFileStr = Replace(srcFileTemplateStr, CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_REPLACE_FIRST_STR, nameSpace)
      srcFileStr = Replace(srcFileStr, CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_REPLACE_SECOND_STR, problemName)
      srcFileStr = Replace(srcFileStr, CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_REPLACE_THIRD_STR, problemURL)
      srcFileStr = Replace(srcFileStr, CommonConstant.ATCODER_SRC_FILE_TEMPLATE_FILE_REPLACE_FORTH_STR, problemNum)
      
      ChDir ThisWorkbook.Path
  
      ads.Charset = "UTF-8"
      ads.Open
      ads.WriteText srcFileStr
      ads.SaveToFile srcFilePath, 2  ' ÉtÉ@ÉCÉãÇ™ë∂ç›Ç∑ÇÈèÍçáÇÕÅAè„èëÇ´
      ads.Close
      
      WSheet.Cells(i, 6).Value = "ÅZ"
    End If
  Next
  
  Set ads = Nothing
  Set fso = Nothing
  
End Sub

