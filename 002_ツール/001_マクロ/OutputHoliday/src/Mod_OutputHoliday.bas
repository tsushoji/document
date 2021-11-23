Attribute VB_Name = "Mod_OutputHoliday"
Option Explicit

Sub OutputHoliday()
    Dim macroWB As Workbook
    Dim tarWB As Workbook
    Dim macroWBHolidaySheet As Worksheet
    Dim tarWBSheet As Worksheet
    Dim fso As New Scripting.FileSystemObject
    Dim tarDate As String
    Dim tarYear As String
    Dim tarDay As String
    Dim tarWorkBookPath As String
    Dim tarWorkSheetName As String
    Dim tarOutputType As String
    Dim convertTarWorkBookPath As String
    Dim tarEndColNum As String
    Dim writeRowNum As Long: writeRowNum = 1
    Dim convertTarRange As String
    Dim convertCsvStr As String
    
    Set macroWB = ThisWorkbook
    
    On Error GoTo ErrorHandler
    
    With macroWB.Sheets(MACRO_IUPUT_FORM_SHEET_NAME)
        tarYear = .Range(IUPUT_YEAR_RANGE).Value
        tarDay = .Range(IUPUT_DAY_RANGE).Value
        tarWorkBookPath = .Range(IUPUT_TARGET_PATH_RANGE).Value
        tarWorkSheetName = .Range(IUPUT_TARGET_SHEET_NAME_RANGE).Value
        tarOutputType = .Range(IUPUT_TARGET_OUTPUT_TYPE_RANGE).Value
    End With
    
    If IsEmpty(tarYear) Or IsEmpty(tarYear) Or IsEmpty(tarWorkBookPath) Or IsEmpty(tarWorkSheetName) Or IsEmpty(tarOutputType) Then
        MsgBox ERR_INPUT_REQUIRED_ITEM_MSG
        Exit Sub
    End If
    
    If Not tarOutputType = OUTPUT_TYPE_XLSX And Not tarOutputType = OUTPUT_TYPE_XLSX_AND_CSV Then
        MsgBox ERR_INPUT_OUTPUT_TYPE_MSG
        Exit Sub
    End If
    
    tarDate = tarYear & "/" & tarDay & "/1"
    
    If Not IsDate(tarDate) Then
        MsgBox ERR_INPUT_DATE_FORMAT_MSG
        Exit Sub
    End If
    
    Set macroWBHolidaySheet = macroWB.Sheets(MACRO_IUPUT_DATA_INIT_SHEET_NAME & tarYear)
    
    If Not fso.FolderExists(fso.GetParentFolderName(tarWorkBookPath)) Then
        MsgBox ERR_INPUT_PARENT_FOLDER_MSG
        Exit Sub
    End If
    
    If Not fso.GetExtensionName(tarWorkBookPath) = OUTPUT_TYPE_XLSX Then
        MsgBox ERR_INPUT_EXTENSION_MSG
        Exit Sub
    End If
    
    If Dir(tarWorkBookPath) = "" Then
        Set tarWB = Workbooks.Add
        Set tarWBSheet = tarWB.Sheets("Sheet1")
    Else
        If MsgBox(CHECK_APPEND_OUTPUT_MSG, vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        Set tarWB = Workbooks.Open(tarWorkBookPath)
        If Mod_Common.isExistWorksheet(tarWB, tarWorkSheetName) Then
            Set tarWBSheet = tarWB.Sheets(tarWorkSheetName)
        End If
        tarWBSheet.Cells.Clear
    End If
    Application.DisplayAlerts = False
    
    Call writeHeader(tarWBSheet, writeRowNum)
    Call writeBusinessHoliday(tarDate, tarWBSheet, writeRowNum)
    Call writeHoliday(tarDate, tarWBSheet, macroWBHolidaySheet, writeRowNum)
    tarEndColNum = tarWBSheet.Cells(1, Columns.Count).End(xlToLeft).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    If writeRowNum > 2 Then
        Call sortWritingRowInAscending(tarWBSheet, writeRowNum, tarEndColNum)
        Call deleteDuplicateDateData(tarWBSheet, writeRowNum, tarEndColNum)
        ' ソート、重複データ削除完了後番号行にデータ番号を書き込む
        Call writeNumRow(tarWBSheet)
    End If
    
    tarWBSheet.Range("A1:" & tarEndColNum).EntireColumn.AutoFit
    
    tarWB.SaveAs (tarWorkBookPath)
    
    If tarOutputType = OUTPUT_TYPE_XLSX_AND_CSV And writeRowNum > 1 Then
        convertTarWorkBookPath = Left(tarWorkBookPath, InStrRev(tarWorkBookPath, XLSX_EXTENSION) - 1) & CSV_EXTENSION
        If Not Dir(convertTarWorkBookPath) = "" Then
            If MsgBox(CHECK_APPEND_OUTPUT_CSV_MSG, vbYesNo + vbQuestion) = vbNo Then
                Exit Sub
            End If
        End If
        convertTarRange = "A1:" & Left(tarEndColNum, Len(tarEndColNum) - 1) & writeRowNum - 1
        ' 改行コードLF
        convertCsvStr = getConvertCsvStr(tarWBSheet, convertTarRange)
        ' BOMなしUTF-8で書き込む
        Call writeCsv(convertCsvStr, convertTarWorkBookPath)
    End If
    
    Application.DisplayAlerts = True
    
Finally:
    tarWB.Close
    If Not tarWBSheet Is Nothing Then
        Set tarWBSheet = Nothing
    End If
    If Not tarWB Is Nothing Then
        Set tarWB = Nothing
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print (Err.Number)
    GoTo Finally
End Sub

Sub writeHeader(ByRef tarWBSheet As Worksheet, ByRef rowNum As Long)
    Dim headerNameAry() As Variant
    Dim i As Long
    
    On Error GoTo ErrorHandler
    
    headerNameAry = Array(FIRST_HEADER_NAME, SECOND_HEADER_NAME, THIRD_HEADER_NAME)
    
    For i = 0 To UBound(headerNameAry)
        tarWBSheet.Cells(rowNum, i + 1).Value = headerNameAry(i)
    Next
    
    rowNum = rowNum + 1
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Sub writeBusinessHoliday(ByVal tarDate As String, ByRef tarWBSheet As Worksheet, ByRef rowNum As Long)
    Dim i As Long
    Dim tmpDate As String
    Dim lastDay As Date
    Dim tarDateVal As String
    
    On Error GoTo ErrorHandler
    
    If Month(tarDate) = 12 Then
        tmpDate = (Year(tarDate) + 1) & "/1/1"
    Else
        tmpDate = Format(DateAdd("m", 1, tarDate), "yyyy/m/1")
    End If
    
    lastDay = DateAdd("d", -1, tmpDate)
    
    For i = 1 To Day(lastDay)
        tarDateVal = Year(tarDate) & "/" & Month(tarDate) & "/" & i
        If Weekday(tarDateVal) = vbSunday Or Weekday(tarDateVal) = vbSaturday Then
            With tarWBSheet
                .Cells(rowNum, 2).Value = tarDateVal
                .Cells(rowNum, 3).Value = OUTPUT_BUSINESS_HOLIDAY_NAME
            End With
            rowNum = rowNum + 1
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Sub writeHoliday(ByVal tarDate As String, ByRef tarWBSheet As Worksheet, ByVal macroWBSheet As Worksheet, ByRef rowNum As Long)
    Dim i As Long
    Dim tarVal As String
    Dim cmptarDateMonth As String
    
    On Error GoTo ErrorHandler
    
    ' 形式「dd」で取得
    cmptarDateMonth = Right("00" & Month(tarDate), 2)
    For i = 2 To macroWBSheet.Cells(Rows.Count, 1).End(xlUp).Row
        ' 形式「yyyy/mm/dd」で取得
        tarVal = macroWBSheet.Cells(i, 1).Value
        ' 0埋めで判定する
        If Mid(tarVal, 6, 2) = cmptarDateMonth Then
            With tarWBSheet
                .Cells(rowNum, 2).Value = tarVal
                .Cells(rowNum, 3).Value = macroWBSheet.Cells(i, 2).Value
            End With
            rowNum = rowNum + 1
        End If
    Next
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Sub sortWritingRowInAscending(ByRef tarWBSheet As Worksheet, ByVal rowNum As Long, ByVal endColRange As String)
    Dim cmpColRange As String
    Dim sortRange As String
    Dim i As Long
    Dim tarSerialVal As Long
    
    On Error GoTo ErrorHandler
    
    ' 日付を並び替えるため、日付→シリアル値に変換した列を追加
    For i = 2 To tarWBSheet.Cells(Rows.Count, 2).End(xlUp).Row
        ' 日付→シリアル値
        tarSerialVal = tarWBSheet.Cells(i, 2).Value
        tarWBSheet.Cells(i, 4).Value = tarSerialVal
    Next
    
    cmpColRange = tarWBSheet.Range(Left(endColRange, Len(endColRange) - 1) & "1").Offset(RowOffset:=0, ColumnOffset:=1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    ' 日付から変換したシリアル値で並び替えを行う
    sortRange = "A2" & ":" & Left(cmpColRange, Len(cmpColRange) - 1) & rowNum - 1
    With tarWBSheet.Sort
        With .SortFields
            .Clear
            .Add Key:=Range(cmpColRange), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .SetRange Range(sortRange)
        .Apply
    End With
    
    ' 並び替え用の列を削除
    tarWBSheet.Range(cmpColRange).EntireColumn.Delete
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Sub deleteDuplicateDateData(ByRef tarWBSheet As Worksheet, ByVal rowNum As Long, ByVal endColRange As String)
    Dim deleteRange As String
    
    On Error GoTo ErrorHandler
    
    deleteRange = "A2" & ":" & Left(endColRange, Len(endColRange) - 1) & rowNum - 1
    tarWBSheet.Range(deleteRange).RemoveDuplicates (Array(DELETE_DUPLICATE_DATA_COL_NUM))
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Sub writeNumRow(ByRef tarWBSheet As Worksheet)
    Dim i As Long
    Dim maxRowNum As Long
    
    On Error GoTo ErrorHandler
    
    maxRowNum = tarWBSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    For i = 2 To maxRowNum
        tarWBSheet.Cells(i, 1).Value = i - 1
    Next
    
    Exit Sub
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Sub

Function getConvertCsvStr(ByVal tarWBSheet As Worksheet, ByVal tarRange As String) As String
    Dim outputAry As Variant
    Dim i As Long
    Dim j As Long
    Dim trgtVal As String
    Dim output As String: output = ""
    
    On Error GoTo ErrorHandler
    
    outputAry = tarWBSheet.Range(tarRange)
    
    For i = LBound(outputAry, 1) To UBound(outputAry, 1)
        For j = LBound(outputAry, 2) To UBound(outputAry, 2)
            trgtVal = outputAry(i, j)
            If j = UBound(outputAry, 2) Then
                '最終列の場合
                output = output & trgtVal & vbLf
            Else
                '最終列でない場合
                output = output & trgtVal & ","
            End If
        Next
    Next

    getConvertCsvStr = output
    
    Exit Function
    
ErrorHandler:
    Call Err.Raise(Err.Number)
End Function

Sub writeCsv(ByVal tarVal As String, ByVal tarfilePath As String)
    Dim adoSt As Object
    Dim tmpAdoSt As Object
    Dim tmpByteData As Variant
    
    On Error GoTo ErrorHandler
    
    Set adoSt = CreateObject("ADODB.Stream")
    Set tmpAdoSt = CreateObject("ADODB.Stream")
    
    With tmpAdoSt
        .Charset = "UTF-8"
        .Open
        .WriteText (tarVal)
        ' ストリームの位置を0にする
        .Position = 0
        ' データの種類をバイナリデータに変更
        .Type = 1
        ' ストリームの位置を3にする
        .Position = 3
        
        ' ストリームの内容を一時格納用変数に保存
        tmpByteData = .Read()
    End With
    
    With adoSt
        .Open
        ' データの種類をバイナリデータに変更
        .Type = 1
        .Write (tmpByteData)
        ' ファイルが存在した場合、上書き
        .savetofile tarfilePath, 2
    End With
    
Finally:
    tmpAdoSt.Close
    adoSt.Close
    If Not tmpAdoSt Is Nothing Then
        Set tmpAdoSt = Nothing
    End If
    If Not adoSt Is Nothing Then
        Set adoSt = Nothing
    End If
    
    Exit Sub
    
ErrorHandler:
    Debug.Print (Err.Number)
    GoTo Finally
End Sub
