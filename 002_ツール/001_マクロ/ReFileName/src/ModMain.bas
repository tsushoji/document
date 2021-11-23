Attribute VB_Name = "ModMain"
'20190612 辻追加 Start
'ファイル名マスタシート列定義
Public Enum Column

BeforeChangeFileName = 1
AfterChangeFileName = 2
Error = 3

End Enum
'20190612 辻追加 End

Sub FileNameChange()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String
Dim strBeforeFilePath As String
Dim strAfterFilePath As String
Dim strMessage As String
Dim intCount As Integer
Dim i As Integer
Dim blnErrFlag As Boolean

'エラーフラグにTrueをセット
blnErrFlag = True

'20190612 辻追加 Start
With Sheet1
'20190612 辻追加 End

'20190612 辻修正 Start
'メインシートにて、名前を変更するファイルのフォルダパスを指定しているか確認
'If Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName1) = ModDeclare.strBlank Then
If .Range("FilePath") = ModDeclare.strBlank Then
'20190612 辻修正 End

    '20190612 辻修正 Start
    'エラーメッセージを出力
    'Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName2) = ModDeclare.strErrMessage5
    .Range("Message") = ModDeclare.strErrMessage5
    '20190612 辻修正 End
    'フォルダパスが指定されていないため、処理終了
    Exit Sub

Else

    '20190612 辻修正 Start
    '名前を変更するファイルのフォルダパスを取得
    'strFolderPath = Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName1)
    strFolderPath = .Range("FilePath")
    '20190612 辻修正 End

End If

'20190612 辻追加 Start
End With
'20190612 辻追加 End

'20190612 辻修正 Start
'With Sheets(ModDeclare.Sheet2)
With Sheet2

'ファイル名マスタの最終行を取得
'intCount = .Cells(Rows.Count, ModDeclare.intColumnsNum1).End(xlUp).Row
intCount = .Cells(Rows.Count, Column.BeforeChangeFileName).End(xlUp).Row
'20190612 辻修正 End

'ファイル名マスタにデータがあるか確認
If intCount > 1 Then

    'ファイル名変更
    For i = 2 To intCount
    
        '20190612 辻修正 Start
        '変更前ファイルパス取得
        'strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, ModDeclare.intColumnsNum1)
        strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.BeforeChangeFileName)
        '変更後ファイルパス取得
        'strAfterFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, ModDeclare.intColumnsNum2)
        strAfterFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.AfterChangeFileName)
           
            'ファイル名マスタにて、ファイル名を指定しているか確認
            'If .Cells(i, ModDeclare.intColumnsNum1) <> ModDeclare.strBlank And .Cells(i, ModDeclare.intColumnsNum2) <> ModDeclare.strBlank Then
            If .Cells(i, Column.BeforeChangeFileName) <> ModDeclare.strBlank And .Cells(i, Column.AfterChangeFileName) <> ModDeclare.strBlank Then
            '20190612 辻修正 End
            
                'ファイル名マスタにて、指定したファイル名がフォルダに存在するか確認
                If Dir(strBeforeFilePath) <> ModDeclare.strBlank Then
                
                    '存在する場合、ファイル名変更
                    Name strBeforeFilePath As strAfterFilePath
                    
                    '20190612 辻修正 Start
                    '存在しない場合、ファイル名マスタに"OK"を書き出す
                    '.Cells(i, ModDeclare.intColumnsNum3) = ModDeclare.strEvaluation1
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation1
                    '20190612 辻修正 End
                
                Else
                    
                    '20190612 辻修正 Start
                    '.Cells(i, ModDeclare.intColumnsNum3) = ModDeclare.strEvaluation2
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                    '20190612 辻修正 End
                    'エラーフラグにFalseをセット
                    blnErrFlag = False
                
                End If
            
            Else
            
                '20190612 辻修正 Start
                '指定していない場合、ファイル名マスタに"NG"を書き出す
                .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                '20190612 辻修正 End
                'エラーフラグにFalseをセット
                blnErrFlag = False
            
            End If
            
    Next

Else
 
    'エラーフラグにFalseをセット
    blnErrFlag = False

End If

End With

'20190612 辻修正 Start
'With Sheets(ModDeclare.Sheet1)
With Sheet1
'20190612 辻修正 End

'メインシートに処理結果出力
If blnErrFlag Then

    '20190612 辻修正 Start
    'エラーフラグがTrueの場合
    '.Range(ModDeclare.strObjName2) = ModDeclare.strMessage1
    .Range("Message") = ModDeclare.strMessage1
    '20190612 辻修正 End

Else

    '20190612 辻修正 Start
    'エラーフラグがFalseの場合
    '.Range(ModDeclare.strObjName2) = ModDeclare.strErrMessage1 & vbLf & ModDeclare.strErrMessage2
    .Range("Message") = ModDeclare.strErrMessage1 & vbLf & ModDeclare.strErrMessage2
    '20190612 辻修正 End

End If

End With

'異常でない場合、処理終了
Exit Sub

'異常時、メッセージボックス
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub


