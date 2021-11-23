Attribute VB_Name = "output_file_name_main"
Sub OutputFileName()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String
Dim strBeforeFilePath As String
Dim strFileName As String
Dim j As Integer

With Sheet1

'メインシートにて、名前を変更するファイルのフォルダパスを指定しているか確認
If .Range("FilePath") = constant.strBlank Then
    'エラーメッセージを出力
    .Range("Message") = constant.strErrMessage5
    'フォルダパスが指定されていないため、処理終了
    Exit Sub

Else

    '名前を変更するファイルのフォルダパスを取得
    strFolderPath = .Range("FilePath")

End If

End With

'変更前ファイルパス取得
strBeforeFilePath = strFolderPath & constant.strDollarMark & constant.strAsteriskMark & constant.strDotMark & constant.strAsteriskMark

'対象セルにファイル名を出力
strFileName = Dir(strBeforeFilePath)
j = constant.intRowsNum1

Do While strFileName <> constant.strBlank

j = j + 1
Sheet2.Cells(j, Column.BeforeChangeFileName) = strFileName
strFileName = Dir()

Loop

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub


