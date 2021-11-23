Attribute VB_Name = "output_file_name_main"
Sub OutputFileName()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String

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

Dim fileSystemObject As fileSystemObject
Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

If (fileSystemObject.FolderExists(strFolderPath) = False) Then
    'エラーメッセージを出力
    .Range("Message") = constant.strErrMessage6
    'フォルダパスが存在していないため、処理終了
    Exit Sub
End If

End With

Dim folder As folder
Set folder = fileSystemObject.GetFolder(strFolderPath)

Dim file As file

Dim j As Integer
j = constant.intRowsNum1

For Each file In folder.Files
    j = j + 1
    '変更前ファイルカラムに取得したファイル名を取得
    Sheet2.Cells(j, Column.BeforeChangeFileName) = file.Name
Next

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub

