Attribute VB_Name = "get_folder_path_main"
Sub GetDialogFolderPath()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String

With Application.FileDialog(msoFileDialogFolderPicker)

    'ダイアログボックスを表示
    If .Show Then
    
        'OKボタンが押された場合、フォルダパスを取得
        strFolderPath = .SelectedItems(1)
    
    End If

End With

'対象セルにフォルダパスを出力
Range("FilePath") = strFolderPath

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub
