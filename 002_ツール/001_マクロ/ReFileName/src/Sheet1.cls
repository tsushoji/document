VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub GetFolderPath()

Call GetDialogFolderPath

End Sub
Sub GetFileName()

Call OutputFileName

End Sub

Private Sub GetDialogFolderPath()

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

'20190612 辻修正 Start
'対象セルにフォルダパスを出力
'Range(ModDeclare.strObjName1) = strFolderPath
Range("FilePath") = strFolderPath
'20190612 辻修正 End

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub
Private Sub OutputFileName()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String
Dim strBeforeFilePath As String
Dim strFileName As String
Dim j As Integer

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

'変更前ファイルパス取得
strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & ModDeclare.strAsteriskMark & ModDeclare.strDotMark & ModDeclare.strAsteriskMark

'対象セルにファイル名を出力
strFileName = Dir(strBeforeFilePath)
j = ModDeclare.intRowsNum1

Do While strFileName <> ModDeclare.strBlank

j = j + 1
'20190612 辻修正 Start
'Sheets(ModDeclare.Sheet2).Cells(j, ModDeclare.intColumnsNum1) = strFileName
Sheet2.Cells(j, Column.BeforeChangeFileName) = strFileName
'20190612 辻修正 End
strFileName = Dir()

Loop

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub

