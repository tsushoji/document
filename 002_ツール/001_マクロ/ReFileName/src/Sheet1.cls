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

'対象セルにフォルダパスを出力
Range("FilePath") = strFolderPath

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

With Sheet1

'メインシートにて、名前を変更するファイルのフォルダパスを指定しているか確認
If .Range("FilePath") = ModDeclare.strBlank Then
    'エラーメッセージを出力
    .Range("Message") = ModDeclare.strErrMessage5
    'フォルダパスが指定されていないため、処理終了
    Exit Sub

Else

    '名前を変更するファイルのフォルダパスを取得
    strFolderPath = .Range("FilePath")

End If

End With

'変更前ファイルパス取得
strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & ModDeclare.strAsteriskMark & ModDeclare.strDotMark & ModDeclare.strAsteriskMark

'対象セルにファイル名を出力
strFileName = Dir(strBeforeFilePath)
j = ModDeclare.intRowsNum1

Do While strFileName <> ModDeclare.strBlank

j = j + 1
Sheet2.Cells(j, Column.BeforeChangeFileName) = strFileName
strFileName = Dir()

Loop

'異常でない場合、処理終了
Exit Sub

'異常時、対応
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub

