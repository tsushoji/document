Attribute VB_Name = "rename_file_name_main"
'ファイル名マスタシート列定義
Public Enum Column

BeforeChangeFileName = 1
AfterChangeFileName = 2
Error = 3

End Enum

Sub FileNameChange()

'異常時、対応
On Error GoTo errHandler

Dim strFolderPath As String
Dim strMessage As String
Dim blnErrFlag As Boolean
                
'エラーフラグにTrueをセット
blnErrFlag = True

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

Dim intCount As Integer
Dim i As Integer
Dim strBeforeFilePath As String
Dim strAfterFilePath As String
Dim strAfterFileName As String

With Sheet2

'ファイル名マスタの最終行を取得
intCount = .Cells(Rows.Count, Column.BeforeChangeFileName).End(xlUp).Row

'ファイル名マスタにデータがあるか確認
If intCount > 1 Then

    Dim fileSystemObject As fileSystemObject
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    Dim backupFolder As String
    Dim afterChangeFolder As String
    
    'ファイル名変更
    For i = 2 To intCount
    
        '変更前ファイルパス取得
        strBeforeFilePath = fileSystemObject.BuildPath(strFolderPath, .Cells(i, Column.BeforeChangeFileName))
        '変更後ファイルパス取得
        strAfterFilePath = fileSystemObject.BuildPath(strFolderPath, .Cells(i, Column.AfterChangeFileName))
           
            'ファイル名マスタにて、ファイル名を指定しているか確認
            If .Cells(i, Column.BeforeChangeFileName) <> constant.strBlank And .Cells(i, Column.AfterChangeFileName) <> constant.strBlank Then
            
                '変更後ファイルパスにフォルダパスが指定されていないか確認
                If InStr(fileSystemObject.GetFileName(strAfterFilePath), constant.strDotMark) = 0 Then
                 
                    .Cells(i, Column.Error) = constant.strEvaluation2
                    'エラーフラグにFalseをセット
                    blnErrFlag = False
                    GoTo file_name_change_rename_continue
                
                End If
                
                'ファイル名マスタにて、指定したファイル名がフォルダに存在するか確認
                If fileSystemObject.FileExists(strBeforeFilePath) Then
                    
                    'バックアップファイル作成
                    backupFolder = fileSystemObject.BuildPath(strFolderPath, constant.strBackupPathName)
                    
                    If fileSystemObject.FolderExists(backupFolder) = False Then
                        fileSystemObject.CreateFolder (backupFolder)
                    End If
                    
                    FileCopy strBeforeFilePath, fileSystemObject.BuildPath(backupFolder, .Cells(i, Column.BeforeChangeFileName))
                
                    afterChangeFolder = fileSystemObject.GetParentFolderName(strAfterFilePath)
                    If fileSystemObject.FolderExists(afterChangeFolder) = False Then
                        fileSystemObject.CreateFolder (afterChangeFolder)
                    End If
                    
                    'ファイル名変更
                    'コピー(上書き)
                    FileCopy strBeforeFilePath, strAfterFilePath
                    '元ファイル削除
                    Kill strBeforeFilePath
                    
                    '成功した場合、ファイル名マスタに"OK"を書き出す
                    .Cells(i, Column.Error) = constant.strEvaluation1
                
                Else
                    
                    .Cells(i, Column.Error) = constant.strEvaluation2
                    'エラーフラグにFalseをセット
                    blnErrFlag = False
                    GoTo file_name_change_rename_continue
                
                End If
            
            Else
            
                '指定していない場合、ファイル名マスタに"NG"を書き出す
                .Cells(i, Column.Error) = constant.strEvaluation2
                'エラーフラグにFalseをセット
                blnErrFlag = False
                GoTo file_name_change_rename_continue
            
            End If
            
file_name_change_rename_continue:
    
    Next

Else
 
    'エラーフラグにFalseをセット
    blnErrFlag = False

End If

End With

With Sheet1

'メインシートに処理結果出力
If blnErrFlag Then

    'エラーフラグがTrueの場合
    .Range("Message") = constant.strMessage1

Else

    'エラーフラグがFalseの場合
    .Range("Message") = constant.strErrMessage1 & vbLf & constant.strErrMessage2

End If

End With

'異常でない場合、処理終了
Exit Sub

'異常時、メッセージボックス
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub



