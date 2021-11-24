Attribute VB_Name = "rename_file_name_main"
'�t�@�C�����}�X�^�V�[�g���`
Public Enum Column

BeforeChangeFileName = 1
AfterChangeFileName = 2
Error = 3

End Enum

Sub FileNameChange()

'�ُ펞�A�Ή�
On Error GoTo errHandler

Dim strFolderPath As String
Dim strMessage As String
Dim blnErrFlag As Boolean
                
'�G���[�t���O��True���Z�b�g
blnErrFlag = True

With Sheet1

'���C���V�[�g�ɂāA���O��ύX����t�@�C���̃t�H���_�p�X���w�肵�Ă��邩�m�F
If .Range("FilePath") = constant.strBlank Then

    '�G���[���b�Z�[�W���o��
    .Range("Message") = constant.strErrMessage5
    '�t�H���_�p�X���w�肳��Ă��Ȃ����߁A�����I��
    Exit Sub

Else

    '���O��ύX����t�@�C���̃t�H���_�p�X���擾
    strFolderPath = .Range("FilePath")
End If

End With

Dim intCount As Integer
Dim i As Integer
Dim strBeforeFilePath As String
Dim strAfterFilePath As String
Dim strAfterFileName As String

With Sheet2

'�t�@�C�����}�X�^�̍ŏI�s���擾
intCount = .Cells(Rows.Count, Column.BeforeChangeFileName).End(xlUp).Row

'�t�@�C�����}�X�^�Ƀf�[�^�����邩�m�F
If intCount > 1 Then

    Dim fileSystemObject As fileSystemObject
    Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
    
    Dim backupFolder As String
    Dim afterChangeFolder As String
    
    '�t�@�C�����ύX
    For i = 2 To intCount
    
        '�ύX�O�t�@�C���p�X�擾
        strBeforeFilePath = fileSystemObject.BuildPath(strFolderPath, .Cells(i, Column.BeforeChangeFileName))
        '�ύX��t�@�C���p�X�擾
        strAfterFilePath = fileSystemObject.BuildPath(strFolderPath, .Cells(i, Column.AfterChangeFileName))
           
            '�t�@�C�����}�X�^�ɂāA�t�@�C�������w�肵�Ă��邩�m�F
            If .Cells(i, Column.BeforeChangeFileName) <> constant.strBlank And .Cells(i, Column.AfterChangeFileName) <> constant.strBlank Then
            
                '�ύX��t�@�C���p�X�Ƀt�H���_�p�X���w�肳��Ă��Ȃ����m�F
                If InStr(fileSystemObject.GetFileName(strAfterFilePath), constant.strDotMark) = 0 Then
                 
                    .Cells(i, Column.Error) = constant.strEvaluation2
                    '�G���[�t���O��False���Z�b�g
                    blnErrFlag = False
                    GoTo file_name_change_rename_continue
                
                End If
                
                '�t�@�C�����}�X�^�ɂāA�w�肵���t�@�C�������t�H���_�ɑ��݂��邩�m�F
                If fileSystemObject.FileExists(strBeforeFilePath) Then
                    
                    '�o�b�N�A�b�v�t�@�C���쐬
                    backupFolder = fileSystemObject.BuildPath(strFolderPath, constant.strBackupPathName)
                    
                    If fileSystemObject.FolderExists(backupFolder) = False Then
                        fileSystemObject.CreateFolder (backupFolder)
                    End If
                    
                    FileCopy strBeforeFilePath, fileSystemObject.BuildPath(backupFolder, .Cells(i, Column.BeforeChangeFileName))
                
                    afterChangeFolder = fileSystemObject.GetParentFolderName(strAfterFilePath)
                    If fileSystemObject.FolderExists(afterChangeFolder) = False Then
                        fileSystemObject.CreateFolder (afterChangeFolder)
                    End If
                    
                    '�t�@�C�����ύX
                    '�R�s�[(�㏑��)
                    FileCopy strBeforeFilePath, strAfterFilePath
                    '���t�@�C���폜
                    Kill strBeforeFilePath
                    
                    '���������ꍇ�A�t�@�C�����}�X�^��"OK"�������o��
                    .Cells(i, Column.Error) = constant.strEvaluation1
                
                Else
                    
                    .Cells(i, Column.Error) = constant.strEvaluation2
                    '�G���[�t���O��False���Z�b�g
                    blnErrFlag = False
                    GoTo file_name_change_rename_continue
                
                End If
            
            Else
            
                '�w�肵�Ă��Ȃ��ꍇ�A�t�@�C�����}�X�^��"NG"�������o��
                .Cells(i, Column.Error) = constant.strEvaluation2
                '�G���[�t���O��False���Z�b�g
                blnErrFlag = False
                GoTo file_name_change_rename_continue
            
            End If
            
file_name_change_rename_continue:
    
    Next

Else
 
    '�G���[�t���O��False���Z�b�g
    blnErrFlag = False

End If

End With

With Sheet1

'���C���V�[�g�ɏ������ʏo��
If blnErrFlag Then

    '�G���[�t���O��True�̏ꍇ
    .Range("Message") = constant.strMessage1

Else

    '�G���[�t���O��False�̏ꍇ
    .Range("Message") = constant.strErrMessage1 & vbLf & constant.strErrMessage2

End If

End With

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A���b�Z�[�W�{�b�N�X
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub



