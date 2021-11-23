Attribute VB_Name = "ModMain"
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
Dim strBeforeFilePath As String
Dim strAfterFilePath As String
Dim strMessage As String
Dim intCount As Integer
Dim i As Integer
Dim blnErrFlag As Boolean

'�G���[�t���O��True���Z�b�g
blnErrFlag = True

With Sheet1

'���C���V�[�g�ɂāA���O��ύX����t�@�C���̃t�H���_�p�X���w�肵�Ă��邩�m�F
If .Range("FilePath") = ModDeclare.strBlank Then

    '�G���[���b�Z�[�W���o��
    .Range("Message") = ModDeclare.strErrMessage5
    '�t�H���_�p�X���w�肳��Ă��Ȃ����߁A�����I��
    Exit Sub

Else

    '���O��ύX����t�@�C���̃t�H���_�p�X���擾
    strFolderPath = .Range("FilePath")
End If

End With

With Sheet2

'�t�@�C�����}�X�^�̍ŏI�s���擾
intCount = .Cells(Rows.Count, Column.BeforeChangeFileName).End(xlUp).Row

'�t�@�C�����}�X�^�Ƀf�[�^�����邩�m�F
If intCount > 1 Then

    '�t�@�C�����ύX
    For i = 2 To intCount
    
        '�ύX�O�t�@�C���p�X�擾
        strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.BeforeChangeFileName)
        '�ύX��t�@�C���p�X�擾
        strAfterFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.AfterChangeFileName)
           
            '�t�@�C�����}�X�^�ɂāA�t�@�C�������w�肵�Ă��邩�m�F
            If .Cells(i, Column.BeforeChangeFileName) <> ModDeclare.strBlank And .Cells(i, Column.AfterChangeFileName) <> ModDeclare.strBlank Then
            
                '�t�@�C�����}�X�^�ɂāA�w�肵���t�@�C�������t�H���_�ɑ��݂��邩�m�F
                If Dir(strBeforeFilePath) <> ModDeclare.strBlank Then
                
                    '���݂���ꍇ�A�t�@�C�����ύX
                    Name strBeforeFilePath As strAfterFilePath
                    
                    '���݂��Ȃ��ꍇ�A�t�@�C�����}�X�^��"OK"�������o��
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation1
                
                Else
                    
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                    '�G���[�t���O��False���Z�b�g
                    blnErrFlag = False
                
                End If
            
            Else
            
                '�w�肵�Ă��Ȃ��ꍇ�A�t�@�C�����}�X�^��"NG"�������o��
                .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                '�G���[�t���O��False���Z�b�g
                blnErrFlag = False
            
            End If
            
    Next

Else
 
    '�G���[�t���O��False���Z�b�g
    blnErrFlag = False

End If

End With

With Sheet1
'20190612 �ҏC�� End

'���C���V�[�g�ɏ������ʏo��
If blnErrFlag Then

    '�G���[�t���O��True�̏ꍇ
    .Range("Message") = ModDeclare.strMessage1

Else

    '�G���[�t���O��False�̏ꍇ
    .Range("Message") = ModDeclare.strErrMessage1 & vbLf & ModDeclare.strErrMessage2

End If

End With

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A���b�Z�[�W�{�b�N�X
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub


