Attribute VB_Name = "ModMain"
'20190612 �Ғǉ� Start
'�t�@�C�����}�X�^�V�[�g���`
Public Enum Column

BeforeChangeFileName = 1
AfterChangeFileName = 2
Error = 3

End Enum
'20190612 �Ғǉ� End

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

'20190612 �Ғǉ� Start
With Sheet1
'20190612 �Ғǉ� End

'20190612 �ҏC�� Start
'���C���V�[�g�ɂāA���O��ύX����t�@�C���̃t�H���_�p�X���w�肵�Ă��邩�m�F
'If Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName1) = ModDeclare.strBlank Then
If .Range("FilePath") = ModDeclare.strBlank Then
'20190612 �ҏC�� End

    '20190612 �ҏC�� Start
    '�G���[���b�Z�[�W���o��
    'Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName2) = ModDeclare.strErrMessage5
    .Range("Message") = ModDeclare.strErrMessage5
    '20190612 �ҏC�� End
    '�t�H���_�p�X���w�肳��Ă��Ȃ����߁A�����I��
    Exit Sub

Else

    '20190612 �ҏC�� Start
    '���O��ύX����t�@�C���̃t�H���_�p�X���擾
    'strFolderPath = Sheets(ModDeclare.Sheet1).Range(ModDeclare.strObjName1)
    strFolderPath = .Range("FilePath")
    '20190612 �ҏC�� End

End If

'20190612 �Ғǉ� Start
End With
'20190612 �Ғǉ� End

'20190612 �ҏC�� Start
'With Sheets(ModDeclare.Sheet2)
With Sheet2

'�t�@�C�����}�X�^�̍ŏI�s���擾
'intCount = .Cells(Rows.Count, ModDeclare.intColumnsNum1).End(xlUp).Row
intCount = .Cells(Rows.Count, Column.BeforeChangeFileName).End(xlUp).Row
'20190612 �ҏC�� End

'�t�@�C�����}�X�^�Ƀf�[�^�����邩�m�F
If intCount > 1 Then

    '�t�@�C�����ύX
    For i = 2 To intCount
    
        '20190612 �ҏC�� Start
        '�ύX�O�t�@�C���p�X�擾
        'strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, ModDeclare.intColumnsNum1)
        strBeforeFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.BeforeChangeFileName)
        '�ύX��t�@�C���p�X�擾
        'strAfterFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, ModDeclare.intColumnsNum2)
        strAfterFilePath = strFolderPath & ModDeclare.strDollarMark & .Cells(i, Column.AfterChangeFileName)
           
            '�t�@�C�����}�X�^�ɂāA�t�@�C�������w�肵�Ă��邩�m�F
            'If .Cells(i, ModDeclare.intColumnsNum1) <> ModDeclare.strBlank And .Cells(i, ModDeclare.intColumnsNum2) <> ModDeclare.strBlank Then
            If .Cells(i, Column.BeforeChangeFileName) <> ModDeclare.strBlank And .Cells(i, Column.AfterChangeFileName) <> ModDeclare.strBlank Then
            '20190612 �ҏC�� End
            
                '�t�@�C�����}�X�^�ɂāA�w�肵���t�@�C�������t�H���_�ɑ��݂��邩�m�F
                If Dir(strBeforeFilePath) <> ModDeclare.strBlank Then
                
                    '���݂���ꍇ�A�t�@�C�����ύX
                    Name strBeforeFilePath As strAfterFilePath
                    
                    '20190612 �ҏC�� Start
                    '���݂��Ȃ��ꍇ�A�t�@�C�����}�X�^��"OK"�������o��
                    '.Cells(i, ModDeclare.intColumnsNum3) = ModDeclare.strEvaluation1
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation1
                    '20190612 �ҏC�� End
                
                Else
                    
                    '20190612 �ҏC�� Start
                    '.Cells(i, ModDeclare.intColumnsNum3) = ModDeclare.strEvaluation2
                    .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                    '20190612 �ҏC�� End
                    '�G���[�t���O��False���Z�b�g
                    blnErrFlag = False
                
                End If
            
            Else
            
                '20190612 �ҏC�� Start
                '�w�肵�Ă��Ȃ��ꍇ�A�t�@�C�����}�X�^��"NG"�������o��
                .Cells(i, Column.Error) = ModDeclare.strEvaluation2
                '20190612 �ҏC�� End
                '�G���[�t���O��False���Z�b�g
                blnErrFlag = False
            
            End If
            
    Next

Else
 
    '�G���[�t���O��False���Z�b�g
    blnErrFlag = False

End If

End With

'20190612 �ҏC�� Start
'With Sheets(ModDeclare.Sheet1)
With Sheet1
'20190612 �ҏC�� End

'���C���V�[�g�ɏ������ʏo��
If blnErrFlag Then

    '20190612 �ҏC�� Start
    '�G���[�t���O��True�̏ꍇ
    '.Range(ModDeclare.strObjName2) = ModDeclare.strMessage1
    .Range("Message") = ModDeclare.strMessage1
    '20190612 �ҏC�� End

Else

    '20190612 �ҏC�� Start
    '�G���[�t���O��False�̏ꍇ
    '.Range(ModDeclare.strObjName2) = ModDeclare.strErrMessage1 & vbLf & ModDeclare.strErrMessage2
    .Range("Message") = ModDeclare.strErrMessage1 & vbLf & ModDeclare.strErrMessage2
    '20190612 �ҏC�� End

End If

End With

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A���b�Z�[�W�{�b�N�X
errHandler:

MsgBox ModDeclare.strErrMessage3 & vbLf & ModDeclare.strErrMessage4 & Err.Description

End Sub


