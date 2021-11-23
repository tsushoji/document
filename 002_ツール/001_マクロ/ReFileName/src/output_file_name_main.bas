Attribute VB_Name = "output_file_name_main"
Sub OutputFileName()

'�ُ펞�A�Ή�
On Error GoTo errHandler

Dim strFolderPath As String
Dim strBeforeFilePath As String
Dim strFileName As String
Dim j As Integer

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

'�ύX�O�t�@�C���p�X�擾
strBeforeFilePath = strFolderPath & constant.strDollarMark & constant.strAsteriskMark & constant.strDotMark & constant.strAsteriskMark

'�ΏۃZ���Ƀt�@�C�������o��
strFileName = Dir(strBeforeFilePath)
j = constant.intRowsNum1

Do While strFileName <> constant.strBlank

j = j + 1
Sheet2.Cells(j, Column.BeforeChangeFileName) = strFileName
strFileName = Dir()

Loop

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A�Ή�
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub


