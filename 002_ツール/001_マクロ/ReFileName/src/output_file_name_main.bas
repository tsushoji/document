Attribute VB_Name = "output_file_name_main"
Sub OutputFileName()

'�ُ펞�A�Ή�
On Error GoTo errHandler

Dim strFolderPath As String

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

Dim fileSystemObject As fileSystemObject
Set fileSystemObject = CreateObject("Scripting.FileSystemObject")

If (fileSystemObject.FolderExists(strFolderPath) = False) Then
    '�G���[���b�Z�[�W���o��
    .Range("Message") = constant.strErrMessage6
    '�t�H���_�p�X�����݂��Ă��Ȃ����߁A�����I��
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
    '�ύX�O�t�@�C���J�����Ɏ擾�����t�@�C�������擾
    Sheet2.Cells(j, Column.BeforeChangeFileName) = file.Name
Next

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A�Ή�
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub

