Attribute VB_Name = "get_folder_path_main"
Sub GetDialogFolderPath()

'�ُ펞�A�Ή�
On Error GoTo errHandler

Dim strFolderPath As String

With Application.FileDialog(msoFileDialogFolderPicker)

    '�_�C�A���O�{�b�N�X��\��
    If .Show Then
    
        'OK�{�^���������ꂽ�ꍇ�A�t�H���_�p�X���擾
        strFolderPath = .SelectedItems(1)
    
    End If

End With

'�ΏۃZ���Ƀt�H���_�p�X���o��
Range("FilePath") = strFolderPath

'�ُ�łȂ��ꍇ�A�����I��
Exit Sub

'�ُ펞�A�Ή�
errHandler:

MsgBox constant.strErrMessage3 & vbLf & constant.strErrMessage4 & Err.Description

End Sub
