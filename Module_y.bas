Attribute VB_Name = "GetFileName"
Public Target_cell_for_output_file As String
Sub GetOutputFolder()
    Target_cell_for_import_file = "B10"
    ' �t�H���_�����擾
    Dim outputFolder As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then ' �L�����Z���{�^��������
            Exit Sub
        End If
        outputFolder = .SelectedItems(1)
    End With
    
    ' �t�@�C�������o��
    Range(Target_cell_for_import_file) = outputFolder
End Sub

