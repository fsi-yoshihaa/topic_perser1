Attribute VB_Name = "Module1"
Public Target_cell_for_import_file As String
Sub GetinputFile()
    Target_cell_for_import_file = "B7"
    ' �t�@�C�������擾
    Dim inputFile As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = False Then
            Exit Sub
        End If
        inputFile = .SelectedItems(1)
    End With
    
    ' �t�@�C�������o��
    Range(Target_cell_for_import_file) = inputFile
End Sub
