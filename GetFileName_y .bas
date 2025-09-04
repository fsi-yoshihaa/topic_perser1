Attribute VB_Name = "Module1"
Public Target_cell_for_import_file As String
Sub GetinputFile()
    Target_cell_for_import_file = "B7"
    ' ファイル名を取得
    Dim inputFile As Variant
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = False Then
            Exit Sub
        End If
        inputFile = .SelectedItems(1)
    End With
    
    ' ファイル名を出力
    Range(Target_cell_for_import_file) = inputFile
End Sub
