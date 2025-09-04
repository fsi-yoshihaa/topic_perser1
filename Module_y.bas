Attribute VB_Name = "GetFileName"
Public Target_cell_for_output_file As String
Sub GetOutputFolder()
    Target_cell_for_import_file = "B10"
    ' フォルダ名を取得
    Dim outputFolder As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then ' キャンセルボタン押下時
            Exit Sub
        End If
        outputFolder = .SelectedItems(1)
    End With
    
    ' ファイル名を出力
    Range(Target_cell_for_import_file) = outputFolder
End Sub

