Attribute VB_Name = "Module1"
Sub GetOutputFolder()
    ' フォルダ名を取得
    Dim outputFolder As Variant
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = False Then ' キャンセルボタン押下時
            Exit Sub
        End If
        outputFolder = .SelectedItems(1)
    End With
    
    ' ファイル名を出力
    Range("B5") = outputFolder
End Sub
