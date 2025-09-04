Attribute VB_Name = "GetFileName"
Sub GetinputFile()

    ' ファイル名を取得
    Dim inputFile
    With Application.FileDialog(msoFileDialogFilePicker)
        If .Show = False Then
            Exit Sub
        End If
        inputFile = .SelectedItems(1)
    End With
    
    ' ファイル名を出力
    Range("B2") = inputFile
End Sub
