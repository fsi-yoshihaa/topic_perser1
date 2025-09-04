Attribute VB_Name = "DumpLog"
Sub DumpLog()
    ' 読み込みファイルのパスを取得
    Dim inputFile As String
    inputFile = Range("B2")
    ' 出力先フォルダのパスを取得
    Dim outputFolderPath As String
    outputFolderPath = Range("B5")
    ' 空白を削除する
    inputFile = Replace(inputFile, " ", "")
    inputFile = Replace(inputFile, "　", "")
    
     ' 入力がされていない、指定されたファイルが存在しない場合の処理
    If inputFile = "" Or Dir(inputFile) = "" Then
        MsgBox "読み込みファイルが存在しません。", vbExclamation
        'セルにフォーカスを移動する
        Range("B2").Select
        Exit Sub
    End If
    
    If outputFolderPath = "" Then ' 出力先フォルダが空白の場合
        ' 読み込みファイルと同じフォルダを出力先フォルダとする
        outputFolderPath = Left(inputFile, (InStrRev(inputFile, "\") - 1))
        Range("B5") = outputFolderPath
    End If
    
    Dim inputFileLen As Integer ' 読み込みファイルのパスの長さ
    inputFileLen = Len(inputFile)
    Dim outputFilePath As String ' 出力ファイルのパス
    Dim outputFileName As String ' 出力ファイル名
    outputFileName = makeFileName()
    outputFilePath = outputFolderPath & "\" & outputFileName
    Dim outputSheetName As String ' 出力するシート名
    outputSheetName = "ダンプ"
    Dim fOutputObj As Object
    Dim outputWb As Workbook  ' 出力するワークブック
    Dim outputWs As Worksheet ' 出力するシート
    Dim tempFilePath As String '　新しい出力フォルダー
    
    If Dir(inputFile) <> "" Then ' B2で指定された読み込みファイルが存在する場合
        If Dir(outputFolderPath, vbDirectory) <> "" Then ' B5で指定された出力先フォルダが存在する場合
            ' ブックを新規作成
            Set outputWb = Workbooks.Add
            ' シート名を変更
            Set outputWs = outputWb.Sheets(1)
            outputWs.Name = outputSheetName
            
             '　新しい出力フォルダー
            tempFilePath = outputFolderPath & "\Temp_CRLF.txt"
            ' 元ファイルをコピーして改行コード変換
            Call LfToCrlfCopy(inputFile, tempFilePath)
    
            ' ダンプ処理（元ファイルではなく一時ファイルを使用）
            Call OutputDumpData(tempFilePath, outputWs)

            ' ブックを保存
            outputWb.SaveAs outputFilePath

            ' ブックを閉じる
            outputWb.Close
            
            ' 一時ファイルを削除
            If Dir(tempFilePath) <> "" Then
                 Kill tempFilePath
            End If

            
             ' 正常に出力されたことを示すメッセージを表示!!!!!!!'
            Call ShowCompletionMessage
            
        Else
            ' エラーメッセージを出力
            MsgBox "出力先フォルダが見つかりません", vbExclamation
        End If
        
    Else  ' B2で指定された読み込みファイルが存在しない場合
        ' エラーメッセージを出力
        MsgBox "読み込みファイルが見つかりません", vbExclamation
        
    End If
    
End Sub


Function makeFileName() As String
    ' 日付_時刻を取得
    Dim dateTime
    dateTime = Now()
    
    ' 文字列に変換
    Dim retStr As String
    retStr = Format(dateTime, "yyyymmdd_hmmss")
    
    makeFileName = retStr & "_log_dump.xlsx"
End Function

'#### 出力 ####'
Function OutputDumpData(ByVal tempFilePath As String, ByVal outputWs As Worksheet)
    ' 出力先のシートをクリアする
    outputWs.Cells.Clear
    
    ' 変数宣言
    Dim serchStr As String              ' 出力するデータかを判断する目印
    Dim flgRead As Boolean              ' 読み込みデータフラグ
    Dim readLine As String              ' 読み込んだ列の文字列
    Dim headerList As New Collection    ' ヘッダの文字列のリスト
    Dim lastKey As String               ' 最後のkey（ヘッダの項目）
    Dim clmMax As Integer               ' 最大列数
    Dim row As Integer                  ' 行
    Dim rowHeader As Integer            ' ヘッダを出力する行
    Dim clm As Integer                  ' 列
    Dim clmNo As Integer                ' Noを表示する列
    Dim cntBlock As Integer             ' 出力ブロックのカウント
    Dim cntData As Integer              ' データのカウント
    
    serchStr = "---"
    flgRead = False
    cntBlock = 0
    row = 3         ' 何行目から出力するかを設定
    rowHeader = row
    clmNo = 2       ' 何列目から出力するかを設定
    cntData = 1
    clmMax = clmNo
    
    ' 入力ファイルオープン
    Open tempFilePath For Input As #1
    
    Do Until EOF(1)
        ' 1列ずつ読み込む
        Line Input #1, readLine
        
        If readLine = serchStr Then ' --- の場合
            ' 読み込みフラグをONにする
            flgRead = True
            
            cntBlock = cntBlock + 1
            cntData = 1
            row = row + 1
            clm = clmNo + 1
            
            If cntBlock = 1 Then ' 1つ目のブロックの場合
                ' [No](ヘッダ)を出力
                outputWs.Cells(rowHeader, clmNo).NumberFormatLocal = "@"
                outputWs.Cells(rowHeader, clmNo) = "No"
            End If
            
            ' No列を出力
            outputWs.Cells(row, clmNo) = cntBlock
            
        ElseIf flgRead Then ' [---] 以降の場合
            Dim key As String
            Dim item As String
            
            If Left(readLine, 1) = "-" Then ' -から始まる文字列の場合
            
                If cntBlock = 1 Then ' 1つ目のブロックの場合
                    ' ヘッダを出力
                    outputWs.Cells(rowHeader, clm).NumberFormatLocal = "@"
                    outputWs.Cells(rowHeader, clm) = lastKey & "_" & cntData
                End If
                
                ' アイテムを出力
                item = Replace(readLine, " ", "")
                item = Replace(item, "　", "")
                outputWs.Cells(row, clm).NumberFormatLocal = "@"
                outputWs.Cells(row, clm) = item
                
                If clm > clmMax Then ' 出力した列が最大列数より多い場合
                    clmMax = clm
                End If
                clm = clm + 1
                cntData = cntData + 1
                
            ElseIf InStr(readLine, ":") > 0 Then '[:]を含む文字列の場合
                item = GetItemStr(readLine)
                key = GetKeyStr(readLine)
                lastKey = key
                cntData = 1
                
                If cntBlock = 1 Then ' 1つ目のブロックの場合
                    ' ヘッダを出力
                    outputWs.Cells(rowHeader, clm).NumberFormatLocal = "@"
                    outputWs.Cells(rowHeader, clm) = key
                End If
                ' アイテムを出力
                outputWs.Cells(row, clm).NumberFormatLocal = "@"
                outputWs.Cells(row, clm) = item
                
                If clm > clmMax Then  ' 出力した列が最大列数より多い場合
                    clmMax = clm
                End If
                clm = clm + 1
            End If
        End If
    Loop
    
    ' 枠を追加
    outputWs.Range(outputWs.Cells(rowHeader, clmNo), outputWs.Cells(row, clmMax)).Borders.LineStyle = xlContinuous
    ' 列幅調整
    outputWs.Columns.AutoFit
    
    ' ファイルクローズ
    Close #1
    
End Function


'#### Keyを取得 ####'
Function GetKeyStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Left(inputStr, InStr(inputStr, ":") - 1)
    '空白文字列は削除する
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "　", "")
    GetKeyStr = retStr
End Function


'#### itemを取得 ####'
Function GetItemStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Mid(inputStr, InStr(inputStr, ":") + 1)
    '空白文字列は削除する
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "　", "")
    GetItemStr = retStr
End Function


'### 正常に出力できたか表示###'
Sub ShowCompletionMessage()
    Dim res As Integer
    res = MsgBox("完了しました", vbOKOnly)
End Sub


'#### 改行コード変換 ####'

Function LfToCrlfCopy(ByVal inputFile As String, ByVal tempFilePath As String)
    Dim FileNum As Integer
    Dim FileContent As String
    Dim NewContent As String

    ' 元ファイルを読み込む
    FileNum = FreeFile
    Open inputFile For Input As #FileNum
    FileContent = Input(LOF(FileNum), #FileNum)
    Close #FileNum

    ' 改行コードを変換（LF → CRLF）
    NewContent = Replace(FileContent, vbLf, vbCrLf)

    ' 一時ファイルに保存
    FileNum = FreeFile
    Open tempFilePath For Output As #FileNum
    Print #FileNum, NewContent
    Close #FileNum
End Function

