Attribute VB_Name = "DumpLog"
Public Target_cell_for_import_file As String
Public Target_cell_for_output_file As String
Sub DumpLog()
Target_cell_for_import_file = "B7"
Target_cell_for_output_file = "B10"
    ' 読み込みファイルのパスを取得
    Dim inputFile As String
    inputFile = Range(Target_cell_for_import_file)
    ' 出力先フォルダのパスを取得
    Dim outputFolderPath As String
    outputFolderPath = Range(Target_cell_for_output_file)
    ' 空白を削除する
    inputFile = Replace(inputFile, " ", "")
    inputFile = Replace(inputFile, "　", "")
    
     ' 入力がされていない、指定されたファイルが存在しない場合の処理
    If inputFile = "" Or Dir(inputFile) = "" Then
        MsgBox "読み込みファイルが存在しません。", vbExclamation
        ' セルにフォーカスを移動する
        Range(Target_cell_for_import_file).Select
        Exit Sub
    End If
    
    If outputFolderPath = "" Then ' 出力先フォルダが空白の場合
        ' 読み込みファイルと同じフォルダを出力先フォルダとする
        outputFolderPath = Left(inputFile, (InStrRev(inputFile, "\") - 1))
        Range(Target_cell_for_output_file) = outputFolderPath
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
    Dim outputws As Worksheet ' 出力するシート
    Dim tempFilePath As String '　新しい出力フォルダー
    
    If Dir(inputFile) <> "" Then ' B2で指定された読み込みファイルが存在する場合
        If Dir(outputFolderPath, vbDirectory) <> "" Then ' B5で指定された出力先フォルダが存在する場合
            ' ブックを新規作成
            Set outputWb = Workbooks.Add
            ' シート名を変更
            Set outputws = outputWb.Sheets(1)
            outputws.Name = outputSheetName
            
             ' 新しい出力フォルダー
            tempFilePath = outputFolderPath & "\Temp_CRLF.txt"
            ' 元ファイルをコピーして改行コード変換
            Call LfToCrlfCopy(inputFile, tempFilePath)
    
            ' ダンプ処理（元ファイルではなく一時ファイルを使用）
            Call OutputDumpData(tempFilePath, outputws)

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
Function OutputDumpData(ByVal tempFilePath As String, ByVal outputws As Worksheet)
    ' 出力先のシートをクリアする
    outputws.Cells.Clear
    
    ' 変数宣言
    Dim serchStr As String              ' 出力するデータかを判断する目印
    Dim flgRead As Boolean              ' 読み込みデータフラグ
    Dim readLine As String              ' 読み込んだ列の文字列
    Dim headerDict As Object            ' 出力したヘッダ名を含む辞書
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim lastKey As String               ' 最後のkey（ヘッダの項目）
    Dim clmMax As Integer               ' 最大列数
    Dim row As Integer                  ' 行
    Dim rowHeader As Integer            ' ヘッダを出力する行
    Dim clm As Integer                  ' 列
    Dim clmNo As Integer                ' Noを表示する列
    Dim cntBlock As Integer             ' 出力ブロックのカウント
    Dim cntData As Integer              ' データのカウント
    Dim secVal As Double                ' 秒数（sec)
    Dim nanoVal As Double               ' ナノ秒(nanosec)
    Dim flgsec As Boolean               ' secが読み込まれたかのフラグ
    Dim flgnanosec As Boolean           ' nanosecが読み込まれたかのフラグ
    Dim key As String                   ' ヘッダ
    Dim item As String                  ' データ
    
    serchStr = "---"
    flgRead = False
    cntBlock = 0
    row = 3         ' 何行目から出力するかを設定
    rowHeader = row
    clmNo = 2       ' 何列目から出力するかを設定
    cntData = 1
    clmMax = clmNo
    secVal = 0
    nanoVal = 0
    flgsec = False
    flgnanosec = False
    
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
                outputws.Cells(rowHeader, clmNo).NumberFormatLocal = "@"
                outputws.Cells(rowHeader, clmNo) = "No"
            End If
            
            ' No列を出力
            outputws.Cells(row, clmNo).NumberFormatLocal = "@"
            outputws.Cells(row, clmNo) = cntBlock
            
        ElseIf flgRead Then ' [---] 以降の場合
            
            If Left(LTrim(readLine), 1) = "-" Then   ' 最初の非空白文字が "-" の場合
                key = lastKey & "_" & cntData
                ' 同じキーが複数回出現する場合に備えて、キー名に連番を付けて一意に識別
                ' 例: sec_1, sec_2, sec_3 など
                clm = RegisterHeader(headerDict, key, clmNo, rowHeader, outputws)
                ' アイテムを出力
                item = GetItemStr(readLine)
                item = Trim(Mid(item, 2)) ' 先頭の "-" を除去
                Call WriteItem(outputws, row, clm, item)
                
                If clm > clmMax Then ' 出力した列が最大列数より多い場合
                    clmMax = clm
                End If
                    clm = clm + 1
                    cntData = cntData + 1
                
            ElseIf InStr(readLine, ":") > 0 Then ' [:]を含む文字列の場合
                Debug.Print readLine
                '先頭の空白数を数える
                NumberofLeadingspace = CountLeadingSpaces(readLine)
                Debug.Print "空白数:" & NumberofLeadingspace
                
                item = GetItemStr(readLine)
                key = GetKeyStr(readLine)
                lastKey = key
                cntData = 1
                
                
                '新しいヘッダ登録
                clm = RegisterHeader(headerDict, key, clmNo, rowHeader, outputws)
              
                ' アイテムを出力
                Call WriteItem(outputws, row, clm, item)
                
                ' sec/nanosec を記録
                If key = "sec" Then ' ヘッダがsecだったら
                    If IsNumeric(item) Then
                        secVal = CDbl(item) ' 文字列から数値に変換
                        flgsec = True
                        Cells(row, clm).NumberFormatLocal = "0_ "  ' セルを数値にする
                    End If
                ElseIf key = "nanosec" Then ' ヘッダがnanosecだったら
                    If IsNumeric(item) Then
                        nanoVal = CDbl(item) ' 文字列から数値に変換
                        flgnanosec = True
                        Cells(row, clm).NumberFormatLocal = "0_ " ' セルを数値にする
                    End If
                End If
                
                ' 両方揃ったら timestamp を出力
                If flgsec And flgnanosec Then
                    ' 日時形式に変換して出力
                    clm = WriteTimestamp(headerDict, clmNo, rowHeader, row, secVal, nanoVal, outputws)
                    
                If clm > clmMax Then  ' 出力した列が最大列数より多い場合
                    clmMax = clm
                End If
                clm = clm + 1
                
                ' フラグをリセット（次の行に備える）
                flgsec = False
                flgnanosec = False
                
                End If
            End If
        End If
    Loop
    
    ' 枠を追加
    outputws.Range(outputws.Cells(rowHeader, clmNo), outputws.Cells(row, clmMax)).Borders.LineStyle = xlContinuous
    ' 列幅調整
    outputws.Columns.AutoFit

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

' #### 新ヘッダー登録 #### '
Function RegisterHeader(headerDict As Object, key As String, clmNo As Integer, rowHeader As Integer, outputws As Worksheet) As Integer
    Dim clm As Integer
    If Not headerDict.exists(key) Then          ' 新しいデータ項目の場合
        clm = headerDict.Count + clmNo + 1      ' 新しいデータ項目に対するヘッダを生成
        headerDict.Add key, clm
        ' ヘッダ出力
        outputws.Cells(rowHeader, clm).NumberFormatLocal = "@"
        outputws.Cells(rowHeader, clm) = key
    Else
        clm = headerDict(key)
    End If
    RegisterHeader = clm
End Function

' #### timestampを生成・出力 #### '
Function WriteTimestamp(headerDict As Object, clmNo As Integer, rowHeader As Integer, row As Integer, secVal As Double, nanoVal As Double, outputws As Worksheet) As Integer
    Dim timestamp As Double             ' UNIXタイムスタンプ（秒＋ナノ秒）
    Dim timestampStr As String          ' 日時形式に変換した文字列
    Dim clm As Integer
    
    ' 1970年1月1日からの経過秒数を日時形式に変換して出力
    timestamp = secVal + nanoVal / 1000000000#
    timestampStr = Format(DateAdd("s", timestamp, #1/1/1970#), "yyyy-mm-dd HH:MM:SS")
    If Not headerDict.exists("timestamp") Then ' 新しいキー項目の場合
        clm = headerDict.Count + clmNo + 1
        headerDict.Add "timestamp", clm
        ' ヘッダ出力
        outputws.Cells(rowHeader, clm).NumberFormatLocal = "@"
        outputws.Cells(rowHeader, clm) = "timestamp"
        Else
        clm = headerDict("timestamp") ' clmを返す
    End If
    ' アイテムを出力
    outputws.Cells(row, headerDict("timestamp")).NumberFormatLocal = "yyyy-mm-dd HH:MM:SS" ' セルの表示形式を「年月日時分秒」に設定（例: 2025-08-04 09:00:00）
    outputws.Cells(row, headerDict("timestamp")) = timestampStr ' 計算されたUNIXタイムスタンプ（1970年1月1日からの経過秒数）を日時文字列として出力
End Function

' #### セルにデータを出力 #### '
Function WriteItem(outputws As Worksheet, row As Integer, clm As Integer, item As String) As Boolean
    On Error GoTo ErrHandler
    outputws.Cells(row, clm).NumberFormatLocal = "@"
    outputws.Cells(row, clm) = item
    WriteItem = True
    Exit Function
ErrHandler:
    WriteItem = False
End Function

' #### 先頭の空白数を数える #### '
Function CountLeadingSpaces(readLine As String) As Long
    Dim Spacecount As Long  ' 空白の数
    Dim FirstCharacter As String ' 1文字目
    
    Spacecount = 1 ' 初期値１
    
    Do While Spacecount <= Len(readLine)
        FirstCharacter = Mid(readLine, Spacecount, 1)
        If FirstCharacter = " " Then  ' 半角スペースの場合
            Spacecount = Spacecount + 1
        Else
            Exit Do ' 文字が出現したら
        End If
    Loop
    CountLeadingSpaces = Spacecount - 1
End Function
