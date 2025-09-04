Attribute VB_Name = "DumpLog"
Public Target_cell_for_import_file As String
Public Target_cell_for_output_file As String
Sub DumpLog()
Target_cell_for_import_file = "B7"
Target_cell_for_output_file = "B10"
    ' �ǂݍ��݃t�@�C���̃p�X���擾
    Dim inputFile As String
    inputFile = Range(Target_cell_for_import_file)
    ' �o�͐�t�H���_�̃p�X���擾
    Dim outputFolderPath As String
    outputFolderPath = Range(Target_cell_for_output_file)
    ' �󔒂��폜����
    inputFile = Replace(inputFile, " ", "")
    inputFile = Replace(inputFile, "�@", "")
    
     ' ���͂�����Ă��Ȃ��A�w�肳�ꂽ�t�@�C�������݂��Ȃ��ꍇ�̏���
    If inputFile = "" Or Dir(inputFile) = "" Then
        MsgBox "�ǂݍ��݃t�@�C�������݂��܂���B", vbExclamation
        ' �Z���Ƀt�H�[�J�X���ړ�����
        Range(Target_cell_for_import_file).Select
        Exit Sub
    End If
    
    If outputFolderPath = "" Then ' �o�͐�t�H���_���󔒂̏ꍇ
        ' �ǂݍ��݃t�@�C���Ɠ����t�H���_���o�͐�t�H���_�Ƃ���
        outputFolderPath = Left(inputFile, (InStrRev(inputFile, "\") - 1))
        Range(Target_cell_for_output_file) = outputFolderPath
    End If
    
    Dim inputFileLen As Integer ' �ǂݍ��݃t�@�C���̃p�X�̒���
    inputFileLen = Len(inputFile)
    Dim outputFilePath As String ' �o�̓t�@�C���̃p�X
    Dim outputFileName As String ' �o�̓t�@�C����
    outputFileName = makeFileName()
    outputFilePath = outputFolderPath & "\" & outputFileName
    Dim outputSheetName As String ' �o�͂���V�[�g��
    outputSheetName = "�_���v"
    Dim fOutputObj As Object
    Dim outputWb As Workbook  ' �o�͂��郏�[�N�u�b�N
    Dim outputws As Worksheet ' �o�͂���V�[�g
    Dim tempFilePath As String '�@�V�����o�̓t�H���_�[
    
    If Dir(inputFile) <> "" Then ' B2�Ŏw�肳�ꂽ�ǂݍ��݃t�@�C�������݂���ꍇ
        If Dir(outputFolderPath, vbDirectory) <> "" Then ' B5�Ŏw�肳�ꂽ�o�͐�t�H���_�����݂���ꍇ
            ' �u�b�N��V�K�쐬
            Set outputWb = Workbooks.Add
            ' �V�[�g����ύX
            Set outputws = outputWb.Sheets(1)
            outputws.Name = outputSheetName
            
             ' �V�����o�̓t�H���_�[
            tempFilePath = outputFolderPath & "\Temp_CRLF.txt"
            ' ���t�@�C�����R�s�[���ĉ��s�R�[�h�ϊ�
            Call LfToCrlfCopy(inputFile, tempFilePath)
    
            ' �_���v�����i���t�@�C���ł͂Ȃ��ꎞ�t�@�C�����g�p�j
            Call OutputDumpData(tempFilePath, outputws)

            ' �u�b�N��ۑ�
            outputWb.SaveAs outputFilePath

            ' �u�b�N�����
            outputWb.Close
            
            ' �ꎞ�t�@�C�����폜
            If Dir(tempFilePath) <> "" Then
                 Kill tempFilePath
            End If

            
             ' ����ɏo�͂��ꂽ���Ƃ��������b�Z�[�W��\��!!!!!!!'
            Call ShowCompletionMessage
            
        Else
            ' �G���[���b�Z�[�W���o��
            MsgBox "�o�͐�t�H���_��������܂���", vbExclamation
        End If
        
    Else  ' B2�Ŏw�肳�ꂽ�ǂݍ��݃t�@�C�������݂��Ȃ��ꍇ
        ' �G���[���b�Z�[�W���o��
        MsgBox "�ǂݍ��݃t�@�C����������܂���", vbExclamation
        
    End If
    
End Sub


Function makeFileName() As String
    ' ���t_�������擾
    Dim dateTime
    dateTime = Now()
    
    ' ������ɕϊ�
    Dim retStr As String
    retStr = Format(dateTime, "yyyymmdd_hmmss")
    
    makeFileName = retStr & "_log_dump.xlsx"
End Function

'#### �o�� ####'
Function OutputDumpData(ByVal tempFilePath As String, ByVal outputws As Worksheet)
    ' �o�͐�̃V�[�g���N���A����
    outputws.Cells.Clear
    
    ' �ϐ��錾
    Dim serchStr As String              ' �o�͂���f�[�^���𔻒f����ڈ�
    Dim flgRead As Boolean              ' �ǂݍ��݃f�[�^�t���O
    Dim readLine As String              ' �ǂݍ��񂾗�̕�����
    Dim headerDict As Object            ' �o�͂����w�b�_�����܂ގ���
    Set headerDict = CreateObject("Scripting.Dictionary")
    Dim lastKey As String               ' �Ō��key�i�w�b�_�̍��ځj
    Dim clmMax As Integer               ' �ő��
    Dim row As Integer                  ' �s
    Dim rowHeader As Integer            ' �w�b�_���o�͂���s
    Dim clm As Integer                  ' ��
    Dim clmNo As Integer                ' No��\�������
    Dim cntBlock As Integer             ' �o�̓u���b�N�̃J�E���g
    Dim cntData As Integer              ' �f�[�^�̃J�E���g
    Dim secVal As Double                ' �b���isec)
    Dim nanoVal As Double               ' �i�m�b(nanosec)
    Dim flgsec As Boolean               ' sec���ǂݍ��܂ꂽ���̃t���O
    Dim flgnanosec As Boolean           ' nanosec���ǂݍ��܂ꂽ���̃t���O
    Dim key As String                   ' �w�b�_
    Dim item As String                  ' �f�[�^
    
    serchStr = "---"
    flgRead = False
    cntBlock = 0
    row = 3         ' ���s�ڂ���o�͂��邩��ݒ�
    rowHeader = row
    clmNo = 2       ' ����ڂ���o�͂��邩��ݒ�
    cntData = 1
    clmMax = clmNo
    secVal = 0
    nanoVal = 0
    flgsec = False
    flgnanosec = False
    
    ' ���̓t�@�C���I�[�v��
    Open tempFilePath For Input As #1
    
    Do Until EOF(1)
        ' 1�񂸂ǂݍ���
        Line Input #1, readLine
        
        If readLine = serchStr Then ' --- �̏ꍇ
            ' �ǂݍ��݃t���O��ON�ɂ���
            flgRead = True
            
            cntBlock = cntBlock + 1
            cntData = 1
            row = row + 1
            clm = clmNo + 1
            
            If cntBlock = 1 Then ' 1�ڂ̃u���b�N�̏ꍇ
                ' [No](�w�b�_)���o��
                outputws.Cells(rowHeader, clmNo).NumberFormatLocal = "@"
                outputws.Cells(rowHeader, clmNo) = "No"
            End If
            
            ' No����o��
            outputws.Cells(row, clmNo).NumberFormatLocal = "@"
            outputws.Cells(row, clmNo) = cntBlock
            
        ElseIf flgRead Then ' [---] �ȍ~�̏ꍇ
            
            If Left(LTrim(readLine), 1) = "-" Then   ' �ŏ��̔�󔒕����� "-" �̏ꍇ
                key = lastKey & "_" & cntData
                ' �����L�[��������o������ꍇ�ɔ����āA�L�[���ɘA�Ԃ�t���Ĉ�ӂɎ���
                ' ��: sec_1, sec_2, sec_3 �Ȃ�
                clm = RegisterHeader(headerDict, key, clmNo, rowHeader, outputws)
                ' �A�C�e�����o��
                item = GetItemStr(readLine)
                item = Trim(Mid(item, 2)) ' �擪�� "-" ������
                Call WriteItem(outputws, row, clm, item)
                
                If clm > clmMax Then ' �o�͂����񂪍ő�񐔂�葽���ꍇ
                    clmMax = clm
                End If
                    clm = clm + 1
                    cntData = cntData + 1
                
            ElseIf InStr(readLine, ":") > 0 Then ' [:]���܂ޕ�����̏ꍇ
                Debug.Print readLine
                '�擪�̋󔒐��𐔂���
                NumberofLeadingspace = CountLeadingSpaces(readLine)
                Debug.Print "�󔒐�:" & NumberofLeadingspace
                
                item = GetItemStr(readLine)
                key = GetKeyStr(readLine)
                lastKey = key
                cntData = 1
                
                
                '�V�����w�b�_�o�^
                clm = RegisterHeader(headerDict, key, clmNo, rowHeader, outputws)
              
                ' �A�C�e�����o��
                Call WriteItem(outputws, row, clm, item)
                
                ' sec/nanosec ���L�^
                If key = "sec" Then ' �w�b�_��sec��������
                    If IsNumeric(item) Then
                        secVal = CDbl(item) ' �����񂩂琔�l�ɕϊ�
                        flgsec = True
                        Cells(row, clm).NumberFormatLocal = "0_ "  ' �Z���𐔒l�ɂ���
                    End If
                ElseIf key = "nanosec" Then ' �w�b�_��nanosec��������
                    If IsNumeric(item) Then
                        nanoVal = CDbl(item) ' �����񂩂琔�l�ɕϊ�
                        flgnanosec = True
                        Cells(row, clm).NumberFormatLocal = "0_ " ' �Z���𐔒l�ɂ���
                    End If
                End If
                
                ' ������������ timestamp ���o��
                If flgsec And flgnanosec Then
                    ' �����`���ɕϊ����ďo��
                    clm = WriteTimestamp(headerDict, clmNo, rowHeader, row, secVal, nanoVal, outputws)
                    
                If clm > clmMax Then  ' �o�͂����񂪍ő�񐔂�葽���ꍇ
                    clmMax = clm
                End If
                clm = clm + 1
                
                ' �t���O�����Z�b�g�i���̍s�ɔ�����j
                flgsec = False
                flgnanosec = False
                
                End If
            End If
        End If
    Loop
    
    ' �g��ǉ�
    outputws.Range(outputws.Cells(rowHeader, clmNo), outputws.Cells(row, clmMax)).Borders.LineStyle = xlContinuous
    ' �񕝒���
    outputws.Columns.AutoFit

    ' �t�@�C���N���[�Y
    Close #1

End Function


'#### Key���擾 ####'
Function GetKeyStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Left(inputStr, InStr(inputStr, ":") - 1)
    '�󔒕�����͍폜����
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "�@", "")
    GetKeyStr = retStr
End Function

'#### item���擾 ####'
Function GetItemStr(ByVal inputStr As String) As String
    Dim retStr As String
    retStr = Mid(inputStr, InStr(inputStr, ":") + 1)
    '�󔒕�����͍폜����
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "�@", "")
    GetItemStr = retStr
End Function

'### ����ɏo�͂ł������\��###'
Sub ShowCompletionMessage()
    Dim res As Integer
    res = MsgBox("�������܂���", vbOKOnly)
End Sub

'#### ���s�R�[�h�ϊ� ####'
Function LfToCrlfCopy(ByVal inputFile As String, ByVal tempFilePath As String)
    Dim FileNum As Integer
    Dim FileContent As String
    Dim NewContent As String

    ' ���t�@�C����ǂݍ���
    FileNum = FreeFile
    Open inputFile For Input As #FileNum
    FileContent = Input(LOF(FileNum), #FileNum)
    Close #FileNum

    ' ���s�R�[�h��ϊ��iLF �� CRLF�j
    NewContent = Replace(FileContent, vbLf, vbCrLf)

    ' �ꎞ�t�@�C���ɕۑ�
    FileNum = FreeFile
    Open tempFilePath For Output As #FileNum
    Print #FileNum, NewContent
    Close #FileNum
End Function

' #### �V�w�b�_�[�o�^ #### '
Function RegisterHeader(headerDict As Object, key As String, clmNo As Integer, rowHeader As Integer, outputws As Worksheet) As Integer
    Dim clm As Integer
    If Not headerDict.exists(key) Then          ' �V�����f�[�^���ڂ̏ꍇ
        clm = headerDict.Count + clmNo + 1      ' �V�����f�[�^���ڂɑ΂���w�b�_�𐶐�
        headerDict.Add key, clm
        ' �w�b�_�o��
        outputws.Cells(rowHeader, clm).NumberFormatLocal = "@"
        outputws.Cells(rowHeader, clm) = key
    Else
        clm = headerDict(key)
    End If
    RegisterHeader = clm
End Function

' #### timestamp�𐶐��E�o�� #### '
Function WriteTimestamp(headerDict As Object, clmNo As Integer, rowHeader As Integer, row As Integer, secVal As Double, nanoVal As Double, outputws As Worksheet) As Integer
    Dim timestamp As Double             ' UNIX�^�C���X�^���v�i�b�{�i�m�b�j
    Dim timestampStr As String          ' �����`���ɕϊ�����������
    Dim clm As Integer
    
    ' 1970�N1��1������̌o�ߕb��������`���ɕϊ����ďo��
    timestamp = secVal + nanoVal / 1000000000#
    timestampStr = Format(DateAdd("s", timestamp, #1/1/1970#), "yyyy-mm-dd HH:MM:SS")
    If Not headerDict.exists("timestamp") Then ' �V�����L�[���ڂ̏ꍇ
        clm = headerDict.Count + clmNo + 1
        headerDict.Add "timestamp", clm
        ' �w�b�_�o��
        outputws.Cells(rowHeader, clm).NumberFormatLocal = "@"
        outputws.Cells(rowHeader, clm) = "timestamp"
        Else
        clm = headerDict("timestamp") ' clm��Ԃ�
    End If
    ' �A�C�e�����o��
    outputws.Cells(row, headerDict("timestamp")).NumberFormatLocal = "yyyy-mm-dd HH:MM:SS" ' �Z���̕\���`�����u�N���������b�v�ɐݒ�i��: 2025-08-04 09:00:00�j
    outputws.Cells(row, headerDict("timestamp")) = timestampStr ' �v�Z���ꂽUNIX�^�C���X�^���v�i1970�N1��1������̌o�ߕb���j�����������Ƃ��ďo��
End Function

' #### �Z���Ƀf�[�^���o�� #### '
Function WriteItem(outputws As Worksheet, row As Integer, clm As Integer, item As String) As Boolean
    On Error GoTo ErrHandler
    outputws.Cells(row, clm).NumberFormatLocal = "@"
    outputws.Cells(row, clm) = item
    WriteItem = True
    Exit Function
ErrHandler:
    WriteItem = False
End Function

' #### �擪�̋󔒐��𐔂��� #### '
Function CountLeadingSpaces(readLine As String) As Long
    Dim Spacecount As Long  ' �󔒂̐�
    Dim FirstCharacter As String ' 1������
    
    Spacecount = 1 ' �����l�P
    
    Do While Spacecount <= Len(readLine)
        FirstCharacter = Mid(readLine, Spacecount, 1)
        If FirstCharacter = " " Then  ' ���p�X�y�[�X�̏ꍇ
            Spacecount = Spacecount + 1
        Else
            Exit Do ' �������o��������
        End If
    Loop
    CountLeadingSpaces = Spacecount - 1
End Function
