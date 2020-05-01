Attribute VB_Name = "GitCommit"

'##################
'### GitLab Ŀ��
'##################


Private Sub commitRepo()
Debug.Print "commitRepo Start : " & VBA.Now

    GitLab.Select

    Dim filesLoc As Variant
    Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
   
    filesLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , ExcelFile, , True) '//��ȣ ���� ���ϵ� ��� ����
     
     
    If VBA.IsEmpty(filesLoc) Then
        Exit Sub
    End If
        
    Dim gitlabWS As Worksheet: Set gitlabWS = GitLab
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim fnt, rear As Long
    Dim i As Long
    Dim inputRow As Long: inputRow = 2
    
    '// sFilePath: ���ϸ�(���Ϲ�ȣ, ���α׷� �̸�, Ȯ���� ����), sFileNum: ���Ϲ�ȣ, sFileName: ���α׷� �̸�, foldName: ���ϸ��� Ȯ���ڸ� ����
    Dim sFilePath, sFileName, sFileNum, foldName As String
    Dim splitResults As Variant
    Dim sDir As String  '// ��� ���� ���丮
    
    sFilePath = filesLoc(0)
    splitResults = VBA.Split(sFilePath, "\")
    If LBound(splitResults) >= 0 Then
        sFilePath = VBA.Replace(sFilePath, splitResults(UBound(splitResults)), vbNullString)
        mainWS.Cells(32, 13).Value = sFilePath  '// Main ȭ�鿡 ������ ��� ���
    End If
    
    
    
    For i = LBound(filesLoc, 1) To UBound(filesLoc, 1)
        sFilePath = filesLoc(i)
        splitResults = VBA.Split(sFilePath, "\")    '// \ ���� �������� ����
        If LBound(splitResults) >= 0 Then
            sFilePath = splitResults(UBound(splitResults))
            
            fnt = VBA.InStr(sFilePath, "(")
            rear = VBA.InStr(sFilePath, ")")
            
            If rear < fnt - 1 Then
                GoTo nextIndex
            End If
            sFileNum = VBA.Trim(VBA.Mid(sFilePath, fnt + 1, rear - fnt - 1))    '// ���� ��ȣ
            
            If VBA.InStr(sFilePath, ".xls") = 0 Then
                GoTo nextIndex
            End If
            foldName = VBA.Trim(VBA.Left(sFilePath, VBA.InStr(sFilePath, ".xls") - 1))    '// Ȯ���� ����(������ ���� �̸�)
                                    
            sFileName = VBA.Trim(VBA.Mid(foldName, VBA.InStr(foldName, ")") + 1)) '// foldName���� ��ȣ ����(���α׷� �̸�)
                    
            sDir = VBA.Replace(filesLoc(i), sFilePath, vbNullString, fnt)
            sDir = VBA.Trim(VBA.Left(sDir, VBA.Len(sDir) - 1))    '// ��� ���� ���丮
        End If
        
        Dim dDir As String: dDir = "C:\CookieGitlab\Solution"
                    
        If VBA.Dir(sDir, vbDirectory) = foldName Then
            Dim tempFile As Object
            Dim tempFolder As Object
            'C:\CookieGitlab\Solution\cookie_solution5331
            'Set tempFolder = fso.GetFolder(dDir & "\cookie_solution" & sFileNum)

            Dim moveDestiPath As String: moveDestiPath = dDir & "\cookie_solution" & sFileNum & "\"
            
            Dim targetFilesPath As Variant
            Dim tempArr As Object: Set tempArr = getCreateObject("ArrayList")
            Dim j As Long
            targetFilesPath = getFilesFromDir(sDir, tempArr, True, True)
            
            Dim splitResult As Variant
            Dim lastSplit2 As String
            
            If Not VBA.IsEmpty(targetFilesPath) Then
                For j = LBound(targetFilesPath) To UBound(targetFilesPath)
                    splitResult = VBA.Split(targetFilesPath(j), "\")
                    
                    If LBound(splitResult) >= 0 Then
                        lastSplit2 = splitResult(UBound(splitResult))
                        
                        '���Ͼ���; getDir
'                        '����
'                        If VBA.InStr(1, lastSplit2, ".") > 0 Then
'                            fso.MoveFile targetFilesPath(j), moveDestiPath   '// ����ҿ� �����(clone) ������ ���� �̵�
'                        '���丮
'                        Else
                            If VBA.Dir(moveDestiPath & lastSplit2, vbDirectory) = vbNullString Then
                                If Not VBA.StrComp(lastSplit2, foldName) = 0 Then
                                    fso.CreateFolder (moveDestiPath & lastSplit2)   '// class, form, normal,
                                End If
                            End If
'                        End If
                    End If
                Next j
            End If
            
            Set tempFolder = fso.GetFolder(sDir)
            Dim moveFileDestiPath As String
            Dim currentFolder As Object
            
            For Each tempFile In tempFolder.Files
                If VBA.Dir(moveDestiPath & "\" & tempFile.Name, vbNormal) <> vbNullString Then
                    fso.DeleteFile (moveDestiPath & "\" & tempFile.Name)
                End If
                fso.MoveFile tempFile.path, moveDestiPath   '// ����ҿ� �����(clone) ������ ���� �̵�
            Next
                
            For Each currentFolder In tempFolder.SubFolders
                For Each tempFile In currentFolder.Files
                    'fso.MoveFile sDir & "\", tempFile.path   '// ����ҿ� �����(clone) ������ ���� �̵�
                    splitResult = VBA.Split(tempFile.path, "\")
                    If LBound(splitResult) >= 0 Then
                        moveFileDestiPath = moveDestiPath & splitResult(UBound(splitResult) - 1) & "\"
                    Else
                        moveFileDestiPath = moveDestiPath
                    End If
                    If VBA.Dir(moveFileDestiPath & "\" & tempFile.Name, vbNormal) <> vbNullString Then
                        fso.DeleteFile (moveFileDestiPath & "\" & tempFile.Name)
                    End If
                    fso.MoveFile tempFile.path, moveFileDestiPath   '// ����ҿ� �����(clone) ������ ���� �̵�
                Next
            Next
            
        End If
        '================= ���� �̵� �Ϸ�
        
        
       
        '====>> ���� ��¥�� Commit ����
        
        Dim cmitScript As String
        Dim shellComd1 As String
        Dim shellComd2 As String
        cmitScript = Format(Date, "yyyy-mm-dd")
        shellComd1 = "git add ."
        shellComd2 = "git commit -m " & """" & cmitScript & """"
        
        ShellChangeCurrentDirectory ("C:\CookieGitlab\Solution\cookie_solution" & sFileNum)
        ShellRun.ShellRun (shellComd1)
        gitlabWS.Cells(inputRow, 5).Value = ShellRun.ShellRun(shellComd2)
        If gitlabWS.Cells(inputRow, 5).Value = "Failed" Then
            gitlabWS.Cells(inputRow, 5).Font.Bold = True
            gitlabWS.Cells(inputRow, 5).Font.Color = RGB(25, 100, 126)
        End If
        
        
        inputRow = inputRow + 1
nextIndex:
    Next i
    
    
    Set fso = Nothing
Debug.Print "commitRepo End : " & VBA.Now

End Sub

