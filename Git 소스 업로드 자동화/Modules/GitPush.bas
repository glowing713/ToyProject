Attribute VB_Name = "GitPush"

'##################
'### GitLab 커밋
'##################


Private Sub pushRepo()
Debug.Print "pushRepo Start : " & VBA.Now

    GitLab.Select

    Dim filesLoc As Variant
    Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
   
    filesLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , ExcelFile, , True) '//암호 없는 파일들 경로 저장
     
     
    If VBA.IsEmpty(filesLoc) Then
        Exit Sub
    End If
        
    Dim gitlabWS As Worksheet: Set gitlabWS = GitLab
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim fnt, rear As Long
    Dim i As Long
    Dim inputRow As Long: inputRow = 2
    
    '// sFilePath: 파일명(파일번호, 프로그램 이름, 확장자 포함), sFileNum: 파일번호, sFileName: 프로그램 이름, foldName: 파일명에서 확장자만 제거
    Dim sFilePath, sFileName, sFileNum, foldName As String
    Dim splitResults As Variant
    Dim sDir As String  '// 결과 폴더 디렉토리
    
    sFilePath = filesLoc(0)
    splitResults = VBA.Split(sFilePath, "\")
    If LBound(splitResults) >= 0 Then
        sFilePath = VBA.Replace(sFilePath, splitResults(UBound(splitResults)), vbNullString)
        mainWS.Cells(39, 13).Value = sFilePath  '// Main 화면에 오픈한 경로 기록
    End If
    
    
    
    For i = LBound(filesLoc, 1) To UBound(filesLoc, 1)
        sFilePath = filesLoc(i)
        splitResults = VBA.Split(sFilePath, "\")    '// \ 문자 기준으로 분할
        If LBound(splitResults) >= 0 Then
            sFilePath = splitResults(UBound(splitResults))
            
            fnt = VBA.InStr(sFilePath, "(")
            rear = VBA.InStr(sFilePath, ")")
            
            If rear < fnt - 1 Then
                GoTo nextIndex
            End If
            sFileNum = VBA.Trim(VBA.Mid(sFilePath, fnt + 1, rear - fnt - 1))    '// 파일 번호
            
            If VBA.InStr(sFilePath, ".xls") = 0 Then
                GoTo nextIndex
            End If
            foldName = VBA.Trim(VBA.Left(sFilePath, VBA.InStr(sFilePath, ".xls") - 1))    '// 확장자 제거(생성할 폴더 이름)
                                    
            sFileName = VBA.Trim(VBA.Mid(foldName, VBA.InStr(foldName, ")") + 1)) '// foldName에서 번호 제거(프로그램 이름)
        End If
        
        
       
        '====>> Push 수행
        
        Dim cmitScript As String
        Dim shellComd4 As String
        shellComd4 = "git push -u origin master"
        
        ShellChangeCurrentDirectory ("C:\CookieGitlab\Solution\cookie_solution" & sFileNum)
        gitlabWS.Cells(inputRow, 6).Value = ShellRun.ShellRun(shellComd4)
        If gitlabWS.Cells(inputRow, 6).Value = "Failed" Then
            gitlabWS.Cells(inputRow, 6).Font.Bold = True
            gitlabWS.Cells(inputRow, 6).Font.Color = RGB(25, 100, 126)
        End If
    
        
        inputRow = inputRow + 1
nextIndex:
    Next i
    
    
    Set fso = Nothing
Debug.Print "pushRepo End : " & VBA.Now

End Sub
