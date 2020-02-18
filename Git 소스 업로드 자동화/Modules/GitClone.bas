Attribute VB_Name = "GitClone"

'##################
'### GitLab 클론
'##################


Private Sub cloneRepo()
Debug.Print "cloneRepo Start : " & VBA.Now

    GitLab.Select

    Dim gitlabWS As Worksheet: Set gitlabWS = GitLab
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim destLoc As Variant
    Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
    Dim M, fnt, rear As Long
    Dim indexRow As Long: indexRow = 2
    Dim dFilePath, fileNum As String
    Dim dsplitResult As Variant
 
    destLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True)
    
    dFilePath = destLoc(0)
    dsplitResults = VBA.Split(dFilePath, "\")
    If LBound(dsplitResults) >= 0 Then
        dFilePath = VBA.Replace(dFilePath, dsplitResults(UBound(dsplitResults)), vbNullString)
        mainWS.Cells(25, 13).Value = dFilePath  '// Main 화면에 오픈한 경로 기록
    End If
    

    For M = LBound(destLoc, 1) To UBound(destLoc, 1)
        dFilePath = destLoc(M)
    
        dsplitResults = VBA.Split(dFilePath, "\")    '// \ 문자 기준으로 분할
        If LBound(dsplitResults) >= 0 Then
            dFilePath = dsplitResults(UBound(dsplitResults))
            
            fnt = VBA.InStr(dFilePath, "(")
            rear = VBA.InStr(dFilePath, ")")
            fileNum = VBA.Trim(VBA.Mid(dFilePath, fnt + 1, rear - fnt - 1))   '// 파일 번호 추출
        End If
        
        '====>> C:\CookieGitlab\Solution\ 으로 클론하기
        Dim foldAddrs, gitAddrs, shellComd As String
        foldAddrs = "C:\CookieGitlab\Solution\cookie_solution" & fileNum
        gitAddrs = "http://www.coukey.co.kr:5001/solution/cookie_solution" & fileNum
        'shellComd = "cmd /c git clone " & gitAddrs & " " & foldAddrs
        shellComd = "git clone " & gitAddrs & " " & foldAddrs
        gitlabWS.Cells(indexRow, 4).Value = ShellRun.ShellRun(shellComd)
        If gitlabWS.Cells(indexRow, 4).Value = "Failed" Then
            gitlabWS.Cells(indexRow, 4).Font.Bold = True
            gitlabWS.Cells(indexRow, 4).Font.Color = RGB(25, 100, 126)
        End If
    
        indexRow = indexRow + 1
    
    Next M
    
    Set fso = Nothing
Debug.Print "cloneRepo End : " & VBA.Now

End Sub
