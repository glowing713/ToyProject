Attribute VB_Name = "GitRepo"


'#######################
'### GitLab 저장소 생성
'#######################

Private Sub createRepo()
    
    GitLab.Select
      
    Dim filesLoc As Variant
    Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
     
    filesLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True)
     
    If VBA.IsEmpty(filesLoc) Then
        Exit Sub
    End If
        
    Dim fnt, rear As Long
    Dim i As Long
    Dim inputRow As Long: inputRow = 2
    Dim sFilePath, sFileName, sFileNum, foldName As String
    Dim splitResults As Variant
    
    Dim gitlabWS As Worksheet: Set gitlabWS = GitLab
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim URL As String
    Dim JSONString As String
    Dim objHTTP As New WinHttpRequest
    sFilePath = filesLoc(0)
    splitResults = VBA.Split(sFilePath, "\")
    If LBound(splitResults) >= 0 Then
        sFilePath = VBA.Replace(sFilePath, splitResults(UBound(splitResults)), vbNullString)
        mainWS.Cells(18, 13).Value = sFilePath  '// Main 화면에 오픈한 경로 기록
    End If
    
    Call Main.clearSheet(gitlabWS)  '// 이전에 기록된 결과화면 지우기

    
    
    For i = LBound(filesLoc, 1) To UBound(filesLoc, 1)
        sFilePath = filesLoc(i)
            
        splitResults = VBA.Split(sFilePath, "\")    '// \ 문자 기준으로 분할
        If LBound(splitResults) >= 0 Then
            sFilePath = splitResults(UBound(splitResults))
            
            fnt = VBA.InStr(sFilePath, "(")
            rear = VBA.InStr(sFilePath, ")")
            sFileNum = VBA.Trim(VBA.Mid(sFilePath, fnt + 1, rear - fnt - 1))
            
            foldName = VBA.Trim(VBA.Left(sFilePath, VBA.InStr(sFilePath, ".xls") - 1))    '// 확장자 제거
                                    
            sFileName = VBA.Trim(VBA.Mid(foldName, VBA.InStr(foldName, ")") + 1)) '// foldName에서 번호 제거(프로그램 이름)
            
        End If
        
        
        '// https://docs.gitlab.com/ee/user/profile/personal_access_tokens.html#personal-access-tokens 참고
        '// 저장소 생성 권한을 가진 개인 토큰을 발급받아야 한다.
        URL = "http://coukey.co.kr:5001/api/v4/projects?private_token=???"
        
        objHTTP.Open "POST", URL, False
        objHTTP.SetRequestHeader "Content-Type", "application/json; charset=UTF-8"
        
        
        Dim tC As String: tC = """"
        JSONString = "{" & tC & "name" & tC & ": " & tC & "Cookie_Solution" & sFileNum & tC & ", " & tC & "namespace_id" & tC & ": " & "13, " & tC & "description" & tC & ":" & tC & sFileName & tC & "}"
            
        objHTTP.Send JSONString '// http request 전송
        gitlabWS.Cells(inputRow, 1).Value = sFileNum
        gitlabWS.Cells(inputRow, 2).Value = sFileName
        gitlabWS.Cells(inputRow, 3).Value = objHTTP.StatusText
        
                
        inputRow = inputRow + 1
        
    Next i
    

End Sub
