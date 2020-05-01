Attribute VB_Name = "GitUpload"


'########################  추가한 레퍼런스  ##########################
'###  1. Microsoft WinHTTP Services, version 5.1
'###  2. Windows Script Host Object Model
'#####################################################################



'============================  진행 순서  ============================
     
' 1. GitUpload의 getFileList() 실행 => 소스 분할
' 2. GitRepo의 createRepo() 실행 => 저장소 생성
' 3. GitClone의 cloneRepo() 실행 => C:\CookieGitlab\Solution 내에 cookie_solution[파일번호]로 클론
' 4. GitCommit의 commitRepo() 실행 => 분할 파일 클론된 폴더로 이동, 커밋, 푸시까지 완료
'=====================================================================

Dim targetWS As Worksheet



'#########################
'### 파일 리스트 불러오기
'#########################


Private Sub getFileList()


    소스분할.Select
     
     Dim filesLoc, cryptLoc As Variant
     Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
   
    '### 암호 없는 파일 폴더 생성 후 이동,
    '### 폴더 내에서 소스 분리까지.
     
     
    filesLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True) '//암호 없는 파일들 경로 저장
     
     
    If VBA.IsEmpty(filesLoc) Then
        Exit Sub
    End If
        
    Set targetWS = 소스분할
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim fnt, rear As Long
    Dim i As Long
    Dim inputRow As Long: inputRow = 2
        
    Call Main.clearSheet(targetWS)  '// 이전에 기록된 결과화면 지우기
    
    
    '// sFilePath: 파일명(파일번호, 프로그램 이름, 확장자 포함), sFileNum: 파일번호, sFileName: 프로그램 이름, foldName: 파일명에서 확장자만 제거
    Dim sFilePath, sFileName, sFileNum, foldName As String
    Dim splitResults As Variant
    Dim sDir As String  '// 결과 폴더 디렉토리
    sFilePath = filesLoc(0)
    splitResults = VBA.Split(sFilePath, "\")
    If LBound(splitResults) >= 0 Then
        sFilePath = VBA.Replace(sFilePath, splitResults(UBound(splitResults)), vbNullString)
        mainWS.Cells(11, 14).Value = sFilePath  '// Main 화면에 오픈한 경로 기록
    End If
    
    For i = LBound(filesLoc, 1) To UBound(filesLoc, 1)
        sFilePath = filesLoc(i)
        
        targetWS.Cells(inputRow, 1).Value = sFilePath   '// 파일 경로 출력
    
        splitResults = VBA.Split(sFilePath, "\")    '// \ 문자 기준으로 분할
        If LBound(splitResults) >= 0 Then
            sFilePath = splitResults(UBound(splitResults))
            
            targetWS.Cells(inputRow, 2).Value = sFilePath   '// 파일명 출력
            
            fnt = VBA.InStr(sFilePath, "(")
            rear = VBA.InStr(sFilePath, ")")
            sFileNum = VBA.Trim(VBA.Mid(sFilePath, fnt + 1, rear - fnt - 1))
            targetWS.Cells(inputRow, 3).Value = sFileNum '// 파일 번호 출력
            
            foldName = VBA.Trim(VBA.Left(sFilePath, VBA.InStr(sFilePath, ".xls") - 1))    '// 확장자 제거(생성할 폴더 이름)
                                    
            sFileName = VBA.Trim(VBA.Mid(foldName, VBA.InStr(foldName, ")") + 1)) '// foldName에서 번호 제거(프로그램 이름)
            targetWS.Cells(inputRow, 4).Value = sFileName '// 프로그램 이름 출력
                    
            sDir = VBA.Replace(filesLoc(i), sFilePath, vbNullString, fnt)
            sDir = VBA.Left(sDir, VBA.Len(sDir) - 1)    '// 결과 폴더 디렉토리
            
            
        End If
        
        
        If VBA.Dir(sDir & "\" & foldName, vbDirectory) = vbNullString Then
            fso.CreateFolder (sDir & "\" & foldName)   '// 디렉토리 내에 없으면 폴더 생성
        End If
            
        If VBA.Dir(sDir & "\" & foldName, vbDirectory) = foldName Then
            fso.MoveFile filesLoc(i), sDir & "\" & foldName & "\" & sFilePath   '// 생성한 폴더로 파일 이동
            'Call ExportForGIT(sDir & "\" & sFilePath)    '// 모듈 파일 분리
            Call makeFolder(sDir & "\" & foldName)
            Call exportModule(sDir & "\" & foldName, sFilePath, inputRow)
        End If
        
        inputRow = inputRow + 1
        
    Next i
    
    
    '################################
    '### 암호 잠긴 파일로 덮어씌우기
    '################################
    
    cryptLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True)  '// 암호 잠긴 파일들 경로 저장
    
    
    If VBA.IsEmpty(cryptLoc) Then
        Exit Sub
    End If
    
    Dim cFilePath, destFold As String
    Dim csplitResult As Variant
    Dim M As Long
    Dim indexRow As Long: indexRow = 2
    cFilePath = cryptLoc(0)
    
    csplitResults = VBA.Split(cFilePath, "\")
    If LBound(csplitResults) >= 0 Then
        cFilePath = VBA.Replace(cFilePath, csplitResults(UBound(csplitResults)), vbNullString)
        mainWS.Cells(13, 14).Value = cFilePath
    End If
    
    For M = LBound(cryptLoc, 1) To UBound(cryptLoc, 1)
        cFilePath = cryptLoc(M)
    
        csplitResults = VBA.Split(cFilePath, "\")    '// \ 문자 기준으로 분할
        If LBound(csplitResults) >= 0 Then
            cFilePath = csplitResults(UBound(csplitResults))
            
            destFold = VBA.Trim(VBA.Left(cFilePath, VBA.InStr(cFilePath, ".xls") - 1))    '// 찾아야하는 폴더 이름
        End If
    
        If VBA.Dir(sDir & "\" & destFold, vbDirectory) = destFold Then
            fso.copyfile cryptLoc(M), sDir & "\" & destFold & "\", True   '// 생성한 폴더로 파일 이동
            If targetWS.Cells(indexRow, 5) = "" Then
                targetWS.Cells(indexRow, 5).Value = "작업 완료"
            End If
        End If
    
    
        indexRow = indexRow + 1
    
    Next M
        
    Set fso = Nothing
 
End Sub




'##################################################
'* Excel 모듈 파일 자동 내보내기
'* 1. [ExcelVBAFileExport.bas] 파일 다운로드
'* 2. [파일 가져오기]하여 모듈에 추가
'* 3. ExportForGIT()메소드에서 Dubug.Print에 커서 위치시킴
'* 4. 실행{F5}을 눌러 실행하세요.
'##################################################
'Private Sub ExportForGIT(xFullPath As String)
'
'    Debug.Print "엑셀 모듈 파일 내보내기"
'    '=====================================
'    Dim sFolderName As String: sFolderName = VBA.Trim(VBA.Left(xFullPath, VBA.InStr(xFullPath, ".xls") - 1))
'    '=====================================
'    '//1.참조 추가
'    If addReferences = False Then
'        Exit Sub
'    End If
'    '=====================================
'    '//2.폴더 삭제 및 생성
'    Call makeFolder(sFolderName)
'    '=====================================
'    '//3.내보내기 및 삭제
'    Call exportModule(xFullPath)
'    '=====================================
'End Sub

'========================================
'*  폴더 삭제 및 생성
'========================================
Private Sub makeFolder(filePath As String)
    Dim fso As Object:  Set fso = CreateObject("Scripting.FileSystemObject")
    Dim mainFolder As Object: Set mainFolder = fso.GetFolder(filePath)
    Dim folders As Object

Try:
    On Error GoTo Catch
    '폴더 삭제
    If mainFolder.SubFolders.Count > 0 Then
        Dim folderName As String
        For Each folders In mainFolder.SubFolders
            folderName = folders.Name
            folderName = Split(folderName, "\")(UBound(Split(folderName, "\")))

            If (folderName = "class") Or (folderName = "form") Or (folderName = "normal") Or (folderName = "sheet") Then
                fso.deletefolder (folders)
            End If
        Next
    End If

    GoTo Finally
Catch:
    'do nothing

Finally:
    '폴더 생성
    fso.CreateFolder (filePath & "\class")
    fso.CreateFolder (filePath & "\form")
    fso.CreateFolder (filePath & "\normal")
    fso.CreateFolder (filePath & "\sheet")

    Set fso = Nothing
    Set mainFolder = Nothing
End Sub

'=====================================
'* 모듈 파일 내보내기
'=====================================
Private Function exportModule(ByVal xFolderPath As String, ByVal xFileName As String, inputRow As Long) As Boolean
On Error GoTo ErrHANDLER
    Application.DisplayAlerts = False
    Dim targetWB As Workbook
    Set targetWB = Workbooks.Open(xFolderPath & "\" & xFileName, False)

    If ProtectedVBProject(targetWB) = True Then
        targetWS.Cells(inputRow, 5).Value = "LOCKED"
        targetWS.Cells(inputRow, 5).Font.Bold = True
        targetWS.Cells(inputRow, 5).Font.Color = RGB(25, 100, 126)
        GoTo exitSub
    End If
    
    Dim vbProj As VBIDE.VBProject:  Set vbProj = targetWB.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim moduleFileName As String

    '//모듈별 저장하기
    For Each vbComp In vbProj.VBComponents
        '=====================================
        '1) 확장자설정
        Select Case vbComp.Type
            'Case vbext_ct_StdModule
            Case 1
                moduleFileName = xFolderPath & "\normal\" & vbComp.Name & ".bas"
            'Case vbext_ct_ClassModule
            Case 2
                moduleFileName = xFolderPath & "\class\" & vbComp.Name & ".cls"
            'Case vbext_ct_MSForm
            Case 3
                moduleFileName = xFolderPath & "\form\" & vbComp.Name & ".frm"
            'Case vbext_ct_Document
            Case 100
                '현재_통합_문서
                moduleFileName = xFolderPath & "\sheet\" & vbComp.Name & ".cls"
        End Select
        '=====================================
        '2) 내보내기 및 삭제
        If moduleFileName <> "" Then
            vbComp.Export (moduleFileName)
        End If
    Next
    exportModule = True
    GoTo exitSub
ErrHANDLER:
If Err.Number = 0 Then
    Resume
Else
    Debug.Print Err.Number & " : " & Err.Description
End If
exitSub:
If Not targetWB Is Nothing Then
    Call targetWB.Close(False)
End If

Application.DisplayAlerts = True
End Function

Private Function ProtectedVBProject(ByVal wb As Workbook) As Boolean
' returns TRUE if the VB project in the active document is protected
    Dim VBC As Integer: VBC = -1
    
    On Error Resume Next
    VBC = wb.VBProject.VBComponents.Count
    
    On Error GoTo 0
    If VBC = -1 Then
        ProtectedVBProject = True
    Else
        ProtectedVBProject = False
    End If
End Function


'##################################################
'* Add Reference
'##################################################
Private Function addReferences() As Boolean
    If programmaticAccessAllowed = False Then
        MsgBox ("보안센터 - 매크로설정 - VBA 객체에 안전하게 액세스할 수 있음 설정값을 확인해주세요. (값변경 후 워크북을 다시 오픈하세요.)")
    Else
        'VBA extensibilities 5.3
        Call addRefFromGUID("{0002E157-0000-0000-C000-000000000046}", 5, 3)
        addReferences = True
    End If
End Function

Private Sub addRefFromGUID(strGUID As String, majorVersion As Long, minorVersion As Long)
'엑셀설정 중 보안센터 - 매크로설정 - VBA 객체에 안전하게 액세스할 수 있음을 체크되야함
On Error GoTo exitSub
    Call removeMissingRefs

    If hasRef(strGUID) Then
        Exit Sub
    End If

    On Error Resume Next

     'Add the reference
    Call ThisWorkbook.VBProject.References.AddFromGuid(guid:=strGUID, Major:=majorVersion, Minor:=minorVersion)

     'If an error was encountered, inform the user
    Select Case Err.Number
        Case Is = 32813
             'Reference already in use.  No action necessary
        Case Is = vbNullString
            'MsgBox ("..")
             'Reference added without issue
        Case Else
             'An unknown error was encountered, so alert the user
            MsgBox "A problem was encountered trying to" & vbNewLine _
            & "add or remove a reference in this file" & vbNewLine & "Please check the " _
            & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"

    End Select
    On Error GoTo 0
exitSub:
End Sub

Private Sub removeMissingRefs()
On Error GoTo ErrHANDLER

    Dim i As Long
    Dim theRef As Variant

    '엑셀2007, 2010 버전에서 ThisWorkbook.VBProject.References.count를 못읽는 버그가 있음 무한루프
    For i = ThisWorkbook.VBProject.References.Count To 1 Step -1
        Set theRef = ThisWorkbook.VBProject.References.Item(i)
        If theRef.isbroken = True Then
            ThisWorkbook.VBProject.References.Remove theRef
        End If
    Next i

    Err.Clear
    Exit Sub

ErrHANDLER:
    If Err.Number = 0 Then
        Resume
    Else
        Debug.Print Err.Number & " : " & Err.Description & " at removeMissingRefs"
    End If
End Sub

Public Function hasRef(xGUID As String) As Boolean
On Error GoTo exitSub
    Dim i As Integer
    hasRef = False

    With Application.ThisWorkbook.VBProject.References
        For i = 1 To .Count
            If .Item(i).guid = xGUID Then
                hasRef = True
                Exit Function
            End If
        Next i
    End With
exitSub:
End Function

'========================================
'* From cookieDevCommon
'* 엑셀설정 중 보안센터 - 매크로설정 - VBA 객체에 안전하게 액세스할 수 있음을 체크되야함
'* X VBA extense 5.3 참조걸려있는지 체크
'* 테스트 완료
'========================================
Private Function programmaticAccessAllowed() As Boolean
On Error Resume Next
    Dim vbTest As Object
    Set vbTest = ThisWorkbook.VBProject
    If Err.Number = 0 Then
        programmaticAccessAllowed = True
    End If
End Function
