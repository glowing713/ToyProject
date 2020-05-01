Attribute VB_Name = "GitUpload"


'########################  �߰��� ���۷���  ##########################
'###  1. Microsoft WinHTTP Services, version 5.1
'###  2. Windows Script Host Object Model
'#####################################################################



'============================  ���� ����  ============================
     
' 1. GitUpload�� getFileList() ���� => �ҽ� ����
' 2. GitRepo�� createRepo() ���� => ����� ����
' 3. GitClone�� cloneRepo() ���� => C:\CookieGitlab\Solution ���� cookie_solution[���Ϲ�ȣ]�� Ŭ��
' 4. GitCommit�� commitRepo() ���� => ���� ���� Ŭ�е� ������ �̵�, Ŀ��, Ǫ�ñ��� �Ϸ�
'=====================================================================

Dim targetWS As Worksheet



'#########################
'### ���� ����Ʈ �ҷ�����
'#########################


Private Sub getFileList()


    �ҽ�����.Select
     
     Dim filesLoc, cryptLoc As Variant
     Dim fso As Object: Set fso = getCreateObject("Scripting.FileSystemObject")
   
    '### ��ȣ ���� ���� ���� ���� �� �̵�,
    '### ���� ������ �ҽ� �и�����.
     
     
    filesLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True) '//��ȣ ���� ���ϵ� ��� ����
     
     
    If VBA.IsEmpty(filesLoc) Then
        Exit Sub
    End If
        
    Set targetWS = �ҽ�����
    Dim mainWS As Worksheet: Set mainWS = Main
    Dim fnt, rear As Long
    Dim i As Long
    Dim inputRow As Long: inputRow = 2
        
    Call Main.clearSheet(targetWS)  '// ������ ��ϵ� ���ȭ�� �����
    
    
    '// sFilePath: ���ϸ�(���Ϲ�ȣ, ���α׷� �̸�, Ȯ���� ����), sFileNum: ���Ϲ�ȣ, sFileName: ���α׷� �̸�, foldName: ���ϸ��� Ȯ���ڸ� ����
    Dim sFilePath, sFileName, sFileNum, foldName As String
    Dim splitResults As Variant
    Dim sDir As String  '// ��� ���� ���丮
    sFilePath = filesLoc(0)
    splitResults = VBA.Split(sFilePath, "\")
    If LBound(splitResults) >= 0 Then
        sFilePath = VBA.Replace(sFilePath, splitResults(UBound(splitResults)), vbNullString)
        mainWS.Cells(11, 14).Value = sFilePath  '// Main ȭ�鿡 ������ ��� ���
    End If
    
    For i = LBound(filesLoc, 1) To UBound(filesLoc, 1)
        sFilePath = filesLoc(i)
        
        targetWS.Cells(inputRow, 1).Value = sFilePath   '// ���� ��� ���
    
        splitResults = VBA.Split(sFilePath, "\")    '// \ ���� �������� ����
        If LBound(splitResults) >= 0 Then
            sFilePath = splitResults(UBound(splitResults))
            
            targetWS.Cells(inputRow, 2).Value = sFilePath   '// ���ϸ� ���
            
            fnt = VBA.InStr(sFilePath, "(")
            rear = VBA.InStr(sFilePath, ")")
            sFileNum = VBA.Trim(VBA.Mid(sFilePath, fnt + 1, rear - fnt - 1))
            targetWS.Cells(inputRow, 3).Value = sFileNum '// ���� ��ȣ ���
            
            foldName = VBA.Trim(VBA.Left(sFilePath, VBA.InStr(sFilePath, ".xls") - 1))    '// Ȯ���� ����(������ ���� �̸�)
                                    
            sFileName = VBA.Trim(VBA.Mid(foldName, VBA.InStr(foldName, ")") + 1)) '// foldName���� ��ȣ ����(���α׷� �̸�)
            targetWS.Cells(inputRow, 4).Value = sFileName '// ���α׷� �̸� ���
                    
            sDir = VBA.Replace(filesLoc(i), sFilePath, vbNullString, fnt)
            sDir = VBA.Left(sDir, VBA.Len(sDir) - 1)    '// ��� ���� ���丮
            
            
        End If
        
        
        If VBA.Dir(sDir & "\" & foldName, vbDirectory) = vbNullString Then
            fso.CreateFolder (sDir & "\" & foldName)   '// ���丮 ���� ������ ���� ����
        End If
            
        If VBA.Dir(sDir & "\" & foldName, vbDirectory) = foldName Then
            fso.MoveFile filesLoc(i), sDir & "\" & foldName & "\" & sFilePath   '// ������ ������ ���� �̵�
            'Call ExportForGIT(sDir & "\" & sFilePath)    '// ��� ���� �и�
            Call makeFolder(sDir & "\" & foldName)
            Call exportModule(sDir & "\" & foldName, sFilePath, inputRow)
        End If
        
        inputRow = inputRow + 1
        
    Next i
    
    
    '################################
    '### ��ȣ ��� ���Ϸ� ������
    '################################
    
    cryptLoc = fileDialog_cdc(msoFileDialogFolderPicker, , , fileType.ExcelFile, , True)  '// ��ȣ ��� ���ϵ� ��� ����
    
    
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
    
        csplitResults = VBA.Split(cFilePath, "\")    '// \ ���� �������� ����
        If LBound(csplitResults) >= 0 Then
            cFilePath = csplitResults(UBound(csplitResults))
            
            destFold = VBA.Trim(VBA.Left(cFilePath, VBA.InStr(cFilePath, ".xls") - 1))    '// ã�ƾ��ϴ� ���� �̸�
        End If
    
        If VBA.Dir(sDir & "\" & destFold, vbDirectory) = destFold Then
            fso.copyfile cryptLoc(M), sDir & "\" & destFold & "\", True   '// ������ ������ ���� �̵�
            If targetWS.Cells(indexRow, 5) = "" Then
                targetWS.Cells(indexRow, 5).Value = "�۾� �Ϸ�"
            End If
        End If
    
    
        indexRow = indexRow + 1
    
    Next M
        
    Set fso = Nothing
 
End Sub




'##################################################
'* Excel ��� ���� �ڵ� ��������
'* 1. [ExcelVBAFileExport.bas] ���� �ٿ�ε�
'* 2. [���� ��������]�Ͽ� ��⿡ �߰�
'* 3. ExportForGIT()�޼ҵ忡�� Dubug.Print�� Ŀ�� ��ġ��Ŵ
'* 4. ����{F5}�� ���� �����ϼ���.
'##################################################
'Private Sub ExportForGIT(xFullPath As String)
'
'    Debug.Print "���� ��� ���� ��������"
'    '=====================================
'    Dim sFolderName As String: sFolderName = VBA.Trim(VBA.Left(xFullPath, VBA.InStr(xFullPath, ".xls") - 1))
'    '=====================================
'    '//1.���� �߰�
'    If addReferences = False Then
'        Exit Sub
'    End If
'    '=====================================
'    '//2.���� ���� �� ����
'    Call makeFolder(sFolderName)
'    '=====================================
'    '//3.�������� �� ����
'    Call exportModule(xFullPath)
'    '=====================================
'End Sub

'========================================
'*  ���� ���� �� ����
'========================================
Private Sub makeFolder(filePath As String)
    Dim fso As Object:  Set fso = CreateObject("Scripting.FileSystemObject")
    Dim mainFolder As Object: Set mainFolder = fso.GetFolder(filePath)
    Dim folders As Object

Try:
    On Error GoTo Catch
    '���� ����
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
    '���� ����
    fso.CreateFolder (filePath & "\class")
    fso.CreateFolder (filePath & "\form")
    fso.CreateFolder (filePath & "\normal")
    fso.CreateFolder (filePath & "\sheet")

    Set fso = Nothing
    Set mainFolder = Nothing
End Sub

'=====================================
'* ��� ���� ��������
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

    '//��⺰ �����ϱ�
    For Each vbComp In vbProj.VBComponents
        '=====================================
        '1) Ȯ���ڼ���
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
                '����_����_����
                moduleFileName = xFolderPath & "\sheet\" & vbComp.Name & ".cls"
        End Select
        '=====================================
        '2) �������� �� ����
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
        MsgBox ("���ȼ��� - ��ũ�μ��� - VBA ��ü�� �����ϰ� �׼����� �� ���� �������� Ȯ�����ּ���. (������ �� ��ũ���� �ٽ� �����ϼ���.)")
    Else
        'VBA extensibilities 5.3
        Call addRefFromGUID("{0002E157-0000-0000-C000-000000000046}", 5, 3)
        addReferences = True
    End If
End Function

Private Sub addRefFromGUID(strGUID As String, majorVersion As Long, minorVersion As Long)
'�������� �� ���ȼ��� - ��ũ�μ��� - VBA ��ü�� �����ϰ� �׼����� �� ������ üũ�Ǿ���
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

    '����2007, 2010 �������� ThisWorkbook.VBProject.References.count�� ���д� ���װ� ���� ���ѷ���
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
'* �������� �� ���ȼ��� - ��ũ�μ��� - VBA ��ü�� �����ϰ� �׼����� �� ������ üũ�Ǿ���
'* X VBA extense 5.3 �����ɷ��ִ��� üũ
'* �׽�Ʈ �Ϸ�
'========================================
Private Function programmaticAccessAllowed() As Boolean
On Error Resume Next
    Dim vbTest As Object
    Set vbTest = ThisWorkbook.VBProject
    If Err.Number = 0 Then
        programmaticAccessAllowed = True
    End If
End Function
