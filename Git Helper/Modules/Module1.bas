Attribute VB_Name = "Module1"
Option Explicit

Public Enum returnType
    integerType = 1
    rangeType = 2
    TwoD_ArrayType = 3
    NumericType = 4
    stringType = 5
End Enum

Public Enum displayMode
    all = 1
    custom = 2
    pageNumer = 3
End Enum

Public Enum fileType
    ExcelFile = 1
    ImageFile = 2
    all = 3
    Directory = 4
End Enum

Private Const clipboardSheetName = "clipboard"
Private Const conditionSheetName = "conditions"
Private Const dropDownSheetName = "드롭다운리스트"

Private Const mgmtModuleName = "cookieDevCommon"
Private Const sD As String = "9EEA15E5-2B08-436E-8C7B-8AEA4D5D6837"
Public g_FSO As Object

' DLL Declare
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    
    '64비트/32비트 시스템에 따라 파라미터가 다른경우에는 아래구분에 서술합니다.
    #If Win64 Then
    
    #Else
    
    #End If
#Else
    'VBA6에는 32Bit 버전만 존재합니다.
    Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
#End If

Public Function fileDialog_cdc(xMsoFileDialogType As MsoFileDialogType, Optional xTitle As String, Optional xAllowMultiSelect As Boolean = False, _
Optional xFileType As fileType = fileType.all, Optional xInitialFileName As String, Optional findToSubFolders As Boolean = False)

     With Application.FileDialog(xMsoFileDialogType)
        .title = xTitle
        
        If xMsoFileDialogType = msoFileDialogFilePicker Then
            .Filters.Clear
            If xFileType = fileType.ExcelFile Then
                .Filters.Add Description:="Excel Workbooks(*.xlsm; *.xlsx; *.xlsb; *.xls)", Extensions:="*.xlsm;*.xlsx;*.xlsb;*.xls"
                .FilterIndex = 0
            ElseIf xFileType = fileType.ImageFile Then
                .Filters.Add Description:="Image Files(*.bmp; *.gif; *.jpg(jpeg); *.png, *.ico, *.cur)", Extensions:="*.bmp;*.gif;*.jpg;*.jpeg;*.png;*.ico,*.cur"
                .FilterIndex = 0
            End If
            .allowMultiSelect = xAllowMultiSelect
            .InitialFileName = xInitialFileName
        End If
        
        Select Case .show
            ' 선택하지 않음
            Case 0
                MsgBox "취소되었습니다."
                Exit Function
        End Select
        
        If .SelectedItems.Count < 1 Then
            Exit Function
        End If
        
        Dim returnVariant As Variant
        Dim i As Long
        
        If xMsoFileDialogType = msoFileDialogFilePicker Then
            ReDim returnVariant(1 To .SelectedItems.Count)
            For i = 1 To UBound(returnVariant)
                returnVariant(i) = .SelectedItems(i)
            Next i
        ElseIf xMsoFileDialogType = msoFileDialogFolderPicker Then
            Dim arrList As Object
            Set arrList = getCreateObject("System.Collections.ArrayList")
            For i = 1 To .SelectedItems.Count
                If xFileType <> fileType.Directory And xFileType <> fileType.all Then
                    returnVariant = getFilesFromDir(.SelectedItems(i) & "\", arrList, findToSubFolders)
                ElseIf xFileType = fileType.Directory Or xFileType = fileType.all Then
                    returnVariant = getFilesFromDir(.SelectedItems(i) & "\", arrList, findToSubFolders, True)
                Else
                    Debug.Print "PAram 확인"
                    Exit Function
                End If
            Next i
        End If
    End With
    
    If Not VBA.IsEmpty(returnVariant) Then
        Dim returnArrList As Object
        Set returnArrList = getCreateObject("System.Collections.ArrayList")
        Dim addFlag As Boolean
        For i = LBound(returnVariant) To UBound(returnVariant)
            addFlag = False
            
            If xFileType <> fileType.ExcelFile Then
                addFlag = True
            ElseIf xFileType = fileType.ExcelFile And isExcelFilePath(VBA.CStr(returnVariant(i))) Then
                addFlag = True
            End If
            
            If addFlag = True Then
                Call returnArrList.Add(returnVariant(i))
            End If
        Next i
        
        'fileDialog_cdc = returnArrList.toArray()
        returnVariant = returnArrList.ToArray()
        If UBound(returnVariant, 1) = -1 Then
            fileDialog_cdc = Empty
        ElseIf UBound(returnVariant, 1) = LBound(returnVariant) Then
            'fileDialog_cdc = returnVariant(LBound(returnVariant))
            fileDialog_cdc = returnVariant
        Else
            fileDialog_cdc = returnVariant
        End If
    End If

End Function

Public Function getFilesFromDir(xDirPath As String, ByRef arrList As Object, Optional findToSubFolders As Boolean, Optional getDir As Boolean = False)
    
    If g_FSO Is Nothing Then
        Set g_FSO = getCreateObject("scripting.filesystemobject")
    End If
    
    Dim currentFolder As Object: Set currentFolder = g_FSO.GetFolder(xDirPath)
    Dim subFolder As Object
    Dim currentFile As Object
    
    If getDir = False Then
        ' xDirPath 파일 순환
        For Each currentFile In currentFolder.Files
            Call arrList.Add(currentFile.path)
        Next currentFile
    ElseIf getDir = True Then
        Call arrList.Add(currentFolder.path)
    End If
    
     
    If findToSubFolders = True Then
        ' xDirPath 내 하위폴더를 순환
        For Each subFolder In currentFolder.SubFolders
            '재귀
            Call getFilesFromDir(subFolder.path, arrList, findToSubFolders, getDir)
        Next subFolder
    End If
    
    Dim returnVariant As Variant
    returnVariant = arrList.ToArray()
    getFilesFromDir = returnVariant
End Function

Public Function getCreateObject(xCalssStr As String) As Object
On Error GoTo ErrHANDLER
    
    If VBA.UCase(xCalssStr) = VBA.UCase("ArrayList") Then
        Set getCreateObject = getCreateObject("System.Collections.ArrayList")
    ElseIf VBA.UCase(xCalssStr) = VBA.UCase("Dic") Or VBA.UCase(xCalssStr) = VBA.UCase("Dictionary") Then
        Set getCreateObject = getCreateObject("Scripting.Dictionary")
    Else
        Set getCreateObject = CreateObject(xCalssStr)
    End If
    
    Exit Function
ErrHANDLER:
If Err.Number = 0 Then
    Resume
Else
    MsgBox ("프로그램 오류 : 닷넷프레임워크 2.0이 설치되어있는지 확인하시기 바랍니다.")
    ThisWorkbook.FollowHyperlink "https://www.microsoft.com/en-us/download/confirmation.aspx?id=1639"
End If

End Function

Public Function isExcelFilePath(xFilePath As String) As Boolean
    isExcelFilePath = False
    Dim dotIndex As Long: dotIndex = getDotIndexFromFilePath(xFilePath)
    
    If dotIndex = 0 Then
        Exit Function
    End If
    
    Dim extensionStr As String
    extensionStr = VBA.Right(xFilePath, dotIndex)
    
    If VBA.InStr(1, extensionStr, "xls") > 0 Then
        isExcelFilePath = True
    End If

End Function

Private Function getDotIndexFromFilePath(xFilePath As String, Optional checkRightDepth As Long = 5)

    Dim i As Long
    For i = 1 To checkRightDepth
        If VBA.Mid(xFilePath, VBA.Len(xFilePath) - i, 1) = "." Then
            getDotIndexFromFilePath = i
            Exit Function
        End If
    Next i

End Function
