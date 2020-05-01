Attribute VB_Name = "ShellRun"
Option Explicit
Private oShell As WshShell

Public Function ShellRun(sCmd As String) As String
    'reference "Windows Script Host Object Model"
    
    'Dim oShell As Object
    'Dim oShell As WshShell
    
    'Set oShell = CreateObject("WScript.Shell")
    If oShell Is Nothing Then
        Set oShell = New WshShell
    End If
    
    'oShell.Run ("cmd cd C:\Users\Lee\Documents\cookie_solution5497")
    'oShell.CurrentDirectory = "C:\CookieGitlab\Solution\cookie_solution5520"
    'oShell.CurrentDirectory = oShell.SpecialFolders("Desktop")
    
    'Run 메서드는 ExitCode만 반환
    'Exec 메서드는 WshExec를 반환받아 다른 처리가능
    
    'Dim oExec As Object
    Dim oExec As WshExec
    'Dim oOutput As Object
    Dim oOutput As TextStream
    
    Set oExec = oShell.Exec(sCmd)
    Sleep (500)
    Set oOutput = oExec.StdOut
    
    If oExec.ExitCode > 0 Then
        ShellRun = "Failed"
    Else
        ShellRun = "Success"
    End If
    
'    Dim errorCode As Variant
'    Const waitOnReturn As Boolean = True
'    errorCode = oShell.Run(sCmd, 1, waitOnReturn)

    'handle the results as they are written to and read from the StdOut object
'    Dim S As String
'    Dim sLine As String
'    S = oOutput.ReadAll
'    While Not oOutput.AtEndOfStream
'        sLine = oOutput.ReadLine
'        If sLine <> "" Then S = S & sLine & vbCrLf
'    Wend

'    ShellRun = S
End Function

Public Sub ShellChangeCurrentDirectory(xDir As String)
    If oShell Is Nothing Then
        Set oShell = New WshShell
    End If
    
    oShell.CurrentDirectory = xDir
End Sub

Private Sub test()
    
    Dim sRun As String
    sRun = "git clone http://www.coukey.co.kr:5001/solution/cookie_solution5520 C:\CookieGitlab\Solution\cookie_solution5520"
    'sRun = "ipconfig"
    
    Debug.Print ShellRun(sRun)
    
    'sRun = "ipconfig"
    
'    Debug.Print ShellRun(sRun)
'
'    Dim ret_val As Variant
'    ret_val = Shell(sRun, vbNormalFocus)
'    Debug.Print ret_val
    
End Sub
