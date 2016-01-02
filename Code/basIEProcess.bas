Attribute VB_Name = "basIEProcess"
Option Explicit

'-------------------------------------------------------------------------------------------------------
Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Public Const NORMAL_PRIORITY_CLASS = &H20&
Public Const INFINITE = -1&
Public Const SW_SHOW& = 5 'sets nShowCmd param in ShellExecute API function
Public Const SW_HIDE& = 0
Public Const STARTF_USESHOWWINDOW& = &H1

Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliSeconds As Long) As Long
'-------------------------------------------------------------------------------------------------------

Private Declare Function GetTickCount Lib "kernel32" () As Long

Function IESpawn(bVisible As Boolean) As InternetExplorer
    Const BLANK_PAGE = "about:blank"
    Const PROGRAM_FILES_VAR = "%programfiles%"
    Dim objShellWins As New SHDocVw.ShellWindows
    Dim BlankURL As String

    ' Get the location of IE
    Dim IE_Path As String
    IE_Path = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\IE Setup\SETUP\", "Path")
    IE_Path = IE_Path & "\iexplore.exe"

    If (1 = InStr(IE_Path, PROGRAM_FILES_VAR)) Then
        Dim PF_Path As String
        PF_Path = ReadRegistry(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\", "ProgramFilesDir")
        IE_Path = Replace(IE_Path, PROGRAM_FILES_VAR, PF_Path)
    End If

    ' Create unique blank page
    BlankURL = BLANK_PAGE & "?" & GetTickCount()
    
    ' Note save before web address
    Call ExecuteApplication(IE_Path, " " & BlankURL, bVisible)
    
    ' There might be multiple IE windows open
    Dim objIE As InternetExplorer
    For Each objIE In objShellWins
        If (objIE.LocationURL = BlankURL) Then
            Set IESpawn = objIE
            Set objShellWins = Nothing
            Exit Function
        End If
    Next
    
    MsgBox ("Could not run IE using CreateProcess for " & IE_Path & ", last error: " & Err.LastDllError)
    
    ' Clean up
    Set objShellWins = Nothing

    ' Could not use CreateProcess
    Set IESpawn = New InternetExplorer
    
End Function

Function ExecuteApplication(ByVal strApplication As String, Optional ByVal strCommandLine As String, Optional bVisible As Boolean = True, Optional ByVal bWait As Boolean = True) As Long

    Dim udtStartup As STARTUPINFO
    Dim udtProcess As PROCESS_INFORMATION
    Dim lPID As Long

    ' Initialize the STARTUPINFO structure:
    udtStartup.cb = Len(udtStartup)
   
    If (Not bVisible) Then
        udtStartup.dwFlags = STARTF_USESHOWWINDOW 'forces function to use wShowWindow param below
        udtStartup.wShowWindow = SW_HIDE 'shelled process is HIDDEN. Use SW_SHOW etc, to run visible
    End If

    lPID = CreateProcess(strApplication, strCommandLine, 0, 0, False, 0, ByVal 0&, vbNullString, udtStartup, udtProcess)
    If lPID <> 0 Then

        ' ajm test
        ' Give IE a chance to go to the busy state
        Sleep 1000

        CloseHandle udtProcess.hThread
        If bWait Then
            WaitForInputIdle udtProcess.hProcess, &HFFFF
        End If

    End If

    ExecuteApplication = lPID

End Function

'-------------------------------------------------------------------------------------------------------

