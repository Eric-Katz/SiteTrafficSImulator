Attribute VB_Name = "basMS_Patch"
Option Explicit

Sub MS_DismissDialogs()
    ' If a security or refresh alert shows, get rid of it
    MS_ClickButton "Security Alert", "&Yes"
    MS_ClickButton "Security Alert", "OK"
    MS_ClickButton "Microsoft Internet Explorer", "&Retry"
    
    ' Need to dismiss "Microsoft is conducting online usability studies..."
    MS_ClickButton "Security Alert", "Cancel"
End Sub

Sub MS_FocusControl(blnState As Boolean)
    On Error Resume Next
    Static wHwnd As Long
    Dim ii As Integer
    
    ' Get or release the focus
    If (blnState) Then
        wHwnd = GetForegroundWindow()
    Else
        ' Allow the previous operation to complete before resetting the focus
        ' Reset the focus and yield for smooth operation
        If (wHwnd <> 0) Then
            For ii = 1 To 5
                DoEvents
                If (wHwnd <> GetForegroundWindow()) Then SetForegroundWindow (wHwnd)
            Next
        End If
    End If
    
End Sub

Public Function MS_ClickButton(WindowTitle As String, ButtonText As String) As Boolean
    Dim wHwnd As Long
    Dim cHwnd As Long

    ' Find the current focus
    MS_FocusControl (True)
    
    ' Does the window exist?
    wHwnd = FindWindow(vbNullString, WindowTitle)
    If wHwnd = 0 Then GoTo Exit_Tag
    cHwnd = FindWindowEx(wHwnd, 0, vbNullString, ButtonText)
    If cHwnd = 0 Then GoTo Exit_Tag
    
    ' Bring the target window to the top and (MSDN recommended)
    ' BringWindowToTop wHwnd
    
    ' Click the button
    SendMessage cHwnd, BM_CLICK, 0, 0
    DoEvents
    
    ' Indicate success
    MS_ClickButton = True
    
Exit_Tag:
    ' Restore the focus to the previous window
    MS_FocusControl (False)
    Exit Function

End Function

