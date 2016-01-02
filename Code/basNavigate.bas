Attribute VB_Name = "basNavigate"
Option Explicit

Global Const TIME_TO_COMPLETE = 30     ' 30 Seconds max to complete operation
Global Const TIME_TO_GOBUSY = 2       '  2 seconds max to go to busy state

'typedef enum BrowserNavConstants {
'    navOpenInNewWindow = 0x1,
'    navNoHistory = 0x2,
'    navNoReadFromCache = 0x4,
'    navNoWriteToCache = 0x8,
'    navAllowAutosearch = 0x10,
'    navBrowserBar = 0x20,
'    navHyperlink = 0x40
'} BrowserNavConstants;
Const navOpenInNewWindow = &H1
Const navNoReadFromCache = &H4

Function SimulateGoToAddress(brwBrowser As InternetExplorer, strAddress As String, strPhrase As String) As Boolean
    On Error GoTo Error_Tag
    
    ' Find the current focus
    MS_FocusControl (True)
    
    ' Try to get to address
    brwBrowser.Navigate strAddress
    
' Separate Process?
' Static second As Boolean
' If (second) Then brwBrowser.Navigate strAddress Else brwBrowser.Navigate strAddress, 0, "_BLANK"
    
    ' Force a refresh
    If (Not WaitForComplete(brwBrowser)) Then Exit Function
    brwBrowser.Refresh
    
    ' Restore the focus to the previous window
    MS_FocusControl (False)
    
    ' Browser ready?
    If (Not WaitForComplete(brwBrowser)) Then Exit Function

    ' Success?
    SimulateGoToAddress = FindTestPhrase(brwBrowser, strPhrase)
    Exit Function
    
Error_Tag:
    ' Indicate failure
    SimulateGoToAddress = False
    Exit Function
End Function

Function SimulateClick(brwBrowser As InternetExplorer, strLink As String, strPhrase As String) As Boolean
    Dim link As HTMLLinkElement
    Dim lnkSelected As HTMLLinkElement
    On Error GoTo Error_Tag
    
    ' Get the links for this page
    Dim document As HTMLDocument
    Set document = brwBrowser.document

    ' Search the links
    For Each link In document.Links
        ' Debug.Print link.outerText, link.toString
        ' Look for link text visible to the user ("" for img link)
        If (strEndsWith(link.outerText, strLink)) Then
            Set lnkSelected = link
            Exit For
        End If
        
        ' Look for href text (note that href property does not seem to work)
        If (strEndsWith(link.toString, strLink)) Then
            Set lnkSelected = link
            Exit For
        End If

    Next link
    
    If (lnkSelected Is Nothing) Then
        ' Indicate failure
        SimulateClick = False
        Exit Function
    End If
    
    ' Activate link
    link.Click
    If (Not WaitForComplete(brwBrowser)) Then Exit Function
    
    ' Success?
    SimulateClick = FindTestPhrase(brwBrowser, strPhrase)
    Exit Function
    
Error_Tag:
    ' Indicate failure
    SimulateClick = False
    Exit Function
End Function

Function SimulateEntry(brwBrowser As InternetExplorer, strField As String, strValue As String) As Boolean
    On Error GoTo Error_Tag
    
    ' Get the input buttons for this form
    Dim document As HTMLDocument
    Set document = brwBrowser.document

    Dim buttons As IHTMLElementCollection
    Set buttons = document.All.tags("INPUT")

'    Testing
'    Dim button As HTMLButtonElement
'    For Each button In buttons
'        Debug.Print button.outerHTML
'    Next button

    ' Enter the field data
    buttons(strField).value = strValue
    SimulateEntry = True
    Exit Function
    
Error_Tag:
    ' Indicate failure
    SimulateEntry = False
    Exit Function
End Function

Function SimulateSubmit(brwBrowser As InternetExplorer, strSubmit As String, strPhrase As String) As Boolean
    On Error GoTo Error_Tag
    
    ' Get the input buttons for this form
    Dim document As HTMLDocument
    Set document = brwBrowser.document

    Dim buttons As IHTMLElementCollection
    Set buttons = document.All.tags("INPUT")

    ' Find the current focus
    MS_FocusControl (True)
    
    ' Execute the submit
'    buttons(strSubmit).Click
    Dim ii As Integer
    For ii = 1 To buttons.length
        If (buttons(ii).value = strSubmit) Then
            buttons(ii).Click
            Exit For
        End If
    Next ii
    
    ' Restore the focus to the previous window
    MS_FocusControl (False)

    ' Wait for browser to finish
    If (Not WaitForComplete(brwBrowser)) Then Exit Function
 
     ' Success?
    SimulateSubmit = FindTestPhrase(brwBrowser, strPhrase)
    Exit Function
    
Error_Tag:
    ' Indicate failure
    SimulateSubmit = False
    Exit Function
End Function

Function WaitForComplete(brwBrowser As InternetExplorer)
    ' Wait for the browser to go to busy before looking for complete
    WaitForBusyState brwBrowser, True, TIME_TO_GOBUSY
    WaitForComplete = WaitForBusyState(brwBrowser, False, TIME_TO_COMPLETE)
End Function

Private Function WaitForBusyState(brwBrowser As InternetExplorer, Optional blnState As Boolean = False, Optional lngSecTimeout As Long = TIME_TO_COMPLETE) As Boolean
    Const SLEEP_MS = 10                         ' Default sleep time in milliseconds
    Const STABLECOUNT = 10                      ' 100 ms of stable state is sufficient
    Dim datStartTime As Date
    Dim lngBusyCount As Long
    Dim lngHwnd As Long
    
    ' Start timeout timer
    datStartTime = Now()
    
    ' Wait for completion
    Do While (lngBusyCount < STABLECOUNT)
        
        ' Find the current focus
        lngHwnd = GetForegroundWindow()
        
        ' Yield the processor
        DoEvents
        
        ' Restore the focus to the previous window
        If (lngHwnd <> 0) Then SetForegroundWindow (lngHwnd)
        
        ' If a security or refresh alert shows, get rid of it
        Call MS_DismissDialogs
        
        If (Now() > (datStartTime + TimeSerial(0, 0, lngSecTimeout))) Then
'        Debug.Print "Start: " & datStartTime & ", End: " & (datStartTime + TimeSerial(0, 0, lngSecTimeout))
            Exit Do
        End If
        
        ' The Busy Property oscillates up and down, so look for a steady state
        ' Restart count if state is not the one expected
        lngBusyCount = IIf((brwBrowser.Busy = blnState), lngBusyCount + 1, 0)
        
        ' Take a nap
        Sleep SLEEP_MS
    Loop
    
    ' Return True if browser succeeded
    WaitForBusyState = CBool(brwBrowser.Busy = blnState)
'    Debug.Print "End Wait, Return: " & WaitForBusyState & ", Busy: " & brwBrowser.Busy
    Exit Function
End Function

Function FindTestPhrase(brwBrowser As InternetExplorer, strPhrase As String) As Boolean
    Dim strGetHTML As String
    
    On Error GoTo Browser_Err
    
    ' Don't accept empty search strings
    If (strPhrase = "") Then Exit Function
    
    ' Get HTML text
    ' Note: This line can throw an Runtime error '91': Object variable or With block variable not set
    strGetHTML = brwBrowser.document.Body.innerHTML
    If (strGetHTML = "") Then Exit Function

    ' Search for test phrase
    If (InStr(LCase(strGetHTML), LCase(strPhrase)) = 0) Then Exit Function
    
    ' Success
    FindTestPhrase = True
    Exit Function
    
Browser_Err:
    ' Browser Error
    FindTestPhrase = False
    Exit Function

End Function

