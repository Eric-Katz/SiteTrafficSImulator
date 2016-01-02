VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Viper Simulator"
   ClientHeight    =   1905
   ClientLeft      =   3315
   ClientTop       =   2520
   ClientWidth     =   4980
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4980
   Begin VB.CommandButton cmdFLATest 
      Caption         =   "Sign In"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdIEBugTest 
      Caption         =   "IE Bug Test"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer tmrSiteCheck 
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdMSGoTo 
      Caption         =   "Go To MS"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "v1.22"
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lblState 
      Alignment       =   2  'Center
      Caption         =   "Idle"
      Height          =   255
      Left            =   150
      TabIndex        =   7
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label lblLastFail 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   150
      TabIndex        =   5
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label lblLastSuccess 
      Alignment       =   2  'Center
      Caption         =   "No Sign-in"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblTestCount 
      Alignment       =   2  'Center
      Caption         =   "Count: 0"
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Menu mnuConfigure 
      Caption         =   "Configure"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Constants
Const TIMER_INTERVAL = 6000     ' 6000 milliseconds = 1/10 minute
Const TICK_FAIL = 1             ' 1 tick  = 1 * Period * Interval minutes
Const TICK_ALERT = 5            ' 5 ticks = 5 * Period * Interval minutes before alert
Const DBG_VISIBLE = False       ' Debug Mode

' Test addresses and strings
Const FLA_SIGNIN_ADDRESS = "http://www.funla.org/social/index.php?page=sign_in"
Const FLA_SIGNIN_VERIFY = "Your username"
Const FLA_USERFIELD = "login"
Const FLA_PASSWORDFIELD = "pswd"
Const FLA_SIGNIN_BUTTON = "Sign in"
Const FLA_SIGNIN_BUTTON_VERIFY = "Members Area"
Const FLA_SIGNIN_PAGENAME = "FLA sign-in page"
Const FLA_SIGNOFF_BUTTON = "Logout"
Const FLA_SIGNOFF_BUTTON_VERIFY = "Free Registration"
Const FLA_SITE_NAME = "FunLA Social"
Const MS_ADDRESS = "http://www.microsoft.com"
Const MS_ADDRESS_VERIFY = "Microsoft Corporation"
Const MS_PAGENAME = "Microsoft Home page"

Dim mpsSession As MAPISession
Dim mpmMessage As MAPIMessages

' Local variables
Private WithEvents objIE As InternetExplorer
Attribute objIE.VB_VarHelpID = -1

Private Sub Form_Load()
    tmrSiteCheck.Interval = TIMER_INTERVAL
    
    ' Associate the events object with the browser control
    ' Set objIE = New InternetExplorer
    
    ' Use CreateProcess to ensure a separate session ID
    Set objIE = IESpawn(DBG_VISIBLE)
    
    ' Load the list of users
    UserListLoad
         
    ' Set manual buttons
    cmdFLATest.Visible = DBG_VISIBLE
    cmdMSGoTo.Visible = DBG_VISIBLE
    cmdIEBugTest.Visible = DBG_VISIBLE

    ' Set operating mode to silent (no dialogs)
    objIE.Visible = DBG_VISIBLE
    objIE.Silent = Not DBG_VISIBLE
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    objIE.Quit
    Set objIE = Nothing
    End
End Sub

Private Sub cmdMSGoTo_Click()
    ' Enable/Disable other controls
    cmdFLATest.Enabled = False
    
    ' Go to a known site
    Call Page_GoTo(MS_PAGENAME, MS_ADDRESS, MS_ADDRESS_VERIFY)
    
    ' Enable/Disable other controls
    cmdFLATest.Enabled = True

End Sub

Private Sub cmdFLATest_Click()
    ' Enable/Disable other controls
    cmdMSGoTo.Enabled = False
    
    ' Test full sign-in, sign-out cycle
    Call FLA_CycleTest(True)
    
    ' Enable/Disable other controls
    cmdMSGoTo.Enabled = True

End Sub
 
Private Sub cmdIEBugTest_Click()
    frmTest.Show
End Sub

Private Sub mnuMain_Click()

End Sub

Private Sub tmrSiteCheck_Timer()
    Static blnActive As Boolean
    Static blnSiteState As Boolean
    Static blnAlertState As Boolean
    Static lngFailCount As Long

    ' Do not permit re-entrancy
    If (blnActive) Then Exit Sub
    
    ' Indicate function is active
    blnActive = True
    
    ' Time to check?
    If (Not IsTimeToCheck(blnSiteState)) Then
        If (InStr(lblState, "Idle") = 0) Then lblState = lblState + " (Idle)"
        blnActive = False
        Exit Sub
    End If
    
    ' Sign In based on login frequency
    Dim bLogin As Boolean
    bLogin = frmConfigure.GetLoginDecision()
    blnSiteState = FLA_CycleTest(bLogin)
    
    ' Is site operational? - inconclusive if not logging in
    If (Not bLogin) Then
        ' Indicate entry available
        blnActive = False
        Exit Sub
    End If
        
    ' Increment failure count?
    lngFailCount = IIf(blnSiteState, 0, lngFailCount + 1)
    
    ' Indicate failure state (will go to "grayed" if site is active)
    ShowFailureState blnSiteState, lngFailCount
    
    ' Was check successful?
    If (blnSiteState) Then
        ' Did an alert email get sent?
        If (blnAlertState) Then
            ' Back up and running
            SendAlert FLA_SITE_NAME & " Site Operational", "Site operational on " & Now()
            blnAlertState = False
        End If
    Else
        ' Failed after x tries
        If (lngFailCount = TICK_ALERT) Then
            SendAlert FLA_SITE_NAME & " Site Failed", lblLastFail
            blnAlertState = True
        End If
    End If

    ' Indicate state change
    blnSiteState = blnSiteState

    ' Indicate entry available
    blnActive = False
   
End Sub

Function FLA_CycleTest(bLogin As Boolean) As Boolean
    Static lngTestCount As Long
    
    ' Dismiss any outstanding dialogs
    Call MS_DismissDialogs
    
    ' Always start from known page
    Call Page_GoTo(FLA_SIGNIN_PAGENAME, FLA_SIGNIN_ADDRESS, FLA_SIGNIN_VERIFY)
    
    ' Sign Off (this version is used for site activity generation)
    Call FLA_SignOff

    ' Go to the Site
    FLA_CycleTest = Page_GoTo(FLA_SIGNIN_PAGENAME, FLA_SIGNIN_ADDRESS, FLA_SIGNIN_VERIFY)
    If (Not FLA_CycleTest) Then Exit Function

    ' Sign in based on login frequency
    If (Not bLogin) Then Exit Function
    
    ' Sign in
    FLA_CycleTest = FLA_SignIn(UserListGetRandom)
    If (Not FLA_CycleTest) Then Exit Function
    
    ' Sign Off (this version is used for site activity check)
    ' Call FLA_SignOff

    ' Update success count
    lngTestCount = CountUpdate(lngTestCount)
    lblLastSuccess = Now()
    
End Function

Private Function Page_GoTo(PageName As String, Address As String, Verify As String) As Boolean
    ' Dismiss any outstanding dialogs
    Call MS_DismissDialogs
    
    ' Try to get to site page
    lblState = "Going to " & PageName & "..."
    Page_GoTo = SimulateGoToAddress(objIE, Address, Verify)
    
    ' Success?
    If (Page_GoTo) Then
        lblState = "Reached " & PageName
    Else
        lblState = "Failed to reach " & PageName
    End If
End Function

Private Function FLA_SignIn(Entry As UserInfo) As Boolean
    ' Fill out the log in fields
    SimulateEntry objIE, FLA_USERFIELD, Entry.UserName
    SimulateEntry objIE, FLA_PASSWORDFIELD, Entry.Password
    
    ' Submit and look for expected response
    lblState = "Signing in as " & Entry.UserName & "..."
    FLA_SignIn = SimulateSubmit(objIE, FLA_SIGNIN_BUTTON, FLA_SIGNIN_BUTTON_VERIFY)
    If (FLA_SignIn) Then
        lblState = "Reached sign in page as " & Entry.UserName
    Else
        lblState = "Failed to reach sign in page as " & Entry.UserName
    End If
    
End Function

Private Function FLA_SignOff() As Boolean
    ' Sign out
    lblState = "Signing out from site..."
    FLA_SignOff = SimulateClick(objIE, FLA_SIGNOFF_BUTTON, FLA_SIGNOFF_BUTTON_VERIFY)
    If (FLA_SignOff) Then
        lblState = "Signed out from site"
    Else
        lblState = "Failed to sign out from site"
    End If
End Function

Private Function CountUpdate(lngCount As Long)
    lngCount = lngCount + 1
    lblTestCount = "Count: " & lngCount
    CountUpdate = lngCount
End Function

Private Function IsTimeToCheck(blnSiteState As Boolean) As Boolean
    Static lngPeriod As Long
    Static lngTick As Long
    
    ' Period expired?
    If (lngPeriod Mod frmConfigure.GetPeriod() <> 0) Then
        lngPeriod = lngPeriod + 1
        Exit Function
    Else
        lngPeriod = lngPeriod + 1
    End If
    
    ' Was last check successful?
    If (blnSiteState) Then
        ' Is it time to check the site?
        If (lngTick >= frmConfigure.GetSuccessPeriodicity()) Then
            IsTimeToCheck = True
            lngTick = 0
        End If
    Else
        ' Is it time to check the site?
        If (lngTick >= TICK_FAIL) Then
            IsTimeToCheck = True
            lngTick = 0
        End If
    End If
    
    ' Tick the timer
    lngTick = lngTick + 1
    
End Function

Private Sub ShowFailureState(blnSiteState As Boolean, lngFailCount As Long)
    ' Indicate site back up or failure
    If (blnSiteState) Then
        lblLastFail.ForeColor = &H80000011
    Else
        lblLastFail = AlertText(lngFailCount) & vbCrLf & AnalyzeError()
        lblLastFail.ForeColor = &HFF&
    End If

End Sub

Private Function AnalyzeError() As String
    ' Database error?
    If (FindTestPhrase(objIE, "system database is not available")) Then
        AnalyzeError = "Database not available."
        Exit Function
    End If
      
    ' Proxy or other possible local error?
    If (Not Page_GoTo(MS_PAGENAME, MS_ADDRESS, MS_ADDRESS_VERIFY)) Then
        AnalyzeError = "Can reach other sites."
        Exit Function
    End If
    
    ' Probably IIS error
    AnalyzeError = "Other sites are reachable."

End Function

Private Function SendAlert(strSubject As String, strMessage As String) As Boolean
    On Error GoTo Error:
    ' Debug.Print (strSubject & vbCrLf & strMessage & vbCrLf)
    ' Microsoft MAPI Mail
    ' SendAlert = SendMail_MAPI(mpsSession, mpmMessage, "*eCommerce Ping", "EP", strSubject, strMessage)
    ' SendAlert = SendMail_MAPI(mpsSession, mpmMessage, "Andrew Michalik", "Michalik", strSubject, strMessage)
    
    ' Lotus Notes
    Dim varRecipient(1) As Variant
    Dim varCopyTo(1) As Variant
    varRecipient(0) = "andrew_michalik@SPE.Sony.com"

    ' SendNotesMail_Multiple varRecipient(), varCopyTo(), strSubject, strMessage, "", "", False
    ' SendNotesMail_Single "andrew michalik/LA/SPE", strSubject, strMessage, "", False

    ' Normal exit
    Exit Function
    
Error:
    Debug.Print (strSubject & vbCrLf & strMessage & vbCrLf)
    
End Function

Private Function AlertText(lngFailCount As Long) As String
    AlertText = "Failed after " & lngFailCount & IIf(lngFailCount > 1, " tries", " try") & " at " & Now()
End Function

Private Sub mnuConfigure_Click()
    frmConfigure.Show
End Sub


