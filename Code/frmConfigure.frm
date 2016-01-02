VERSION 5.00
Begin VB.Form frmConfigure 
   Caption         =   "Configure"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4380
   ScaleHeight     =   2010
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame frmLoginSetttings 
      Caption         =   "Login Settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtSuccessPeriodicity 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "5"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtLoginFreq 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         Text            =   "50"
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtPeriod 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Text            =   "50"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblSuccessPeriodicity 
         Caption         =   "Success Periodicity"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   765
         Width           =   1575
      End
      Begin VB.Label lblLoginFreq 
         Caption         =   "Login Freq (%)"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label lblTestPeriod 
         Caption         =   "Period (minutes/10)"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   405
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Const FILE_INI = "ViperSim.ini"             ' Local INI file
Const INI_CONFIG = "Configuration"          ' Configuration section
Const INI_PERIOD = "Period"                 ' Period in minutes
Const INI_SUCCESS = "SuccessPeriodicity"    ' Period in minutes
Const INI_LOGINFREQ = "LoginFreq"           ' Login Frequency (in percent)

Const TIMER_PERIODMIN = 1           ' 1/10 minute
Const TICK_SUCCESSPERIODICITY = 1   ' 1 ticks = 1 * Period * Interval minutes
Const LOGIN_FREQDEF = 100           ' Login Frequency (in percent)

Private Sub Form_Load()
    Dim FilePath As String
    FilePath = App.Path
    If (Not strEndsWith(FilePath, "\")) Then FilePath = FilePath & "\"
    FilePath = FilePath & FILE_INI
    
    txtPeriod = ProfileGet(FilePath, INI_CONFIG, INI_PERIOD, txtPeriod)
    txtSuccessPeriodicity = ProfileGet(FilePath, INI_CONFIG, INI_SUCCESS, txtSuccessPeriodicity)
    txtLoginFreq = ProfileGet(FilePath, INI_CONFIG, INI_LOGINFREQ, txtLoginFreq)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim FilePath As String
    FilePath = App.Path
    If (Not strEndsWith(FilePath, "\")) Then FilePath = FilePath & "\"
    FilePath = FilePath & FILE_INI
    
    Call ProfileSet(FilePath, INI_CONFIG, INI_PERIOD, txtPeriod)
    Call ProfileSet(FilePath, INI_CONFIG, INI_SUCCESS, txtSuccessPeriodicity)
    Call ProfileSet(FilePath, INI_CONFIG, INI_LOGINFREQ, txtLoginFreq)
End Sub

Public Function GetPeriod() As Long
    If (CLng(txtPeriod) < TIMER_PERIODMIN) Then GetPeriod = TIMER_PERIODMIN Else GetPeriod = CLng(txtPeriod)
End Function

Public Function GetSuccessPeriodicity() As Long
    If (CLng(txtSuccessPeriodicity) < TICK_SUCCESSPERIODICITY) Then GetSuccessPeriodicity = TICK_SUCCESSPERIODICITY Else GetSuccessPeriodicity = CLng(txtSuccessPeriodicity)
End Function

Public Function GetLoginDecision() As Boolean
    Dim Random As Integer
    Dim Frequency As Integer
    
    ' Get the login probability
    If ((CInt(txtLoginFreq) < 0) Or (CInt(txtLoginFreq) > 100)) Then Frequency = LOGIN_FREQDEF Else Frequency = CInt(txtLoginFreq)
    
    ' Generate a random number between 0 and 99
    Random = (100 * Rnd())

    ' Return the decision based on Login Frequency
    GetLoginDecision = (Frequency > Random)

End Function

Private Function ProfileSet(FileName As String, SectionName As String, KeyName As String, Value As String) As Boolean
    WritePrivateProfileString SectionName, KeyName, Value, FileName
End Function

Private Function ProfileGet(FileName As String, SectionName As String, KeyName As String, Optional Default As String = "") As String
    Dim Buffer As String

    'call the API. If passing null as the parameter for the lpKeyName
    ' parameter the will API to return a list of all keys under that
    ' section. Pad the passed string large enough to hold the data.
    Buffer = Space$(2048)
    If (GetPrivateProfileString(SectionName, KeyName, Default, Buffer, Len(Buffer), FileName)) Then
        ProfileGet = Buffer
    Else
        ProfileGet = Default
    End If
End Function

Private Function strEndsWith(aString As String, sufix As String) As Boolean
    strEndsWith = InStrRev(aString, sufix) = Len(aString) - Len(sufix) + 1
End Function

