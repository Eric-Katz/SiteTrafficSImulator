VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   6525
   ClientTop       =   3870
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   6585
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   3240
      TabIndex        =   11
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton cmdEnumWinProc 
      Caption         =   "Enum"
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   5
      Left            =   360
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Text            =   "Text4"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "Test Text"
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "198.147.91.2"
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ping"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblCount 
      Caption         =   "Count"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function EnumWindows _
   Lib "user32" (ByVal lpEnumFunc As Long, _
                 ByVal lParam As Long) As Long


Private Sub Form_Load()

   ReDim TabArray(0 To 2) As Long
   
   TabArray(0) = 0
   TabArray(1) = -142
   TabArray(2) = 154
   
  'clear any existing tabs
   Call SendMessage(List1.hwnd, LB_SETTABSTOPS, 0&, ByVal 0&)
   
  'set list tabstops
   Call SendMessage(List1.hwnd, LB_SETTABSTOPS, 3&, TabArray(0))
   
End Sub


Private Sub cmdEnd_Click()

   Unload Me

End Sub


Private Sub cmdEnumWinProc_Click()

   List1.Clear

  'enumerate the windows passing the AddressOf the
  'callback function.  This example doesn't use the
  'lParam member.
   Call EnumWindows(AddressOf EnumWindowProc, &H0)

  'show the window count
   lblCount = List1.ListCount & " windows found."

End Sub



' Form code
'To a form add a command button (Command1), two text boxes (Text1, Text2) to the top of the form, and six text boxes in a control array (Text4(0) - Text4(5)) below. The labels are optional. Add the following to the form:
Private Sub Command1_Click()
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim pos As Long
   Dim success As Long
   
   If SocketsInitialize() Then
   
     'ping the ip passing the address, text
     'to send, and the ECHO structure.
      success = Ping((Text1.Text), (Text2.Text), ECHO)
      
     'display the results
      Text4(0).Text = GetStatusCode(success)
      Text4(1).Text = ECHO.Address
      Text4(2).Text = ECHO.RoundTripTime & " ms"
      Text4(3).Text = ECHO.DataSize & " bytes"
      
      If Left$(ECHO.Data, 1) <> Chr$(0) Then
         pos = InStr(ECHO.Data, Chr$(0))
         Text4(4).Text = Left$(ECHO.Data, pos - 1)
      End If
   
      Text4(5).Text = ECHO.DataPointer
      
      SocketsCleanup
      
   Else
   
        MsgBox "Windows Sockets for 32 bit Windows " & _
               "environments is not successfully responding."
   
   End If
   
End Sub




