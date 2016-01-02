Attribute VB_Name = "basEnumWin"
Option Explicit

Public Const MAX_PATH = 260
Public Const LB_SETTABSTOPS As Long = &H192
Public Const BM_CLICK As Long = &HF5

Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function FindWindow Lib "user32" _
    Alias "FindWindowA" _
   (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" _
  Alias "FindWindowExA" _
  (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
  ByVal lpsz1 As String, ByVal lpsz2 As Any) As Long

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetFocusAPI Lib "user32" _
    Alias "SetFocus" _
   (ByVal hWnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" _
   (ByVal hWnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" _
   (ByVal hWnd As Long) As Long

Public Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" _
   (ByVal hWnd As Long, _
    ByVal lpClassName As String, _
    ByVal nMaxCount As Long) As Long

Public Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" _
   (ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" _
    Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" _
   (ByVal hWnd As Long) As Long

Public Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
   
  'working vars
   Dim nSize As Long
   Dim sTitle As String
   Dim sClass As String
   Dim pos As Integer
   
  'set up the strings to receive the class and
  'window text. You could use GetWindowTextLength,
  'but I'll cheat and use MAX_PATH instead.
   sTitle = Space$(MAX_PATH)
   sClass = Space$(MAX_PATH)
   
   Call GetClassName(hWnd, sClass, MAX_PATH)
   Call GetWindowText(hWnd, sTitle, MAX_PATH)

  'strip the trailing chr$(0)'s from the strings
  'returned above and add the window data to the list
   sTitle = TrimNull(sTitle)
   sClass = TrimNull(sClass)
  
   frmTest.List1.AddItem sTitle & vbTab & sClass & vbTab & hWnd
                       
  'to continue enumeration, we must return True
  '(in C that's 1).  If we wanted to stop (perhaps
  'using if this as a specialized FindWindow method,
  'comparing a known class and title against the
  'returned values, and a match was found, we'd need
  'to return False (0) to stop enumeration. When 1 is
  'returned, enumeration continues until there are no
  'more windows left.
   EnumWindowProc = 1

End Function

Private Function TrimNull(item As String)

  'remove string before the terminating null(s)
   Dim pos As Integer
   
   pos = InStr(item, Chr$(0))
   
   If pos Then
         TrimNull = Left$(item, pos - 1)
   Else: TrimNull = item
   End If
   
End Function

