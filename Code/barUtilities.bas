Attribute VB_Name = "basUtilities"
Option Explicit

' Win API's
Public Declare Function SleepEx Lib "kernel32" (ByVal dwMilliSeconds As Long, ByVal bAlertable As Long) As Long

Function strEndsWith(aString As String, sufix As String) As Boolean
    strEndsWith = InStrRev(aString, sufix) = Len(aString) - Len(sufix) + 1
End Function

Function Sleep(MilliSeconds As Long, Optional Resolution As Long = 10, Optional bAlertable As Long = 0)
    ' Force minimum 10 ms resolution
    If (Resolution <= 0) Then Resolution = 10
    
    ' Release the processor during sleep
    If (MilliSeconds > 2 * Resolution) Then
        ' Round up to the highest RESOLUTION ms
        MilliSeconds = Ceil(MilliSeconds / CDbl(Resolution)) * Resolution
        Dim ii As Integer
        For ii = 1 To (MilliSeconds / Resolution)
            SleepEx Resolution, bAlertable
            DoEvents
        Next ii
    Else
        SleepEx MilliSeconds, bAlertable
    End If
End Function

Function Ceil(value As Double)
    Ceil = CDbl(CLng(value + 0.5))
End Function
