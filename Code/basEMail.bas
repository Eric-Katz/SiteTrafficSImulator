Attribute VB_Name = "basEMail"
Option Explicit

Function SendMail_MAPI_POP(strDisplayName As String, strAddress As String, strSubject As String, Optional strMessage As String) As Boolean
    Dim objMail As Object
    
    Set objMail = CreateObject("Persits.MailSender")
    objMail.host = "FunLA.com"                  ' Specify a valid SMTP server
    objMail.From = "FLA_Admin@FunLA.com"        ' Specify sender's address
    objMail.FromName = "Viper Simulator"        ' Specify sender's name

    objMail.UserName = "FLA_Admin"
    objMail.Password = "strAddress"

    objMail.AddAddress strDisplayName, strAddress
    objMail.Subject = strSubject
    objMail.Body = strMessage

    On Error Resume Next
    objMail.Send
    If Err <> 0 Then
        ' Note: Error = 4 may indicate proxy server is denying port access
        Debug.Print "Error encountered: " & Err.Description
    End If

End Function

Function SendMail_MAPI(mpsSession As MAPISession, mpmMessage As MAPIMessages, strDisplayName As String, strAddress As String, strSubject As String, strMessage As String) As Boolean
    
    ' Sign in using session object
    On Error GoTo errLogInFail
    mpsSession.UserName = "webtrend"
    mpsSession.Password = "ecomrpt"
    mpsSession.SignOn
    mpmMessage.SessionID = mpsSession.SessionID
    
    On Error GoTo ErrMessageSend
    'Compose new message
    mpmMessage.Compose

    'Address message
    mpmMessage.RecipDisplayName = strDisplayName
    mpmMessage.RecipAddress = strAddress

    ' Resolve recipient name
    mpmMessage.AddressResolveUI = True
    mpmMessage.ResolveName

    'Create the message
    mpmMessage.MsgSubject = strSubject
    mpmMessage.MsgNoteText = strMessage

    'Send the message
    mpmMessage.Send False

    ' Successfull exit
    SendMail_MAPI = True
    Exit Function
    
errLogInFail:
    If (Err.Number = 32050) Then
        ' Already logged in, continue
        Resume Next
    End If
    If (Err.Number = 32003) Then
        MsgBox "Canceled Login"
        SendMail_MAPI = False
    End If
    Debug.Print Err.Number, Err.Description
    Exit Function

ErrMessageSend:
    Debug.Print Err.Number, Err.Description
    Exit Function

End Function

