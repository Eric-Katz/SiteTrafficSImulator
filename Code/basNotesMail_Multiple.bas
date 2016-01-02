Attribute VB_Name = "basNotesMail_Multiple"
Option Explicit

' Here is the code that sends the e-mail.
' The remarked last input variable, varBRecip() is for a bcc, which I never use.

Public Sub SendNotesMail_Multiple( _
    varRecipient() As Variant, _
    varCopyTo() As Variant, _
    strSubject As String, _
    strBodyTxt As String, _
    strPath As String, _
    strAttachment As String, _
    blnSaveIt As Boolean)
'    varBRecip() As Variant, _

On Error GoTo Err_Handler

'Set up the objects required for Automation into lotus notes
    Dim Maildb As Object
    Dim UserName As String
    Dim MailDbName As String
    Dim MailDoc As Object
    Dim AttachME As Object
    Dim Session As Object
    Dim EmbedObj As Object
    
'Start a session to notes
    Set Session = CreateObject("Notes.NotesSession")
    UserName = Session.UserName
    MailDbName = Left$(UserName, 1) & Right$(UserName, _
        (Len(UserName) - InStr(1, UserName, " "))) & ".nsf"
    Set Maildb = Session.GETDATABASE("", MailDbName)
    If Maildb.isOpen = False Then
        Maildb.OPENMAIL
    End If
    Set MailDoc = Maildb.CREATEDOCUMENT
    MailDoc.Form = "Memo"
    MailDoc.sendto = varRecipient()
    If varCopyTo(0) <> "" Then
      MailDoc.CopyTo = varCopyTo()
    End If
'    MailDoc.BlindCopyTo = varBRecip()
    MailDoc.Subject = strSubject
    MailDoc.Body = strBodyTxt
    MailDoc.PostedDate = Now()
    MailDoc.SAVEMESSAGEONSEND = blnSaveIt

'   Set up the embedded object and attachment and attach it
    If strAttachment <> "" Then
      Set AttachME = MailDoc.CREATERICHTEXTITEM("Attachment")
      Set EmbedObj = AttachME.EMBEDOBJECT(1454, "", strPath & strAttachment, _
          "Attachment")
    End If
    MailDoc.Send False, varRecipient()
    
Exit_Here:
'Clean Up
On Error Resume Next
    Set Maildb = Nothing
    Set MailDoc = Nothing
    Set AttachME = Nothing
    Set Session = Nothing
    Set EmbedObj = Nothing
    Exit Sub

Err_Handler:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Here

End Sub


