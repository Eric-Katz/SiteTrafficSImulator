Attribute VB_Name = "basNotesMail_Single"
'Public Sub SendNotesMail(Subject as string, attachment as string,
'recipient as string, bodytext as string,saveit as Boolean)
'This public sub will send a mail and attachment if neccessary to the
'recipient including the body text.
'Requires that notes client is installed on the system.
Public Sub SendNotesMail_Single(Recipient As String, Subject As String, BodyText As String, Attachment As String, SaveIt As Boolean)
'Set up the objects required for Automation into lotus notes
    Dim Maildb As Object 'The mail database
    Dim UserName As String 'The current users notes name
    Dim MailDbName As String 'THe current users notes mail database name
    Dim MailDoc As Object 'The mail document itself
    Dim AttachME As Object 'The attachment richtextfile object
    Dim Session As Object 'The notes session
    Dim EmbedObj As Object 'The embedded object (Attachment)
    'Start a session to notes
    Set Session = CreateObject("Notes.NotesSession")
    'Next line only works with 5.x and above. Replace password with your password
    ' Session.Initialize ("password")
    'Get the sessions username and then calculate the mail file name
    'You may or may not need this as for MailDBname with some systems you
    'can pass an empty string or using above password you can use other mailboxes.
    UserName = Session.UserName
    MailDbName = Left$(UserName, 1) & Right$(UserName, (Len(UserName) - InStr(1, UserName, " "))) & ".nsf"
    'Open the mail database in notes
    Set Maildb = Session.GETDATABASE("", MailDbName)
     If Maildb.isOpen = True Then
          'Already open for mail
     Else
         Maildb.OPENMAIL
     End If
    'Set up the new mail document
    Set MailDoc = Maildb.CREATEDOCUMENT
    MailDoc.Form = "Memo"
    MailDoc.sendto = Recipient
    MailDoc.Subject = Subject
    MailDoc.Body = BodyText
    MailDoc.SAVEMESSAGEONSEND = SaveIt
    'Set up the embedded object and attachment and attach it
    If Attachment <> "" Then
        Set AttachME = MailDoc.CREATERICHTEXTITEM("Attachment")
        Set EmbedObj = AttachME.EMBEDOBJECT(1454, "", Attachment, "Attachment")
        MailDoc.CREATERICHTEXTITEM ("Attachment")
    End If
    'Send the document
    MailDoc.PostedDate = Now() 'Gets the mail to appear in the sent items folder
    MailDoc.Send 0, Recipient
    'Clean Up
    Set Maildb = Nothing
    Set MailDoc = Nothing
    Set AttachME = Nothing
    Set Session = Nothing
    Set EmbedObj = Nothing
End Sub
  


