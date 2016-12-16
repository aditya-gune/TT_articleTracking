Attribute VB_Name = "emailModule"
Option Compare Database

Public Function sendEmail( _
    MessageTo As String, _
    cc As String, _
    subject As String, _
    MessageBody As String, _
    attachment As String)

Dim oApp As Outlook.Application
Dim oMail As MailItem
Dim myAttachments As Outlook.Attachments
Set oApp = CreateObject("Outlook.application")

Set oMail = oApp.CreateItem(olMailItem)
Set myAttachments = oMail.Attachments

oMail.Body = MessageBody
oMail.subject = subject
oMail.To = MessageTo
oMail.cc = cc
myAttachments.add (attachment)
'DoCmd.SetWarnings False
oMail.Display
'DoCmd.SetWarnings True
Set oMail = Nothing
Set oApp = Nothing
End Function
