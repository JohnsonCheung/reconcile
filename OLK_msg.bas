Attribute VB_Name = "OLK_msg"
Option Compare Database

Property Get KerryFdr() As Outlook.Folder
Dim A As New Outlook.Application
Dim Fdr As Outlook.Folder
Set Fdr = A.Session.Folders("johnson.cheung@godiva.com")
Set Fdr = Fdr.Folders("Filing")
Set Fdr = Fdr.Folders("AutoEmail Received")
Set KerryFdr = Fdr.Folders("Kerry EDI")
End Property
Sub ExpKerryFdrMailAttachment()
ExpFdrMailAttachment KerryFdr
End Sub
Sub ExpFdrMailAttachment(Fdr As Outlook.Folder)
Dim I, M As MailItem
For Each I In Fdr.Items
    If TypeName(I) = "MailItem" Then
        Set M = I
        ExpMailAttachment M
    End If
Next
End Sub
Sub ExpMailAttachment(M As MailItem)
Dim A As Outlook.Attachment
For Each A In M.Attachments
    A.SaveAsFile "C:\temp\" & A.FileName
    'CvEDI F, KillCsv:=True
Next
End Sub
