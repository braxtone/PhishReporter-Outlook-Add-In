Imports Microsoft.Office.Tools.Ribbon
'Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook



Public Class HOME

    Dim objItem As Outlook.MailItem
    Dim objMsg As Outlook.MailItem
    Dim app As Outlook.Application
    Dim exp As Outlook.Explorer
    Dim sel As Outlook.Selection
    Dim Application As Outlook.Application
    Dim attachments As Outlook.Attachments
    Dim objOutlookAtt As Outlook.Attachment



    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles PHISHING.Click

        
        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()
        If exp.Selection.Count Then
            Dim response = MsgBox("The selected message will be forwarded to phishing@company.com" & vbCrLf & " and removed from your inbox.  Would you like to continue?", MsgBoxStyle.YesNo, "PhishReporter Report Phishing")
            If response = MsgBoxResult.Yes Then
                Dim selectedMail As Outlook.MailItem = exp.Selection(1)
                Dim newMail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)
                newMail.Attachments.Add(selectedMail, Outlook.OlAttachmentType.olEmbeddeditem)
                newMail.Subject = "[SPAM/PHISHING]"
                newMail.To = "phishing@company.com"
                newMail.Send()
                selectedMail.Delete()
            Else
            End If
        Else
            MsgBox("Please select a message to continue.", MsgBoxStyle.OkOnly, "PhishReporter - No E-Mail Message Selected")
        End If

    End Sub

End Class




