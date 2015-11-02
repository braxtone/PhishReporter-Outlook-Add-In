Imports Microsoft.Office.Tools.Ribbon


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

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Phishing.Click

        Dim exp As Outlook.Explorer = Globals.ThisAddIn.Application.ActiveExplorer()

        If exp.Selection.Count Then
            Dim response = MsgBox("The selected message will be forwarded to " & PhishReporterConfig.SecurityTeamEmailAlias & vbCrLf & " and removed from your inbox.  Would you like to continue?", MsgBoxStyle.YesNo, "Report Phishing To Your Security Team")
            If response = MsgBoxResult.Yes Then
                ' TODO: Be able to handle multiple selected messages rather than just the first one.
                Dim phishEmail As Outlook.MailItem = exp.Selection(1)
                Dim reportEmail As Outlook.MailItem = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem)

                reportEmail.Attachments.Add(phishEmail, Outlook.OlAttachmentType.olEmbeddeditem)
                reportEmail.Subject = PhishReporterConfig.ReportEmailSubject & " - '" & phishEmail.Subject & "'"
                reportEmail.To = PhishReporterConfig.SecurityTeamEmailAlias
                reportEmail.Body = "This is a user-submitted report of a phishing email delivered by the PhishReporter Outlook plugin. Please review the attached phishing email"

                If String.IsNullOrEmpty(PhishReporterConfig.RunbookURL) Then
                    reportEmail.Body = reportEmail.Body & "."
                Else
                    reportEmail.Body = reportEmail.Body & "according to the process defined in " & PhishReporterConfig.RunbookURL
                End If

                reportEmail.Send()
                phishEmail.Delete()
            Else
            End If
        Else
            MsgBox("Please Select a message To Continue.", MsgBoxStyle.OkOnly, "PhishReporter - No E-Mail Message Selected")
        End If

    End Sub

End Class



