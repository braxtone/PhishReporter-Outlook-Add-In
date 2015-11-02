Public Class PhishReporterConfig
    ' Report Configuration
    Public Shared Property SecurityTeamEmailAlias As String = "securityteam@example.com"
    Public Shared Property ReportEmailSubject As String = "[PhishReporter] Phishing Email Report"
    ' Link to the security team's runbook for handling phishing emails. If the variable is empty or not defined, defaults to a simplified message"
    Public Shared Property RunbookURL As String = "https://corporatewiki/path/to/runbook"

    ' Ribbon Group Config
    Public Shared Property RibbonGroupName As String = "Report Security Issues"

    ' Button Config
    Public Shared Property ButtonName As String = "Report Phishing"
    Public Shared Property ButtonHoverDescription As String = "Report a suspicious email to the $COMPANY Information Security Team."
    Public Shared Property ButtonScreenTip As String = "Report phishing emails"
    Public Shared Property ButtonSuperTip As String = "Use this button to report suspicious emails to the $COMPANY Information Security team."

End Class

