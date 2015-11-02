

Public Class ThisAddIn

    Private WithEvents inspectors As Outlook.Inspectors

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        inspectors = Me.Application.Inspectors
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
