Public Class ctlControlCenter

    Dim PageURL As Uri

    Private Sub btnHome_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHome.Click

        Try
            WebBrowser1.GoHome()

        Catch ex As Exception
            LogError(ex, "ctlControlCenter btnHome_Click")
        End Try

    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click

        Try
            WebBrowser1.GoBack()

        Catch ex As Exception
            LogError(ex, "ctlControlCenter btnBack_Click")
        End Try

    End Sub

    Private Sub btnForward_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnForward.Click

        Try
            WebBrowser1.GoForward()

        Catch ex As Exception
            LogError(ex, "ctlControlCenter btnForward_Click")
        End Try

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        Try
            WebBrowser1.Refresh(WebBrowserRefreshOption.Completely)

        Catch ex As Exception
            LogError(ex, "ctlControlCenter btnRefresh_Click")
        End Try

    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click

        Try
            WebBrowser1.Navigate(PageURL)

        Catch ex As Exception
            LogError(ex, "ctlControlCenter btnGo_Click")
        End Try

    End Sub

    Private Sub WebBrowser1_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted

    End Sub

    Private Sub txtURL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtURL.Click

    End Sub
End Class
