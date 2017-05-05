Public Class ctlControlCenter

    Dim PageURL As Uri

    Overloads Sub OnLoad() Handles Me.Load

        Try
            If CCurl <> "" Then
                PageURL = New Uri(CCurl)
                txtURL.Text = PageURL.AbsoluteUri
            Else
                PageURL = New Uri("http://www.sqdata.com/")
                txtURL.Text = PageURL.AbsoluteUri
            End If
            WebBrowser1.Navigate(PageURL)
            TSSLabel1.Text = WebBrowser1.StatusText

        Catch ex As Exception
            LogError(ex, "ctlControlCenter OnLoad")
            WebBrowser1.GoHome()
        End Try

    End Sub

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

        Try
            txtURL.Text = WebBrowser1.Url.AbsoluteUri
            TSSLabel1.Text = WebBrowser1.StatusText

        Catch ex As Exception
            LogError(ex, "ctlControlCenter WebBrowser1_DocumentCompleted")
        End Try

    End Sub

    Private Sub txtURL_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtURL.TextChanged

        Try
            CCurl = txtURL.Text
            PageURL = New Uri(CCurl)

        Catch ex As Exception
            LogError(ex, "ctlControlCenter txtURL_TextChanged")
        End Try

    End Sub

    Private Sub txtURL_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtURL.KeyDown

        Try
            Select Case e.KeyCode
                Case Keys.BrowserBack
                    btnBack_Click(Me, New EventArgs)
                Case Keys.BrowserForward
                    btnForward_Click(Me, New EventArgs)
                Case Keys.BrowserHome
                    btnHome_Click(Me, New EventArgs)
                Case Keys.BrowserRefresh
                    btnRefresh_Click(Me, New EventArgs)
                Case Keys.Enter
                    btnGo_Click(Me, New EventArgs)
                Case Else
                    '// do nothing
            End Select

        Catch ex As Exception
            LogError(ex, "ctlControlCenter txtURL_TextChanged")
        End Try

    End Sub

    Private Sub WebBrowser1_ProgressChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserProgressChangedEventArgs) Handles WebBrowser1.ProgressChanged

        Try
            ProgressBar1.Maximum = e.MaximumProgress
            If ProgressBar1.Value > 0 Then
                ProgressBar1.Value = e.CurrentProgress
            End If
            TSSLabel1.Text = WebBrowser1.StatusText

        Catch ex As Exception
            LogError(ex, "ctlControlCenter WebBrowser1_ProgressChanged")
        End Try

    End Sub

End Class
