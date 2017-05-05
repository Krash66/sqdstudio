Public Class frmDS_MTDdescriptions

    Dim filename As String
    Dim List As New Collection

    Public Function EditObj(ByVal InDS As clsDatastore) As Collection

        Try
            Dim IsChk As Boolean
            clbList.Items.Clear()
            List.Clear()

            txtFileName.Text = InDS.MTDfile

            For Each dssel As clsDSSelection In InDS.ObjSelections
                IsChk = InDS.DescList.Contains(dssel.SelectionName)
                clbList.Items.Add(dssel.SelectionName, IsChk)
            Next

doAgain:
            Select Case Me.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    EditObj = List
                Case Windows.Forms.DialogResult.Retry
                    GoTo doAgain
                Case Else
                    Return Nothing
            End Select

        Catch ex As Exception
            LogError(ex, "frmDS_MTDdescriptions EditObj")
            Return Nothing
        End Try

    End Function

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Try
            For Each itm As Object In clbList.CheckedItems
                List.Add(itm.ToString, itm.ToString)
            Next

            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex, "frmDS_MTDdescriptions cmdOk_Click_1")
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Try


            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()

        Catch ex As Exception
            LogError(ex, "frmDS_MTDdescriptions cmdCancel_Click_1")
        End Try
        
    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

    End Sub

End Class
