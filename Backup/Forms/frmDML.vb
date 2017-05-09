Public Class frmDML

    Dim FilePath As String = ""
    Dim ReturnString As String = ""
    Dim ImportFileName As String = ""
    Dim ObjEnv As clsEnvironment

    Function NewObj(ByVal Env As clsEnvironment) As String

        ObjEnv = Env

        SetComboConn()

        cbConn.Focus()
        cbConn.Select()
        Windows.Forms.Cursor.Show()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = ReturnString
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                NewObj = Nothing
        End Select

    End Function

    Sub SetComboConn()

        cbConn.Items.Clear()

        For Each conn As clsConnection In ObjEnv.Connections
            cbConn.Items.Add(New Mylist(conn.ConnectionName, conn))
        Next
        cbConn.SelectedIndex = -1

    End Sub

    Private Sub btnImport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImport.Click

        Dim SQLtext As New System.Text.StringBuilder()

        Try
            OpenFileDialog1.ShowDialog(Me)

            If OpenFileDialog1.FileName <> "" Then
                Dim Ofs As System.IO.FileStream = CType(OpenFileDialog1.OpenFile(), System.IO.FileStream)
                Dim objreader As New System.IO.StreamReader(Ofs)

                While (objreader.Peek() > -1)
                    SQLtext.AppendLine(objreader.ReadLine())
                End While

                txtSQL.Text = SQLtext.ToString
                objreader.Close()
                Ofs.Close()
            End If

        Catch ex As Exception
            LogError(ex, "frmDML btnImport_click")
        End Try

    End Sub

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim i As Integer = 0
        Dim SQLArray As Char()

        Try
            If cbConn.SelectedItem Is Nothing Then
                MsgBox("A Database Connection Must be chosen", MsgBoxStyle.Information, MsgTitle)
                cbConn.Focus()
                cbConn.Select()
                Windows.Forms.Cursor.Show()
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Try
            End If

            SaveFileDialog1.ShowDialog(Me)

            If SaveFileDialog1.FileName <> "" Then
                Dim fs As System.IO.FileStream = CType(SaveFileDialog1.OpenFile(), System.IO.FileStream)
                Dim objWriter As New System.IO.StreamWriter(fs)

                SQLArray = txtSQL.Text.ToCharArray

                While i < txtSQL.TextLength
                    objWriter.Write(SQLArray(i))
                    i += 1
                End While

                objWriter.WriteLine()
                objWriter.Close()
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Try
            End If

            FilePath = SaveFileDialog1.FileName
            ReturnString = CType(cbConn.SelectedItem, Mylist).Name & "." & FilePath
            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex, "frmDML cmdOK")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Relational_DML_File)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click_1(sender, e)
            Case Keys.Enter
                cmdOk_Click_1(sender, e)
            Case Keys.F1
                cmdHelp_Click_1(sender, New EventArgs)
        End Select
    End Sub

End Class
