Public Class frmChooseColumns
    Inherits frmBlank

    Dim Objthis As clsDMLinfo
    Dim TableArray As ArrayList

    Function getColumns(ByVal TableInfo As clsDMLinfo) As ArrayList

        Objthis = TableInfo
        ObjThis.TryAgain = False
        gbColumns.Text = "Columns in Table " & Objthis.DSNname & " with Schema " & Objthis.Schema
        If Load_ColumnTable() = False Then
            If MsgBox("Would you like to try again?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                TableInfo.TryAgain = True
                TableInfo.TableName = ""
            Else
                TableInfo.TryAgain = False
                TableInfo.TableName = ""
            End If
            Return Nothing
            Me.Close()
        End If

        DataGridView1.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return TableArray
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select
    End Function

    Private Function Load_ColumnTable() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim cnn As New System.Data.Odbc.OdbcConnection(Objthis.MetaConnectionString)
        Dim sql As String = ""

        dt.Clear()
        Try
            cnn.Open()

            Select Case Objthis.DSNtype
                Case enumODBCtype.DB2
                    sql = "SELECT colname, typename, length, scale FROM syscat.columns WHERE tabschema = '" & Objthis.Schema & "' AND tabname = '" & Objthis.TableName & "'"
                Case enumODBCtype.ORACLE
                    sql = "select table_name,tablespace_name from user_tables order by table_name"
            End Select

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            If Not da.Fill(dt) > 0 Then
                DataGridView1.DataSource = dt
                DataGridView1.Visible = True
                DataGridView1.Enabled = True
                DataGridView1.Show()
                DataGridView1.ClearSelection()

                cmdOk.Enabled = True
                MsgBox("The table you have chosen contains no columns" & (Chr(13)) & "Please cancel and select a different Table")
                DataGridView1.Enabled = False
            Else
                DataGridView1.DataSource = dt
                DataGridView1.Visible = True
                DataGridView1.Enabled = True
                DataGridView1.Show()
                DataGridView1.ClearSelection()

                'TableName = DataGridView1.Item(0, 0).Value.ToString
                'txtTable.Text = TableName

                cmdOk.Enabled = True
            End If

            Return True

        Catch ex As Exception
            If ex.ToString.Contains("Oracle") = True Then
                If ex.ToString.Contains("null password") = True Then
                    MsgBox("Password cannot be left blank")
                ElseIf ex.ToString.Contains("password") = True Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.")
                Else
                    MsgBox("You have chosen an invalid SQMetaData source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.")
                End If
            ElseIf ex.ToString.Contains("IBM") = True Then
                If ex.ToString.Contains("PASSWORD MISSING") = True Then
                    MsgBox("You have entered an incorrect Password.")
                ElseIf ex.ToString.Contains("24") = True Then
                    MsgBox("You have entered an incorrect Username and/or Password.")
                ElseIf ex.ToString.Contains("PROJECTS") = True Then
                    MsgBox("You have entered an incorrect Table Schema.")
                Else
                    MsgBox("You have chosen an invalid SQMetaData source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.")
                End If
            Else
                MsgBox("You have chosen an invalid SQMetaData source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.")
            End If

            Return False
        Finally
            cnn.Close()
        End Try

    End Function

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim rownum As Integer = 0
        Dim column As String

        Try
            ListView1.SmallImageList = ImageList1

            rownum = DataGridView1.CurrentCell.RowIndex
            column = DataGridView1.Item(0, rownum).Value.ToString

            If ListView1.Items.ContainsKey(column) = False Then
                ListView1.Items.Add(column, column, 0)
            End If


        Catch ex As Exception
            LogError(ex, "frmChooseCol DGV1_SelectionChg")
        End Try

    End Sub

    Private Sub DataGridView1_CellContentDblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        DataGridView1_CellContentClick(sender, e)

    End Sub

    Private Sub mnuRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemove.Click

        If ListView1.FocusedItem IsNot Nothing Then
            ListView1.FocusedItem.Remove()
        End If

    End Sub

    Private Sub mnuMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMoveUp.Click

        If ListView1.FocusedItem IsNot Nothing Then

        End If

    End Sub

    Private Sub mnuMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMoveDown.Click

    End Sub
End Class