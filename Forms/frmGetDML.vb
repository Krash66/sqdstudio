Public Class frmGetDML

    Dim ObjThis As clsDMLinfo
    Dim TableName As String
    Dim DSNname As String
    Dim SchemaName As String

    Function GetTable(ByRef ObjDMLinfo As clsDMLinfo) As clsDMLinfo

        ObjThis = ObjDMLinfo
        ObjThis.TryAgain = False
        txtDSN.Text = ObjThis.DSNname

        If Load_Table() = False Then
            If MsgBox("Would you like to try again?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ObjDMLinfo.TryAgain = True
                ObjDMLinfo.TableName = ""
            Else
                ObjDMLinfo.TryAgain = False
                ObjDMLinfo.TableName = ""
            End If
            Return ObjDMLinfo
            Me.Close()
        End If

        DataGridView1.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                ObjDMLinfo = ObjThis
                Return ObjDMLinfo
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Private Function Load_Table() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim cnn As New System.Data.Odbc.OdbcConnection(ObjThis.MetaConnectionString)
        Dim sql As String = ""

        dt.Clear()
        Try
            cnn.Open()

            Select Case ObjThis.DSNtype
                Case enumODBCtype.DB2
                    sql = "select tabname,tabschema from syscat.tables order by tabname"
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
                MsgBox("The source you have chosen contains no SQDStudio projects" & (Chr(13)) & "Please cancel and select a different ODBC Source, or" & (Chr(13)) & "Go to main program to create a new Project")
                DataGridView1.Enabled = False
            Else
                DataGridView1.DataSource = dt
                DataGridView1.Visible = True
                DataGridView1.Enabled = True
                DataGridView1.Show()
                DataGridView1.ClearSelection()

                TableName = DataGridView1.Item(0, 0).Value.ToString
                txtTable.Text = TableName

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

        If TableName <> "" Then
            ObjThis.TableName = TableName
            ObjThis.Schema = SchemaName
            
            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK
        Else
            DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Dim rowNum As Integer = 0

        rowNum = DataGridView1.CurrentCell.RowIndex
        TableName = DataGridView1.Item(0, rowNum).Value.ToString
        SchemaName = DataGridView1.Item(1, rowNum).Value.ToString
        txtTable.Text = TableName

    End Sub

    Private Sub DataGridView1_CellContentDblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick

        Dim rowNum As Integer = 0

        rowNum = DataGridView1.CurrentCell.RowIndex
        TableName = DataGridView1.Item(0, rowNum).Value.ToString
        SchemaName = DataGridView1.Item(1, rowNum).Value.ToString
        txtTable.Text = TableName

        cmdOk_Click_1(sender, New EventArgs)

    End Sub

    Private Sub frmChangeProj_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            cmdOk_Click_1(sender, New EventArgs)
        ElseIf e.KeyCode = Keys.F1 Then
            '/// show help
        End If

    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged

        Dim rowNum As Integer = 0

        Try
            If DataGridView1.CurrentCell IsNot Nothing Then
                rowNum = DataGridView1.CurrentCell.RowIndex
                TableName = DataGridView1.Item(0, rowNum).Value.ToString
                SchemaName = DataGridView1.Item(1, rowNum).Value.ToString
                txtTable.Text = TableName
            End If
            
        Catch ex As Exception
            LogError(ex, "frmGetdml DataGV1_selChanged")
        End Try

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

    End Sub
End Class
