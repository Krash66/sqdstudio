Public Class frmSelectDSN

    Dim objDMLinfo As New clsDMLinfo
    Dim Ulogin As clsLogin

    Private ODBCHelper As New CRODBCHelper()

    Dim DSNname As String = ""
    Dim DSNtype As enumODBCtype
    Dim DSNschema As String = ""
    Dim DSNuid As String = ""
    Dim DSNpwd As String = ""
    Dim TableName As String = ""

    Dim IsCodeEvent As Boolean
    Dim DoLogin As Boolean
    Dim DSNDesc As String

    Dim mDSNTable As System.Data.DataTable

    Function getDSNlist() As Boolean
        'GET THE DSN LIST
        Dim ret As Boolean

        ret = ODBCHelper.BuildDSNDataTable()
        'mDSNTable.Clear()
        If ret = True Then
            mDSNTable = ODBCHelper.DSNDataTable
            ' Use the Datatable to show in a grid or loop through it to get the data...
            DataGridView1.DataSource = mDSNTable
            DataGridView1.Visible = True
            DataGridView1.Enabled = True
            DataGridView1.Show()
            Return True
        Else
            'MsgBox("(" + Str(ODBCHelper.ErrorCode) + ")" + ODBCHelper.ErrorMessage)
            Return False
        End If

    End Function

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim frmL As frmODBCLogin

        If DSNtype <> enumODBCtype.DB2 And DSNtype <> enumODBCtype.ORACLE Then
            MsgBox("DML modeling can only be done from Oracle or DB2." & Chr(13) & "Please select a DB2 or Oracle Source DSN.")
            DialogResult = Windows.Forms.DialogResult.Retry
            Exit Sub
        End If

        Try
            IsCodeEvent = True
            If txtSchema.Enabled = True Then
                objDMLinfo.Schema = GetSchema(txtSchema.Text)
            End If
            IsCodeEvent = False

            If DoLogin = True Then
                frmL = New frmODBCLogin
                Ulogin = frmL.getLogin()
                If Ulogin IsNot Nothing Then
                    objDMLinfo.UID = Ulogin.UID
                    objDMLinfo.PWD = Ulogin.PWD
                    If Ulogin.UID = "" And Ulogin.PWD = "" Then
                        DoLogin = False
                    End If
                Else
                    If Ulogin Is Nothing Then
                        DialogResult = Windows.Forms.DialogResult.Retry
                        Exit Sub
                    End If
                End If
            End If

            'If ValidateDatabase() = False Then
            '    DialogResult = Windows.Forms.DialogResult.Retry
            '    Exit Sub
            'End If

            If objDMLinfo IsNot Nothing Then
                If DSNname <> "" Then
                    objDMLinfo.DSNname = DSNname
                    objDMLinfo.DSNtype = DSNtype
                    objDMLinfo.Schema = txtSchema.Text
                    Me.Close()
                    DialogResult = Windows.Forms.DialogResult.OK
                Else
                    MsgBox("Please Select an ODBC Source", MsgBoxStyle.Information)
                    DialogResult = Windows.Forms.DialogResult.Retry
                End If
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
            End If

        Catch ex As Exception
            LogError(ex, "frmProjOpen cmdOKclick")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cbLoginYes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLoginYes.CheckedChanged

        If IsCodeEvent = False Then
            If cbLoginYes.Checked = True Then
                cbLoginNo.Checked = False
                DoLogin = True
            Else
                cbLoginNo.Checked = True
                DoLogin = False
            End If
        End If

    End Sub

    Private Sub cbLoginNo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbLoginNo.CheckedChanged

        If IsCodeEvent = False Then
            If cbLoginNo.Checked = True Then
                cbLoginYes.Checked = False
                DoLogin = False
            Else
                cbLoginYes.Checked = True
                DoLogin = True
            End If
        End If

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Dim RowIdx As Integer = 0

        cmdOk.Enabled = True
        RowIdx = DataGridView1.CurrentCell.RowIndex
        DSNname = DataGridView1.Item(0, RowIdx).Value.ToString
        DSNdesc = DataGridView1.Item(1, RowIdx).Value.ToString
        DSNtype = GetDsnType(RowIdx)
        txtDSN.Text = DSNname
        txtODBCtype.Text = DSNtype.ToString

        '//for debug
        'MsgBox("DSN type >> " & DSNtype)

    End Sub

    Private Sub DataGridView1_CellContentdblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDoubleClick

        Dim RowIdx As Integer = 0

        cmdOk.Enabled = True
        RowIdx = DataGridView1.CurrentCell.RowIndex
        DSNname = DataGridView1.Item(0, RowIdx).Value.ToString
        DSNdesc = DataGridView1.Item(1, RowIdx).Value.ToString
        DSNtype = GetDsnType(RowIdx)
        txtDSN.Text = DSNname
        txtODBCtype.Text = DSNtype.ToString

        '//for debug
        'MsgBox("DSN type >> " & DSNtype)

        cmdOk_Click_1(sender, New EventArgs)


    End Sub

    Function GetDsnType(ByVal RowIdx As Integer) As enumODBCtype

        Dim TypeString As String = DataGridView1.Item(1, rowidx).Value.ToString

        If TypeString.Contains("access") Or TypeString.Contains("Access") Then
            Return enumODBCtype.ACCESS
        ElseIf TypeString.Contains("SQL") Then
            Return enumODBCtype.SQL_SERVER
        ElseIf TypeString.Contains("Oracle") Or TypeString.Contains("oracle") Or TypeString.Contains("ORACLE") Then
            Return enumODBCtype.ORACLE
        ElseIf TypeString.Contains("DB2") Then
            Return enumODBCtype.DB2
        Else
            Return enumODBCtype.OTHER
        End If

    End Function

    Public Function GetInfo() As clsDMLinfo

        cmdOk.Enabled = False

        cbLoginYes.Checked = True
        DoLogin = True

        If getDSNlist() = True Then
            DSNname = DataGridView1.Item(0, 0).Value.ToString
            DSNdesc = DataGridView1.Item(1, 0).Value.ToString
            DSNtype = GetDsnType(0)
            txtDSN.Text = DSNname
            cmdOk.Enabled = True
        End If

        DataGridView1.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return objDMLinfo
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Private Sub frmChangeDSN_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            cmdOk_Click_1(sender, New EventArgs)
        ElseIf e.KeyCode = Keys.F1 Then
            '/// show help
        End If

    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged

        Dim RowIdx As Integer = 0

        Try
            If DataGridView1.CurrentCell IsNot Nothing Then
                cmdOk.Enabled = True
                RowIdx = DataGridView1.CurrentCell.RowIndex
                DSNname = DataGridView1.Item(0, RowIdx).Value.ToString
                DSNDesc = DataGridView1.Item(1, RowIdx).Value.ToString
                DSNtype = GetDsnType(RowIdx)
                txtDSN.Text = DSNname
                txtODBCtype.Text = DSNtype.ToString
            End If
            
        Catch ex As Exception
            LogError(ex, "frmselectDSN DGV1_selChng")
        End Try
        

    End Sub

    Private Sub txtSchema_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSchema.TextChanged

        If IsCodeEvent = True Then Exit Sub

        If objDMLinfo IsNot Nothing Then
            objDMLinfo.Schema = txtSchema.Text
        End If

    End Sub

    Private Function GetSchema(ByVal input As String) As String

        If input.EndsWith(".") Then
            input = input.Remove(input.LastIndexOf("."))
            input = input.Trim()
        End If
        Return input

    End Function

    'Function ValidateDatabase() As Boolean

    '    Dim da As System.Data.Odbc.OdbcDataAdapter
    '    Dim dt As New System.Data.DataTable("temp1")
    '    Dim cnn As New System.Data.Odbc.OdbcConnection(objDMLinfo.MetaConnectionString)

    '    dt.Clear()
    '    Try
    '        cnn.Open()
    '        Dim sql As String = "select * from " & m_objthis.tblProjects
    '        da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
    '        Log(sql)
    '        If Not da.Fill(dt) > 0 Then
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        MsgBox("You have entered an incorrect Table Schema or an incorrect Username and Password.")
    '        Return False
    '    Finally
    '        cnn.Close()
    '    End Try

    'End Function

End Class
