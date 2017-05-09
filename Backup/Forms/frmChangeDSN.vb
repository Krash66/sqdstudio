Public Class frmChangeDSN

    Dim ODBCsource As New clsODBCinfo

    Private ODBCHelper As New CRODBCHelper()

    Dim DSNname As String
    Dim DSNdesc As String
    Dim DSNtype As enumODBCtype

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

        Try
            If DSNtype <> enumODBCtype.OTHER Then
                ODBCsource.DSNname = DSNname
                ODBCsource.DSNdesc = DSNdesc
                ODBCsource.ODBCtype = DSNtype
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.OK
            Else
                MsgBox("You have chosen a Datasource that is not designed for SQData," & Chr(13) & "Please select a different Datasource.", MsgBoxStyle.Information, MsgTitle)
                DialogResult = Windows.Forms.DialogResult.Retry
            End If

        Catch ex As Exception
            LogError(ex, "frmChangeDSN OK_click")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_ChangeDSN)
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Dim RowIdx As Integer = 0

        cmdOk.Enabled = True
        RowIdx = DataGridView1.CurrentCell.RowIndex
        DSNname = DataGridView1.Item(0, RowIdx).Value.ToString
        DSNdesc = DataGridView1.Item(1, RowIdx).Value.ToString
        DSNtype = GetDsnType(RowIdx)
        txtDSNname.Text = DSNname
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
        txtDSNname.Text = DSNname
        '//for debug
        'MsgBox("DSN type >> " & DSNtype)

        cmdOk_Click_1(sender, New EventArgs)


    End Sub

    Function GetDsnType(ByVal RowIdx As Integer) As enumODBCtype

        Dim TypeString As String = DataGridView1.Item(1, rowidx).Value.ToString

        If TypeString.Contains("access") Or TypeString.Contains("Access") Or TypeString.Contains("ACCESS") Then
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

    Public Function ChgDSN() As clsODBCinfo

        cmdOk.Enabled = False

        If getDSNlist() = True Then
            DSNname = DataGridView1.Item(0, 0).Value.ToString
            DSNdesc = DataGridView1.Item(1, 0).Value.ToString
            DSNtype = GetDsnType(0)
            txtDSNname.Text = DSNname
            
            cmdOk.Enabled = True
        End If

        DataGridView1.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return ODBCsource
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
            cmdHelp_Click_1(sender, New EventArgs)
        End If

    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged

        Dim RowIdx As Integer = 0

        Try
            cmdOk.Enabled = True
            RowIdx = DataGridView1.CurrentCell.RowIndex
            DSNname = DataGridView1.Item(0, RowIdx).Value.ToString
            DSNdesc = DataGridView1.Item(1, RowIdx).Value.ToString
            DSNtype = GetDsnType(RowIdx)
            txtDSNname.Text = DSNname
        Catch ex As Exception
            LogError(ex, "frmChngDSN DGV1_selChg")
        End Try

    End Sub

End Class
