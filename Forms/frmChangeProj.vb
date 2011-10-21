Public Class frmChangeProj
    Inherits frmBlank

    Dim ObjThis As clsProjTblData

    Private ProjName As String = ""
    Private Projdesc As String = ""
    Private ProjVer As String = ""
    Private ProjCdate As String = ""
    Private SecATTR As String = ""
    Private MetaVer As enumMetaVersion

    Public Event ShowErrorLog(ByVal sender As System.Object, ByVal e As Object)

    Private Sub DataGridView2_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseClick

        Dim rowNum As Integer = 0

        rowNum = DataGridView2.CurrentCell.RowIndex
        If DataGridView2.ColumnCount = 4 Then
            ProjName = DataGridView2.Item(0, rowNum).Value.ToString
            Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
            ProjVer = DataGridView2.Item(2, rowNum).Value.ToString
            ProjCdate = DataGridView2.Item(3, rowNum).Value.ToString
            MetaVer = enumMetaVersion.V2
        Else
            ProjName = DataGridView2.Item(0, rowNum).Value.ToString
            Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
            SecATTR = DataGridView2.Item(2, rowNum).Value.ToString
            MetaVer = enumMetaVersion.V3
        End If
       
        txtProj.Text = ProjName

    End Sub

    Private Sub DataGridView2_CellContentdblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseDoubleClick

        Dim rowNum As Integer = 0

        rowNum = DataGridView2.CurrentCell.RowIndex
        If DataGridView2.ColumnCount = 4 Then
            ProjName = DataGridView2.Item(0, rowNum).Value.ToString
            Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
            ProjVer = DataGridView2.Item(2, rowNum).Value.ToString
            ProjCdate = DataGridView2.Item(3, rowNum).Value.ToString
            MetaVer = enumMetaVersion.V2
        Else
            ProjName = DataGridView2.Item(0, rowNum).Value.ToString
            Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
            SecATTR = DataGridView2.Item(2, rowNum).Value.ToString
            MetaVer = enumMetaVersion.V3
        End If
        txtProj.Text = ProjName

        cmdOk_Click_1(sender, New EventArgs)

    End Sub


    Public Function ChangeProject(ByRef TempObj As clsProjTblData) As clsProjTblData

        ObjThis = TempObj
        txtDSN.Text = ObjThis.ProjectDSN
        If Load_Table() = False Then
            Return Nothing
            Me.Close()
        End If
        DataGridView2.ClearSelection()
        DataGridView2.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return ObjThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Private Function Load_Table() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")

        Dim cnn As New System.Data.Odbc.OdbcConnection(ObjThis.MetaConnString)

        'ToolStripProgressBar1.Enabled = True
        'ToolStripProgressBar1.MarqueeAnimationSpeed = 100
        
        dt.Clear()
        Try
            Me.Cursor = Cursors.WaitCursor

            cnn.Open()

            Dim sql As String = "select * from " & ObjThis.ProjectTableName

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            If Not da.Fill(dt) > 0 Then
                DataGridView2.DataSource = dt
                DataGridView2.Visible = True
                DataGridView2.Enabled = True
                DataGridView2.Show()
                DataGridView2.ClearSelection()

                cmdOk.Enabled = True
                Me.Cursor = Cursors.Default
                MsgBox("The source you have chosen contains no projects" & (Chr(13)) & "Please cancel and select a different ODBC Source, or" & (Chr(13)) & "Go to main program to create a new Project", MsgBoxStyle.Information, MsgTitle)
                DataGridView2.Enabled = False
            Else
                DataGridView2.DataSource = dt
                DataGridView2.Visible = True
                DataGridView2.Enabled = True
                DataGridView2.Show()
                DataGridView2.ClearSelection()

                If DataGridView2.ColumnCount = 4 Then
                    ProjName = DataGridView2.Item(0, 0).Value.ToString
                    Projdesc = DataGridView2.Item(1, 0).Value.ToString
                    ProjVer = DataGridView2.Item(2, 0).Value.ToString
                    ProjCdate = DataGridView2.Item(3, 0).Value.ToString
                Else
                    ProjName = DataGridView2.Item(0, 0).Value.ToString
                    Projdesc = DataGridView2.Item(1, 0).Value.ToString
                    SecATTR = DataGridView2.Item(2, 0).Value.ToString
                    ProjCdate = DataGridView2.Item(3, 0).Value.ToString
                End If

                txtProj.Text = ProjName
                Me.Cursor = Cursors.Default
                cmdOk.Enabled = True
            End If

            Return True
            'ToolStripProgressBar1.Enabled = False

        Catch OE As Odbc.OdbcException
            Me.Cursor = Cursors.Default
            LogODBCError(OE, "frmChangeProj Load_Table")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            Return False

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmChangeProj Load_Table")
            If ex.ToString.Contains("Oracle") = True Then
                If ex.ToString.Contains("null password") = True Then
                    MsgBox("Password cannot be left blank", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("password") = True Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.", MsgBoxStyle.Information, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid SQMetaData source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.Information, MsgTitle)
                End If
            ElseIf ex.ToString.Contains("IBM") = True Then
                If ex.ToString.Contains("PASSWORD MISSING") = True Then
                    MsgBox("You have entered an incorrect Password.", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("24") = True Then
                    MsgBox("You have entered an incorrect Username and/or Password.", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("PROJECTS") = True Then
                    MsgBox("You have entered an incorrect Table Schema.", MsgBoxStyle.Information, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid MetaData source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.Information, MsgTitle)
                End If
            Else
                MsgBox("A Windows exception error occured: " & Chr(13) & _
                    ex.Message.ToString & Chr(13) & Chr(13) & _
                    "For more information, see the Error Log" & Chr(13) & _
                    "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            End If
            Return False
        Finally
            cnn.Close()
        End Try

    End Function

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        If ProjName <> "" Then
            ObjThis.ProjectName = ProjName
            ObjThis.ProjectDesc = Projdesc
            ObjThis.ProjectVer = ProjVer
            ObjThis.ProjectCDate = ProjCdate
            ObjThis.MetaVersion = MetaVer

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
        ShowHelp(modHelp.HHId.H_ChangeProject)
    End Sub

    Private Sub frmChangeProj_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            cmdOk_Click_1(sender, New EventArgs)
        ElseIf e.KeyCode = Keys.F1 Then
            '/// show help
            cmdHelp_Click_1(sender, New EventArgs)
        End If

    End Sub

    Private Sub DataGridView2_keydown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DataGridView2.SelectionChanged

        Dim rowNum As Integer = 0

        Try
            rowNum = DataGridView2.CurrentCell.RowIndex
            If DataGridView2.ColumnCount = 4 Then
                ProjName = DataGridView2.Item(0, rowNum).Value.ToString
                Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
                ProjVer = DataGridView2.Item(2, rowNum).Value.ToString
                ProjCdate = DataGridView2.Item(3, 0).Value.ToString
                MetaVer = enumMetaVersion.V2
            Else
                ProjName = DataGridView2.Item(0, rowNum).Value.ToString
                Projdesc = DataGridView2.Item(1, rowNum).Value.ToString
                SecATTR = DataGridView2.Item(2, rowNum).Value.ToString
                MetaVer = enumMetaVersion.V3
            End If
            
            txtProj.Text = ProjName

        Catch ex As Exception
            LogError(ex, "frmCngProj DGV1_selChanged")
        End Try

    End Sub

End Class