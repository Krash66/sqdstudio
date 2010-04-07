Public Class frmNewProj
    Inherits frmBlank

    Public objThis As New clsProject

    '/// Locals
    Private m_objthis As New clsProject
    Private ODBCsource As New clsODBCinfo
    Private Ulogin As clsLogin
    
    Dim LastOpened As Boolean
    Dim DoLogin As Boolean = False
    Dim IsCodeEvent As Boolean = False
    
    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim frmL As frmODBCLogin

        Try
            If txtSchema.Enabled = True Then
                m_objthis.SchemaReq = True
                m_objthis.TablePrefix = txtSchema.Text
                If txtSchema.Text = "" Then
                    m_objthis.SchemaReq = False
                End If
            End If

            If DoLogin = True Then
                m_objthis.LoginReq = True
                frmL = New frmODBCLogin
                Ulogin = frmL.getLogin()
                If Ulogin IsNot Nothing Then
                    If Ulogin.LoggedInChange = True Then
                        m_objthis.ProjectMetaDSNUID = Ulogin.UID
                        m_objthis.ProjectMetaDSNPWD = Ulogin.PWD
                    End If
                    If Ulogin.UID = "" And Ulogin.PWD = "" Then
                        m_objthis.LoginReq = False
                    End If
                Else
                    If Ulogin Is Nothing Then
                        Exit Sub
                    End If
                End If
            Else
                m_objthis.ProjectMetaDSNUID = ""
                m_objthis.ProjectMetaDSNPWD = ""
                m_objthis.LoginReq = False
            End If

            If ValidateDatabase() = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            If m_objthis IsNot Nothing Then
                If ValidateNewName(txtProj.Text) = True Then
                    objThis.ProjectName = txtProj.Text
                    objThis.ODBCtype = m_objthis.ODBCtype
                    objThis.ProjectCreationDate = Now()
                    objThis.ProjectDesc = m_objthis.ProjectDesc
                    objThis.ProjectMetaDSN = m_objthis.ProjectMetaDSN
                    objThis.ProjectMetaDSNPWD = m_objthis.ProjectMetaDSNPWD
                    objThis.ProjectMetaDSNUID = m_objthis.ProjectMetaDSNUID
                    objThis.ProjectMetaVersion = m_objthis.ProjectMetaVersion
                    objThis.TablePrefix = m_objthis.TablePrefix
                    objThis.MainSeparatorX = m_objthis.MainSeparatorX
                    objThis.LoginReq = objThis.LoginReq
                    objThis.SchemaReq = objThis.SchemaReq
                    objThis.ProjectVersion = Application.ProductVersion


                    cnn = New System.Data.Odbc.OdbcConnection(objThis.MetaConnectionString)
                    cnn.Open()

                    If objThis.ValidateNewObject() = True Then
                        If objThis.AddNew() = False Then
                            DialogResult = Windows.Forms.DialogResult.Retry
                            cnn.Close()
                            Exit Sub
                        End If
                    Else
                        DialogResult = Windows.Forms.DialogResult.Retry
                        cnn.Close()
                        Exit Sub
                    End If
                Else
                    DialogResult = Windows.Forms.DialogResult.Retry
                    cnn.Close()
                    Exit Sub
                End If
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
                cnn.Close()
                Exit Sub
            End If

            cnn.Close()
            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex, "frmNewProj cmdOKclick")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_Creating_a_Project)
    End Sub

    Private Sub btnDSN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDSN.Click

        Dim frm As frmChangeDSN

        Try
            frm = New frmChangeDSN
            ODBCsource = frm.ChgDSN()

            txtDSN.Text = ODBCsource.DSNname

            Select Case ODBCsource.ODBCtype
                Case enumODBCtype.ACCESS
                    txtODBCtype.Text = "MS Access"
                    cbLoginYes.Checked = False
                    txtSchema.Enabled = False
                Case enumODBCtype.DB2
                    txtODBCtype.Text = "IBM DB2"
                    cbLoginYes.Checked = True
                    txtSchema.Enabled = True
                Case enumODBCtype.ORACLE
                    txtODBCtype.Text = "Oracle"
                    cbLoginYes.Checked = True
                    txtSchema.Enabled = True
                Case enumODBCtype.SQL_SERVER
                    txtODBCtype.Text = "MS SQL Server"
                    cbLoginYes.Checked = True
                    txtSchema.Enabled = False
                Case enumODBCtype.OTHER
                    txtODBCtype.Text = "Currently Unsupported Type"
                    cbLoginYes.Checked = False
                    txtSchema.Enabled = False
                Case Else
                    txtODBCtype.Text = "Currently Unsupported Type"
                    cbLoginYes.Checked = False
                    txtSchema.Enabled = False
            End Select

            m_objthis.ProjectMetaDSN = ODBCsource.DSNname
            m_objthis.ODBCtype = CInt(ODBCsource.ODBCtype)
            m_objthis.ProjectName = ""
            txtProj.Text = ""
            txtSchema.Text = ""

            cmdOk.Focus()

        Catch ex As Exception
            LogError(ex, "frmProjOpen btnDSNclick")
        End Try

    End Sub

    Function fillLastProjPanel() As Boolean

        Try
            IsCodeEvent = True

            m_objthis.ProjectName = ""

            txtDSN.Text = m_objthis.ProjectMetaDSN

            txtSchema.Text = m_objthis.TablePrefix

            Select Case m_objthis.ODBCtype
                Case enumODBCtype.ACCESS
                    txtODBCtype.Text = "MS Access"
                    txtSchema.Enabled = False
                    txtSchema.Text = ""
                Case enumODBCtype.DB2
                    txtODBCtype.Text = "IBM DB2"
                    txtSchema.Enabled = True
                Case enumODBCtype.ORACLE
                    txtODBCtype.Text = "Oracle"
                    txtSchema.Enabled = True
                Case enumODBCtype.SQL_SERVER
                    txtODBCtype.Text = "MS SQL Server"
                    txtSchema.Enabled = False
                Case enumODBCtype.OTHER
                    txtODBCtype.Text = "Currently Unsupported Type"
                    txtSchema.Enabled = False
                Case Else
                    txtODBCtype.Text = "Currently Unsupported Type"
                    txtSchema.Enabled = False
            End Select

            If m_objthis.LoginReq = True Then
                DoLogin = True
                cbLoginYes.Checked = True
                cbLoginNo.Checked = False
            Else
                DoLogin = False
                cbLoginYes.Checked = False
                cbLoginNo.Checked = True
            End If

            IsCodeEvent = False
            cmdOk.Enabled = True

            Return True

        Catch ex As Exception
            LogError(ex, "frmNewProj fillLastProjPanel")
            Return False
        End Try

    End Function

    Function ValidateDatabase() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim cnn As New System.Data.Odbc.OdbcConnection(m_objthis.MetaConnectionString)

        dt.Clear()
        Try
            cnn.Open()
            Dim sql As String = "select * from " & m_objthis.tblProjects
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            Log(sql)
            If Not da.Fill(dt) > 0 Then
            End If
            If dt.Columns.Count = 4 Then
                m_objthis.ProjectMetaVersion = enumMetaVersion.V2
            Else
                m_objthis.ProjectMetaVersion = enumMetaVersion.V3
            End If

            Return True

        Catch ex As Exception
            MsgBox("You have chosen an invalid SQMetaData source," & Chr(13) & _
            "entered an incorrect Table Schema or an incorrect Username and Password.", MsgBoxStyle.Information, MsgTitle)
            Return False
        Finally
            cnn.Close()
        End Try

    End Function

    Function NewProj() As clsProject

        '// get the registry values for the last project opened
        LastOpened = m_objthis.RetrieveFromRegistry()
        If LastOpened = True Then
            fillLastProjPanel()
        End If

        txtProj.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                NewProj = objThis
            Case Windows.Forms.DialogResult.Retry
                txtProj.Text = ""
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Private Sub frmNewProj_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            cmdOk_Click_1(sender, New EventArgs)
        ElseIf e.KeyCode = Keys.F1 Then
            '/// show help
            cmdHelp_Click_1(sender, New EventArgs)
        End If

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

End Class