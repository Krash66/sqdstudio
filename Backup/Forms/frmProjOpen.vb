Public Class frmProjOpen
    Inherits SQDStudio.frmBlank

    Public objThis As New clsProject

    '/// Locals
    Private m_objthis As New clsProject
    Private ODBCsource As New clsODBCinfo
    Private Ulogin As clsLogin
    Private ODBCHelper As New CRODBCHelper()

    Dim LastOpened As Boolean
    Dim mDSNTable As System.Data.DataTable
    Dim dt As New System.Data.DataTable("temp")
    Dim createDSN As Boolean = False
    Dim DoLogin As Boolean = False
    Dim IsCodeEvent As Boolean = False

    '/// clsProject Variables
    Dim DSNname As String = ""
    Dim DSNdesc As String = ""
    Dim DSNtype As enumODBCtype = enumODBCtype.ACCESS
    Dim UID As String = ""
    Dim PWD As String = ""
    Dim Projname As String = ""
    Dim projDesc As String = ""
    Dim projVer As String = ""
    Dim projCdate As String = ""


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
            txtProj.Text = "No Project Selected"
            txtProjDate.Text = "No Project Selected"

            IsCodeEvent = True
            txtSchema.Text = ""
            IsCodeEvent = False

            btnProj.Focus()

        Catch ex As Exception
            LogError(ex, "frmProjOpen btnDSNclick")
        End Try

    End Sub

    Private Sub btnProj_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProj.Click

        Dim frmL As frmODBCLogin
        Dim frmP As frmChangeProj
        Dim ProjectDBdata As clsProjTblData

        Try
            If txtSchema.Enabled = True And txtSchema.Text <> "" Then
                m_objthis.SchemaReq = True
                m_objthis.TablePrefix = GetSchema(txtSchema.Text)
            Else
                m_objthis.SchemaReq = False
                m_objthis.TablePrefix = ""
            End If

            If DoLogin = True Then
                m_objthis.LoginReq = True
                frmL = New frmODBCLogin
                Ulogin = frmL.getLogin()
                Me.Refresh()
                Me.Cursor = Cursors.WaitCursor
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
                Me.Cursor = Cursors.WaitCursor
                m_objthis.ProjectMetaDSNUID = ""
                m_objthis.ProjectMetaDSNPWD = ""
                m_objthis.LoginReq = False
            End If

            '/// initialize new obj with data
            ProjectDBdata = New clsProjTblData
            ProjectDBdata.MetaConnString = m_objthis.MetaConnectionString
            ProjectDBdata.ProjectTableName = m_objthis.tblProjects
            ProjectDBdata.ProjectName = m_objthis.ProjectName
            ProjectDBdata.ProjectDSN = m_objthis.ProjectMetaDSN
            ProjectDBdata.ProjectDesc = m_objthis.ProjectDesc
            ProjectDBdata.ProjectCDate = m_objthis.ProjectCreationDate
            ProjectDBdata.ProjectUDate = m_objthis.ProjectLastUpdated
            ProjectDBdata.ProjectVer = m_objthis.ProjectVersion

            '// open new form to pick a project
            frmP = New frmChangeProj
            ProjectDBdata = frmP.ChangeProject(ProjectDBdata)
            Me.Cursor = Cursors.WaitCursor
            If ProjectDBdata IsNot Nothing Then
                DoLogin = False
                If ProjectDBdata.ProjectName <> m_objthis.ProjectName Then
                    m_objthis.ProjectName = ProjectDBdata.ProjectName
                    m_objthis.ProjectDesc = ProjectDBdata.ProjectDesc
                    m_objthis.ProjectCreationDate = ProjectDBdata.ProjectCDate
                    If ProjectDBdata.ProjectVer = "" Then
                        ProjectDBdata.ProjectVer = Application.ProductVersion
                        'm_objthis.ChangeVersion = True
                    End If
                    m_objthis.ProjectVersion = ProjectDBdata.ProjectVer
                    m_objthis.ProjectMetaVersion = ProjectDBdata.MetaVersion
                    m_objthis.ProductVersion = Application.ProductVersion
                End If
            End If

            If m_objthis.ProjectName <> "" Then
                txtProj.Text = m_objthis.ProjectName
                txtProjDate.Text = m_objthis.ProjectCreationDate
            Else
                txtProj.Text = "No Project Selected"
                txtProjDate.Text = "No Project Selected"
            End If
            Me.Cursor = Cursors.Default
            cmdOk.Focus()

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmProjOpen btnProjClick")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim frmL As frmODBCLogin

        Try
            IsCodeEvent = True
            If txtSchema.Enabled = True Then
                m_objthis.SchemaReq = True
                m_objthis.TablePrefix = GetSchema(txtSchema.Text)
                If m_objthis.TablePrefix = "" Then
                    m_objthis.SchemaReq = False
                End If
            End If
            IsCodeEvent = False

            If DoLogin = True Then
                m_objthis.LoginReq = True
                frmL = New frmODBCLogin
                Ulogin = frmL.getLogin()
                Me.Refresh()
                Me.Cursor = Cursors.WaitCursor
                If Ulogin IsNot Nothing Then
                    m_objthis.ProjectMetaDSNUID = Ulogin.UID
                    m_objthis.ProjectMetaDSNPWD = Ulogin.PWD
                    If Ulogin.UID = "" And Ulogin.PWD = "" Then
                        m_objthis.LoginReq = False
                    End If
                Else
                    If Ulogin Is Nothing Then
                        DialogResult = Windows.Forms.DialogResult.Retry
                        Exit Sub
                    End If
                End If
            End If

            If ValidateDatabase() = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            If m_objthis IsNot Nothing Then
                If m_objthis.ProjectName <> "" Then
                    objThis.ODBCtype = m_objthis.ODBCtype
                    objThis.ProjectCreationDate = m_objthis.ProjectCreationDate
                    objThis.ProjectLastUpdated = m_objthis.ProjectLastUpdated
                    objThis.ProjectDesc = m_objthis.ProjectDesc
                    objThis.ProjectMetaDSN = m_objthis.ProjectMetaDSN
                    objThis.ProjectMetaDSNPWD = m_objthis.ProjectMetaDSNPWD
                    objThis.ProjectMetaDSNUID = m_objthis.ProjectMetaDSNUID
                    objThis.ProjectName = m_objthis.ProjectName
                    objThis.TablePrefix = m_objthis.TablePrefix
                    objThis.MainSeparatorX = m_objthis.MainSeparatorX
                    objThis.LoginReq = m_objthis.LoginReq
                    objThis.SchemaReq = m_objthis.SchemaReq
                    objThis.MapListPath = m_objthis.MapListPath
                    objThis.ProjectVersion = m_objthis.ProjectVersion
                    objThis.ChangeVersion = m_objthis.ChangeVersion
                    objThis.ProjectMetaVersion = m_objthis.ProjectMetaVersion
                    objThis.ProductVersion = m_objthis.ProductVersion
                    If objThis.ProjectVersion = "" Then
                        objThis.ProjectVersion = Application.ProductVersion
                    End If

                    'objThis.LoadMe()


                    Me.Close()
                    DialogResult = Windows.Forms.DialogResult.OK
                Else
                    MsgBox("Please Select a Project", MsgBoxStyle.Information, MsgTitle)
                    DialogResult = Windows.Forms.DialogResult.Retry
                End If
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
            End If
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmProjOpen cmdOKclick")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Function fillLastProjPanel() As Boolean

        Try
            IsCodeEvent = True

            If m_objthis.ProjectName <> "" Then
                txtProj.Text = m_objthis.ProjectName
                txtProjDate.Text = m_objthis.ProjectCreationDate
            Else
                txtProj.Text = "No Project Selected"
                txtProjDate.Text = "No Project Selected"
            End If

            txtDSN.Text = m_objthis.ProjectMetaDSN

            txtSchema.Text = m_objthis.TablePrefix

            Select Case m_objthis.ODBCtype
                Case enumODBCtype.ACCESS
                    txtODBCtype.Text = "MS Access"
                    txtSchema.Enabled = False
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
            LogError(ex, "frmProjOpen fillLastProjPanel")
            Return False
        End Try

    End Function

    Public Function OpenProj() As clsProject

        '// get the registry values for the last project opened
        'LastOpened = m_objthis.RetrieveFromRegistry()
        LastOpened = m_objthis.RetrieveFromXML

        If LastOpened = True Then
            fillLastProjPanel()
        End If

        btnDSN.Focus()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                OpenProj = objThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

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
        ShowHelp(modHelp.HHId.H_Open_a_Project)
    End Sub

    Private Sub frmProjOpen_keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            cmdOk_Click_1(sender, New EventArgs)
        ElseIf e.KeyCode = Keys.F1 Then
            '/// show help
            cmdHelp_Click_1(sender, New EventArgs)
        End If

    End Sub

    Private Sub txtSchema_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSchema.TextChanged

        If IsCodeEvent = True Then Exit Sub

        If m_objthis IsNot Nothing Then
            m_objthis.TablePrefix = txtSchema.Text
            If txtSchema.Text = "" Then
                m_objthis.SchemaReq = False
            Else
                m_objthis.SchemaReq = True
            End If
        End If

    End Sub

    Private Function GetSchema(ByVal input As String) As String

        If input.EndsWith(".") Then
            input = input.Remove(input.LastIndexOf("."))
            input = input.Trim()
        End If
        Return input

    End Function

    Function ValidateDatabase() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim cnn As New System.Data.Odbc.OdbcConnection(m_objthis.MetaConnectionString)

        dt.Clear()
        Try
            Me.Cursor = Cursors.WaitCursor

            cnn.Open()

            Dim sql As String = "select * from " & m_objthis.tblProjects
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            Log(sql)
            If Not da.Fill(dt) > 0 Then
            End If

            If dt.Columns.Count = 4 Then
                m_objthis.ProjectMetaVersion = enumMetaVersion.V2
                Me.Cursor = Cursors.Default
                MsgBox("You are attempting to use a version of Metadata that is no longer supported by Design Studio. Please use a Version 3 Metadata format", MsgBoxStyle.OkOnly, "Incorrect Metadata Version")
                Return False
            Else
                m_objthis.ProjectMetaVersion = enumMetaVersion.V3
            End If
            Me.Cursor = Cursors.Default
            Return True

        Catch OE As Odbc.OdbcException
            Me.Cursor = Cursors.Default
            LogODBCError(OE, "frmProjOpen ValidateDatabase")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            Return False

        Catch ex As Exception
            LogError(ex, "frmProjOpen ValidateDatabase")
            Me.Cursor = Cursors.Default
            MsgBox("A Windows exception error occured: " & Chr(13) & _
                   ex.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            Return False

        Finally
            cnn.Close()
        End Try

    End Function

End Class
