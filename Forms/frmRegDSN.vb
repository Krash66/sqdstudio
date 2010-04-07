Public Class frmRegDSN
    Inherits SQStudio.frmBlank

    Private ODBCHelper As New CRODBCHelper()

    Dim IsModified As Boolean = False
    Dim DSNname As String = ""
    Dim UID As String = ""
    Dim PWD As String = ""
    Dim Desc As String = ""
    Dim FullDBname As String = ""
    Dim DBname As String = ""
    Dim DBpath As String = ""
    Dim SQLSvrName As String = ""
    Dim ChosenDSNname As String = ""
    Dim ChosenDriver As String = ""
    Dim RmvDSNname As String = ""
    Dim RmvDSNdvr As String = ""
    Dim mDSNTable As System.Data.DataTable
    Dim DSNcreated As Boolean = False
    Dim DSNremoved As Boolean = False
    Dim strDriver As String = ""
    Dim CreateDB As Boolean = False
    Dim dataVal As String = ""
    Dim IsSysDSN As Boolean = False
    Dim UseWinAuth As Boolean = False

    Private Const ODBC_ADD_DSN As Integer = 1               ' Add data source
    Private Const ODBC_CONFIG_DSN As Integer = 2            ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN As Integer = 3            ' Remove data source
    Private Const ODBC_ADD_SYS_DSN As Integer = 4           ' Add a SYSTEM DSN
    Private Const ODBC_CONFIG_SYS_DSN As Integer = 5
    Private Const ODBC_REMOVE_SYS_DSN As Integer = 6
    Private Const vbAPINull As Long = 0                     ' NULL Pointer

    Public Enum verifyRet
        VerNoProj = 1
        VerHasProj = 2
        VerNotComp = 3
        VerNotSQD = 4
    End Enum

    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
         (ByVal hwndParent As Integer, ByVal fRequest As Integer, _
         ByVal lpszDriver As String, ByVal lpszAttributes As String) As Integer

#Region "Form Events"

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDSN.TextChanged, txtDBN.TextChanged, txtDBpath.TextChanged, txtDesc.TextChanged, txtPWD.TextChanged, txtChosenDSN.TextChanged, txtUID.TextChanged

        If IsEventFromCode = True Then Exit Sub
        IsModified = True
        IsEventFromCode = False

    End Sub

    Function OpenRegDSN(Optional ByVal NewPWD As String = "", Optional ByVal NewUID As String = "") As Boolean

        StartLoad()
        setcombo()
        txtPWD.Text = NewPWD
        PWD = NewPWD
        txtUID.Text = NewUID
        UID = NewUID
        EndLoad()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return True
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return False
        End Select

    End Function

    Private Sub setcombo()
        cmbDBtype.Items.Add(New Mylist("Microsoft Access Driver (*.mdb)", "Access"))
        cmbDBtype.Items.Add(New Mylist("SQL Server", "SQLServer"))
        cmbDBtype.Items.Add(New Mylist("IBM DB2 ODBC DRIVER", "DB2"))
        cmbDBtype.SelectedItem = cmbDBtype.Items.Item(0)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        cmdOk.Enabled = True
        txtChosenDSN.Enabled = False
        btnRemoveDSN.Enabled = False
        btnAddDSN.Enabled = False
        btnVerify.Enabled = False
        cmdBrowse.Enabled = False
        gbDesc.Enabled = False
        txtDBpath.Enabled = False
        txtDBN.Enabled = False
        txtSQLSvrName.Enabled = False
        cbxWinAuth.Checked = True
        cbxSysDSN.Checked = False
        getDSNlist()
        txtDBN.Text = ""
        'gbCreateTables.Enabled = False

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

#End Region

#Region "Add DSN Events"

    Private Sub cmbDBtype_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDBtype.SelectedIndexChanged

        txtDBpath.Text = ""
        txtDBN.Text = ""
        dataVal = CType(cmbDBtype.Items(cmbDBtype.SelectedIndex), Mylist).ItemData.ToString

        Select Case dataVal
            Case "Access"
                lbltgtLocation.Text = "Browse for name and location of .mdb file"
                txtSQLSvrName.Enabled = False
                lblSQLServer.Enabled = False
                txtDBN.Text = ""
                cmdBrowse.Enabled = True
                cmdBrowse.Text = "..." '"MetaData Target Path"
                txtDBpath.Enabled = True
                btnAddDSN.Enabled = False
                gbUIDpwd.Enabled = True
                cbxWinAuth.Enabled = False
                'gbCreateTables.Enabled = False


            Case "SQLServer"
                lbltgtLocation.Text = "Select Database (.mdf) to store MetaData"
                lblSQLServer.Enabled = True
                txtSQLSvrName.Enabled = True
                'txtSQLSvrName.Text = "(Local)"
                cmdBrowse.Text = "..." '"Select SQLServer File"
                cmdBrowse.Enabled = True
                txtDBpath.Enabled = True
                'txtDBN.Text = "(Default)"
                btnAddDSN.Enabled = True
                gbUIDpwd.Enabled = True
                cbxWinAuth.Enabled = True
                'gbCreateTables.Enabled = True

            Case "DB2"

        End Select
        gbDesc.Enabled = True
        txtDBN.Enabled = True

    End Sub

    Private Sub cbxSysDSN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxSysDSN.CheckedChanged
        If cbxSysDSN.Checked = True Then
            IsSysDSN = True
        Else
            IsSysDSN = False
        End If
    End Sub

    Private Sub cbxWinAuth_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbxWinAuth.CheckedChanged
        If cbxWinAuth.Checked = True Then
            UseWinAuth = True
            lblUID.Enabled = False
            txtUID.Enabled = False
            lblPWD.Enabled = False
            txtPWD.Enabled = False
        Else
            UseWinAuth = False
            lblUID.Enabled = True
            txtUID.Enabled = True
            lblPWD.Enabled = True
            txtPWD.Enabled = True
        End If
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click

        Dim strlength As Integer = 0
        Dim dlgBrowse As OpenFileDialog = Me.OpenFileDialog1

        Select Case dataVal
            Case "Access"
                dlgBrowse.InitialDirectory = GetAppPath()
                dlgBrowse.Title = "Choose Filename and Path to copy SQDMeta.mdb file"
                dlgBrowse.DefaultExt = "*.mdb"
                dlgBrowse.FileName = "SQDMetaData"
                dlgBrowse.Filter = "MS Access Files|*.mdb|All Files|*.*"
                dlgBrowse.AddExtension = True

                If dlgBrowse.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    FullDBname = dlgBrowse.FileName
                    DBname = GetFileNameOnly(FullDBname)
                    strlength = FullDBname.Length - DBname.Length
                    DBpath = Strings.Left(FullDBname, strlength)
                    txtDBN.Text = DBname
                    txtDBpath.Text = DBpath
                    btnAddDSN.Enabled = True
                End If

            Case "SQLServer"
                If System.IO.Directory.Exists("c:\Program Files\Microsoft SQL Server\") Then
                    dlgBrowse.InitialDirectory = "c:\Program Files\Microsoft SQL Server\"
                Else
                    dlgBrowse.InitialDirectory = GetAppPath()
                End If
                dlgBrowse.Title = "Choose SQDMetaData SQLserver DataBase"
                dlgBrowse.DefaultExt = "*.mdf"
                dlgBrowse.FileName = ""
                dlgBrowse.Filter = "MS SQLServer DB Files|*.mdf|All Files|*.*"
                dlgBrowse.AddExtension = True


                If dlgBrowse.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    FullDBname = dlgBrowse.FileName
                    DBname = GetFileNameOnly(FullDBname)
                    strlength = FullDBname.Length - DBname.Length
                    DBpath = Strings.Left(FullDBname, strlength)
                    txtDBN.Text = DBname
                    txtDBpath.Text = DBpath
                    btnAddDSN.Enabled = True
                    btnAddDSN.Enabled = True
                End If

            Case "DB2"
                '//// TODO////
            Case Else
                '////TODO/////
        End Select

    End Sub

    Private Sub btnAddDSN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddDSN.Click

        Dim intRet As Integer
        Dim strAttributes As String = ""
        Dim verRet As verifyRet

        Try
            If ValidateNewName(txtDSN.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            Else
                '/// If name is Valid save form data to Variables
                DBname = txtDBN.Text
                DBpath = txtDBpath.Text
                DSNname = txtDSN.Text
                Desc = txtDesc.Text
                UID = txtUID.Text
                PWD = txtPWD.Text
                SQLSvrName = txtSQLSvrName.Text

                '//// add DSN Based on Database Type Selected
                Select Case dataVal
                    '///*************************************
                    '///*** ACCESS **************************
                    '///*************************************
                    Case "Access"
                        If System.IO.File.Exists(FullDBname) = True Then
                            MsgBox(FullDBname & " already exists. Browse again to enter a different value for Database Name and/or Target Path")
                            txtDBN.Focus()
                            Exit Sub
                        End If

                        '/// User entries OK now copy blank metadata to new location
                        System.IO.File.Copy(System.Windows.Forms.Application.StartupPath() & "\SQDMeta.mdb", FullDBname)

                        '//For microsoft Access set DSN parameters
                        strDriver = CType(cmbDBtype.Items(cmbDBtype.SelectedIndex), Mylist).Name() & Chr(0)
                        strAttributes = "DSN=" & DSNname & Chr(0) & "Uid=" & UID & Chr(0) & "pwd=" & PWD & Chr(0) & "DBQ=" & FullDBname & Chr(0) & "Description=" & Desc & Chr(0)
                        '*******************************
                        '**** Now add the DSN Entry ****
                        '*******************************
                        If IsSysDSN = True Then
                            intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, strDriver, strAttributes)
                        Else
                            intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
                        End If

                        If intRet Then
                            MsgBox(DSNname & " registered successfully")
                            DSNcreated = True
                        Else
                            MsgBox("Registration of " & DSNname & " Failed with error " & intRet, MsgBoxStyle.Critical)
                            Log(DSNname & "creation failed with return code " & intRet)
                            Exit Select
                        End If

                        '///*************************************
                        '///*** SQL Server **********************
                        '///*************************************
                    Case "SQLServer"
                        If DSNname = "" Or SQLSvrName = "" Or DBname = "" Then
                            MsgBox("DSN Name, Server Name, and Database Name are all reguired fields", MsgBoxStyle.Information)
                            Exit Sub
                        End If
                        If DBname.EndsWith(".mdf") Then
                            DBname = Strings.Left(DBname, DBname.Length - 4)
                        End If

                        Dim attributes As New System.Text.StringBuilder()

                        attributes.Append("DSN=")
                        attributes.Append(DSNname)
                        attributes.Append(Chr(0))
                        attributes.Append("Server=")
                        attributes.Append(SQLSvrName)
                        attributes.Append(Chr(0))
                        attributes.Append("Database=")
                        attributes.Append(DBname)
                        attributes.Append(Chr(0))
                        attributes.Append("AnsiNPW=No")
                        attributes.Append(Chr(0))
                        If UseWinAuth = True Then
                            attributes.Append("trusted_connection=yes")
                            attributes.Append(Chr(0))
                            attributes.Append(Chr(0))
                            attributes.Append(Chr(0))
                        Else
                            attributes.Append("UID=")
                            attributes.Append(UID)
                            attributes.Append(Chr(0))
                            attributes.Append("Pwd=")
                            attributes.Append(PWD)
                            attributes.Append(Chr(0))
                        End If
                        strAttributes = attributes.ToString

                        strDriver = CType(cmbDBtype.Items(cmbDBtype.SelectedIndex), Mylist).Name() & Chr(0)
                        If IsSysDSN = True Then
                            intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_SYS_DSN, strDriver, strAttributes)
                        Else
                            intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)
                        End If


                        If intRet Then
                            MsgBox(DSNname & " registered successfully")
                            verRet = VerifyTables(DSNname, dataVal)
                            Select Case verRet
                                Case verifyRet.VerHasProj
                                    MsgBox(DSNname & " is a valid SQData MetaData DSN" & (Chr(10)) & "and it has SQD Studio Projects in it", MsgBoxStyle.Information)
                                Case verifyRet.VerNoProj
                                    MsgBox(DSNname & " is a valid SQData MetaData DSN" & (Chr(10)) & "but there are NO SQD Studio Projects in it", MsgBoxStyle.Information)
                                Case verifyRet.VerNotComp
                                    MsgBox(DSNname & " is an SQData MetaData DSN" & (Chr(10)) & "but there are tables missing from it", MsgBoxStyle.Exclamation)
                                Case verifyRet.VerNotSQD
                                    MsgBox(DSNname & " is NOT a valid SQData MetaData DSN", MsgBoxStyle.Critical)
                            End Select
                            DSNcreated = True
                        Else
                            MsgBox("Registration of " & DSNname & " Failed with error " & intRet, MsgBoxStyle.Critical)
                            Log(DSNname & "creation failed with return code " & intRet)
                            Exit Select
                        End If

                        '///*************************************
                        '///*** IBM DB2 *************************
                        '///*************************************
                    Case "IBM DB2 ODBC DRIVER"
                        '***************
                        ' //// ToDo ////
                        '***************
                End Select
            End If
            '//  reset form
            StartLoad()

        Catch ex As Exception
            LogError(ex, ex.Message)
            DSNcreated = False
        End Try
        

    End Sub

#End Region

#Region "Remove DSN Events"

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        Dim RowIdx As Integer = 0

        txtChosenDSN.Enabled = True
        btnRemoveDSN.Enabled = True
        btnVerify.Enabled = True
        RowIdx = DataGridView1.CurrentCell.RowIndex
        ChosenDSNname = DataGridView1.Item(0, RowIdx).Value.ToString
        ChosenDriver = DataGridView1.Item(1, RowIdx).Value.ToString
        txtChosenDSN.Text = ChosenDSNname

    End Sub

    Private Sub DataGridView1_CellContentdblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick, DataGridView1.CellMouseDoubleClick, DataGridView1.CellMouseDown

        Dim RowIdx As Integer = 0

        txtChosenDSN.Enabled = True
        btnRemoveDSN.Enabled = True
        btnVerify.Enabled = True
        RowIdx = DataGridView1.CurrentCell.RowIndex
        ChosenDSNname = DataGridView1.Item(0, RowIdx).Value.ToString
        ChosenDriver = DataGridView1.Item(1, RowIdx).Value.ToString
        txtChosenDSN.Text = ChosenDSNname

    End Sub

    Private Sub btnRemoveDSN_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveDSN.Click

        RmvDSNname = ChosenDSNname
        RmvDSNdvr = ChosenDriver

        If MsgBox("Are you sure you want to REMOVE this DSN ? -> " & RmvDSNname, MsgBoxStyle.YesNo) = MsgBoxResult.No Then
            Exit Sub
        Else
            If VerifyProjTable(False) Then
                If RemoveDSN(RmvDSNname, RmvDSNdvr) = True Then
                    MsgBox("DSN " & RmvDSNname & " was removed successfully", MsgBoxStyle.Information)
                    getDSNlist()
                Else
                    MsgBox("DSN " & RmvDSNname & " was NOT removed", MsgBoxStyle.Information)
                End If
                getDSNlist()
                txtChosenDSN.Text = ""
                RmvDSNname = ""
            Else
                If MsgBox("This is not a valid SQDMetaData Source," & (Chr(10)) & "would you like to remove it anyway?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If RemoveDSN(RmvDSNname, RmvDSNdvr) = True Then
                        MsgBox("DSN " & RmvDSNname & " was removed successfully", MsgBoxStyle.Information)
                        getDSNlist()
                    Else
                        MsgBox("DSN " & RmvDSNname & " was NOT removed", MsgBoxStyle.Information)
                    End If
                    getDSNlist()
                    txtChosenDSN.Text = ""
                    RmvDSNname = ""
                End If
            End If
        End If

    End Sub

    Function RemoveDSN(ByVal rmvDsnName As String, ByVal RmvDSNdvr As String) As Boolean

        Dim intRet As Integer
        Dim strDriver As String
        Dim strAttributes As String

        Try
            strDriver = RmvDSNdvr
            strAttributes = "DSN=" & rmvDsnName & Chr(0)

            intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)
            If intRet Then
                Return True
            Else
                intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_SYS_DSN, strDriver, strAttributes)
                If intRet Then
                    Return True
                Else
                    Return False
                End If
            End If

        Catch ex As Exception
            LogError(ex)
            Return False
        End Try

    End Function

#End Region

#Region "Click and Help events"

    Public Overrides Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            If DSNcreated = True Then
                DialogResult = Windows.Forms.DialogResult.OK
            Else
                DialogResult = Windows.Forms.DialogResult.None
            End If

        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        Finally
            Me.Close()
        End Try

    End Sub

    Public Overrides Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If IsModified = False Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Exit Sub
        End If

        If MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    Private Sub btnVerify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVerify.Click

        Dim RetMsg As verifyRet

        RetMsg = VerifyTables(ChosenDSNname)
        Select Case RetMsg
            Case verifyRet.VerHasProj
                MsgBox("This is a valid SQData MetaData DSN" & (Chr(10)) & "and it has SQD Studio Projects in it", MsgBoxStyle.Information)
            Case verifyRet.VerNoProj
                MsgBox("This is a valid SQData MetaData DSN" & (Chr(10)) & "but there are NO SQD Studio Projects in it", MsgBoxStyle.Information)
            Case verifyRet.VerNotComp
                MsgBox("This is an SQData MetaData DSN" & (Chr(10)) & "but there are tables missing from it", MsgBoxStyle.Exclamation)
            Case verifyRet.VerNotSQD
                MsgBox("This is NOT a valid SQData MetaData DSN", MsgBoxStyle.Critical)
        End Select

    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        getDSNlist()
    End Sub


    '////// Help and Shortcut Keys ///////
    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ShowHelp(modHelp.HHId.H_Installation)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click(sender, e)
            Case Keys.Enter
                cmdOk_Click(sender, e)
            Case Keys.F1
                cmdHelp_Click(sender, e)
        End Select
    End Sub

#End Region

#Region "Methods"

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

        Else
            MsgBox("(" + Str(ODBCHelper.ErrorCode) + ")" + ODBCHelper.ErrorMessage)
        End If
    End Function

    Function VerifyProjTable(Optional ByVal msg As Boolean = True) As Boolean
        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim sql As String = "select * from PROJECTS"
        Dim cnnstr As String = "DSN=" & ChosenDSNname & "; UID=" & txtUID.Text & "; PWD=" & txtPWD.Text
        'Provider=MSDASQL; 
        Dim cnn As New System.Data.Odbc.OdbcConnection(cnnstr)

        dt.Clear()
        Try
            cnn.Open()
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            If Not da.Fill(dt) > 0 Then
                If msg Then MsgBox("The source you have chosen is valid," & (Chr(10)) & "but contains no SQDStudio projects")
            Else
                If msg Then MsgBox("The source you have chosen is valid")
            End If
            Return True
        Catch ex As Exception
            If msg Then MsgBox("You have chosen an invalid SQMetaData source")
            Return False
        Finally
            dt.Clear()
            cnn.Close()
        End Try
    End Function

    Function VerifyTables(ByVal dsn As String, Optional ByVal datavalue As String = "access") As verifyRet

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp2")
        Dim cnnstr As String

        Select Case datavalue
            Case "access"
                If txtUID.Text = "" Then
                    If txtPWD.Text = "" Then
                        cnnstr = "DSN=" & dsn & ";"
                    Else
                        cnnstr = "DSN=" & dsn & "; UID=" & UID & ";"
                    End If
                Else
                    cnnstr = "DSN=" & dsn & "; UID=" & UID & "; PWD=" & PWD
                End If

            Case "SQLServer"
                Dim attributes As New System.Text.StringBuilder()

                attributes.Append("DSN=")
                attributes.Append(DSNname)
                attributes.Append(";")
                attributes.Append("Driver={SQL Server};")
                attributes.Append("Server=")
                attributes.Append(SQLSvrName)
                attributes.Append(";")
                attributes.Append("Database=")
                attributes.Append(DBname)
                attributes.Append(";")
                If UseWinAuth = True Then
                    attributes.Append("trusted_connection=yes")
                    attributes.Append(";")
                Else
                    attributes.Append("UID=")
                    attributes.Append(UID)
                    attributes.Append(";")
                    attributes.Append("Pwd=")
                    attributes.Append(PWD)
                    attributes.Append(";")
                End If
                cnnstr = attributes.ToString
            Case Else
                cnnstr = "DSN=" & DSNname & "; UID=" & UID & "; PWD=" & PWD
        End Select

        Dim cnn As New System.Data.Odbc.OdbcConnection(cnnstr)
        Dim sql As String
        Dim tablenumber As Integer = 1
        Dim retVal As verifyRet = verifyRet.VerHasProj

        Try
            cnn.Open()
            If CheckBox1.Checked = True Then
                MsgBox("Connection string is >>  " & cnnstr)
            End If
            While tablenumber < 17
                Select Case tablenumber
                    Case 1
                        sql = "select * from PROJECTS"
                    Case 2
                        sql = "select * from CONNECTIONS"
                    Case 3
                        sql = "select * from DATASTORES"
                    Case 4
                        sql = "select * from DSSELECTIONS"
                    Case 5
                        sql = "select * from DSSELFIELDS"
                    Case 6
                        sql = "select * from ENGINES"
                    Case 7
                        sql = "select * from ENVIRONMENTS"
                    Case 8
                        sql = "select * from STRSELFIELDS"
                    Case 9
                        sql = "select * from STRUCTFIELDS"
                    Case 10
                        sql = "select * from STRUCTSEL"
                    Case 11
                        sql = "select * from STRUCTURES"
                    Case 12
                        sql = "select * from SYSTEMS"
                    Case 13
                        sql = "select * from TASKDS"
                    Case 14
                        sql = "select * from TASKMAP"
                    Case 15
                        sql = "select * from TASKS"
                    Case Else
                        sql = "select * from VARIABLES"
                End Select

                If CheckBox1.Checked = True Then
                    MsgBox("sql for table verification is >>  " & sql)
                End If
                da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
                da.Fill(dt)
                If (tablenumber = 1) And (dt.Rows.Count < 1) Then
                    retVal = verifyRet.VerNoProj
                End If

                dt.Clear()
                tablenumber += 1
            End While

            Return retVal

        Catch ex As Exception
            LogError(ex)
            If tablenumber > 1 Then
                Return verifyRet.VerNotComp
            Else
                Return verifyRet.VerNotSQD
            End If
        Finally
            cnn.Close()
        End Try

    End Function

#End Region

End Class