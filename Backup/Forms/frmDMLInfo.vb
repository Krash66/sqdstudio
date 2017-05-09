Public Class frmDMLInfo
    Inherits frmBlank

    '/// This Form Created by Tom Karasch June 2007
    '********* OverHauled July 25 - 28, 2007 by TK to use Connection Objects ***************
    '********* Overhauled AGAIN August 2007 by TK to enable multiselect of DML Structures **************
    '/// to load relational DMLs from ODBC relational Databases
    '/// For now, Only DB2, Oracle, and MS SQLserver are set up.

    Dim TableObj As New clsDMLinfo  '// Master DML Obj to hold Connection, schema, etc. from 1st tab
    Dim Ulogin As clsLogin          '/// for ODBC login information
    Dim DMLinfoArray As New Collection
    Public PageArray As Collection
    Dim Connection As clsConnection

    Private ODBCHelper As New CRODBCHelper()   '// gets ODBC list

    Dim ODBCConnected As Boolean   '/// tells if ODBC was connected or not
    Dim TableChosen As Boolean     '/// tells if user chose a table from the Database
    Dim restart As Boolean = False '/// tells if user wants to start over
    Dim ObjEnv As clsEnvironment
    Dim Selectpage As TabPage
    Dim PageControl As ctlSelectTab

    Dim mDSNTable As System.Data.DataTable        '/// Table of ODBC sources
    Dim dt As New System.Data.DataTable("temp1")  '/// Table of tables from DB system catalog


#Region "Populate Tables"

    Private Function getDSNlist() As Boolean
        'GET THE DSN LIST
        Dim ret As Boolean

        ret = ODBCHelper.BuildDSNDataTable()
        If ret = True Then
            mDSNTable = ODBCHelper.DSNDataTable
            Return True
        Else
            Return False
        End If

    End Function

    Private Function CheckConn() As Boolean

        Dim TempODBCTypeStr As String
        Dim tempDBName As String
        Dim TempType As enumODBCtype

        Try
            If getDSNlist() = True Then
                For Each dr As DataRow In mDSNTable.Rows
                    tempDBName = dr(0).ToString
                    TempODBCTypeStr = dr(1).ToString
                    TempType = GetDsnType(TempODBCTypeStr)
                    If TableObj.Connection.Database = tempDBName.Trim Then
                        If TempType <> enumODBCtype.ACCESS And TempType <> enumODBCtype.OTHER Then
                            TableObj.DSNname = tempDBName.Trim
                            TableObj.DSNtype = TempType
                            Return True
                            Exit Try
                        Else
                            MsgBox("This ODBC Connection is not complete or is not DB2, Oracle, or MS SQLserver." & Chr(13) & "Please choose a different Connection or change the ODBC datasource for this Connection", MsgBoxStyle.OkOnly, MsgTitle)
                        End If
                        Exit For
                    End If
                Next
            End If

            Return False

        Catch ex As Exception
            LogError(ex, "frmDMLinfo CheckConn")
            Return False
        End Try

    End Function

    Function GetDsnType(ByVal TypeString As String) As enumODBCtype

        If TypeString <> "" Then
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
        End If

    End Function

    '/// Loads Table of Tables from System Catalog
    Private Function Load_Table() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim cnn As New System.Data.Odbc.OdbcConnection(TableObj.MetaConnectionString)
        Dim sql As String = ""

        dt.Clear()
        dt.Columns.Clear()

        Try
            Me.Cursor = Cursors.WaitCursor
            cnn.Open()

            Select Case TableObj.DSNtype

                Case enumODBCtype.DB2

                    '////// UNUSED ... this uses default views instead of acual tables *************
                    '***************************************************************************
                    'If TableObj.PartSchemaName <> "" Then
                    '    If TableObj.PartSchemaName.Contains("%") Then
                    '        If TableObj.PartTabName <> "" Then
                    '            If TableObj.PartTabName.Contains("%") Then
                    '                sql = "select tabname,tabschema from syscat.tables where tabschema like '" & TableObj.PartSchemaName & "' AND tabname like '" & TableObj.PartTabName & "' order by tabname"
                    '            Else
                    '                sql = "select tabname,tabschema from syscat.tables where tabschema like '" & TableObj.PartSchemaName & "' AND tabname= '" & TableObj.PartTabName & "' order by tabname"
                    '            End If
                    '        Else
                    '            sql = "select tabname,tabschema from syscat.tables where tabschema like '" & TableObj.PartSchemaName & "' order by tabname"
                    '        End If
                    '    Else
                    '        If TableObj.PartTabName <> "" Then
                    '            If TableObj.PartTabName.Contains("%") Then
                    '                sql = "select tabname,tabschema from syscat.tables where tabschema= '" & TableObj.PartSchemaName & "' AND tabname like '" & TableObj.PartTabName & "' order by tabname"
                    '            Else
                    '                sql = "select tabname,tabschema from syscat.tables where tabschema= '" & TableObj.PartSchemaName & "' AND tabname= '" & TableObj.PartTabName & "' order by tabname"
                    '            End If
                    '        Else
                    '            sql = "select tabname,tabschema from syscat.tables where tabschema= '" & TableObj.PartSchemaName & "' order by tabname"
                    '        End If
                    '    End If
                    'Else
                    '    If TableObj.PartTabName = "" Then
                    '        sql = "select tabname, tabschema from syscat.tables order by tabname"
                    '    Else
                    '        If TableObj.PartTabName.Contains("%") Then
                    '            sql = "select tabname, tabschema from syscat.tables where tabname like '" & TableObj.PartTabName & "' order by tabname"
                    '        Else
                    '            sql = "select tabname, tabschema from syscat.tables where tabname= '" & TableObj.PartTabName & "' order by tabname"
                    '        End If
                    '    End If
                    'End If
                    If TableObj.PartSchemaName <> "" Then
                        If TableObj.PartSchemaName.Contains("%") Then
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select name,creator from sysibm.systables where creator like '" & TableObj.PartSchemaName & "' AND name like '" & TableObj.PartTabName & "' order by name"
                                Else
                                    sql = "select name,creator from sysibm.systables where creator like '" & TableObj.PartSchemaName & "' AND name= '" & TableObj.PartTabName & "' order by name"
                                End If
                            Else
                                sql = "select name,creator from sysibm.systables where creator like '" & TableObj.PartSchemaName & "' order by name"
                            End If
                        Else
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select name,creator from sysibm.systables where creator= '" & TableObj.PartSchemaName & "' AND name like '" & TableObj.PartTabName & "' order by name"
                                Else
                                    sql = "select name,creator from sysibm.systables where creator= '" & TableObj.PartSchemaName & "' AND name= '" & TableObj.PartTabName & "' order by name"
                                End If
                            Else
                                sql = "select name,creator from sysibm.systables where creator= '" & TableObj.PartSchemaName & "' order by name"
                            End If
                        End If
                    Else
                        If TableObj.PartTabName = "" Then
                            sql = "select name,creator from sysibm.systables order by name"
                        Else
                            If TableObj.PartTabName.Contains("%") Then
                                sql = "select name,creator from sysibm.systables where name like '" & TableObj.PartTabName & "' order by name"
                            Else
                                sql = "select name,creator from sysibm.systables where name= '" & TableObj.PartTabName & "' order by name"
                            End If
                        End If
                    End If

                Case enumODBCtype.ORACLE
                    If TableObj.PartSchemaName <> "" Then
                        If TableObj.PartSchemaName.Contains("%") Then
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select table_name,tablespace_name from user_tables where tablespace_name like '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name,tablespace_name from user_tables where tablespace_name like '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name,tablespace_name from user_tables where tablespace_name like '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        Else
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select table_name,tablespace_name from user_tables where tablespace_name= '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name,tablespace_name from user_tables where tablespace_name= '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name,tablespace_name from user_tables where tablespace_name= '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        End If
                    Else
                        If TableObj.PartTabName = "" Then
                            sql = "select table_name,tablespace_name from user_tables order by table_name"
                        Else
                            If TableObj.PartTabName.Contains("%") Then
                                sql = "select table_name,tablespace_name from user_tables where table_name like '" & TableObj.PartTabName & "' order by table_name"
                            Else
                                sql = "select table_name,tablespace_name from user_tables where table_name= '" & TableObj.PartTabName & "' order by table_name"
                            End If
                        End If
                    End If

                Case enumODBCtype.SQL_SERVER
                    If TableObj.PartSchemaName <> "" Then
                        If TableObj.PartSchemaName.Contains("%") Then
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema like '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema like '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema like '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        Else
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema= '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema= '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where Table_Schema= '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        End If
                    Else
                        If TableObj.PartTabName = "" Then
                            sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES order by table_name"
                        Else
                            If TableObj.PartTabName.Contains("%") Then
                                sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where table_name like '" & TableObj.PartTabName & "' order by table_name"
                            Else
                                sql = "select table_name, Table_Schema from INFORMATION_SCHEMA.TABLES where table_name= '" & TableObj.PartTabName & "' order by table_name"
                            End If
                        End If
                    End If

            End Select

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            If Not da.Fill(dt) > 0 Then
                dgvTables.DataSource = dt
                dgvTables.Visible = True
                dgvTables.Enabled = True
                dgvTables.Show()
                dgvTables.ClearSelection()
                Me.Cursor = Cursors.Default
                MsgBox("There are no tables in your Selection" & (Chr(13)) & "Please enter a different Schema and/or Partial Table Name", MsgBoxStyle.OkOnly, MsgTitle)
                dgvTables.Enabled = False
            Else
                dgvTables.DataSource = dt
                dgvTables.Visible = True
                dgvTables.Enabled = True
                Me.Cursor = Cursors.Default
                dgvTables.Show()
                dgvTables.ClearSelection()
            End If
            Me.Cursor = Cursors.Default
            Return True

        Catch OE As Odbc.OdbcException
            Me.Cursor = Cursors.Default
            LogODBCError(OE, "frmDMLinfo Load_Table")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)

            TableObj.DoLogin = True
            ActivateTab(TabSelectTable)

            Return False

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmDMLinfo Load_Table")
            If ex.ToString.Contains("Oracle") = True Then
                If ex.ToString.Contains("null password") = True Then
                    MsgBox("Password cannot be left blank", MsgBoxStyle.OkOnly, MsgTitle)
                ElseIf ex.ToString.Contains("password") = True Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.", MsgBoxStyle.OkOnly, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.OkOnly, MsgTitle)
                End If
            ElseIf ex.ToString.Contains("IBM") = True Then
                If ex.ToString.Contains("PASSWORD MISSING") = True Then
                    MsgBox("You have entered an incorrect Password.", MsgBoxStyle.OkOnly, MsgTitle)
                ElseIf ex.ToString.Contains("24") = True Then
                    MsgBox("You have entered an incorrect Username and/or Password.", MsgBoxStyle.OkOnly, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.OkOnly, MsgTitle)
                End If
            ElseIf ex.ToString.Contains("MS") = True Then
                If ex.ToString.Contains("PASSWORD") Then
                    MsgBox("You have entered an incorrect Password.", MsgBoxStyle.OkOnly, MsgTitle)
                ElseIf ex.ToString.Contains("connection") Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.", MsgBoxStyle.OkOnly, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.OkOnly, MsgTitle)
                End If
            Else
                MsgBox("A Windows exception error occured: " & Chr(13) & _
                   ex.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            End If

            TableObj.DoLogin = True
            ActivateTab(TabSelectTable)

            Return False
        Finally
            cnn.Close()
        End Try

    End Function

#End Region

#Region "Form Functions"

    '/// Main Fuction to get DML info Data from this form
    Public Function GetInfo(ByVal Envobj As clsEnvironment) As Collection

        ObjEnv = Envobj
        If SetComboConn() = False Then
            Return Nothing
            Me.Close()
            Exit Function
        End If
        StartLoad()
        EndLoad()

doAgain:
        Select Case Me.ShowDialog()
            Case Windows.Forms.DialogResult.OK
                Return DMLinfoArray
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    '///Clears the "Select Table" tab
    Sub clearTabTable()

        dt.Clear()
        txtSchemaTab2.Text = ""
        txtPartTableName.Text = ""

    End Sub

#End Region

#Region "Textbox CheckBox Listview and keydown Events"

    '/// Refreshes Select Tables table when you change search criteria
    Private Sub tabTables_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSchemaTab2.KeyDown, txtPartTableName.KeyDown, txtSchemaTab2.KeyUp, txtPartTableName.KeyUp

        If e.KeyValue = Keys.Enter Then
            btnRefresh_Click(sender, New EventArgs)
        End If

    End Sub

    '// if user goes back to change ODBC source, "Select Tables" and "Select Columns" tables are cleared
    Private Sub txtDSN_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        clearTabTable()

    End Sub

    '/// Makes sure schema is in a usable form for the GUI
    Private Function GetSchema(ByVal input As String) As String

        If input.EndsWith(".") Then
            input = input.Remove(input.LastIndexOf("."))
            input = input.Trim()
        End If
        Return input

    End Function

    '/// clear the table if the user is changing schema search pattern
    Private Sub txtSchemaTab2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSchemaTab2.TextChanged

        TableObj.PartSchemaName = txtSchemaTab2.Text
        dt.Clear()

    End Sub

    '/// clear table if user is changing table name search pattern
    Private Sub txtPartTableName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPartTableName.TextChanged

        TableObj.PartTabName = txtPartTableName.Text
        dt.Clear()

    End Sub

#End Region

#Region "Button and Menu Events"

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Try
            Me.Cursor = Cursors.WaitCursor
            If TableObj.Connection Is Nothing Then
                TableObj.Connection = CType(cbConn.SelectedItem, Mylist).ItemData
                Call CheckSelectedConn()
            End If

            For Each Page As TabPage In PageArray
                If CType(Page.Controls(0), ctlSelectTab).SaveColumns() = True Then
                    DMLinfoArray.Add(CType(Page.Controls(0), ctlSelectTab).objthis)
                    TabControl1.TabPages.Remove(Page)
                End If
            Next

            If TabControl1.TabPages.Count > 1 Then
                Me.Cursor = Cursors.Default
                DialogResult = Windows.Forms.DialogResult.Retry
                MsgBox("There was a problem creating structures for the Tables in the Tabs remaining" & Chr(13) & "Please check these selections and try again", MsgBoxStyle.OkOnly, MsgTitle)
            Else
                Me.Cursor = Cursors.Default
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.OK
            End If

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmDMLinfo cmdOK_click")
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Relational_DML)

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

    '/// Goes to the next appropriate tab
    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click

        Try
            Me.Cursor = Cursors.WaitCursor
            Dim i As Integer
            PageArray = New Collection
            For Each page As TabPage In TabControl1.TabPages
                If page.Name <> "TabSelectTable" Then
                    TabControl1.TabPages.Remove(page)
                End If
            Next
            For Each row As DataGridViewRow In dgvTables.SelectedRows

                Selectpage = New TabPage(row.Cells(1).Value.ToString.Trim & "." & row.Cells(0).Value.ToString.Trim)
                PageControl = New ctlSelectTab

                PageControl.objthis.TableName = row.Cells(0).Value.ToString
                PageControl.objthis.Schema = row.Cells(1).Value.ToString
                PageControl.objthis.Connection = TableObj.Connection
                PageControl.objthis.DSNname = TableObj.DSNname
                PageControl.objthis.DSNtype = TableObj.DSNtype

                Selectpage.Controls.Add(PageControl)

                TabControl1.TabPages.Add(Selectpage)

                With Selectpage
                    .Name = PageControl.objthis.TableName
                End With
                With PageControl
                    .Parent = Selectpage
                    .Dock = DockStyle.Fill
                    .gbColumns.Text = Selectpage.Text
                    .Tag = PageControl.objthis
                End With
                PageArray.Add(Selectpage, Selectpage.Text)
            Next

            If TabControl1.TabCount > 1 Then
                For i = 1 To TabControl1.TabPages.Count - 1
                    TabControl1.SelectedIndex = i
                    ActivateTab(TabControl1.TabPages(i))
                Next
            End If
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            LogError(ex, "frmDMLinfo btnNext_Click")
        End Try

    End Sub

    Private Sub btnRemoveTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveTab.Click

        Call RemoveTab(TabControl1.SelectedTab)

    End Sub

    '/// this refreshes the list of tables from system catalog
    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click

        If CheckConn() = True Then
            If Load_Table() = True Then
                ODBCConnected = True
            Else
                ODBCConnected = False
            End If
        End If

    End Sub

    Private Sub cbConn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbConn.SelectedIndexChanged

        If IsEventFromCode = False Then
            If cbConn.Text <> "" Then
                TableObj.Connection = CType(cbConn.SelectedItem, Mylist).ItemData
                Call CheckSelectedConn()
            End If
        End If

    End Sub

    Function CheckSelectedConn() As Boolean

        Try
            Dim MBtext As String = ""

            If TableObj.Connection.Database.Trim = "" Then
                MBtext = "The Connection you have selected does not have a Database defined"
                GoTo DoMsg
            End If
            If TableObj.Connection.UserId.Trim = "" Then
                MBtext = "The Connection you have selected does not have a UserID defined"
                GoTo DoMsg
            End If
            If TableObj.Connection.Password.Trim = "" Then
                MBtext = "The Connection you have selected does not have a Password defined"
            End If
DoMsg:      If MBtext <> "" Then
                MsgBox(MBtext & Chr(13) & "Please complete this Connection's information or choose a different Connection", MsgBoxStyle.OkOnly, MsgTitle)
                Return False
                Exit Try
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLinfo CheckSelectedConn")
            Return False
        End Try

    End Function

#End Region

#Region "DataGridView Events"

    Private Sub dgvTables_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvTables.CellMouseClick



        If dgvTables.SelectedRows.Count > 0 Then
            TableChosen = True
            btnNext.Enabled = True
            btnNext.Visible = True
        Else
            TableChosen = False
            btnNext.Enabled = False
            btnNext.Visible = False
        End If

    End Sub

    Private Sub dgvTables_CellContentDblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles dgvTables.CellMouseDoubleClick

        If dgvTables.SelectedRows.Count > 0 Then
            TableChosen = True
            btnNext.Enabled = True
            btnNext.Visible = True

            btnNext_Click(sender, New EventArgs)
        Else
            TableChosen = False
            btnNext.Enabled = False
            btnNext.Visible = False
        End If

    End Sub

#End Region

#Region "Tab Events"

    '/// function to activate each tab based on what info user has entered or chosen
    Function ActivateTab(ByVal Tab As TabPage) As Boolean

        Try
            Dim idx As Integer
            If Tab Is Nothing Then
                Tab = TabControl1.TabPages(0)
                TabControl1.SelectedTab = Tab
            End If

            idx = TabControl1.TabPages.IndexOf(Tab)

            Select Case idx

                Case 0
                    cmdOk.Enabled = False
                    gbNarrowSearch.Enabled = True
                    gbTableList.Enabled = True
                    txtSchemaTab2.Text = TableObj.PartSchemaName
                    txtPartTableName.Text = TableObj.PartTabName
                    dgvTables.ClearSelection()
                    btnNext.Enabled = False
                    btnNext.Visible = True
                    btnRemoveTab.Enabled = False
                    btnRemoveTab.Visible = False
                    Tab.Show()



                Case Is > 0
                    'Dim obj As clsDMLinfo = CType(Tab.Tag, clsDMLinfo)
                    Dim ctl As ctlSelectTab = Tab.Controls(0)

                    If ODBCConnected = True Then
                        If TableChosen = True Then
                            ctl.gbColumns.Enabled = True
                            ctl.gblvColumns.Enabled = True
                            If ctl.objthis.TableName <> "" And ctl.objthis.DSNname <> "" Then
                                ctl.gbColumns.Text = "Columns in Table " & ctl.objthis.TableName & " with Schema " & ctl.objthis.Schema
                                btnNext.Enabled = False
                                btnNext.Visible = False
                                btnRemoveTab.Enabled = True
                                btnRemoveTab.Visible = True
                                If ctl.Load_ColumnTable() = False Then
                                    If MsgBox("Would you like to try again?", MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
                                        ctl.objthis.TryAgain = True
                                        ctl.objthis.TableName = ""
                                        restart = True
                                        ActivateTab(TabSelectTable)
                                    Else
                                        ctl.objthis.TryAgain = False
                                        ctl.objthis.TableName = ""
                                        restart = False
                                        DialogResult = Windows.Forms.DialogResult.Cancel
                                        Me.Close()
                                    End If
                                Else
                                    cmdOk.Enabled = True
                                End If
                            Else
                                TabControl1.SelectedIndex = 0
                                ActivateTab(TabSelectTable)
                            End If
                        Else
                            TabControl1.SelectedIndex = 0
                            ActivateTab(TabSelectTable)
                        End If
                    End If

            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLinfo ActivateTab")
            Return False
        End Try

    End Function

    Private Sub TabControl1_selectedChanged(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected

        ActivateTab(TabControl1.SelectedTab)

    End Sub

    Public Function RemoveTab(ByRef Tab As TabPage) As Boolean

        Try
            PageArray.Remove(Tab.Text)
            TabControl1.TabPages.Remove(Tab)

            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLinfo RemoveTab")
            Return False
        End Try
        
    End Function

#End Region

#Region "Load Events"

    '/// Initializes the form, clears all tables
    Private Sub StartLoad()

        IsEventFromCode = True

        gbChseConn.Enabled = True
        gbNarrowSearch.Enabled = True
        gbTableList.Enabled = True

        btnNext.Enabled = False
        cmdOk.Enabled = False
        btnRemoveTab.Enabled = False
        btnRemoveTab.Visible = False

        TableObj.DoLogin = True

        ODBCConnected = False
        TableChosen = False

        clearTabTable()

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

    Function SetComboConn() As Boolean

        Try
            cbConn.Items.Clear()

            For Each conn As clsConnection In ObjEnv.Connections
                conn.LoadMe()
                If conn.ConnectionType = "ODBC" Then
                    cbConn.Items.Add(New Mylist(conn.ConnectionName, conn))
                End If
            Next
            If cbConn.Items.Count > 0 Then
                cbConn.SelectedIndex = 0
                TableObj.Connection = CType(cbConn.SelectedItem, Mylist).ItemData
                Return True
            Else
                MsgBox("There are no Connections defined of type 'ODBC'." & Chr(13) & "You must use an ODBC Connection to create this DML type", MsgBoxStyle.OkOnly, MsgTitle)
                Return False
            End If

        Catch ex As Exception
            LogError(ex, "frmDMLInfo SetComboConn")
            Return False
        End Try

    End Function

#End Region
    
   
End Class