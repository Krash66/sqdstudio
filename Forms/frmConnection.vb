Public Class frmConnection
    Inherits SQDStudio.frmBlank

    Public objThis As New clsConnection
    Dim mDSNTable As System.Data.DataTable        '/// Table of ODBC sources
    Private ODBCHelper As New CRODBCHelper()      '// gets ODBC list

    Dim IsNewObj As Boolean

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtConnectionDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmbCnnType As System.Windows.Forms.ComboBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents txtConnectionName As System.Windows.Forms.TextBox
    Friend WithEvents txtDateformat As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbODBC As System.Windows.Forms.ComboBox
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbDB As System.Windows.Forms.GroupBox
    Friend WithEvents gbLogin As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtConnectionDesc = New System.Windows.Forms.TextBox
        Me.cmbCnnType = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDatabase = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtUserId = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtConnectionName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDateformat = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbODBC = New System.Windows.Forms.ComboBox
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbDB = New System.Windows.Forms.GroupBox
        Me.gbLogin = New System.Windows.Forms.GroupBox
        Me.Panel1.SuspendLayout()
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbDB.SuspendLayout()
        Me.gbLogin.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(458, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 423)
        Me.GroupBox1.Size = New System.Drawing.Size(460, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(176, 446)
        Me.cmdOk.TabIndex = 8
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(275, 446)
        Me.cmdCancel.TabIndex = 9
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(368, 446)
        '
        'Label1
        '
        Me.Label1.Text = "Connection definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(382, 39)
        Me.Label2.Text = "Enter a unique connection name within a system."
        '
        'txtConnectionDesc
        '
        Me.txtConnectionDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtConnectionDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtConnectionDesc.MaxLength = 1000
        Me.txtConnectionDesc.Multiline = True
        Me.txtConnectionDesc.Name = "txtConnectionDesc"
        Me.txtConnectionDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConnectionDesc.Size = New System.Drawing.Size(416, 74)
        Me.txtConnectionDesc.TabIndex = 7
        '
        'cmbCnnType
        '
        Me.cmbCnnType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCnnType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCnnType.Location = New System.Drawing.Point(50, 45)
        Me.cmbCnnType.Name = "cmbCnnType"
        Me.cmbCnnType.Size = New System.Drawing.Size(378, 21)
        Me.cmbCnnType.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(34, 14)
        Me.Label7.TabIndex = 36
        Me.Label7.Text = "Type"
        '
        'txtDatabase
        '
        Me.txtDatabase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabase.Location = New System.Drawing.Point(86, 19)
        Me.txtDatabase.MaxLength = 128
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(339, 20)
        Me.txtDatabase.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 14)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Database"
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Location = New System.Drawing.Point(86, 45)
        Me.txtPassword.MaxLength = 128
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(339, 20)
        Me.txtPassword.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 14)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "Password"
        '
        'txtUserId
        '
        Me.txtUserId.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserId.Location = New System.Drawing.Point(86, 19)
        Me.txtUserId.MaxLength = 128
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(339, 20)
        Me.txtUserId.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 14)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Username"
        '
        'txtConnectionName
        '
        Me.txtConnectionName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConnectionName.Location = New System.Drawing.Point(50, 19)
        Me.txtConnectionName.MaxLength = 128
        Me.txtConnectionName.Name = "txtConnectionName"
        Me.txtConnectionName.Size = New System.Drawing.Size(378, 20)
        Me.txtConnectionName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Name"
        '
        'txtDateformat
        '
        Me.txtDateformat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDateformat.Location = New System.Drawing.Point(86, 46)
        Me.txtDateformat.MaxLength = 128
        Me.txtDateformat.Name = "txtDateformat"
        Me.txtDateformat.Size = New System.Drawing.Size(339, 20)
        Me.txtDateformat.TabIndex = 4
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 49)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(73, 14)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Date Format"
        '
        'cmbODBC
        '
        Me.cmbODBC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbODBC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbODBC.FormattingEnabled = True
        Me.cmbODBC.Location = New System.Drawing.Point(86, 19)
        Me.cmbODBC.Name = "cmbODBC"
        Me.cmbODBC.Size = New System.Drawing.Size(339, 21)
        Me.cmbODBC.TabIndex = 58
        '
        'gbName
        '
        Me.gbName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Controls.Add(Me.txtConnectionName)
        Me.gbName.Controls.Add(Me.Label7)
        Me.gbName.Controls.Add(Me.cmbCnnType)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbName.Location = New System.Drawing.Point(12, 74)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(434, 171)
        Me.gbName.TabIndex = 59
        Me.gbName.TabStop = False
        Me.gbName.Text = "Connection Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtConnectionDesc)
        Me.gbDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDesc.Location = New System.Drawing.Point(6, 72)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(422, 93)
        Me.gbDesc.TabIndex = 37
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbDB
        '
        Me.gbDB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDB.Controls.Add(Me.Label5)
        Me.gbDB.Controls.Add(Me.cmbODBC)
        Me.gbDB.Controls.Add(Me.Label9)
        Me.gbDB.Controls.Add(Me.txtDateformat)
        Me.gbDB.Controls.Add(Me.txtDatabase)
        Me.gbDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDB.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDB.Location = New System.Drawing.Point(12, 251)
        Me.gbDB.Name = "gbDB"
        Me.gbDB.Size = New System.Drawing.Size(434, 81)
        Me.gbDB.TabIndex = 60
        Me.gbDB.TabStop = False
        Me.gbDB.Text = "Connection Database Properties"
        '
        'gbLogin
        '
        Me.gbLogin.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbLogin.Controls.Add(Me.Label4)
        Me.gbLogin.Controls.Add(Me.txtUserId)
        Me.gbLogin.Controls.Add(Me.txtPassword)
        Me.gbLogin.Controls.Add(Me.Label6)
        Me.gbLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLogin.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbLogin.Location = New System.Drawing.Point(12, 338)
        Me.gbLogin.Name = "gbLogin"
        Me.gbLogin.Size = New System.Drawing.Size(434, 78)
        Me.gbLogin.TabIndex = 61
        Me.gbLogin.TabStop = False
        Me.gbLogin.Text = "Database Login Properties"
        '
        'frmConnection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(458, 492)
        Me.Controls.Add(Me.gbLogin)
        Me.Controls.Add(Me.gbDB)
        Me.Controls.Add(Me.gbName)
        Me.MaximumSize = New System.Drawing.Size(466, 526)
        Me.MinimumSize = New System.Drawing.Size(466, 526)
        Me.Name = "frmConnection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Connection Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.gbName, 0)
        Me.Controls.SetChildIndex(Me.gbDB, 0)
        Me.Controls.SetChildIndex(Me.gbLogin, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbDB.ResumeLayout(False)
        Me.gbDB.PerformLayout()
        Me.gbLogin.ResumeLayout(False)
        Me.gbLogin.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConnectionDesc.TextChanged, txtConnectionName.TextChanged, txtDatabase.TextChanged, txtDateformat.TextChanged, txtPassword.TextChanged, txtUserId.TextChanged, cmbODBC.SelectedIndexChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub OnCnnTypeChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCnnType.SelectedIndexChanged

        If cmbCnnType.Text = "ODBC" Then
            txtDatabase.Visible = False
            cmbODBC.Visible = True
        Else
            txtDatabase.Visible = True
            cmbODBC.Visible = False
        End If
        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()
        IsEventFromCode = True
        objThis.IsModified = False
    End Sub

    Private Sub SetDefaultName()

        Try
            Dim EnvObj As clsEnvironment = objThis.Environment
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = "Conn_" & Count.ToString '& "_" & Strings.Left(EnvObj.Text, 9)
                For Each TestConn As clsConnection In EnvObj.Connections
                    If TestConn.ConnectionName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtConnectionName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmConn SetDefaultName")
        End Try

    End Sub

    Private Sub EndLoad()
        IsEventFromCode = False
    End Sub

    Public Overrides Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If objThis.IsModified = False Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Exit Sub
        End If

        If MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    '//Returns new Environment object
    Public Function NewObj() As clsConnection

        IsNewObj = True

        StartLoad()

        SetDefaultName()

        InitControls()

        txtConnectionName.SelectAll()

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = objThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If ValidateNewName(txtConnectionName.Text) = False Or objThis.ValidateNewObject(txtConnectionName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.ConnectionDescription = txtConnectionDesc.Text
            objThis.ConnectionType = cmbCnnType.Text
            objThis.ConnectionName = txtConnectionName.Text

            If objThis.ConnectionType <> "ODBC" Then
                objThis.Database = txtDatabase.Text
            Else
                objThis.Database = cmbODBC.Text
            End If

            objThis.UserId = txtUserId.Text
            objThis.Password = txtPassword.Text
            objThis.DateFormat = txtDateformat.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                If objThis.Save() = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            End If

            DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()

        Catch ex As Exception
            LogError(ex, "frmConnection cmdOK_click")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Sub InitControls()

        cmbCnnType.Items.Clear()
        cmbCnnType.Items.Add(New Mylist("ODBC", "ODBC"))
        cmbCnnType.Items.Add(New Mylist("ODBCwithSubstitutionVariable", "ODBCwithSubstitutionVariable"))
        cmbCnnType.Items.Add(New Mylist("DB2", "DB2"))
        cmbCnnType.Items.Add(New Mylist("OCI", "OCI"))
        cmbCnnType.Items.Add(New Mylist("NATIVEORI", "NATIVEORI"))
        cmbCnnType.SelectedIndex = 0

        SetComboConn()

        txtConnectionName.Focus()

    End Sub

    Function SetComboConn() As Boolean

        cmbODBC.Items.Clear()

        Dim tempDBName As String

        Try
            If getDSNlist() = True Then
                For Each dr As DataRow In mDSNTable.Rows
                    If dr(0).ToString <> "" Then
                        tempDBName = dr(0).ToString
                        cmbODBC.Items.Add(New Mylist(tempDBName, tempDBName))
                    End If
                Next
            End If

            cmbODBC.SelectedIndex = 0

            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLInfo SetComboConn")
            Return False
        End Try

    End Function

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

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ShowHelp(modHelp.HHId.H_Add_Connections)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Select Case e.KeyCode
            Case Keys.Escape
                Call cmdCancel_Click(sender, e)
            Case Keys.Enter
                Call cmdOk_Click(sender, e)
            Case Keys.F1
                Call cmdHelp_Click(sender, e)
        End Select
    End Sub

End Class
