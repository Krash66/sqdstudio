Public Class ctlConnection
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    Dim mDSNTable As System.Data.DataTable        '/// Table of ODBC sources
    Private ODBCHelper As New CRODBCHelper()   '// gets ODBC list
    Dim TableObj As New clsDMLinfo  '// Master DML Obj to hold Connection, schema, etc. from 1st tab
    Dim dt As New System.Data.DataTable("temp1")  '/// Table of tables from DB system catalog

    Dim objThis As New clsConnection
    Friend WithEvents lblTest As System.Windows.Forms.Label
    
    Dim IsNewObj As Boolean

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
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
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtDateformat As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtConnectionDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmbCnnType As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtConnectionName As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents cmbODBC As System.Windows.Forms.ComboBox
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbDB As System.Windows.Forms.GroupBox
    Friend WithEvents gbLogin As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbSubVar As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnTest As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtDateformat = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
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
        Me.lblName = New System.Windows.Forms.Label
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmbODBC = New System.Windows.Forms.ComboBox
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbDB = New System.Windows.Forms.GroupBox
        Me.gbLogin = New System.Windows.Forms.GroupBox
        Me.btnTest = New System.Windows.Forms.Button
        Me.gbSubVar = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblTest = New System.Windows.Forms.Label
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbDB.SuspendLayout()
        Me.gbLogin.SuspendLayout()
        Me.gbSubVar.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(403, 485)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 8
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(483, 485)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 9
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 469)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(627, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'txtDateformat
        '
        Me.txtDateformat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDateformat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDateformat.Location = New System.Drawing.Point(85, 45)
        Me.txtDateformat.MaxLength = 128
        Me.txtDateformat.Name = "txtDateformat"
        Me.txtDateformat.Size = New System.Drawing.Size(251, 20)
        Me.txtDateformat.TabIndex = 4
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(73, 14)
        Me.Label9.TabIndex = 54
        Me.Label9.Text = "Date Format"
        '
        'txtConnectionDesc
        '
        Me.txtConnectionDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConnectionDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtConnectionDesc.Location = New System.Drawing.Point(9, 19)
        Me.txtConnectionDesc.MaxLength = 1000
        Me.txtConnectionDesc.Multiline = True
        Me.txtConnectionDesc.Name = "txtConnectionDesc"
        Me.txtConnectionDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConnectionDesc.Size = New System.Drawing.Size(315, 56)
        Me.txtConnectionDesc.TabIndex = 7
        '
        'cmbCnnType
        '
        Me.cmbCnnType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbCnnType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCnnType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCnnType.Location = New System.Drawing.Point(50, 45)
        Me.cmbCnnType.Name = "cmbCnnType"
        Me.cmbCnnType.Size = New System.Drawing.Size(286, 21)
        Me.cmbCnnType.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(34, 14)
        Me.Label7.TabIndex = 50
        Me.Label7.Text = "Type"
        '
        'txtDatabase
        '
        Me.txtDatabase.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatabase.Location = New System.Drawing.Point(85, 18)
        Me.txtDatabase.MaxLength = 128
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(251, 20)
        Me.txtDatabase.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.Window
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 14)
        Me.Label5.TabIndex = 48
        Me.Label5.Text = "Database"
        '
        'txtPassword
        '
        Me.txtPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPassword.Location = New System.Drawing.Point(85, 45)
        Me.txtPassword.MaxLength = 128
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(251, 20)
        Me.txtPassword.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.Window
        Me.Label6.Location = New System.Drawing.Point(6, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(63, 14)
        Me.Label6.TabIndex = 46
        Me.Label6.Text = "Password"
        '
        'txtUserId
        '
        Me.txtUserId.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUserId.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserId.Location = New System.Drawing.Point(85, 19)
        Me.txtUserId.MaxLength = 128
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(251, 20)
        Me.txtUserId.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.Window
        Me.Label4.Location = New System.Drawing.Point(6, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 14)
        Me.Label4.TabIndex = 44
        Me.Label4.Text = "Username"
        '
        'txtConnectionName
        '
        Me.txtConnectionName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtConnectionName.Location = New System.Drawing.Point(50, 19)
        Me.txtConnectionName.MaxLength = 128
        Me.txtConnectionName.Name = "txtConnectionName"
        Me.txtConnectionName.Size = New System.Drawing.Size(286, 20)
        Me.txtConnectionName.TabIndex = 1
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(6, 22)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(38, 14)
        Me.lblName.TabIndex = 42
        Me.lblName.Text = "Name"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(563, 485)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 10
        Me.cmdHelp.Text = "&Help"
        '
        'cmbODBC
        '
        Me.cmbODBC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbODBC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbODBC.Location = New System.Drawing.Point(85, 18)
        Me.cmbODBC.Name = "cmbODBC"
        Me.cmbODBC.Size = New System.Drawing.Size(251, 21)
        Me.cmbODBC.TabIndex = 55
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.lblName)
        Me.gbName.Controls.Add(Me.txtConnectionName)
        Me.gbName.Controls.Add(Me.Label7)
        Me.gbName.Controls.Add(Me.cmbCnnType)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.Color.White
        Me.gbName.Location = New System.Drawing.Point(3, 3)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(343, 168)
        Me.gbName.TabIndex = 56
        Me.gbName.TabStop = False
        Me.gbName.Text = "Connection Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Controls.Add(Me.txtConnectionDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 72)
        Me.gbDesc.MaximumSize = New System.Drawing.Size(331, 86)
        Me.gbDesc.MinimumSize = New System.Drawing.Size(331, 86)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(331, 86)
        Me.gbDesc.TabIndex = 59
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbDB
        '
        Me.gbDB.Controls.Add(Me.txtDatabase)
        Me.gbDB.Controls.Add(Me.Label5)
        Me.gbDB.Controls.Add(Me.cmbODBC)
        Me.gbDB.Controls.Add(Me.Label9)
        Me.gbDB.Controls.Add(Me.txtDateformat)
        Me.gbDB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDB.ForeColor = System.Drawing.Color.White
        Me.gbDB.Location = New System.Drawing.Point(3, 227)
        Me.gbDB.Name = "gbDB"
        Me.gbDB.Size = New System.Drawing.Size(343, 79)
        Me.gbDB.TabIndex = 57
        Me.gbDB.TabStop = False
        Me.gbDB.Text = "Connection Database Properties"
        '
        'gbLogin
        '
        Me.gbLogin.Controls.Add(Me.btnTest)
        Me.gbLogin.Controls.Add(Me.Label4)
        Me.gbLogin.Controls.Add(Me.txtUserId)
        Me.gbLogin.Controls.Add(Me.txtPassword)
        Me.gbLogin.Controls.Add(Me.Label6)
        Me.gbLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLogin.ForeColor = System.Drawing.Color.White
        Me.gbLogin.Location = New System.Drawing.Point(3, 312)
        Me.gbLogin.Name = "gbLogin"
        Me.gbLogin.Size = New System.Drawing.Size(343, 103)
        Me.gbLogin.TabIndex = 58
        Me.gbLogin.TabStop = False
        Me.gbLogin.Text = "Database Login Properties"
        '
        'btnTest
        '
        Me.btnTest.ForeColor = System.Drawing.Color.White
        Me.btnTest.Location = New System.Drawing.Point(116, 71)
        Me.btnTest.Name = "btnTest"
        Me.btnTest.Size = New System.Drawing.Size(172, 23)
        Me.btnTest.TabIndex = 47
        Me.btnTest.Text = "Test ODBC Connection"
        Me.btnTest.UseVisualStyleBackColor = True
        '
        'gbSubVar
        '
        Me.gbSubVar.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.gbSubVar.Controls.Add(Me.Label1)
        Me.gbSubVar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSubVar.ForeColor = System.Drawing.Color.Red
        Me.gbSubVar.Location = New System.Drawing.Point(3, 177)
        Me.gbSubVar.Name = "gbSubVar"
        Me.gbSubVar.Size = New System.Drawing.Size(343, 44)
        Me.gbSubVar.TabIndex = 59
        Me.gbSubVar.TabStop = False
        Me.gbSubVar.Text = "IMPORTANT"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(297, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Properties below may contain substitution variables"
        '
        'lblTest
        '
        Me.lblTest.AutoSize = True
        Me.lblTest.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTest.Location = New System.Drawing.Point(42, 418)
        Me.lblTest.Name = "lblTest"
        Me.lblTest.Size = New System.Drawing.Size(298, 20)
        Me.lblTest.TabIndex = 60
        Me.lblTest.Text = "Please Wait as Connection is tested"
        '
        'ctlConnection
        '
        Me.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.lblTest)
        Me.Controls.Add(Me.gbSubVar)
        Me.Controls.Add(Me.gbLogin)
        Me.Controls.Add(Me.gbDB)
        Me.Controls.Add(Me.gbName)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.Name = "ctlConnection"
        Me.Size = New System.Drawing.Size(643, 517)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbDB.ResumeLayout(False)
        Me.gbDB.PerformLayout()
        Me.gbLogin.ResumeLayout(False)
        Me.gbLogin.PerformLayout()
        Me.gbSubVar.ResumeLayout(False)
        Me.gbSubVar.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtConnectionName.Enabled = True '//added on 7/24 to disable object name editing

        InitControls()

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsConnection

        ClearControls(Me.Controls)

        lblTest.Visible = False

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor

            '// First Check Validity before Saving
            If ValidateNewName(txtConnectionName.Text) = False Then
                Save = False
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            If objThis.ConnectionName <> txtConnectionName.Text Then
                objThis.IsRenamed = RenameConnection(objThis, txtConnectionName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtConnectionName.Text = objThis.ConnectionName
            Else
                objThis.ConnectionName = txtConnectionName.Text
            End If

            objThis.ConnectionDescription = txtConnectionDesc.Text
            objThis.ConnectionType = cmbCnnType.Text

            If objThis.ConnectionType <> "ODBC" Then
                If txtDatabase.Text.StartsWith("%") = True Then
                    txtDatabase.Text = AddToSubstList(txtDatabase.Text)
                End If
                objThis.Database = txtDatabase.Text
            Else
                objThis.Database = cmbODBC.Text
            End If

            If txtUserId.Text.StartsWith("%") = True Then
                txtUserId.Text = AddToSubstList(txtUserId.Text)
            End If
            objThis.UserId = txtUserId.Text
            If txtPassword.Text.StartsWith("%") = True Then
                txtPassword.Text = AddToSubstList(txtPassword.Text)
            End If
            objThis.Password = txtPassword.Text
            objThis.DateFormat = txtDateformat.Text

            If objThis.Save() = False Then
                Save = False
                Me.Cursor = Cursors.Default
                Exit Function
            End If
            Me.Cursor = Cursors.Default

            Save = True

            cmdSave.Enabled = False

            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If
            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex)
            Save = False
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Public Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        Save()
    End Sub

    Public Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Dim ret As MsgBoxResult

        Try
            If objThis.IsModified = True Then
                ret = MsgBox("Do you want to save change(s) made to the opened form?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel, MsgTitle)
                If ret = MsgBoxResult.Yes Then
                    Save()
                ElseIf ret = MsgBoxResult.No Then
                    If IsNewObj = True Then
                        objThis = Nothing
                    Else
                        objThis.IsModified = False
                    End If
                ElseIf ret = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Me.Visible = False
            RaiseEvent Closed(Me, objThis)

        Catch ex As Exception
            LogError(ex)
        End Try
    End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCnnType.SelectedIndexChanged, txtConnectionDesc.TextChanged, txtConnectionName.TextChanged, txtDatabase.TextChanged, txtDateformat.TextChanged, txtPassword.TextChanged, txtUserId.TextChanged, cmbODBC.SelectedIndexChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub OnCnnTypeChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCnnType.SelectedIndexChanged

        If cmbCnnType.Text = "ODBC" Then
            txtDatabase.Visible = False
            cmbODBC.Visible = True
            btnTest.Visible = True
        Else
            txtDatabase.Visible = True
            cmbODBC.Visible = False
            btnTest.Visible = False
        End If

        If cmbCnnType.Text <> "ODBC" Then
            AdjustGroupBoxes(True)
        Else
            AdjustGroupBoxes(False)
        End If

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Public Function EditObj(ByVal obj As INode) As clsConnection

        Try
            Me.Cursor = Cursors.WaitCursor

            IsNewObj = False
            StartLoad()
            objThis = obj '//Load the form env object

            objThis.LoadMe()

            UpdateFields()
            EndLoad()

            EditObj = objThis
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            LogError(ex, "ctlConn EditObj")
            Me.Cursor = Cursors.Default
            EditObj = Nothing
        End Try

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtConnectionName.Text = objThis.ConnectionName
        txtConnectionDesc.Text = objThis.ConnectionDescription
        txtUserId.Text = objThis.UserId
        txtPassword.Text = objThis.Password
        txtDateformat.Text = objThis.DateFormat
        cmbCnnType.Text = objThis.ConnectionType

        If objThis.ConnectionType = "ODBC" Then
            cmbODBC.Text = objThis.Database
        Else
            txtDatabase.Text = objThis.Database
        End If

    End Function

    Sub InitControls()

        cmbCnnType.Items.Clear()
        cmbCnnType.Items.Add(New Mylist("ODBC", "ODBC"))
        cmbCnnType.Items.Add(New Mylist("ODBCwithSubstitutionVariable", "ODBCwithSubstitutionVariable"))
        cmbCnnType.Items.Add(New Mylist("DB2", "DB2"))
        cmbCnnType.Items.Add(New Mylist("OCI", "OCI"))
        cmbCnnType.Items.Add(New Mylist("NATIVEORA", "NATIVEORA"))

        SetComboConn()

        cmbCnnType.SelectedIndex = 0

    End Sub

    Sub AdjustGroupBoxes(ByVal Subst As Boolean)

        If Subst = True Then
            gbSubVar.Visible = True
            gbDB.Top = gbSubVar.Bottom + 1
            gbLogin.Top = gbDB.Bottom + 1
        Else
            gbSubVar.Visible = False
            gbDB.Top = gbName.Bottom + 1
            gbLogin.Top = gbDB.Bottom + 1
        End If

    End Sub

    Function SetComboConn() As Boolean

        cmbODBC.Items.Clear()

        Dim tempDBName As String

        Try
            If getDSNlist() = True Then
                For Each dr As DataRow In mDSNTable.Rows
                    If dr(0).ToString.Trim <> "" Then
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

    Public Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Connections)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbCnnType.KeyDown, txtConnectionDesc.KeyDown, txtConnectionName.KeyDown, txtDatabase.KeyDown, txtDateformat.KeyDown, txtPassword.KeyDown, txtUserId.KeyDown, Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    Private Function CheckConn() As Boolean

        Dim TempODBCTypeStr As String
        Dim tempDBName As String
        Dim TempType As enumODBCtype

        Try
            TableObj.DSNname = CType(cmbODBC.SelectedItem, Mylist).ItemData
            TableObj.Connection = objThis
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
            LogError(ex, "ctlConnection CheckConn")
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

    Private Function Load_Table() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim cnn As New System.Data.Odbc.OdbcConnection(TableObj.MetaConnectionString)
        Dim sql As String = ""

        dt.Clear()
        dt.Columns.Clear()

        Try
            cnn.Open()

            Select Case TableObj.DSNtype

                Case enumODBCtype.DB2
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
                                    sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema like '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema like '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema like '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        Else
                            If TableObj.PartTabName <> "" Then
                                If TableObj.PartTabName.Contains("%") Then
                                    sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema= '" & TableObj.PartSchemaName & "' AND table_name like '" & TableObj.PartTabName & "' order by table_name"
                                Else
                                    sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema= '" & TableObj.PartSchemaName & "' AND table_name= '" & TableObj.PartTabName & "' order by table_name"
                                End If
                            Else
                                sql = "select table_name, Table_Schema from information_schema.tables where Table_Schema= '" & TableObj.PartSchemaName & "' order by table_name"
                            End If
                        End If
                    Else
                        If TableObj.PartTabName = "" Then
                            sql = "select table_name, Table_Schema from information_schema.tables order by table_name"
                        Else
                            If TableObj.PartTabName.Contains("%") Then
                                sql = "select table_name, Table_Schema from information_schema.tables where table_name like '" & TableObj.PartTabName & "' order by table_name"
                            Else
                                sql = "select table_name, Table_Schema from information_schema.tables where table_name= '" & TableObj.PartTabName & "' order by table_name"
                            End If
                        End If
                    End If

            End Select

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            If Not da.Fill(dt) > 0 Then
                'dgvTables.DataSource = dt
                'dgvTables.Visible = True
                'dgvTables.Enabled = True
                'dgvTables.Show()
                'dgvTables.ClearSelection()
                MsgBox("There are no tables in your Selection" & (Chr(13)) & "Please enter a different Schema and/or Partial Table Name", MsgBoxStyle.OkOnly, MsgTitle)
                'dgvTables.Enabled = False
            Else
                'dgvTables.DataSource = dt
                'dgvTables.Visible = True
                'dgvTables.Enabled = True
                'dgvTables.Show()
                'dgvTables.ClearSelection()
            End If

            Return True

        Catch ex As Exception
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
                MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.OkOnly, MsgTitle)
            End If

            TableObj.DoLogin = True
            'ActivateTab(TabSelectTable)

            Return False
        Finally
            cnn.Close()
        End Try

    End Function

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click

        Try
            lblTest.Visible = True
            Me.Refresh()
            If CheckConn() = True Then
                If Load_Table() = True Then
                    lblTest.Visible = False
                    MsgBox("Test Succeeded!", MsgBoxStyle.OkOnly, "Connection Test")
                Else
                    lblTest.Visible = False
                    MsgBox("Test Failed", MsgBoxStyle.Critical, "Connection Test")
                End If
            Else
                lblTest.Visible = False
                MsgBox("Test Failed", MsgBoxStyle.Critical, "Connection Test")
            End If

        Catch ex As Exception
            LogError(ex, "ctlConnection btnTest_Click")
        End Try

    End Sub

End Class
