<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegDSN
    Inherits SQStudio.frmBlank

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmRegDSN))
        Me.txtDSN = New System.Windows.Forms.TextBox
        Me.txtDBN = New System.Windows.Forms.TextBox
        Me.txtDBpath = New System.Windows.Forms.TextBox
        Me.txtDesc = New System.Windows.Forms.TextBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.lblDSNname = New System.Windows.Forms.Label
        Me.lblDBname = New System.Windows.Forms.Label
        Me.lbltgtLocation = New System.Windows.Forms.Label
        Me.lblDBtype = New System.Windows.Forms.Label
        Me.cmbDBtype = New System.Windows.Forms.ComboBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbUIDpwd = New System.Windows.Forms.GroupBox
        Me.cbxWinAuth = New System.Windows.Forms.CheckBox
        Me.lblPWD = New System.Windows.Forms.Label
        Me.lblUID = New System.Windows.Forms.Label
        Me.txtPWD = New System.Windows.Forms.TextBox
        Me.txtUID = New System.Windows.Forms.TextBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.cbxSysDSN = New System.Windows.Forms.CheckBox
        Me.txtSQLSvrName = New System.Windows.Forms.TextBox
        Me.lblSQLServer = New System.Windows.Forms.Label
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtChosenDSN = New System.Windows.Forms.TextBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.btnRemoveDSN = New System.Windows.Forms.Button
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnAddDSN = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.GroupBox9 = New System.Windows.Forms.GroupBox
        Me.btnVerify = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbUIDpwd.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(785, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 526)
        Me.GroupBox1.Size = New System.Drawing.Size(784, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(500, 549)
        Me.cmdOk.TabIndex = 18
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(596, 549)
        Me.cmdCancel.TabIndex = 19
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(692, 549)
        Me.cmdHelp.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.Text = "Add or Remove SQDStudio MetaData DSN"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(72, 21)
        Me.Label2.Size = New System.Drawing.Size(703, 42)
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'txtDSN
        '
        Me.txtDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSN.Location = New System.Drawing.Point(131, 19)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.Size = New System.Drawing.Size(262, 20)
        Me.txtDSN.TabIndex = 1
        '
        'txtDBN
        '
        Me.txtDBN.Location = New System.Drawing.Point(131, 179)
        Me.txtDBN.Name = "txtDBN"
        Me.txtDBN.Size = New System.Drawing.Size(262, 20)
        Me.txtDBN.TabIndex = 6
        '
        'txtDBpath
        '
        Me.txtDBpath.Location = New System.Drawing.Point(9, 153)
        Me.txtDBpath.Name = "txtDBpath"
        Me.txtDBpath.Size = New System.Drawing.Size(350, 20)
        Me.txtDBpath.TabIndex = 5
        '
        'txtDesc
        '
        Me.txtDesc.Location = New System.Drawing.Point(9, 16)
        Me.txtDesc.Multiline = True
        Me.txtDesc.Name = "txtDesc"
        Me.txtDesc.Size = New System.Drawing.Size(384, 58)
        Me.txtDesc.TabIndex = 9
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Location = New System.Drawing.Point(365, 151)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(28, 23)
        Me.cmdBrowse.TabIndex = 4
        Me.cmdBrowse.Text = "..."
        Me.cmdBrowse.UseVisualStyleBackColor = True
        '
        'lblDSNname
        '
        Me.lblDSNname.AutoSize = True
        Me.lblDSNname.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDSNname.Location = New System.Drawing.Point(6, 22)
        Me.lblDSNname.Name = "lblDSNname"
        Me.lblDSNname.Size = New System.Drawing.Size(119, 13)
        Me.lblDSNname.TabIndex = 63
        Me.lblDSNname.Text = "ODBC Name (DSN):"
        '
        'lblDBname
        '
        Me.lblDBname.AutoSize = True
        Me.lblDBname.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBname.Location = New System.Drawing.Point(6, 182)
        Me.lblDBname.Name = "lblDBname"
        Me.lblDBname.Size = New System.Drawing.Size(102, 13)
        Me.lblDBname.TabIndex = 64
        Me.lblDBname.Text = "DataBase Name:"
        '
        'lbltgtLocation
        '
        Me.lbltgtLocation.AutoSize = True
        Me.lbltgtLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbltgtLocation.Location = New System.Drawing.Point(6, 97)
        Me.lbltgtLocation.Name = "lbltgtLocation"
        Me.lbltgtLocation.Size = New System.Drawing.Size(242, 13)
        Me.lbltgtLocation.TabIndex = 65
        Me.lbltgtLocation.Text = "Browse for name and location of .mdb file"
        '
        'lblDBtype
        '
        Me.lblDBtype.AutoSize = True
        Me.lblDBtype.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBtype.Location = New System.Drawing.Point(6, 46)
        Me.lblDBtype.Name = "lblDBtype"
        Me.lblDBtype.Size = New System.Drawing.Size(113, 13)
        Me.lblDBtype.TabIndex = 68
        Me.lblDBtype.Text = "Type of DataBase:"
        '
        'cmbDBtype
        '
        Me.cmbDBtype.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDBtype.FormattingEnabled = True
        Me.cmbDBtype.Location = New System.Drawing.Point(9, 62)
        Me.cmbDBtype.Name = "cmbDBtype"
        Me.cmbDBtype.Size = New System.Drawing.Size(384, 21)
        Me.cmbDBtype.TabIndex = 3
        '
        'gbDesc
        '
        Me.gbDesc.Controls.Add(Me.txtDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.Location = New System.Drawing.Point(12, 342)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(402, 80)
        Me.gbDesc.TabIndex = 70
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "MetaData Description"
        '
        'gbUIDpwd
        '
        Me.gbUIDpwd.Controls.Add(Me.cbxWinAuth)
        Me.gbUIDpwd.Controls.Add(Me.lblPWD)
        Me.gbUIDpwd.Controls.Add(Me.lblUID)
        Me.gbUIDpwd.Controls.Add(Me.txtPWD)
        Me.gbUIDpwd.Controls.Add(Me.txtUID)
        Me.gbUIDpwd.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbUIDpwd.Location = New System.Drawing.Point(12, 428)
        Me.gbUIDpwd.Name = "gbUIDpwd"
        Me.gbUIDpwd.Size = New System.Drawing.Size(240, 96)
        Me.gbUIDpwd.TabIndex = 71
        Me.gbUIDpwd.TabStop = False
        Me.gbUIDpwd.Text = "User Name and Passord if Necessary"
        '
        'cbxWinAuth
        '
        Me.cbxWinAuth.AutoSize = True
        Me.cbxWinAuth.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbxWinAuth.Location = New System.Drawing.Point(6, 70)
        Me.cbxWinAuth.Name = "cbxWinAuth"
        Me.cbxWinAuth.Size = New System.Drawing.Size(196, 17)
        Me.cbxWinAuth.TabIndex = 12
        Me.cbxWinAuth.Text = "Use Windows Authentication?"
        Me.cbxWinAuth.UseVisualStyleBackColor = True
        '
        'lblPWD
        '
        Me.lblPWD.AutoSize = True
        Me.lblPWD.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPWD.Location = New System.Drawing.Point(6, 48)
        Me.lblPWD.Name = "lblPWD"
        Me.lblPWD.Size = New System.Drawing.Size(61, 13)
        Me.lblPWD.TabIndex = 3
        Me.lblPWD.Text = "Password"
        '
        'lblUID
        '
        Me.lblUID.AutoSize = True
        Me.lblUID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUID.Location = New System.Drawing.Point(6, 22)
        Me.lblUID.Name = "lblUID"
        Me.lblUID.Size = New System.Drawing.Size(69, 13)
        Me.lblUID.TabIndex = 2
        Me.lblUID.Text = "User Name"
        '
        'txtPWD
        '
        Me.txtPWD.Location = New System.Drawing.Point(81, 45)
        Me.txtPWD.Name = "txtPWD"
        Me.txtPWD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPWD.Size = New System.Drawing.Size(153, 20)
        Me.txtPWD.TabIndex = 11
        Me.txtPWD.UseSystemPasswordChar = True
        '
        'txtUID
        '
        Me.txtUID.Location = New System.Drawing.Point(81, 19)
        Me.txtUID.Name = "txtUID"
        Me.txtUID.Size = New System.Drawing.Size(153, 20)
        Me.txtUID.TabIndex = 10
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cbxSysDSN)
        Me.GroupBox4.Controls.Add(Me.txtSQLSvrName)
        Me.GroupBox4.Controls.Add(Me.lblSQLServer)
        Me.GroupBox4.Controls.Add(Me.txtDSN)
        Me.GroupBox4.Controls.Add(Me.txtDBN)
        Me.GroupBox4.Controls.Add(Me.cmbDBtype)
        Me.GroupBox4.Controls.Add(Me.lblDSNname)
        Me.GroupBox4.Controls.Add(Me.lblDBname)
        Me.GroupBox4.Controls.Add(Me.lblDBtype)
        Me.GroupBox4.Controls.Add(Me.txtDBpath)
        Me.GroupBox4.Controls.Add(Me.lbltgtLocation)
        Me.GroupBox4.Controls.Add(Me.cmdBrowse)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(12, 81)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(402, 255)
        Me.GroupBox4.TabIndex = 72
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Enter New ODBC Information"
        '
        'cbxSysDSN
        '
        Me.cbxSysDSN.AutoSize = True
        Me.cbxSysDSN.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbxSysDSN.Location = New System.Drawing.Point(63, 232)
        Me.cbxSysDSN.Name = "cbxSysDSN"
        Me.cbxSysDSN.Size = New System.Drawing.Size(274, 17)
        Me.cbxSysDSN.TabIndex = 8
        Me.cbxSysDSN.Text = "Add as System DSN ? (default is User DSN)"
        Me.cbxSysDSN.UseVisualStyleBackColor = True
        '
        'txtSQLSvrName
        '
        Me.txtSQLSvrName.Location = New System.Drawing.Point(131, 205)
        Me.txtSQLSvrName.Name = "txtSQLSvrName"
        Me.txtSQLSvrName.Size = New System.Drawing.Size(262, 20)
        Me.txtSQLSvrName.TabIndex = 7
        '
        'lblSQLServer
        '
        Me.lblSQLServer.AutoSize = True
        Me.lblSQLServer.Location = New System.Drawing.Point(6, 208)
        Me.lblSQLServer.Name = "lblSQLServer"
        Me.lblSQLServer.Size = New System.Drawing.Size(108, 13)
        Me.lblSQLServer.TabIndex = 69
        Me.lblSQLServer.Text = "SQL Server Name"
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox5.Controls.Add(Me.btnRefresh)
        Me.GroupBox5.Controls.Add(Me.DataGridView1)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Controls.Add(Me.txtChosenDSN)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.Location = New System.Drawing.Point(426, 81)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(346, 292)
        Me.GroupBox5.TabIndex = 73
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Current DSN List"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(107, 232)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(128, 23)
        Me.btnRefresh.TabIndex = 15
        Me.btnRefresh.Text = "Refresh Table"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.Control
        Me.DataGridView1.Location = New System.Drawing.Point(6, 16)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(335, 209)
        Me.DataGridView1.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 264)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Chosen DSN"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtChosenDSN
        '
        Me.txtChosenDSN.BackColor = System.Drawing.SystemColors.Info
        Me.txtChosenDSN.Location = New System.Drawing.Point(91, 261)
        Me.txtChosenDSN.Name = "txtChosenDSN"
        Me.txtChosenDSN.Size = New System.Drawing.Size(249, 20)
        Me.txtChosenDSN.TabIndex = 1
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.GroupBox6.Controls.Add(Me.GroupBox8)
        Me.GroupBox6.Controls.Add(Me.btnRemoveDSN)
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox6.Location = New System.Drawing.Point(426, 450)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(347, 74)
        Me.GroupBox6.TabIndex = 74
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "SQDMetaData DSN Removal"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.TextBox2)
        Me.GroupBox8.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.GroupBox8.ForeColor = System.Drawing.Color.Red
        Me.GroupBox8.Location = New System.Drawing.Point(6, 15)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(192, 51)
        Me.GroupBox8.TabIndex = 3
        Me.GroupBox8.TabStop = False
        Me.GroupBox8.Text = "WARNING!!"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.ForeColor = System.Drawing.Color.Red
        Me.TextBox2.Location = New System.Drawing.Point(17, 17)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(161, 29)
        Me.TextBox2.TabIndex = 0
        Me.TextBox2.Text = "This will remove this Metadata DSN from your System"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'btnRemoveDSN
        '
        Me.btnRemoveDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveDSN.Location = New System.Drawing.Point(204, 23)
        Me.btnRemoveDSN.Name = "btnRemoveDSN"
        Me.btnRemoveDSN.Size = New System.Drawing.Size(137, 39)
        Me.btnRemoveDSN.TabIndex = 17
        Me.btnRemoveDSN.Text = "&Remove DSN"
        Me.btnRemoveDSN.UseVisualStyleBackColor = True
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.Label5)
        Me.GroupBox7.Controls.Add(Me.btnAddDSN)
        Me.GroupBox7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox7.Location = New System.Drawing.Point(258, 428)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(156, 96)
        Me.GroupBox7.TabIndex = 75
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Add DSN"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 13)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Add this DSN to ODBC"
        '
        'btnAddDSN
        '
        Me.btnAddDSN.Location = New System.Drawing.Point(10, 49)
        Me.btnAddDSN.Name = "btnAddDSN"
        Me.btnAddDSN.Size = New System.Drawing.Size(137, 39)
        Me.btnAddDSN.TabIndex = 13
        Me.btnAddDSN.Text = "&Add New DSN"
        Me.btnAddDSN.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.CheckFileExists = False
        Me.OpenFileDialog1.SupportMultiDottedExtensions = True
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.CheckBox1)
        Me.GroupBox9.Controls.Add(Me.btnVerify)
        Me.GroupBox9.Controls.Add(Me.Label4)
        Me.GroupBox9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox9.Location = New System.Drawing.Point(427, 379)
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.Size = New System.Drawing.Size(346, 65)
        Me.GroupBox9.TabIndex = 77
        Me.GroupBox9.TabStop = False
        Me.GroupBox9.Text = "Verify All Tables in Chosen DSN"
        '
        'btnVerify
        '
        Me.btnVerify.Location = New System.Drawing.Point(204, 16)
        Me.btnVerify.Name = "btnVerify"
        Me.btnVerify.Size = New System.Drawing.Size(137, 22)
        Me.btnVerify.TabIndex = 16
        Me.btnVerify.Text = "&Verify Tables"
        Me.btnVerify.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 21)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(184, 13)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "To Verify all Tables, Click Here"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(34, 42)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(262, 17)
        Me.CheckBox1.TabIndex = 17
        Me.CheckBox1.Text = "check to view SQL and Connection string"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'frmRegDSN
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(782, 588)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox9)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.gbUIDpwd)
        Me.Controls.Add(Me.gbDesc)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmRegDSN"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ODBC MetaData DSN"
        Me.Controls.SetChildIndex(Me.gbDesc, 0)
        Me.Controls.SetChildIndex(Me.gbUIDpwd, 0)
        Me.Controls.SetChildIndex(Me.GroupBox7, 0)
        Me.Controls.SetChildIndex(Me.GroupBox5, 0)
        Me.Controls.SetChildIndex(Me.GroupBox9, 0)
        Me.Controls.SetChildIndex(Me.GroupBox6, 0)
        Me.Controls.SetChildIndex(Me.GroupBox4, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbUIDpwd.ResumeLayout(False)
        Me.gbUIDpwd.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox8.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.GroupBox9.ResumeLayout(False)
        Me.GroupBox9.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents txtDBN As System.Windows.Forms.TextBox
    Friend WithEvents txtDBpath As System.Windows.Forms.TextBox
    Friend WithEvents txtDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents lblDSNname As System.Windows.Forms.Label
    Friend WithEvents lblDBname As System.Windows.Forms.Label
    Friend WithEvents lbltgtLocation As System.Windows.Forms.Label
    Friend WithEvents lblDBtype As System.Windows.Forms.Label
    Friend WithEvents cmbDBtype As System.Windows.Forms.ComboBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents txtPWD As System.Windows.Forms.TextBox
    Friend WithEvents txtUID As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents lblPWD As System.Windows.Forms.Label
    Friend WithEvents lblUID As System.Windows.Forms.Label
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents btnAddDSN As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnRemoveDSN As System.Windows.Forms.Button
    Friend WithEvents txtChosenDSN As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents GroupBox9 As System.Windows.Forms.GroupBox
    Friend WithEvents btnVerify As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSQLSvrName As System.Windows.Forms.TextBox
    Friend WithEvents lblSQLServer As System.Windows.Forms.Label
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents cbxWinAuth As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cbxSysDSN As System.Windows.Forms.CheckBox
    Friend WithEvents gbUIDpwd As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
End Class
