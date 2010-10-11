Public Class ctlEngine
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    
    Dim objThis As New clsEngine
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmbEngVer As System.Windows.Forms.ComboBox

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
    Friend WithEvents txtDDLLib As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtDTDLib As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIncludeLib As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCopybookLib As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtEngineDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtReportFile As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCommitEvery As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtReportEvery As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEngineName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbLib As System.Windows.Forms.GroupBox
    Friend WithEvents cbConn As System.Windows.Forms.ComboBox
    Friend WithEvents cbForceCommit As System.Windows.Forms.CheckBox
    Friend WithEvents cbDateFormat As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents gbMain As System.Windows.Forms.GroupBox
    Friend WithEvents btnMain As System.Windows.Forms.Button
    Friend WithEvents txtMain As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtDDLLib = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtDTDLib = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtIncludeLib = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtCopybookLib = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtEngineDesc = New System.Windows.Forms.TextBox
        Me.txtReportFile = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtCommitEvery = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtReportEvery = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtEngineName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.cbDateFormat = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cbForceCommit = New System.Windows.Forms.CheckBox
        Me.cbConn = New System.Windows.Forms.ComboBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbLib = New System.Windows.Forms.GroupBox
        Me.gbMain = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btnMain = New System.Windows.Forms.Button
        Me.txtMain = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cmbEngVer = New System.Windows.Forms.ComboBox
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbLib.SuspendLayout()
        Me.gbMain.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(446, 611)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 10
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(526, 611)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 11
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 595)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(670, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'txtDDLLib
        '
        Me.txtDDLLib.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDDLLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDDLLib.Location = New System.Drawing.Point(75, 97)
        Me.txtDDLLib.MaxLength = 255
        Me.txtDDLLib.Name = "txtDDLLib"
        Me.txtDDLLib.Size = New System.Drawing.Size(411, 20)
        Me.txtDDLLib.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 100)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(28, 14)
        Me.Label9.TabIndex = 97
        Me.Label9.Text = "DDL"
        '
        'txtDTDLib
        '
        Me.txtDTDLib.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDTDLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDTDLib.Location = New System.Drawing.Point(75, 71)
        Me.txtDTDLib.MaxLength = 255
        Me.txtDTDLib.Name = "txtDTDLib"
        Me.txtDTDLib.Size = New System.Drawing.Size(411, 20)
        Me.txtDTDLib.TabIndex = 8
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 74)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(28, 14)
        Me.Label10.TabIndex = 95
        Me.Label10.Text = "DTD"
        '
        'txtIncludeLib
        '
        Me.txtIncludeLib.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtIncludeLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIncludeLib.Location = New System.Drawing.Point(75, 45)
        Me.txtIncludeLib.MaxLength = 255
        Me.txtIncludeLib.Name = "txtIncludeLib"
        Me.txtIncludeLib.Size = New System.Drawing.Size(411, 20)
        Me.txtIncludeLib.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(47, 14)
        Me.Label7.TabIndex = 93
        Me.Label7.Text = "Include"
        '
        'txtCopybookLib
        '
        Me.txtCopybookLib.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCopybookLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCopybookLib.Location = New System.Drawing.Point(75, 19)
        Me.txtCopybookLib.MaxLength = 255
        Me.txtCopybookLib.Name = "txtCopybookLib"
        Me.txtCopybookLib.Size = New System.Drawing.Size(411, 20)
        Me.txtCopybookLib.TabIndex = 6
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(6, 22)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(63, 14)
        Me.Label11.TabIndex = 91
        Me.Label11.Text = "Copybook"
        '
        'txtEngineDesc
        '
        Me.txtEngineDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEngineDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEngineDesc.Location = New System.Drawing.Point(9, 19)
        Me.txtEngineDesc.MaxLength = 1000
        Me.txtEngineDesc.Multiline = True
        Me.txtEngineDesc.Name = "txtEngineDesc"
        Me.txtEngineDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEngineDesc.Size = New System.Drawing.Size(349, 51)
        Me.txtEngineDesc.TabIndex = 5
        '
        'txtReportFile
        '
        Me.txtReportFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReportFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReportFile.Location = New System.Drawing.Point(89, 72)
        Me.txtReportFile.MaxLength = 255
        Me.txtReportFile.Name = "txtReportFile"
        Me.txtReportFile.Size = New System.Drawing.Size(282, 20)
        Me.txtReportFile.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 75)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(66, 14)
        Me.Label5.TabIndex = 87
        Me.Label5.Text = "Report File"
        '
        'txtCommitEvery
        '
        Me.txtCommitEvery.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCommitEvery.Location = New System.Drawing.Point(250, 124)
        Me.txtCommitEvery.MaxLength = 128
        Me.txtCommitEvery.Name = "txtCommitEvery"
        Me.txtCommitEvery.Size = New System.Drawing.Size(121, 20)
        Me.txtCommitEvery.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(160, 127)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(84, 14)
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "Commit Every"
        '
        'txtReportEvery
        '
        Me.txtReportEvery.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReportEvery.Location = New System.Drawing.Point(89, 98)
        Me.txtReportEvery.MaxLength = 128
        Me.txtReportEvery.Name = "txtReportEvery"
        Me.txtReportEvery.Size = New System.Drawing.Size(58, 20)
        Me.txtReportEvery.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 101)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 14)
        Me.Label4.TabIndex = 83
        Me.Label4.Text = "Report Every"
        '
        'txtEngineName
        '
        Me.txtEngineName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEngineName.Location = New System.Drawing.Point(89, 19)
        Me.txtEngineName.MaxLength = 128
        Me.txtEngineName.Name = "txtEngineName"
        Me.txtEngineName.Size = New System.Drawing.Size(282, 20)
        Me.txtEngineName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 81
        Me.Label3.Text = "Name"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(606, 611)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 12
        Me.cmdHelp.Text = "&Help"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(8, 595)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(670, 7)
        Me.GroupBox2.TabIndex = 30
        Me.GroupBox2.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(71, 13)
        Me.Label1.TabIndex = 99
        Me.Label1.Text = "Connection"
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.cbDateFormat)
        Me.gbName.Controls.Add(Me.Label2)
        Me.gbName.Controls.Add(Me.cbForceCommit)
        Me.gbName.Controls.Add(Me.txtReportEvery)
        Me.gbName.Controls.Add(Me.Label4)
        Me.gbName.Controls.Add(Me.txtCommitEvery)
        Me.gbName.Controls.Add(Me.Label6)
        Me.gbName.Controls.Add(Me.cbConn)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Controls.Add(Me.Label1)
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.txtEngineName)
        Me.gbName.Controls.Add(Me.Label5)
        Me.gbName.Controls.Add(Me.txtReportFile)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.Color.White
        Me.gbName.Location = New System.Drawing.Point(3, 3)
        Me.gbName.MaximumSize = New System.Drawing.Size(400, 232)
        Me.gbName.MinimumSize = New System.Drawing.Size(346, 232)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(378, 232)
        Me.gbName.TabIndex = 101
        Me.gbName.TabStop = False
        Me.gbName.Text = "Engine Properties"
        '
        'cbDateFormat
        '
        Me.cbDateFormat.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbDateFormat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbDateFormat.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbDateFormat.FormattingEnabled = True
        Me.cbDateFormat.Location = New System.Drawing.Point(250, 97)
        Me.cbDateFormat.Name = "cbDateFormat"
        Me.cbDateFormat.Size = New System.Drawing.Size(121, 21)
        Me.cbDateFormat.TabIndex = 106
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(168, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(76, 13)
        Me.Label2.TabIndex = 105
        Me.Label2.Text = "Date Format"
        '
        'cbForceCommit
        '
        Me.cbForceCommit.AutoSize = True
        Me.cbForceCommit.CheckAlign = System.Drawing.ContentAlignment.BottomRight
        Me.cbForceCommit.Location = New System.Drawing.Point(5, 124)
        Me.cbForceCommit.Name = "cbForceCommit"
        Me.cbForceCommit.Size = New System.Drawing.Size(102, 17)
        Me.cbForceCommit.TabIndex = 104
        Me.cbForceCommit.Text = "Force Commit"
        Me.cbForceCommit.UseVisualStyleBackColor = True
        '
        'cbConn
        '
        Me.cbConn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbConn.FormattingEnabled = True
        Me.cbConn.Location = New System.Drawing.Point(89, 45)
        Me.cbConn.Name = "cbConn"
        Me.cbConn.Size = New System.Drawing.Size(282, 21)
        Me.cbConn.TabIndex = 103
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtEngineDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 146)
        Me.gbDesc.MaximumSize = New System.Drawing.Size(400, 80)
        Me.gbDesc.MinimumSize = New System.Drawing.Size(333, 80)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(365, 80)
        Me.gbDesc.TabIndex = 102
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbLib
        '
        Me.gbLib.Controls.Add(Me.Label11)
        Me.gbLib.Controls.Add(Me.txtCopybookLib)
        Me.gbLib.Controls.Add(Me.txtIncludeLib)
        Me.gbLib.Controls.Add(Me.txtDTDLib)
        Me.gbLib.Controls.Add(Me.txtDDLLib)
        Me.gbLib.Controls.Add(Me.Label9)
        Me.gbLib.Controls.Add(Me.Label7)
        Me.gbLib.Controls.Add(Me.Label10)
        Me.gbLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLib.ForeColor = System.Drawing.Color.White
        Me.gbLib.Location = New System.Drawing.Point(3, 241)
        Me.gbLib.MaximumSize = New System.Drawing.Size(493, 128)
        Me.gbLib.MinimumSize = New System.Drawing.Size(493, 128)
        Me.gbLib.Name = "gbLib"
        Me.gbLib.Size = New System.Drawing.Size(493, 128)
        Me.gbLib.TabIndex = 104
        Me.gbLib.TabStop = False
        Me.gbLib.Text = "Target System Libraries"
        '
        'gbMain
        '
        Me.gbMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbMain.Controls.Add(Me.Label8)
        Me.gbMain.Controls.Add(Me.btnMain)
        Me.gbMain.Controls.Add(Me.txtMain)
        Me.gbMain.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMain.ForeColor = System.Drawing.Color.White
        Me.gbMain.Location = New System.Drawing.Point(4, 475)
        Me.gbMain.Name = "gbMain"
        Me.gbMain.Size = New System.Drawing.Size(674, 117)
        Me.gbMain.TabIndex = 105
        Me.gbMain.TabStop = False
        Me.gbMain.Text = "Main Procedure"
        Me.gbMain.Visible = False
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.Window
        Me.Label8.Location = New System.Drawing.Point(27, 89)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(567, 18)
        Me.Label8.TabIndex = 2
        Me.Label8.Text = "Copy and Paste this into a new Main or General Procedure to use in script"
        '
        'btnMain
        '
        Me.btnMain.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnMain.Enabled = False
        Me.btnMain.ForeColor = System.Drawing.Color.Black
        Me.btnMain.Location = New System.Drawing.Point(8, 88)
        Me.btnMain.Name = "btnMain"
        Me.btnMain.Size = New System.Drawing.Size(152, 23)
        Me.btnMain.TabIndex = 1
        Me.btnMain.Text = "Build Main Procedure"
        Me.btnMain.UseVisualStyleBackColor = True
        Me.btnMain.Visible = False
        '
        'txtMain
        '
        Me.txtMain.AllowDrop = True
        Me.txtMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMain.BackColor = System.Drawing.SystemColors.Window
        Me.txtMain.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMain.Location = New System.Drawing.Point(8, 19)
        Me.txtMain.Multiline = True
        Me.txtMain.Name = "txtMain"
        Me.txtMain.ReadOnly = True
        Me.txtMain.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtMain.Size = New System.Drawing.Size(659, 63)
        Me.txtMain.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmbEngVer)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.White
        Me.GroupBox3.Location = New System.Drawing.Point(3, 375)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(204, 57)
        Me.GroupBox3.TabIndex = 106
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Parser/Engine Version"
        '
        'cmbEngVer
        '
        Me.cmbEngVer.FormattingEnabled = True
        Me.cmbEngVer.Location = New System.Drawing.Point(15, 19)
        Me.cmbEngVer.Name = "cmbEngVer"
        Me.cmbEngVer.Size = New System.Drawing.Size(177, 21)
        Me.cmbEngVer.TabIndex = 0
        '
        'ctlEngine
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.gbMain)
        Me.Controls.Add(Me.gbLib)
        Me.Controls.Add(Me.gbName)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlEngine"
        Me.Size = New System.Drawing.Size(686, 643)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbLib.ResumeLayout(False)
        Me.gbLib.PerformLayout()
        Me.gbMain.ResumeLayout(False)
        Me.gbMain.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtEngineName.Enabled = True '//added on 7/24 to disable object name editing

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsEngine

        ClearControls(Me.Controls)

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    Public Function Save() As Boolean

        Try
            '// First Check Validity before Saving
            If ValidateNewName(txtEngineName.Text, True) = False Then
                Exit Function
            End If

            If objThis.EngineName <> txtEngineName.Text Then
                objThis.IsRenamed = RenameEngine(objThis, txtEngineName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtEngineName.Text = objThis.EngineName
            Else
                objThis.EngineName = txtEngineName.Text
            End If

            objThis.EngineDescription = txtEngineDesc.Text
            objThis.ReportFile = txtReportFile.Text
            objThis.ReportEvery = txtReportEvery.Text
            objThis.CommitEvery = txtCommitEvery.Text
            objThis.CopybookLib = txtCopybookLib.Text
            objThis.IncludeLib = txtIncludeLib.Text
            objThis.DTDLib = txtDTDLib.Text
            objThis.DDLLib = txtDDLLib.Text
            objThis.Main = txtMain.Text
            objThis.EngVersion = cmbEngVer.Text

            Dim temp As Mylist = cbConn.SelectedItem

            If temp Is Nothing Then
                temp = New Mylist("None", "None")
            End If
            If temp.Name = "None" Then
                objThis.Connection = Nothing
            Else
                For Each conn As clsConnection In objThis.ObjSystem.Environment.Connections
                    If conn.ConnectionName = temp.Name Then
                        objThis.Connection = conn
                        Exit For
                    End If
                    objThis.Connection = Nothing
                Next
            End If

            Dim Temp2 As Mylist = cbDateFormat.SelectedItem

            If Temp2.Name.Trim = "" Then
                objThis.DateFormat = " "
            Else
                objThis.DateFormat = Temp2.Name
            End If

            If IsNewObj = True Then
                If objThis.AddNew = False Then Exit Function
            Else
                objThis.Save()
            End If

            'If objThis.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    objThis.SaveEngProps()
            'End If

            Save = True
            cmdSave.Enabled = False
            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If
            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex, "ctlEngine Save")
            Save = False
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
                    objThis.IsModified = False
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

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Engines)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCommitEvery.KeyDown, txtCopybookLib.KeyDown, txtDDLLib.KeyDown, txtDTDLib.KeyDown, txtEngineDesc.KeyDown, txtEngineName.KeyDown, txtReportEvery.KeyDown, txtReportFile.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    Public Sub MyCTL_KeyDown_Main(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMain.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                ShowHelp(modHelp.HHId.H_Main_Procedure)
        End Select
    End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCopybookLib.TextChanged, txtDDLLib.TextChanged, txtDTDLib.TextChanged, txtIncludeLib.TextChanged, txtCommitEvery.TextChanged, txtEngineDesc.TextChanged, txtEngineName.TextChanged, txtReportEvery.TextChanged, txtReportFile.TextChanged, cbConn.SelectedIndexChanged, cbDateFormat.SelectedIndexChanged, txtMain.TextChanged, cmbEngVer.SelectedIndexChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Public Function EditObj(ByVal obj As INode) As clsEngine

        IsNewObj = False

        StartLoad()

        objThis = obj '//Load the form env object
        objThis.LoadMe()

        UpdateFields()

        EndLoad()

        EditObj = objThis

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        Try
            txtEngineName.Text = objThis.EngineName
            txtReportEvery.Text = objThis.ReportEvery
            txtCommitEvery.Text = objThis.CommitEvery
            txtReportFile.Text = objThis.ReportFile
            txtEngineDesc.Text = objThis.EngineDescription
            txtCopybookLib.Text = objThis.CopybookLib
            txtIncludeLib.Text = objThis.IncludeLib
            txtDTDLib.Text = objThis.DTDLib
            txtDDLLib.Text = objThis.DDLLib
            txtMain.Text = objThis.Main

            'If objThis.Main.Trim = "" Then
            gbMain.Visible = False
            'Else
            'gbMain.Visible = True
            'End If

            SetComboConn()
            setComboDate()
            setEngVer()

            If objThis.ForceCommit = True Then
                cbForceCommit.Checked = True
            Else
                cbForceCommit.Checked = False
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlEngine UpdateFields")
            Return False
        End Try

    End Function

    Sub SetComboConn()

        Try
            cbConn.Items.Clear()

            cbConn.Items.Add(New Mylist("None", "none")) ' New clsConnection))
            For Each conn As clsConnection In objThis.ObjSystem.Environment.Connections
                If conn.ConnectionName <> "" Then
                    cbConn.Items.Add(New Mylist(conn.ConnectionName, conn.ConnectionName))
                End If
            Next

            If objThis.Connection IsNot Nothing Then
                SetListItemByValue(cbConn, objThis.Connection.ConnectionName, False)
            Else
                cbConn.SelectedIndex = 0
            End If

        Catch ex As Exception
            LogError(ex, "ctlEngine setcomboconn")
        End Try

    End Sub

    Sub setComboDate()

        Try
            cbDateFormat.Items.Clear()

            cbDateFormat.Items.Add(New Mylist(" ", " "))
            cbDateFormat.Items.Add(New Mylist("ISO", "ISO"))
            cbDateFormat.Items.Add(New Mylist("YYYYMMDD", "YYYYMMDD"))

            If objThis.DateFormat.Trim <> "" Then
                SetListItemByValue(cbDateFormat, objThis.DateFormat, False)
            Else
                cbDateFormat.SelectedIndex = 0
            End If

        Catch ex As Exception
            LogError(ex, "ctlEngine setcomboDateFormat")
        End Try

    End Sub

    Sub setEngVer()

        Try
            cmbEngVer.Items.Clear()

            cmbEngVer.Items.Add(New Mylist("3.7.12", "3.7.12"))
            cmbEngVer.Items.Add(New Mylist("3.7.7", "3.7.7"))
            cmbEngVer.Items.Add(New Mylist("3.7.6", "3.7.6"))
            cmbEngVer.Items.Add(New Mylist("3.6.14", "3.6.14"))

            If objThis.EngVersion.Trim <> "" Then
                SetListItemByValue(cmbEngVer, objThis.EngVersion, False)
            Else
                cmbEngVer.SelectedIndex = 0
            End If

        Catch ex As Exception
            LogError(ex, "ctlEngine setEngVer")
        End Try

    End Sub

    Private Sub cbForceCommit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbForceCommit.CheckedChanged

        If IsEventFromCode = True Then Exit Sub
        If cbForceCommit.Checked = True Then
            If txtCommitEvery.Text = "" Then
                MsgBox("You have chosen to Force Commit." & Chr(13) & _
                "A value must be entered for ""Commit Every""", MsgBoxStyle.OkOnly)
            End If
            objThis.ForceCommit = True
        Else
            objThis.ForceCommit = False
        End If
        objThis.IsModified = True
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub btnMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMain.Click

        txtMain.Text = GetMainText()

    End Sub

    Function insertParens() As Boolean

        Try

        Catch ex As Exception

        End Try

    End Function

    Function removeParens() As Boolean

        Try

        Catch ex As Exception

        End Try

    End Function

    Private Function GetMainText() As String

        Dim sb As New System.Text.StringBuilder
        Dim sel As clsDSSelection
        'Dim i As Integer
        Dim taskName As String
        Dim sDs As clsDatastore = Nothing
        Dim nDs As clsDatastore = Nothing
        Dim count As Integer = 0
        Dim SrcCount As Integer = 0

        Try
            While SrcCount < objThis.Sources.Count

                sDs = objThis.Sources(SrcCount + 1)


                If SrcCount > 0 Then
                    nDs = objThis.Sources(SrcCount)
                    If sDs.IsLookUp = False Then
                        sb.AppendLine("}")   '& inDsNode.Text
                        sb.AppendLine("FROM " & nDs.DatastoreName)
                        sb.AppendLine("UNION")
                        sb.AppendLine("{")
                    End If
                End If
                SrcCount += 1

                'sb.Append("CASE" & vbCrLf)
                'sb.Append("(" & vbCrLf)

                'For i = 0 To sDs.ObjSelections.Count - 1
                '    sel = sDs.ObjSelections(i)
                '    taskName = "P_" & sel.Text
                '    sb.Append(vbTab & IIf(i = 0, "", ",") & "EQ(RECNAME(" & sDs.Text & "),'" & sel.Text & "')" & vbCrLf)
                '    sb.Append(vbTab & ",CALLPROC(" & taskName & ")" & vbCrLf)
                '    sb.Append(vbCrLf)
                'Next

                'sb.Append(")" & vbCrLf)

                'sb.AppendLine("{")
                If sDs.IsLookUp = False Then
                    sb.AppendLine("CASE RECNAME(" & sDs.Text & ")")
                    For Each sel In sDs.ObjSelections
                        'count += 1
                        'If count <= objThis.Procs.Count Then
                        '    taskName = CType(objThis.Procs(count), clsTask).TaskName
                        'Else
                        taskName = ""
                        'End If
                        For Each proc As clsTask In objThis.Procs
                            If proc.TaskName.Contains(sel.Text) = True Then
                                taskName = proc.TaskName
                                Exit For
                            End If
                        Next

                        sb.AppendLine(TAB & "WHEN '" & sel.Text & "'") '& IIf(nd.Index = 0, "", ",")
                        sb.AppendLine(TAB & "DO")
                        sb.AppendLine(TAB & TAB & "CALLPROC(" & taskName & ")")
                        sb.AppendLine(TAB & "END") '& vbCrLf)
                        'sb.Append(vbCrLf)
                    Next
                End If

                'sb.Append(")" & vbCrLf)
                'sb.AppendLine("}")   '& inDsNode.Text
                'sb.AppendLine("FROM " & sDs.DatastoreName)
            End While

            GetMainText = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlEngine GetMainText")
            GetMainText = ""
        End Try

    End Function

#Region "DragDrop"

    Private Sub txtCodeEditor_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtMain.DragEnter

        'See if there is a TreeNode or text being dragged
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Or e.Data.GetDataPresent(DataFormats.Text) Then
            'TreeNode found allow copy effect
            e.Effect = e.AllowedEffect
        End If

    End Sub

    Private Sub txtCodeEditor_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtMain.DragOver

        Dim pt As Point
        Dim ret As Integer

        pt = New Point(e.X, e.Y)
        pt = txtMain.PointToClient(pt)
        txtMain.Focus()
        ret = GetCharFromPos(txtMain, pt)
        txtMain.Select(ret, 0)

    End Sub

    Private Sub txtCodeEditor_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtMain.DragDrop

        Dim prevSel As Integer
        Dim txt As String
        Dim obj As INode

        Try
            prevSel = txtMain.SelectionStart
            If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
                Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
                '//Only accept nodes with INOde interface
                If Not (draggedNode.Tag.GetType.GetInterface("INode") Is Nothing) Then
                    obj = draggedNode.Tag
                    txt = obj.Text
                    txtMain.Text = txtMain.Text.Insert(prevSel, txt)
                    txtMain.SelectionStart = prevSel
                End If
            ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
                txtMain.Text = txtMain.Text.Insert(prevSel, e.Data.GetData(DataFormats.Text))
                txtMain.SelectionStart = prevSel
            End If

        Catch ex As Exception
            LogError(ex, "ctlTask DragDrop")
        End Try

    End Sub

#End Region

End Class
