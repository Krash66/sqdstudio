<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlMain
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.scTask = New System.Windows.Forms.SplitContainer
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.txtTaskName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtTaskDesc = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.scSrc = New System.Windows.Forms.SplitContainer
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabSrc = New System.Windows.Forms.TabPage
        Me.gbSrc = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tvSource = New System.Windows.Forms.TreeView
        Me.TabFunct = New System.Windows.Forms.TabPage
        Me.gbFun = New System.Windows.Forms.GroupBox
        Me.pnlFunctionInfo = New System.Windows.Forms.Panel
        Me.lblFunSyntax = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.lblFunDesc = New System.Windows.Forms.Label
        Me.lblFunName = New System.Windows.Forms.Label
        Me.tvFunctions = New System.Windows.Forms.TreeView
        Me.gbMap = New System.Windows.Forms.GroupBox
        Me.pnlScript = New System.Windows.Forms.Panel
        Me.txtCodeEditor = New System.Windows.Forms.TextBox
        Me.cbIsMain = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtColNum = New System.Windows.Forms.TextBox
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.scTask.Panel1.SuspendLayout()
        Me.scTask.Panel2.SuspendLayout()
        Me.scTask.SuspendLayout()
        Me.gbName.SuspendLayout()
        Me.scSrc.Panel1.SuspendLayout()
        Me.scSrc.Panel2.SuspendLayout()
        Me.scSrc.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabSrc.SuspendLayout()
        Me.gbSrc.SuspendLayout()
        Me.TabFunct.SuspendLayout()
        Me.gbFun.SuspendLayout()
        Me.pnlFunctionInfo.SuspendLayout()
        Me.gbMap.SuspendLayout()
        Me.pnlScript.SuspendLayout()
        Me.SuspendLayout()
        '
        'scTask
        '
        Me.scTask.AllowDrop = True
        Me.scTask.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scTask.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.scTask.Location = New System.Drawing.Point(4, 3)
        Me.scTask.Name = "scTask"
        Me.scTask.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scTask.Panel1
        '
        Me.scTask.Panel1.Controls.Add(Me.gbName)
        '
        'scTask.Panel2
        '
        Me.scTask.Panel2.Controls.Add(Me.scSrc)
        Me.scTask.Size = New System.Drawing.Size(867, 575)
        Me.scTask.SplitterDistance = 69
        Me.scTask.TabIndex = 96
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.txtTaskName)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Controls.Add(Me.txtTaskDesc)
        Me.gbName.Controls.Add(Me.Label2)
        Me.gbName.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.Location = New System.Drawing.Point(0, 0)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(867, 69)
        Me.gbName.TabIndex = 90
        Me.gbName.TabStop = False
        '
        'txtTaskName
        '
        Me.txtTaskName.Location = New System.Drawing.Point(50, 13)
        Me.txtTaskName.MaxLength = 20
        Me.txtTaskName.Name = "txtTaskName"
        Me.txtTaskName.Size = New System.Drawing.Size(174, 20)
        Me.txtTaskName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.Window
        Me.Label3.Location = New System.Drawing.Point(6, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 63
        Me.Label3.Text = "Name"
        '
        'txtTaskDesc
        '
        Me.txtTaskDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTaskDesc.Location = New System.Drawing.Point(307, 13)
        Me.txtTaskDesc.Multiline = True
        Me.txtTaskDesc.Name = "txtTaskDesc"
        Me.txtTaskDesc.Size = New System.Drawing.Size(554, 51)
        Me.txtTaskDesc.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.Window
        Me.Label2.Location = New System.Drawing.Point(230, 17)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 89
        Me.Label2.Text = "Description"
        '
        'scSrc
        '
        Me.scSrc.AllowDrop = True
        Me.scSrc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scSrc.Location = New System.Drawing.Point(0, 0)
        Me.scSrc.Name = "scSrc"
        '
        'scSrc.Panel1
        '
        Me.scSrc.Panel1.Controls.Add(Me.TabControl1)
        '
        'scSrc.Panel2
        '
        Me.scSrc.Panel2.Controls.Add(Me.gbMap)
        Me.scSrc.Size = New System.Drawing.Size(867, 502)
        Me.scSrc.SplitterDistance = 183
        Me.scSrc.TabIndex = 0
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabSrc)
        Me.TabControl1.Controls.Add(Me.TabFunct)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.HotTrack = True
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(183, 502)
        Me.TabControl1.TabIndex = 94
        '
        'TabSrc
        '
        Me.TabSrc.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.TabSrc.Controls.Add(Me.gbSrc)
        Me.TabSrc.Location = New System.Drawing.Point(4, 22)
        Me.TabSrc.Name = "TabSrc"
        Me.TabSrc.Padding = New System.Windows.Forms.Padding(3)
        Me.TabSrc.Size = New System.Drawing.Size(175, 476)
        Me.TabSrc.TabIndex = 0
        Me.TabSrc.Text = "Sources"
        Me.TabSrc.UseVisualStyleBackColor = True
        '
        'gbSrc
        '
        Me.gbSrc.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.gbSrc.Controls.Add(Me.Label4)
        Me.gbSrc.Controls.Add(Me.tvSource)
        Me.gbSrc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbSrc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSrc.Location = New System.Drawing.Point(3, 3)
        Me.gbSrc.Name = "gbSrc"
        Me.gbSrc.Size = New System.Drawing.Size(169, 470)
        Me.gbSrc.TabIndex = 93
        Me.gbSrc.TabStop = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(6, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(102, 16)
        Me.Label4.TabIndex = 69
        Me.Label4.Text = "Select Source(s)"
        '
        'tvSource
        '
        Me.tvSource.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvSource.CheckBoxes = True
        Me.tvSource.HideSelection = False
        Me.tvSource.HotTracking = True
        Me.tvSource.Location = New System.Drawing.Point(3, 28)
        Me.tvSource.Name = "tvSource"
        Me.tvSource.Size = New System.Drawing.Size(163, 439)
        Me.tvSource.TabIndex = 3
        '
        'TabFunct
        '
        Me.TabFunct.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.TabFunct.Controls.Add(Me.gbFun)
        Me.TabFunct.Location = New System.Drawing.Point(4, 22)
        Me.TabFunct.Name = "TabFunct"
        Me.TabFunct.Padding = New System.Windows.Forms.Padding(3)
        Me.TabFunct.Size = New System.Drawing.Size(175, 476)
        Me.TabFunct.TabIndex = 1
        Me.TabFunct.Text = "Functions"
        Me.TabFunct.UseVisualStyleBackColor = True
        '
        'gbFun
        '
        Me.gbFun.Controls.Add(Me.pnlFunctionInfo)
        Me.gbFun.Controls.Add(Me.tvFunctions)
        Me.gbFun.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFun.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFun.Location = New System.Drawing.Point(3, 3)
        Me.gbFun.Name = "gbFun"
        Me.gbFun.Size = New System.Drawing.Size(169, 470)
        Me.gbFun.TabIndex = 91
        Me.gbFun.TabStop = False
        '
        'pnlFunctionInfo
        '
        Me.pnlFunctionInfo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlFunctionInfo.AutoScroll = True
        Me.pnlFunctionInfo.BackColor = System.Drawing.SystemColors.ControlLight
        Me.pnlFunctionInfo.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlFunctionInfo.Controls.Add(Me.lblFunSyntax)
        Me.pnlFunctionInfo.Controls.Add(Me.Label9)
        Me.pnlFunctionInfo.Controls.Add(Me.Label8)
        Me.pnlFunctionInfo.Controls.Add(Me.Label7)
        Me.pnlFunctionInfo.Controls.Add(Me.Label11)
        Me.pnlFunctionInfo.Controls.Add(Me.lblFunDesc)
        Me.pnlFunctionInfo.Controls.Add(Me.lblFunName)
        Me.pnlFunctionInfo.Location = New System.Drawing.Point(0, 344)
        Me.pnlFunctionInfo.Name = "pnlFunctionInfo"
        Me.pnlFunctionInfo.Size = New System.Drawing.Size(169, 123)
        Me.pnlFunctionInfo.TabIndex = 1
        '
        'lblFunSyntax
        '
        Me.lblFunSyntax.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFunSyntax.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFunSyntax.ForeColor = System.Drawing.Color.Blue
        Me.lblFunSyntax.Location = New System.Drawing.Point(3, 33)
        Me.lblFunSyntax.Name = "lblFunSyntax"
        Me.lblFunSyntax.Size = New System.Drawing.Size(164, 19)
        Me.lblFunSyntax.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.Black
        Me.Label9.Location = New System.Drawing.Point(1, 52)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(79, 16)
        Me.Label9.TabIndex = 2
        Me.Label9.Text = "Description :"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.Black
        Me.Label8.Location = New System.Drawing.Point(3, 17)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(53, 16)
        Me.Label8.TabIndex = 1
        Me.Label8.Text = "Syntax :"
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(3, 1)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(46, 16)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Name :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.Blue
        Me.Label11.Location = New System.Drawing.Point(56, 8)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(0, 14)
        Me.Label11.TabIndex = 1
        '
        'lblFunDesc
        '
        Me.lblFunDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFunDesc.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFunDesc.ForeColor = System.Drawing.Color.Maroon
        Me.lblFunDesc.Location = New System.Drawing.Point(1, 68)
        Me.lblFunDesc.Name = "lblFunDesc"
        Me.lblFunDesc.Size = New System.Drawing.Size(166, 53)
        Me.lblFunDesc.TabIndex = 1
        Me.lblFunDesc.Text = "<none>"
        '
        'lblFunName
        '
        Me.lblFunName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFunName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFunName.ForeColor = System.Drawing.Color.Blue
        Me.lblFunName.Location = New System.Drawing.Point(55, 3)
        Me.lblFunName.Name = "lblFunName"
        Me.lblFunName.Size = New System.Drawing.Size(109, 17)
        Me.lblFunName.TabIndex = 1
        '
        'tvFunctions
        '
        Me.tvFunctions.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvFunctions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold)
        Me.tvFunctions.HideSelection = False
        Me.tvFunctions.Location = New System.Drawing.Point(2, 5)
        Me.tvFunctions.Name = "tvFunctions"
        Me.tvFunctions.Size = New System.Drawing.Size(164, 338)
        Me.tvFunctions.TabIndex = 7
        '
        'gbMap
        '
        Me.gbMap.Controls.Add(Me.pnlScript)
        Me.gbMap.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbMap.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMap.ForeColor = System.Drawing.SystemColors.Window
        Me.gbMap.Location = New System.Drawing.Point(0, 0)
        Me.gbMap.Name = "gbMap"
        Me.gbMap.Size = New System.Drawing.Size(680, 502)
        Me.gbMap.TabIndex = 92
        Me.gbMap.TabStop = False
        Me.gbMap.Text = "Script View"
        '
        'pnlScript
        '
        Me.pnlScript.AllowDrop = True
        Me.pnlScript.Controls.Add(Me.txtCodeEditor)
        Me.pnlScript.Controls.Add(Me.cbIsMain)
        Me.pnlScript.Controls.Add(Me.Label1)
        Me.pnlScript.Controls.Add(Me.txtColNum)
        Me.pnlScript.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlScript.Location = New System.Drawing.Point(3, 16)
        Me.pnlScript.Name = "pnlScript"
        Me.pnlScript.Size = New System.Drawing.Size(674, 483)
        Me.pnlScript.TabIndex = 53
        '
        'txtCodeEditor
        '
        Me.txtCodeEditor.AcceptsReturn = True
        Me.txtCodeEditor.AcceptsTab = True
        Me.txtCodeEditor.AllowDrop = True
        Me.txtCodeEditor.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCodeEditor.Location = New System.Drawing.Point(3, 29)
        Me.txtCodeEditor.Multiline = True
        Me.txtCodeEditor.Name = "txtCodeEditor"
        Me.txtCodeEditor.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtCodeEditor.Size = New System.Drawing.Size(668, 451)
        Me.txtCodeEditor.TabIndex = 6
        Me.txtCodeEditor.WordWrap = False
        '
        'cbIsMain
        '
        Me.cbIsMain.AutoSize = True
        Me.cbIsMain.Location = New System.Drawing.Point(168, 5)
        Me.cbIsMain.Name = "cbIsMain"
        Me.cbIsMain.Size = New System.Drawing.Size(204, 17)
        Me.cbIsMain.TabIndex = 5
        Me.cbIsMain.Text = "This is Main Routing Procedure"
        Me.cbIsMain.UseVisualStyleBackColor = True
        Me.cbIsMain.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 6)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(74, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Line Length"
        '
        'txtColNum
        '
        Me.txtColNum.Location = New System.Drawing.Point(83, 3)
        Me.txtColNum.Name = "txtColNum"
        Me.txtColNum.Size = New System.Drawing.Size(65, 20)
        Me.txtColNum.TabIndex = 2
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(643, 584)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 97
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(721, 584)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 98
        Me.cmdClose.Text = "&Close"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(799, 584)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 11
        Me.cmdHelp.Text = "&Help"
        '
        'ctlMain
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.scTask)
        Me.Name = "ctlMain"
        Me.Size = New System.Drawing.Size(874, 611)
        Me.scTask.Panel1.ResumeLayout(False)
        Me.scTask.Panel2.ResumeLayout(False)
        Me.scTask.ResumeLayout(False)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.scSrc.Panel1.ResumeLayout(False)
        Me.scSrc.Panel2.ResumeLayout(False)
        Me.scSrc.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabSrc.ResumeLayout(False)
        Me.gbSrc.ResumeLayout(False)
        Me.TabFunct.ResumeLayout(False)
        Me.gbFun.ResumeLayout(False)
        Me.pnlFunctionInfo.ResumeLayout(False)
        Me.pnlFunctionInfo.PerformLayout()
        Me.gbMap.ResumeLayout(False)
        Me.pnlScript.ResumeLayout(False)
        Me.pnlScript.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents scTask As System.Windows.Forms.SplitContainer
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents txtTaskName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtTaskDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents scSrc As System.Windows.Forms.SplitContainer
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabSrc As System.Windows.Forms.TabPage
    Friend WithEvents gbSrc As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tvSource As System.Windows.Forms.TreeView
    Friend WithEvents TabFunct As System.Windows.Forms.TabPage
    Friend WithEvents gbFun As System.Windows.Forms.GroupBox
    Friend WithEvents pnlFunctionInfo As System.Windows.Forms.Panel
    Friend WithEvents lblFunSyntax As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Private WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents lblFunDesc As System.Windows.Forms.Label
    Friend WithEvents lblFunName As System.Windows.Forms.Label
    Friend WithEvents tvFunctions As System.Windows.Forms.TreeView
    Friend WithEvents gbMap As System.Windows.Forms.GroupBox
    Friend WithEvents pnlScript As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtColNum As System.Windows.Forms.TextBox
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents cbIsMain As System.Windows.Forms.CheckBox
    Friend WithEvents txtCodeEditor As System.Windows.Forms.TextBox

End Class
