<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSelectDSN
    Inherits SQStudio.frmBlank

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.gbDSNdatagrid = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.gbODBC = New System.Windows.Forms.GroupBox
        Me.gbProject = New System.Windows.Forms.GroupBox
        Me.txtSchema = New System.Windows.Forms.TextBox
        Me.cbLoginNo = New System.Windows.Forms.CheckBox
        Me.cbLoginYes = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtODBCtype = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDSN = New System.Windows.Forms.TextBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.gbDSNdatagrid.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbODBC.SuspendLayout()
        Me.gbProject.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(472, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 481)
        Me.GroupBox1.Size = New System.Drawing.Size(474, 7)
        Me.GroupBox1.TabIndex = 44
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(190, 504)
        Me.cmdOk.TabIndex = 10
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(286, 504)
        Me.cmdCancel.TabIndex = 11
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(382, 504)
        Me.cmdHelp.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.Size = New System.Drawing.Size(390, 16)
        Me.Label1.TabIndex = 30
        Me.Label1.Text = "Select ODBC Source"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(377, 37)
        Me.Label2.TabIndex = 31
        Me.Label2.Text = "Choose an ODBC Source to retrieve table names from to be Modeled. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Supported typ" & _
            "es are: IBM DB2 and Oracle."
        '
        'gbDSNdatagrid
        '
        Me.gbDSNdatagrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDSNdatagrid.Controls.Add(Me.DataGridView1)
        Me.gbDSNdatagrid.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDSNdatagrid.Location = New System.Drawing.Point(12, 73)
        Me.gbDSNdatagrid.Name = "gbDSNdatagrid"
        Me.gbDSNdatagrid.Size = New System.Drawing.Size(448, 244)
        Me.gbDSNdatagrid.TabIndex = 32
        Me.gbDSNdatagrid.TabStop = False
        Me.gbDSNdatagrid.Text = "Choose an ODBC Source"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.NullValue = Nothing
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ActiveCaptionText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.DataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DataGridView1.GridColor = System.Drawing.SystemColors.InactiveCaption
        Me.DataGridView1.Location = New System.Drawing.Point(3, 16)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.ShowCellErrors = False
        Me.DataGridView1.ShowCellToolTips = False
        Me.DataGridView1.ShowEditingIcon = False
        Me.DataGridView1.ShowRowErrors = False
        Me.DataGridView1.Size = New System.Drawing.Size(442, 225)
        Me.DataGridView1.StandardTab = True
        Me.DataGridView1.TabIndex = 1
        '
        'gbODBC
        '
        Me.gbODBC.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbODBC.Controls.Add(Me.gbProject)
        Me.gbODBC.Controls.Add(Me.Label6)
        Me.gbODBC.Controls.Add(Me.txtODBCtype)
        Me.gbODBC.Controls.Add(Me.Label5)
        Me.gbODBC.Controls.Add(Me.txtDSN)
        Me.gbODBC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbODBC.Location = New System.Drawing.Point(12, 323)
        Me.gbODBC.Name = "gbODBC"
        Me.gbODBC.Size = New System.Drawing.Size(446, 153)
        Me.gbODBC.TabIndex = 33
        Me.gbODBC.TabStop = False
        Me.gbODBC.Text = "Selected ODBC Source"
        '
        'gbProject
        '
        Me.gbProject.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbProject.Controls.Add(Me.txtSchema)
        Me.gbProject.Controls.Add(Me.cbLoginNo)
        Me.gbProject.Controls.Add(Me.cbLoginYes)
        Me.gbProject.Controls.Add(Me.Label4)
        Me.gbProject.Controls.Add(Me.Label9)
        Me.gbProject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProject.Location = New System.Drawing.Point(11, 71)
        Me.gbProject.Name = "gbProject"
        Me.gbProject.Size = New System.Drawing.Size(423, 69)
        Me.gbProject.TabIndex = 40
        Me.gbProject.TabStop = False
        Me.gbProject.Text = "Database Login properties"
        '
        'txtSchema
        '
        Me.txtSchema.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSchema.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSchema.Location = New System.Drawing.Point(127, 17)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.Size = New System.Drawing.Size(290, 20)
        Me.txtSchema.TabIndex = 2
        '
        'cbLoginNo
        '
        Me.cbLoginNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginNo.AutoSize = True
        Me.cbLoginNo.Location = New System.Drawing.Point(375, 44)
        Me.cbLoginNo.Name = "cbLoginNo"
        Me.cbLoginNo.Size = New System.Drawing.Size(42, 17)
        Me.cbLoginNo.TabIndex = 4
        Me.cbLoginNo.Text = "No"
        Me.cbLoginNo.UseVisualStyleBackColor = True
        '
        'cbLoginYes
        '
        Me.cbLoginYes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginYes.AutoSize = True
        Me.cbLoginYes.Checked = True
        Me.cbLoginYes.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbLoginYes.Location = New System.Drawing.Point(322, 44)
        Me.cbLoginYes.Name = "cbLoginYes"
        Me.cbLoginYes.Size = New System.Drawing.Size(47, 17)
        Me.cbLoginYes.TabIndex = 3
        Me.cbLoginYes.Text = "Yes"
        Me.cbLoginYes.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(110, 13)
        Me.Label4.TabIndex = 41
        Me.Label4.Text = "Database Schema"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 45)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(312, 13)
        Me.Label9.TabIndex = 42
        Me.Label9.Text = "Database requires User Login? (Username, Password)"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(93, 13)
        Me.Label6.TabIndex = 36
        Me.Label6.Text = "Database Type"
        '
        'txtODBCtype
        '
        Me.txtODBCtype.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtODBCtype.BackColor = System.Drawing.SystemColors.Info
        Me.txtODBCtype.Location = New System.Drawing.Point(191, 45)
        Me.txtODBCtype.Name = "txtODBCtype"
        Me.txtODBCtype.Size = New System.Drawing.Size(241, 20)
        Me.txtODBCtype.TabIndex = 37
        Me.txtODBCtype.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(179, 13)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Currently Selected ODBC DSN"
        '
        'txtDSN
        '
        Me.txtDSN.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDSN.BackColor = System.Drawing.SystemColors.Info
        Me.txtDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSN.Location = New System.Drawing.Point(191, 19)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.ReadOnly = True
        Me.txtDSN.Size = New System.Drawing.Size(241, 20)
        Me.txtDSN.TabIndex = 35
        Me.txtDSN.TabStop = False
        '
        'frmSelectDSN
        '
        Me.AutoSize = True
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(472, 543)
        Me.Controls.Add(Me.gbODBC)
        Me.Controls.Add(Me.gbDSNdatagrid)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MinimumSize = New System.Drawing.Size(482, 445)
        Me.Name = "frmSelectDSN"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select ODBC Source"
        Me.Controls.SetChildIndex(Me.gbDSNdatagrid, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.gbODBC, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.gbDSNdatagrid.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbODBC.ResumeLayout(False)
        Me.gbODBC.PerformLayout()
        Me.gbProject.ResumeLayout(False)
        Me.gbProject.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbDSNdatagrid As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents gbODBC As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtODBCtype As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents gbProject As System.Windows.Forms.GroupBox
    Friend WithEvents txtSchema As System.Windows.Forms.TextBox
    Friend WithEvents cbLoginNo As System.Windows.Forms.CheckBox
    Friend WithEvents cbLoginYes As System.Windows.Forms.CheckBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label

End Class
