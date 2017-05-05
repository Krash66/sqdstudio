<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewProj
    Inherits SQDStudio.frmBlank

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
        Me.gbODBC = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtODBCtype = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnDSN = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDSN = New System.Windows.Forms.TextBox
        Me.gbProject = New System.Windows.Forms.GroupBox
        Me.txtSchema = New System.Windows.Forms.TextBox
        Me.cbLoginNo = New System.Windows.Forms.CheckBox
        Me.cbLoginYes = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtProj = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.gbODBC.SuspendLayout()
        Me.gbProject.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(467, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 285)
        Me.GroupBox1.Size = New System.Drawing.Size(469, 7)
        Me.GroupBox1.TabIndex = 11
        '
        'cmdOk
        '
        Me.cmdOk.Enabled = False
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(185, 308)
        Me.cmdOk.TabIndex = 19
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(281, 308)
        Me.cmdCancel.TabIndex = 20
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(377, 308)
        Me.cmdHelp.TabIndex = 21
        '
        'Label1
        '
        Me.Label1.Size = New System.Drawing.Size(385, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Create New Project"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(391, 39)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Use the current or choose a different ODBC Metadata Source, then type in a new Pr" & _
            "oject Name to create a new project."
        '
        'gbODBC
        '
        Me.gbODBC.Controls.Add(Me.Label6)
        Me.gbODBC.Controls.Add(Me.txtODBCtype)
        Me.gbODBC.Controls.Add(Me.Label5)
        Me.gbODBC.Controls.Add(Me.btnDSN)
        Me.gbODBC.Controls.Add(Me.Label3)
        Me.gbODBC.Controls.Add(Me.txtDSN)
        Me.gbODBC.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbODBC.Location = New System.Drawing.Point(12, 74)
        Me.gbODBC.Name = "gbODBC"
        Me.gbODBC.Size = New System.Drawing.Size(443, 104)
        Me.gbODBC.TabIndex = 3
        Me.gbODBC.TabStop = False
        Me.gbODBC.Text = "Selected ODBC Source"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(94, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "MetaData Type"
        '
        'txtODBCtype
        '
        Me.txtODBCtype.BackColor = System.Drawing.SystemColors.Window
        Me.txtODBCtype.Location = New System.Drawing.Point(182, 45)
        Me.txtODBCtype.Name = "txtODBCtype"
        Me.txtODBCtype.Size = New System.Drawing.Size(254, 20)
        Me.txtODBCtype.TabIndex = 13
        Me.txtODBCtype.TabStop = False
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(170, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Currently Selected MetaData"
        '
        'btnDSN
        '
        Me.btnDSN.Location = New System.Drawing.Point(311, 71)
        Me.btnDSN.Name = "btnDSN"
        Me.btnDSN.Size = New System.Drawing.Size(125, 23)
        Me.btnDSN.TabIndex = 14
        Me.btnDSN.Text = "C&hange"
        Me.btnDSN.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 76)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(238, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Click here to select a different MetaData"
        '
        'txtDSN
        '
        Me.txtDSN.BackColor = System.Drawing.SystemColors.Window
        Me.txtDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSN.Location = New System.Drawing.Point(182, 19)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.ReadOnly = True
        Me.txtDSN.Size = New System.Drawing.Size(254, 20)
        Me.txtDSN.TabIndex = 12
        Me.txtDSN.TabStop = False
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
        Me.gbProject.Controls.Add(Me.Label7)
        Me.gbProject.Controls.Add(Me.txtProj)
        Me.gbProject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProject.Location = New System.Drawing.Point(12, 184)
        Me.gbProject.Name = "gbProject"
        Me.gbProject.Size = New System.Drawing.Size(443, 95)
        Me.gbProject.TabIndex = 7
        Me.gbProject.TabStop = False
        Me.gbProject.Text = "Create Project"
        '
        'txtSchema
        '
        Me.txtSchema.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSchema.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSchema.Location = New System.Drawing.Point(124, 17)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.Size = New System.Drawing.Size(313, 20)
        Me.txtSchema.TabIndex = 15
        '
        'cbLoginNo
        '
        Me.cbLoginNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginNo.AutoSize = True
        Me.cbLoginNo.Location = New System.Drawing.Point(395, 43)
        Me.cbLoginNo.Name = "cbLoginNo"
        Me.cbLoginNo.Size = New System.Drawing.Size(42, 17)
        Me.cbLoginNo.TabIndex = 17
        Me.cbLoginNo.Text = "No"
        Me.cbLoginNo.UseVisualStyleBackColor = True
        '
        'cbLoginYes
        '
        Me.cbLoginYes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginYes.AutoSize = True
        Me.cbLoginYes.Location = New System.Drawing.Point(342, 43)
        Me.cbLoginYes.Name = "cbLoginYes"
        Me.cbLoginYes.Size = New System.Drawing.Size(47, 17)
        Me.cbLoginYes.TabIndex = 16
        Me.cbLoginYes.Text = "Yes"
        Me.cbLoginYes.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 13)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "MetaData Schema"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(5, 44)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(313, 13)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "MetaData requires User Login? (Username, Password)"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 70)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(146, 13)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Enter New Project Name"
        '
        'txtProj
        '
        Me.txtProj.AllowDrop = True
        Me.txtProj.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtProj.BackColor = System.Drawing.SystemColors.Window
        Me.txtProj.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProj.Location = New System.Drawing.Point(206, 67)
        Me.txtProj.Name = "txtProj"
        Me.txtProj.Size = New System.Drawing.Size(231, 20)
        Me.txtProj.TabIndex = 18
        '
        'frmNewProj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(467, 347)
        Me.Controls.Add(Me.gbProject)
        Me.Controls.Add(Me.gbODBC)
        Me.MaximumSize = New System.Drawing.Size(475, 496)
        Me.Name = "frmNewProj"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "New Project"
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.gbODBC, 0)
        Me.Controls.SetChildIndex(Me.gbProject, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbODBC.ResumeLayout(False)
        Me.gbODBC.PerformLayout()
        Me.gbProject.ResumeLayout(False)
        Me.gbProject.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbODBC As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtODBCtype As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnDSN As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents gbProject As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtProj As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbLoginNo As System.Windows.Forms.CheckBox
    Friend WithEvents cbLoginYes As System.Windows.Forms.CheckBox
    Friend WithEvents txtSchema As System.Windows.Forms.TextBox
End Class
