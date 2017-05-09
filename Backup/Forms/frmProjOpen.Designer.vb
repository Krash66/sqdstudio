<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProjOpen
    Inherits SQDStudio.frmBlank

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
        Me.gbProject = New System.Windows.Forms.GroupBox
        Me.txtSchema = New System.Windows.Forms.TextBox
        Me.cbLoginNo = New System.Windows.Forms.CheckBox
        Me.cbLoginYes = New System.Windows.Forms.CheckBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtProjDate = New System.Windows.Forms.TextBox
        Me.btnProj = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtProj = New System.Windows.Forms.TextBox
        Me.txtDSN = New System.Windows.Forms.TextBox
        Me.gbODBC = New System.Windows.Forms.GroupBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtODBCtype = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.btnDSN = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.gbProject.SuspendLayout()
        Me.gbODBC.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(465, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 347)
        Me.GroupBox1.Size = New System.Drawing.Size(467, 7)
        Me.GroupBox1.TabIndex = 13
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(183, 370)
        Me.cmdOk.TabIndex = 24
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(279, 370)
        Me.cmdCancel.TabIndex = 25
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(375, 370)
        Me.cmdHelp.TabIndex = 26
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(72, 3)
        Me.Label1.Size = New System.Drawing.Size(383, 14)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Select Project to open"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(72, 20)
        Me.Label2.Size = New System.Drawing.Size(389, 44)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "The project last opened is filled in by default." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Use the buttons to change your " & _
            "MetaData source" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "and Project if necessary."
        '
        'gbProject
        '
        Me.gbProject.Controls.Add(Me.txtSchema)
        Me.gbProject.Controls.Add(Me.cbLoginNo)
        Me.gbProject.Controls.Add(Me.cbLoginYes)
        Me.gbProject.Controls.Add(Me.Label10)
        Me.gbProject.Controls.Add(Me.Label9)
        Me.gbProject.Controls.Add(Me.Label8)
        Me.gbProject.Controls.Add(Me.Label7)
        Me.gbProject.Controls.Add(Me.txtProjDate)
        Me.gbProject.Controls.Add(Me.btnProj)
        Me.gbProject.Controls.Add(Me.Label4)
        Me.gbProject.Controls.Add(Me.txtProj)
        Me.gbProject.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProject.Location = New System.Drawing.Point(13, 185)
        Me.gbProject.Name = "gbProject"
        Me.gbProject.Size = New System.Drawing.Size(442, 159)
        Me.gbProject.TabIndex = 7
        Me.gbProject.TabStop = False
        Me.gbProject.Text = "Selected Project"
        '
        'txtSchema
        '
        Me.txtSchema.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSchema.Location = New System.Drawing.Point(123, 18)
        Me.txtSchema.Name = "txtSchema"
        Me.txtSchema.Size = New System.Drawing.Size(313, 20)
        Me.txtSchema.TabIndex = 18
        '
        'cbLoginNo
        '
        Me.cbLoginNo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginNo.AutoSize = True
        Me.cbLoginNo.Location = New System.Drawing.Point(393, 47)
        Me.cbLoginNo.Name = "cbLoginNo"
        Me.cbLoginNo.Size = New System.Drawing.Size(42, 17)
        Me.cbLoginNo.TabIndex = 20
        Me.cbLoginNo.Text = "No"
        Me.cbLoginNo.UseVisualStyleBackColor = True
        '
        'cbLoginYes
        '
        Me.cbLoginYes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbLoginYes.AutoSize = True
        Me.cbLoginYes.Location = New System.Drawing.Point(341, 47)
        Me.cbLoginYes.Name = "cbLoginYes"
        Me.cbLoginYes.Size = New System.Drawing.Size(47, 17)
        Me.cbLoginYes.TabIndex = 19
        Me.cbLoginYes.Text = "Yes"
        Me.cbLoginYes.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(6, 21)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(111, 13)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "MetaData Schema"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(6, 48)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(313, 13)
        Me.Label9.TabIndex = 9
        Me.Label9.Text = "MetaData requires User Login? (Username, Password)"
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(6, 104)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(151, 13)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Date Project was created"
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 77)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(155, 13)
        Me.Label7.TabIndex = 10
        Me.Label7.Text = "Currently Selected Project"
        '
        'txtProjDate
        '
        Me.txtProjDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtProjDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtProjDate.Location = New System.Drawing.Point(181, 101)
        Me.txtProjDate.Name = "txtProjDate"
        Me.txtProjDate.Size = New System.Drawing.Size(255, 20)
        Me.txtProjDate.TabIndex = 22
        Me.txtProjDate.TabStop = False
        '
        'btnProj
        '
        Me.btnProj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnProj.Location = New System.Drawing.Point(311, 127)
        Me.btnProj.Name = "btnProj"
        Me.btnProj.Size = New System.Drawing.Size(125, 23)
        Me.btnProj.TabIndex = 23
        Me.btnProj.Text = "Cha&nge"
        Me.btnProj.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 132)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(223, 13)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Click here to select a different Project"
        '
        'txtProj
        '
        Me.txtProj.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtProj.BackColor = System.Drawing.SystemColors.Window
        Me.txtProj.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProj.Location = New System.Drawing.Point(181, 74)
        Me.txtProj.Name = "txtProj"
        Me.txtProj.ReadOnly = True
        Me.txtProj.Size = New System.Drawing.Size(255, 20)
        Me.txtProj.TabIndex = 21
        Me.txtProj.TabStop = False
        '
        'txtDSN
        '
        Me.txtDSN.BackColor = System.Drawing.SystemColors.Window
        Me.txtDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDSN.Location = New System.Drawing.Point(182, 19)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.ReadOnly = True
        Me.txtDSN.Size = New System.Drawing.Size(254, 20)
        Me.txtDSN.TabIndex = 15
        Me.txtDSN.TabStop = False
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
        Me.gbODBC.Size = New System.Drawing.Size(442, 105)
        Me.gbODBC.TabIndex = 3
        Me.gbODBC.TabStop = False
        Me.gbODBC.Text = "Selected MetaData Source"
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
        Me.txtODBCtype.TabIndex = 16
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
        Me.btnDSN.TabIndex = 17
        Me.btnDSN.Text = "Ch&ange"
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
        'frmProjOpen
        '
        Me.ClientSize = New System.Drawing.Size(465, 409)
        Me.Controls.Add(Me.gbODBC)
        Me.Controls.Add(Me.gbProject)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(473, 443)
        Me.MinimumSize = New System.Drawing.Size(473, 443)
        Me.Name = "frmProjOpen"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select Project"
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.gbProject, 0)
        Me.Controls.SetChildIndex(Me.gbODBC, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbProject.ResumeLayout(False)
        Me.gbProject.PerformLayout()
        Me.gbODBC.ResumeLayout(False)
        Me.gbODBC.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbProject As System.Windows.Forms.GroupBox
    Friend WithEvents txtProj As System.Windows.Forms.TextBox
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents gbODBC As System.Windows.Forms.GroupBox
    Friend WithEvents btnProj As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnDSN As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtProjDate As System.Windows.Forms.TextBox
    Friend WithEvents txtODBCtype As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cbLoginNo As System.Windows.Forms.CheckBox
    Friend WithEvents cbLoginYes As System.Windows.Forms.CheckBox
    Friend WithEvents txtSchema As System.Windows.Forms.TextBox

End Class
