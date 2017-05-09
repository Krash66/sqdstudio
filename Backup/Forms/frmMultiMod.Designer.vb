<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMultiMod
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
        Me.gbSelectTree = New System.Windows.Forms.GroupBox
        Me.tvModSel = New System.Windows.Forms.TreeView
        Me.gbFileName = New System.Windows.Forms.GroupBox
        Me.txtScriptName = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.gbSelectTree.SuspendLayout()
        Me.gbFileName.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(487, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 409)
        Me.GroupBox1.Size = New System.Drawing.Size(489, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(205, 432)
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(301, 432)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(397, 432)
        '
        'Label1
        '
        Me.Label1.Text = "Table Model Selection"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(411, 39)
        Me.Label2.Text = "Check the boxes representing the tables you would like to model as one large file" & _
            "."
        '
        'gbSelectTree
        '
        Me.gbSelectTree.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbSelectTree.Controls.Add(Me.tvModSel)
        Me.gbSelectTree.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSelectTree.Location = New System.Drawing.Point(12, 74)
        Me.gbSelectTree.Name = "gbSelectTree"
        Me.gbSelectTree.Size = New System.Drawing.Size(463, 275)
        Me.gbSelectTree.TabIndex = 58
        Me.gbSelectTree.TabStop = False
        Me.gbSelectTree.Text = "Select Tables to Model"
        '
        'tvModSel
        '
        Me.tvModSel.CheckBoxes = True
        Me.tvModSel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvModSel.Location = New System.Drawing.Point(3, 16)
        Me.tvModSel.Name = "tvModSel"
        Me.tvModSel.Size = New System.Drawing.Size(457, 256)
        Me.tvModSel.TabIndex = 0
        '
        'gbFileName
        '
        Me.gbFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFileName.Controls.Add(Me.txtScriptName)
        Me.gbFileName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFileName.Location = New System.Drawing.Point(12, 355)
        Me.gbFileName.Name = "gbFileName"
        Me.gbFileName.Size = New System.Drawing.Size(463, 48)
        Me.gbFileName.TabIndex = 59
        Me.gbFileName.TabStop = False
        Me.gbFileName.Text = "Name Your Model Script"
        '
        'txtScriptName
        '
        Me.txtScriptName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtScriptName.Location = New System.Drawing.Point(6, 19)
        Me.txtScriptName.Name = "txtScriptName"
        Me.txtScriptName.Size = New System.Drawing.Size(451, 20)
        Me.txtScriptName.TabIndex = 0
        '
        'frmMultiMod
        '
        Me.ClientSize = New System.Drawing.Size(487, 471)
        Me.Controls.Add(Me.gbSelectTree)
        Me.Controls.Add(Me.gbFileName)
        Me.Name = "frmMultiMod"
        Me.Controls.SetChildIndex(Me.gbFileName, 0)
        Me.Controls.SetChildIndex(Me.gbSelectTree, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbSelectTree.ResumeLayout(False)
        Me.gbFileName.ResumeLayout(False)
        Me.gbFileName.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbSelectTree As System.Windows.Forms.GroupBox
    Friend WithEvents tvModSel As System.Windows.Forms.TreeView
    Friend WithEvents gbFileName As System.Windows.Forms.GroupBox
    Friend WithEvents txtScriptName As System.Windows.Forms.TextBox

End Class
