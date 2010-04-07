<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRplcDescRet
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
        Me.txtNewFlds = New System.Windows.Forms.TextBox
        Me.txtOldFlds = New System.Windows.Forms.TextBox
        Me.txtNewPath = New System.Windows.Forms.TextBox
        Me.txtOldPath = New System.Windows.Forms.TextBox
        Me.lblNewPath = New System.Windows.Forms.Label
        Me.lblOldPath = New System.Windows.Forms.Label
        Me.gbNewDesc = New System.Windows.Forms.GroupBox
        Me.lblNewFlds = New System.Windows.Forms.Label
        Me.gbOldDesc = New System.Windows.Forms.GroupBox
        Me.lblOldFlds = New System.Windows.Forms.Label
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.lblProceed = New System.Windows.Forms.Label
        Me.gbProc = New System.Windows.Forms.GroupBox
        Me.lvProc = New System.Windows.Forms.ListView
        Me.Engine = New System.Windows.Forms.ColumnHeader
        Me.Procedure = New System.Windows.Forms.ColumnHeader
        Me.Panel1.SuspendLayout()
        Me.gbNewDesc.SuspendLayout()
        Me.gbOldDesc.SuspendLayout()
        Me.gbName.SuspendLayout()
        Me.gbProc.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(469, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 577)
        Me.GroupBox1.Size = New System.Drawing.Size(471, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(187, 600)
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(283, 600)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(379, 600)
        '
        'Label1
        '
        Me.Label1.Text = "Description Replacement"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(393, 39)
        Me.Label2.Text = "Summary of fields added to or deleted from a Description during description file " & _
            "replacement"
        '
        'txtNewFlds
        '
        Me.txtNewFlds.BackColor = System.Drawing.SystemColors.Window
        Me.txtNewFlds.Location = New System.Drawing.Point(119, 58)
        Me.txtNewFlds.Multiline = True
        Me.txtNewFlds.Name = "txtNewFlds"
        Me.txtNewFlds.ReadOnly = True
        Me.txtNewFlds.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtNewFlds.Size = New System.Drawing.Size(331, 87)
        Me.txtNewFlds.TabIndex = 58
        '
        'txtOldFlds
        '
        Me.txtOldFlds.BackColor = System.Drawing.SystemColors.Window
        Me.txtOldFlds.Location = New System.Drawing.Point(119, 58)
        Me.txtOldFlds.Multiline = True
        Me.txtOldFlds.Name = "txtOldFlds"
        Me.txtOldFlds.ReadOnly = True
        Me.txtOldFlds.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtOldFlds.Size = New System.Drawing.Size(331, 83)
        Me.txtOldFlds.TabIndex = 59
        '
        'txtNewPath
        '
        Me.txtNewPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtNewPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNewPath.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.txtNewPath.Location = New System.Drawing.Point(6, 32)
        Me.txtNewPath.Name = "txtNewPath"
        Me.txtNewPath.ReadOnly = True
        Me.txtNewPath.Size = New System.Drawing.Size(444, 20)
        Me.txtNewPath.TabIndex = 60
        '
        'txtOldPath
        '
        Me.txtOldPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtOldPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOldPath.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.txtOldPath.Location = New System.Drawing.Point(6, 32)
        Me.txtOldPath.Name = "txtOldPath"
        Me.txtOldPath.ReadOnly = True
        Me.txtOldPath.Size = New System.Drawing.Size(444, 20)
        Me.txtOldPath.TabIndex = 61
        '
        'lblNewPath
        '
        Me.lblNewPath.AutoSize = True
        Me.lblNewPath.Location = New System.Drawing.Point(6, 16)
        Me.lblNewPath.Name = "lblNewPath"
        Me.lblNewPath.Size = New System.Drawing.Size(154, 13)
        Me.lblNewPath.TabIndex = 62
        Me.lblNewPath.Text = "New Full Description Path"
        '
        'lblOldPath
        '
        Me.lblOldPath.AutoSize = True
        Me.lblOldPath.Location = New System.Drawing.Point(6, 16)
        Me.lblOldPath.Name = "lblOldPath"
        Me.lblOldPath.Size = New System.Drawing.Size(148, 13)
        Me.lblOldPath.TabIndex = 63
        Me.lblOldPath.Text = "Old Full Description Path"
        '
        'gbNewDesc
        '
        Me.gbNewDesc.Controls.Add(Me.lblNewFlds)
        Me.gbNewDesc.Controls.Add(Me.lblNewPath)
        Me.gbNewDesc.Controls.Add(Me.txtNewPath)
        Me.gbNewDesc.Controls.Add(Me.txtNewFlds)
        Me.gbNewDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNewDesc.Location = New System.Drawing.Point(7, 126)
        Me.gbNewDesc.Name = "gbNewDesc"
        Me.gbNewDesc.Size = New System.Drawing.Size(456, 151)
        Me.gbNewDesc.TabIndex = 64
        Me.gbNewDesc.TabStop = False
        Me.gbNewDesc.Text = "New Description File"
        '
        'lblNewFlds
        '
        Me.lblNewFlds.Location = New System.Drawing.Point(6, 59)
        Me.lblNewFlds.Name = "lblNewFlds"
        Me.lblNewFlds.Size = New System.Drawing.Size(107, 86)
        Me.lblNewFlds.TabIndex = 63
        Me.lblNewFlds.Text = "Fields added by New Description File. " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Must be manually mapped."
        '
        'gbOldDesc
        '
        Me.gbOldDesc.Controls.Add(Me.lblOldFlds)
        Me.gbOldDesc.Controls.Add(Me.txtOldFlds)
        Me.gbOldDesc.Controls.Add(Me.lblOldPath)
        Me.gbOldDesc.Controls.Add(Me.txtOldPath)
        Me.gbOldDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbOldDesc.Location = New System.Drawing.Point(7, 283)
        Me.gbOldDesc.Name = "gbOldDesc"
        Me.gbOldDesc.Size = New System.Drawing.Size(456, 146)
        Me.gbOldDesc.TabIndex = 65
        Me.gbOldDesc.TabStop = False
        Me.gbOldDesc.Text = "Old Description File"
        '
        'lblOldFlds
        '
        Me.lblOldFlds.Location = New System.Drawing.Point(7, 59)
        Me.lblOldFlds.Name = "lblOldFlds"
        Me.lblOldFlds.Size = New System.Drawing.Size(106, 82)
        Me.lblOldFlds.TabIndex = 64
        Me.lblOldFlds.Text = "Fields that will no longer exist in New Description." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Must be manually remapped" & _
            ""
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.txtName)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.Location = New System.Drawing.Point(7, 74)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(456, 46)
        Me.gbName.TabIndex = 66
        Me.gbName.TabStop = False
        Me.gbName.Text = "Description Name"
        '
        'txtName
        '
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.ForeColor = System.Drawing.SystemColors.MenuHighlight
        Me.txtName.Location = New System.Drawing.Point(6, 19)
        Me.txtName.Name = "txtName"
        Me.txtName.ReadOnly = True
        Me.txtName.Size = New System.Drawing.Size(272, 20)
        Me.txtName.TabIndex = 0
        '
        'lblProceed
        '
        Me.lblProceed.AutoSize = True
        Me.lblProceed.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProceed.Location = New System.Drawing.Point(51, 561)
        Me.lblProceed.Name = "lblProceed"
        Me.lblProceed.Size = New System.Drawing.Size(408, 13)
        Me.lblProceed.TabIndex = 67
        Me.lblProceed.Text = "*** Do you still want to proceed with Description File Replacement? ***"
        '
        'gbProc
        '
        Me.gbProc.Controls.Add(Me.lvProc)
        Me.gbProc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProc.Location = New System.Drawing.Point(7, 435)
        Me.gbProc.Name = "gbProc"
        Me.gbProc.Size = New System.Drawing.Size(456, 123)
        Me.gbProc.TabIndex = 68
        Me.gbProc.TabStop = False
        Me.gbProc.Text = "Procedures Affected by Description Change"
        '
        'lvProc
        '
        Me.lvProc.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Engine, Me.Procedure})
        Me.lvProc.Location = New System.Drawing.Point(9, 19)
        Me.lvProc.Name = "lvProc"
        Me.lvProc.Size = New System.Drawing.Size(441, 98)
        Me.lvProc.TabIndex = 0
        Me.lvProc.UseCompatibleStateImageBehavior = False
        Me.lvProc.View = System.Windows.Forms.View.Details
        '
        'Engine
        '
        Me.Engine.Text = "Engine"
        Me.Engine.Width = 186
        '
        'Procedure
        '
        Me.Procedure.Text = "Procedure"
        Me.Procedure.Width = 223
        '
        'frmRplcDescRet
        '
        Me.ClientSize = New System.Drawing.Size(469, 639)
        Me.Controls.Add(Me.gbName)
        Me.Controls.Add(Me.gbNewDesc)
        Me.Controls.Add(Me.gbOldDesc)
        Me.Controls.Add(Me.gbProc)
        Me.Controls.Add(Me.lblProceed)
        Me.MaximumSize = New System.Drawing.Size(477, 673)
        Me.MinimumSize = New System.Drawing.Size(477, 673)
        Me.Name = "frmRplcDescRet"
        Me.Text = "SQData Studio V3 "
        Me.Controls.SetChildIndex(Me.lblProceed, 0)
        Me.Controls.SetChildIndex(Me.gbProc, 0)
        Me.Controls.SetChildIndex(Me.gbOldDesc, 0)
        Me.Controls.SetChildIndex(Me.gbNewDesc, 0)
        Me.Controls.SetChildIndex(Me.gbName, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbNewDesc.ResumeLayout(False)
        Me.gbNewDesc.PerformLayout()
        Me.gbOldDesc.ResumeLayout(False)
        Me.gbOldDesc.PerformLayout()
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbProc.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtNewFlds As System.Windows.Forms.TextBox
    Friend WithEvents txtOldFlds As System.Windows.Forms.TextBox
    Friend WithEvents txtNewPath As System.Windows.Forms.TextBox
    Friend WithEvents txtOldPath As System.Windows.Forms.TextBox
    Friend WithEvents lblNewPath As System.Windows.Forms.Label
    Friend WithEvents lblOldPath As System.Windows.Forms.Label
    Friend WithEvents gbNewDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbOldDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents lblNewFlds As System.Windows.Forms.Label
    Friend WithEvents lblOldFlds As System.Windows.Forms.Label
    Friend WithEvents lblProceed As System.Windows.Forms.Label
    Friend WithEvents gbProc As System.Windows.Forms.GroupBox
    Friend WithEvents lvProc As System.Windows.Forms.ListView
    Friend WithEvents Engine As System.Windows.Forms.ColumnHeader
    Friend WithEvents Procedure As System.Windows.Forms.ColumnHeader

End Class
