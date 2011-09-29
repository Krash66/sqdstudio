<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBatchFiles
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBatchFiles))
        Me.OK_Button = New System.Windows.Forms.Button
        Me.gbFileList = New System.Windows.Forms.GroupBox
        Me.txtBatFiles = New System.Windows.Forms.TextBox
        Me.cbOverwrite = New System.Windows.Forms.CheckBox
        Me.btnCreate = New System.Windows.Forms.Button
        Me.gbFileList.SuspendLayout()
        Me.SuspendLayout()
        '
        'OK_Button
        '
        Me.OK_Button.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.OK_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.OK_Button.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OK_Button.Location = New System.Drawing.Point(353, 367)
        Me.OK_Button.Name = "OK_Button"
        Me.OK_Button.Size = New System.Drawing.Size(67, 23)
        Me.OK_Button.TabIndex = 1
        Me.OK_Button.Text = "OK"
        '
        'gbFileList
        '
        Me.gbFileList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFileList.Controls.Add(Me.txtBatFiles)
        Me.gbFileList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFileList.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbFileList.Location = New System.Drawing.Point(12, 71)
        Me.gbFileList.Name = "gbFileList"
        Me.gbFileList.Size = New System.Drawing.Size(410, 282)
        Me.gbFileList.TabIndex = 1
        Me.gbFileList.TabStop = False
        Me.gbFileList.Text = "Batch Files Created"
        '
        'txtBatFiles
        '
        Me.txtBatFiles.AcceptsReturn = True
        Me.txtBatFiles.BackColor = System.Drawing.Color.White
        Me.txtBatFiles.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtBatFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBatFiles.ForeColor = System.Drawing.Color.Blue
        Me.txtBatFiles.Location = New System.Drawing.Point(3, 16)
        Me.txtBatFiles.Multiline = True
        Me.txtBatFiles.Name = "txtBatFiles"
        Me.txtBatFiles.ReadOnly = True
        Me.txtBatFiles.Size = New System.Drawing.Size(404, 263)
        Me.txtBatFiles.TabIndex = 0
        '
        'cbOverwrite
        '
        Me.cbOverwrite.AutoSize = True
        Me.cbOverwrite.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbOverwrite.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbOverwrite.Location = New System.Drawing.Point(15, 12)
        Me.cbOverwrite.Name = "cbOverwrite"
        Me.cbOverwrite.Size = New System.Drawing.Size(292, 17)
        Me.cbOverwrite.TabIndex = 2
        Me.cbOverwrite.Text = "Over-write Existing Batch Files? check for YES"
        Me.cbOverwrite.UseVisualStyleBackColor = True
        '
        'btnCreate
        '
        Me.btnCreate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCreate.Location = New System.Drawing.Point(15, 36)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(241, 23)
        Me.btnCreate.TabIndex = 3
        Me.btnCreate.Text = "Create Batch Files"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'frmBatchFiles
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.OK_Button
        Me.ClientSize = New System.Drawing.Size(434, 400)
        Me.Controls.Add(Me.OK_Button)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.cbOverwrite)
        Me.Controls.Add(Me.gbFileList)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmBatchFiles"
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "SQData Batch File Creation"
        Me.gbFileList.ResumeLayout(False)
        Me.gbFileList.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents gbFileList As System.Windows.Forms.GroupBox
    Friend WithEvents txtBatFiles As System.Windows.Forms.TextBox
    Friend WithEvents cbOverwrite As System.Windows.Forms.CheckBox
    Friend WithEvents btnCreate As System.Windows.Forms.Button

End Class
