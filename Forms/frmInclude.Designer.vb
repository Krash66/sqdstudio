<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInclude
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInclude))
        Me.gbInclude = New System.Windows.Forms.GroupBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnBrowse = New System.Windows.Forms.Button
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtFile = New System.Windows.Forms.TextBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.gbInclude.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        'Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        'Me.PictureBox1.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstu" & _
        '    "dio\images\FormTop\sq_skyblue.jpg"
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 233)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(232, 256)
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(328, 256)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(424, 256)
        '
        'Label1
        '
        Me.Label1.Text = "Include Script"
        '
        'Label2
        '
        Me.Label2.Text = "Add a Datastore or Procedure from an existing script. Add the inline script of a " & _
            "datastore or procedure to be referenced in this Project."
        '
        'gbInclude
        '
        Me.gbInclude.Controls.Add(Me.txtName)
        Me.gbInclude.Controls.Add(Me.Label4)
        Me.gbInclude.Controls.Add(Me.btnBrowse)
        Me.gbInclude.Controls.Add(Me.Label3)
        Me.gbInclude.Controls.Add(Me.txtFile)
        Me.gbInclude.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbInclude.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbInclude.Location = New System.Drawing.Point(12, 74)
        Me.gbInclude.Name = "gbInclude"
        Me.gbInclude.Size = New System.Drawing.Size(490, 151)
        Me.gbInclude.TabIndex = 58
        Me.gbInclude.TabStop = False
        '
        'txtName
        '
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.Location = New System.Drawing.Point(6, 36)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(196, 20)
        Me.txtName.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 20)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Name"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(315, 112)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(168, 23)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse for File"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(4, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(217, 13)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Inline Script File to include in Project"
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(6, 86)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.Size = New System.Drawing.Size(477, 20)
        Me.txtFile.TabIndex = 0
        '
        'frmInclude
        '
        Me.ClientSize = New System.Drawing.Size(514, 295)
        Me.Controls.Add(Me.gbInclude)
        Me.Name = "frmInclude"
        Me.Text = "SQData Studio V3 "
        Me.Controls.SetChildIndex(Me.gbInclude, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.gbInclude.ResumeLayout(False)
        Me.gbInclude.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbInclude As System.Windows.Forms.GroupBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label

End Class
