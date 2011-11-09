<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmXMLconv
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
        Me.gbIn = New System.Windows.Forms.GroupBox
        Me.btnbrowseIn = New System.Windows.Forms.Button
        Me.txtInMessage = New System.Windows.Forms.TextBox
        Me.txtInPath = New System.Windows.Forms.TextBox
        Me.gbXMLout = New System.Windows.Forms.GroupBox
        Me.txtXMLout = New System.Windows.Forms.TextBox
        Me.btnbrowseOut = New System.Windows.Forms.Button
        Me.txtOutPath = New System.Windows.Forms.TextBox
        Me.OFD1 = New System.Windows.Forms.OpenFileDialog
        Me.SFD1 = New System.Windows.Forms.SaveFileDialog
        Me.btnConv = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.gbIn.SuspendLayout()
        Me.gbXMLout.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(734, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 475)
        Me.GroupBox1.Size = New System.Drawing.Size(736, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdOk.Location = New System.Drawing.Point(377, 498)
        Me.cmdOk.Size = New System.Drawing.Size(155, 24)
        Me.cmdOk.Text = "&Open DTD File"
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(548, 498)
        Me.cmdCancel.Text = "&Close"
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(644, 498)
        '
        'Label1
        '
        Me.Label1.Text = "XML File Conversion"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(658, 39)
        Me.Label2.Text = "Open an XML message and convert it to an XML DTD description file."
        '
        'gbIn
        '
        Me.gbIn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbIn.Controls.Add(Me.btnbrowseIn)
        Me.gbIn.Controls.Add(Me.txtInMessage)
        Me.gbIn.Controls.Add(Me.txtInPath)
        Me.gbIn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbIn.Location = New System.Drawing.Point(12, 74)
        Me.gbIn.Name = "gbIn"
        Me.gbIn.Size = New System.Drawing.Size(710, 173)
        Me.gbIn.TabIndex = 58
        Me.gbIn.TabStop = False
        Me.gbIn.Text = "XML Message In"
        '
        'btnbrowseIn
        '
        Me.btnbrowseIn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnbrowseIn.Location = New System.Drawing.Point(498, 17)
        Me.btnbrowseIn.Name = "btnbrowseIn"
        Me.btnbrowseIn.Size = New System.Drawing.Size(206, 23)
        Me.btnbrowseIn.TabIndex = 2
        Me.btnbrowseIn.Text = "Input Message"
        Me.btnbrowseIn.UseVisualStyleBackColor = True
        '
        'txtInMessage
        '
        Me.txtInMessage.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInMessage.BackColor = System.Drawing.SystemColors.Window
        Me.txtInMessage.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInMessage.Location = New System.Drawing.Point(6, 45)
        Me.txtInMessage.Multiline = True
        Me.txtInMessage.Name = "txtInMessage"
        Me.txtInMessage.ReadOnly = True
        Me.txtInMessage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInMessage.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInMessage.Size = New System.Drawing.Size(698, 122)
        Me.txtInMessage.TabIndex = 1
        Me.txtInMessage.Text = " "
        Me.txtInMessage.WordWrap = False
        '
        'txtInPath
        '
        Me.txtInPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInPath.Location = New System.Drawing.Point(6, 19)
        Me.txtInPath.Name = "txtInPath"
        Me.txtInPath.Size = New System.Drawing.Size(486, 20)
        Me.txtInPath.TabIndex = 0
        '
        'gbXMLout
        '
        Me.gbXMLout.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbXMLout.Controls.Add(Me.txtXMLout)
        Me.gbXMLout.Controls.Add(Me.btnbrowseOut)
        Me.gbXMLout.Controls.Add(Me.txtOutPath)
        Me.gbXMLout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbXMLout.Location = New System.Drawing.Point(12, 282)
        Me.gbXMLout.Name = "gbXMLout"
        Me.gbXMLout.Size = New System.Drawing.Size(710, 187)
        Me.gbXMLout.TabIndex = 59
        Me.gbXMLout.TabStop = False
        Me.gbXMLout.Text = "XML DTD output"
        '
        'txtXMLout
        '
        Me.txtXMLout.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtXMLout.BackColor = System.Drawing.SystemColors.Window
        Me.txtXMLout.Font = New System.Drawing.Font("Courier New", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtXMLout.Location = New System.Drawing.Point(6, 19)
        Me.txtXMLout.Multiline = True
        Me.txtXMLout.Name = "txtXMLout"
        Me.txtXMLout.ReadOnly = True
        Me.txtXMLout.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtXMLout.Size = New System.Drawing.Size(698, 134)
        Me.txtXMLout.TabIndex = 3
        Me.txtXMLout.WordWrap = False
        '
        'btnbrowseOut
        '
        Me.btnbrowseOut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnbrowseOut.Location = New System.Drawing.Point(498, 159)
        Me.btnbrowseOut.Name = "btnbrowseOut"
        Me.btnbrowseOut.Size = New System.Drawing.Size(206, 23)
        Me.btnbrowseOut.TabIndex = 2
        Me.btnbrowseOut.Text = "Save Output DTD File"
        Me.btnbrowseOut.UseVisualStyleBackColor = True
        '
        'txtOutPath
        '
        Me.txtOutPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutPath.Location = New System.Drawing.Point(6, 161)
        Me.txtOutPath.Name = "txtOutPath"
        Me.txtOutPath.Size = New System.Drawing.Size(486, 20)
        Me.txtOutPath.TabIndex = 1
        '
        'OFD1
        '
        Me.OFD1.Filter = "XML Files|*.xml|All Files|*.*"
        Me.OFD1.Title = "Open XML Message"
        '
        'SFD1
        '
        Me.SFD1.DefaultExt = "DTD"
        Me.SFD1.Filter = "XML DTD Files|*.DTD|All Files|*.*"
        Me.SFD1.Title = "DTD File Save"
        '
        'btnConv
        '
        Me.btnConv.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnConv.Location = New System.Drawing.Point(169, 253)
        Me.btnConv.Name = "btnConv"
        Me.btnConv.Size = New System.Drawing.Size(208, 23)
        Me.btnConv.TabIndex = 61
        Me.btnConv.Text = "Convert"
        Me.btnConv.UseVisualStyleBackColor = True
        '
        'frmXMLconv
        '
        Me.ClientSize = New System.Drawing.Size(734, 537)
        Me.Controls.Add(Me.gbIn)
        Me.Controls.Add(Me.gbXMLout)
        Me.Controls.Add(Me.btnConv)
        Me.Name = "frmXMLconv"
        Me.Controls.SetChildIndex(Me.btnConv, 0)
        Me.Controls.SetChildIndex(Me.gbXMLout, 0)
        Me.Controls.SetChildIndex(Me.gbIn, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbIn.ResumeLayout(False)
        Me.gbIn.PerformLayout()
        Me.gbXMLout.ResumeLayout(False)
        Me.gbXMLout.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbIn As System.Windows.Forms.GroupBox
    Friend WithEvents txtInMessage As System.Windows.Forms.TextBox
    Friend WithEvents txtInPath As System.Windows.Forms.TextBox
    Friend WithEvents gbXMLout As System.Windows.Forms.GroupBox
    Friend WithEvents OFD1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SFD1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnbrowseIn As System.Windows.Forms.Button
    Friend WithEvents btnConv As System.Windows.Forms.Button
    Friend WithEvents txtXMLout As System.Windows.Forms.TextBox
    Friend WithEvents btnbrowseOut As System.Windows.Forms.Button
    Friend WithEvents txtOutPath As System.Windows.Forms.TextBox

End Class
