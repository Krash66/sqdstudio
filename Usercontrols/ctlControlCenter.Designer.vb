<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlControlCenter
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctlControlCenter))
        Me.WebBrowser1 = New System.Windows.Forms.WebBrowser
        Me.gbControlCenter = New System.Windows.Forms.GroupBox
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.btnHome = New System.Windows.Forms.ToolStripButton
        Me.btnBack = New System.Windows.Forms.ToolStripButton
        Me.btnForward = New System.Windows.Forms.ToolStripButton
        Me.txtURL = New System.Windows.Forms.ToolStripTextBox
        Me.btnRefresh = New System.Windows.Forms.ToolStripButton
        Me.btnGo = New System.Windows.Forms.ToolStripButton
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.gbControlCenter.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'WebBrowser1
        '
        Me.WebBrowser1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.WebBrowser1.Location = New System.Drawing.Point(3, 16)
        Me.WebBrowser1.MinimumSize = New System.Drawing.Size(20, 20)
        Me.WebBrowser1.Name = "WebBrowser1"
        Me.WebBrowser1.Size = New System.Drawing.Size(559, 413)
        Me.WebBrowser1.TabIndex = 0
        '
        'gbControlCenter
        '
        Me.gbControlCenter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbControlCenter.Controls.Add(Me.WebBrowser1)
        Me.gbControlCenter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbControlCenter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbControlCenter.Location = New System.Drawing.Point(3, 28)
        Me.gbControlCenter.Name = "gbControlCenter"
        Me.gbControlCenter.Size = New System.Drawing.Size(565, 432)
        Me.gbControlCenter.TabIndex = 1
        Me.gbControlCenter.TabStop = False
        Me.gbControlCenter.Text = "Control Center"
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnHome, Me.btnBack, Me.btnForward, Me.ToolStripLabel1, Me.txtURL, Me.btnRefresh, Me.btnGo})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(571, 25)
        Me.ToolStrip1.TabIndex = 2
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btnHome
        '
        Me.btnHome.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnHome.Image = CType(resources.GetObject("btnHome.Image"), System.Drawing.Image)
        Me.btnHome.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnHome.Name = "btnHome"
        Me.btnHome.Size = New System.Drawing.Size(23, 22)
        Me.btnHome.Text = "Home"
        '
        'btnBack
        '
        Me.btnBack.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnBack.Image = CType(resources.GetObject("btnBack.Image"), System.Drawing.Image)
        Me.btnBack.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnBack.Name = "btnBack"
        Me.btnBack.Size = New System.Drawing.Size(23, 22)
        Me.btnBack.Text = "Back"
        '
        'btnForward
        '
        Me.btnForward.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnForward.Image = CType(resources.GetObject("btnForward.Image"), System.Drawing.Image)
        Me.btnForward.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnForward.Name = "btnForward"
        Me.btnForward.Size = New System.Drawing.Size(23, 22)
        Me.btnForward.Text = "Forward"
        '
        'txtURL
        '
        Me.txtURL.AcceptsReturn = True
        Me.txtURL.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.txtURL.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.AllUrl
        Me.txtURL.AutoSize = False
        Me.txtURL.Name = "txtURL"
        Me.txtURL.Size = New System.Drawing.Size(300, 25)
        '
        'btnRefresh
        '
        Me.btnRefresh.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnRefresh.Image = CType(resources.GetObject("btnRefresh.Image"), System.Drawing.Image)
        Me.btnRefresh.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(23, 22)
        Me.btnRefresh.Text = "Refresh"
        '
        'btnGo
        '
        Me.btnGo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnGo.Image = CType(resources.GetObject("btnGo.Image"), System.Drawing.Image)
        Me.btnGo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(23, 22)
        Me.btnGo.Text = "Go"
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(39, 22)
        Me.ToolStripLabel1.Text = "http://"
        '
        'ctlControlCenter
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.gbControlCenter)
        Me.Name = "ctlControlCenter"
        Me.Size = New System.Drawing.Size(571, 463)
        Me.gbControlCenter.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents WebBrowser1 As System.Windows.Forms.WebBrowser
    Friend WithEvents gbControlCenter As System.Windows.Forms.GroupBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnHome As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnBack As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnForward As System.Windows.Forms.ToolStripButton
    Friend WithEvents txtURL As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents btnRefresh As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnGo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel

End Class
