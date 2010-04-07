'Public Class frmInput
'    Inherits System.Windows.Forms.Form

'    Public Function GetInput(Optional ByVal Title As String = "Enter name of workvariable", Optional ByVal Prompt As String = "Workvariable name", Optional ByVal DefaultVal As String = "Workvar1") As String
'        Me.Text = Title
'        Me.lblPrompt.Text = Prompt
'        Me.txtInput.Text = DefaultVal

'        Select Case Me.ShowDialog
'            Case DialogResult.OK
'                GetInput = Trim(txtInput.Text)
'            Case Else
'                GetInput = ""
'        End Select

'    End Function
'#Region " Windows Form Designer generated code "

'    Public Sub New()
'        MyBase.New()

'        'This call is required by the Windows Form Designer.
'        InitializeComponent()

'        'Add any initialization after the InitializeComponent() call

'    End Sub

'    'Form overrides dispose to clean up the component list.
'    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
'        If disposing Then
'            If Not (components Is Nothing) Then
'                components.Dispose()
'            End If
'        End If
'        MyBase.Dispose(disposing)
'    End Sub

'    'Required by the Windows Form Designer
'    Private components As System.ComponentModel.IContainer

'    'NOTE: The following procedure is required by the Windows Form Designer
'    'It can be modified using the Windows Form Designer.  
'    'Do not modify it using the code editor.
'    Friend WithEvents cmdOk As System.Windows.Forms.Button
'    Friend WithEvents cmdCancel As System.Windows.Forms.Button
'    Friend WithEvents txtInput As System.Windows.Forms.TextBox
'    Friend WithEvents lblPrompt As System.Windows.Forms.Label
'    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
'        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmInput))
'        Me.txtInput = New System.Windows.Forms.TextBox
'        Me.cmdOk = New System.Windows.Forms.Button
'        Me.cmdCancel = New System.Windows.Forms.Button
'        Me.lblPrompt = New System.Windows.Forms.Label
'        Me.SuspendLayout()
'        '
'        'txtInput
'        '
'        Me.txtInput.Location = New System.Drawing.Point(120, 8)
'        Me.txtInput.MaxLength = 128
'        Me.txtInput.Name = "txtInput"
'        Me.txtInput.Size = New System.Drawing.Size(144, 20)
'        Me.txtInput.TabIndex = 0
'        Me.txtInput.Text = "workvariable1"
'        '
'        'cmdOk
'        '
'        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
'        Me.cmdOk.Location = New System.Drawing.Point(96, 72)
'        Me.cmdOk.Name = "cmdOk"
'        Me.cmdOk.Size = New System.Drawing.Size(80, 24)
'        Me.cmdOk.TabIndex = 1
'        Me.cmdOk.Text = "&OK"
'        '
'        'cmdCancel
'        '
'        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
'        Me.cmdCancel.Location = New System.Drawing.Point(184, 72)
'        Me.cmdCancel.Name = "cmdCancel"
'        Me.cmdCancel.Size = New System.Drawing.Size(80, 24)
'        Me.cmdCancel.TabIndex = 1
'        Me.cmdCancel.Text = "&Cancel"
'        '
'        'lblPrompt
'        '
'        Me.lblPrompt.Location = New System.Drawing.Point(8, 16)
'        Me.lblPrompt.Name = "lblPrompt"
'        Me.lblPrompt.Size = New System.Drawing.Size(112, 16)
'        Me.lblPrompt.TabIndex = 2
'        Me.lblPrompt.Text = "Work Variable Name"
'        '
'        'frmInput
'        '
'        Me.AcceptButton = Me.cmdOk
'        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
'        Me.ClientSize = New System.Drawing.Size(272, 109)
'        Me.ControlBox = False
'        Me.Controls.Add(Me.lblPrompt)
'        Me.Controls.Add(Me.cmdOk)
'        Me.Controls.Add(Me.txtInput)
'        Me.Controls.Add(Me.cmdCancel)
'        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
'        Me.MaximizeBox = False
'        Me.MinimizeBox = False
'        Me.Name = "frmInput"
'        Me.ShowInTaskbar = False
'        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
'        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
'        Me.Text = "Enter work variable name"
'        Me.ResumeLayout(False)

'    End Sub

'#End Region

'    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
'        Me.DialogResult = DialogResult.OK
'        Me.Close()
'    End Sub

'    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
'        Me.DialogResult = DialogResult.Cancel
'        Me.Close()
'    End Sub
'End Class
