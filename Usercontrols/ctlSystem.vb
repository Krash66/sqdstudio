Public Class ctlSystem
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)
   
    Dim objThis As New clsSystem
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents gbConn As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbLib As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDBDLib As System.Windows.Forms.TextBox
    Dim IsNewObj As Boolean


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSystemDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmbOSType As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtQMgr As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSystemName As System.Windows.Forms.TextBox
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtCopybookLib As System.Windows.Forms.TextBox
    Friend WithEvents txtIncludeLib As System.Windows.Forms.TextBox
    Friend WithEvents txtDDLLib As System.Windows.Forms.TextBox
    Friend WithEvents txtDTDLib As System.Windows.Forms.TextBox
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtSystemDesc = New System.Windows.Forms.TextBox
        Me.cmbOSType = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtQMgr = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPort = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtHost = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSystemName = New System.Windows.Forms.TextBox
        Me.lblName = New System.Windows.Forms.Label
        Me.txtCopybookLib = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtIncludeLib = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDDLLib = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtDTDLib = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.gbConn = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbLib = New System.Windows.Forms.GroupBox
        Me.txtDBDLib = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox2.SuspendLayout()
        Me.gbConn.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbLib.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(388, 477)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 11
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(468, 477)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 12
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 461)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(612, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'txtSystemDesc
        '
        Me.txtSystemDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSystemDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSystemDesc.Location = New System.Drawing.Point(6, 19)
        Me.txtSystemDesc.MaxLength = 1000
        Me.txtSystemDesc.Multiline = True
        Me.txtSystemDesc.Name = "txtSystemDesc"
        Me.txtSystemDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSystemDesc.Size = New System.Drawing.Size(486, 52)
        Me.txtSystemDesc.TabIndex = 6
        '
        'cmbOSType
        '
        Me.cmbOSType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOSType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbOSType.Location = New System.Drawing.Point(50, 45)
        Me.cmbOSType.Name = "cmbOSType"
        Me.cmbOSType.Size = New System.Drawing.Size(240, 21)
        Me.cmbOSType.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 48)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(34, 14)
        Me.Label7.TabIndex = 41
        Me.Label7.Text = "Type"
        '
        'txtQMgr
        '
        Me.txtQMgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQMgr.Location = New System.Drawing.Point(331, 19)
        Me.txtQMgr.MaxLength = 128
        Me.txtQMgr.Name = "txtQMgr"
        Me.txtQMgr.Size = New System.Drawing.Size(161, 20)
        Me.txtQMgr.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(231, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(94, 14)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Queue Manager"
        '
        'txtPort
        '
        Me.txtPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPort.Location = New System.Drawing.Point(57, 45)
        Me.txtPort.MaxLength = 128
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(152, 20)
        Me.txtPort.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 48)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(30, 14)
        Me.Label6.TabIndex = 37
        Me.Label6.Text = "Port"
        '
        'txtHost
        '
        Me.txtHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHost.Location = New System.Drawing.Point(57, 19)
        Me.txtHost.MaxLength = 128
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(152, 20)
        Me.txtHost.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 22)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 14)
        Me.Label4.TabIndex = 35
        Me.Label4.Text = "Host"
        '
        'txtSystemName
        '
        Me.txtSystemName.Location = New System.Drawing.Point(50, 19)
        Me.txtSystemName.MaxLength = 128
        Me.txtSystemName.Name = "txtSystemName"
        Me.txtSystemName.Size = New System.Drawing.Size(240, 20)
        Me.txtSystemName.TabIndex = 1
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(6, 22)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(38, 14)
        Me.lblName.TabIndex = 33
        Me.lblName.Text = "Name"
        '
        'txtCopybookLib
        '
        Me.txtCopybookLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCopybookLib.Location = New System.Drawing.Point(75, 50)
        Me.txtCopybookLib.MaxLength = 255
        Me.txtCopybookLib.Name = "txtCopybookLib"
        Me.txtCopybookLib.Size = New System.Drawing.Size(423, 20)
        Me.txtCopybookLib.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(63, 14)
        Me.Label1.TabIndex = 45
        Me.Label1.Text = "Copybook"
        '
        'txtIncludeLib
        '
        Me.txtIncludeLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIncludeLib.Location = New System.Drawing.Point(75, 76)
        Me.txtIncludeLib.MaxLength = 255
        Me.txtIncludeLib.Name = "txtIncludeLib"
        Me.txtIncludeLib.Size = New System.Drawing.Size(423, 20)
        Me.txtIncludeLib.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 79)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(47, 14)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Include"
        '
        'txtDDLLib
        '
        Me.txtDDLLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDDLLib.Location = New System.Drawing.Point(75, 128)
        Me.txtDDLLib.MaxLength = 255
        Me.txtDDLLib.Name = "txtDDLLib"
        Me.txtDDLLib.Size = New System.Drawing.Size(423, 20)
        Me.txtDDLLib.TabIndex = 10
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 131)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(28, 14)
        Me.Label9.TabIndex = 51
        Me.Label9.Text = "DDL"
        '
        'txtDTDLib
        '
        Me.txtDTDLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDTDLib.Location = New System.Drawing.Point(75, 102)
        Me.txtDTDLib.MaxLength = 255
        Me.txtDTDLib.Name = "txtDTDLib"
        Me.txtDTDLib.Size = New System.Drawing.Size(423, 20)
        Me.txtDTDLib.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 105)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(28, 14)
        Me.Label10.TabIndex = 49
        Me.Label10.Text = "DTD"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(548, 477)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 13
        Me.cmdHelp.Text = "&Help"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.gbConn)
        Me.GroupBox2.Controls.Add(Me.gbDesc)
        Me.GroupBox2.Controls.Add(Me.txtSystemName)
        Me.GroupBox2.Controls.Add(Me.lblName)
        Me.GroupBox2.Controls.Add(Me.cmbOSType)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(510, 240)
        Me.GroupBox2.TabIndex = 52
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "System Properties"
        '
        'gbConn
        '
        Me.gbConn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbConn.Controls.Add(Me.txtHost)
        Me.gbConn.Controls.Add(Me.Label4)
        Me.gbConn.Controls.Add(Me.txtQMgr)
        Me.gbConn.Controls.Add(Me.Label5)
        Me.gbConn.Controls.Add(Me.Label6)
        Me.gbConn.Controls.Add(Me.txtPort)
        Me.gbConn.ForeColor = System.Drawing.Color.White
        Me.gbConn.Location = New System.Drawing.Point(6, 156)
        Me.gbConn.Name = "gbConn"
        Me.gbConn.Size = New System.Drawing.Size(498, 78)
        Me.gbConn.TabIndex = 43
        Me.gbConn.TabStop = False
        Me.gbConn.Text = "System Connection Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtSystemDesc)
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 72)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(498, 81)
        Me.gbDesc.TabIndex = 42
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbLib
        '
        Me.gbLib.Controls.Add(Me.Label3)
        Me.gbLib.Controls.Add(Me.txtDBDLib)
        Me.gbLib.Controls.Add(Me.txtCopybookLib)
        Me.gbLib.Controls.Add(Me.Label1)
        Me.gbLib.Controls.Add(Me.txtIncludeLib)
        Me.gbLib.Controls.Add(Me.Label9)
        Me.gbLib.Controls.Add(Me.txtDDLLib)
        Me.gbLib.Controls.Add(Me.Label10)
        Me.gbLib.Controls.Add(Me.txtDTDLib)
        Me.gbLib.Controls.Add(Me.Label2)
        Me.gbLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLib.ForeColor = System.Drawing.Color.White
        Me.gbLib.Location = New System.Drawing.Point(3, 243)
        Me.gbLib.Name = "gbLib"
        Me.gbLib.Size = New System.Drawing.Size(510, 162)
        Me.gbLib.TabIndex = 53
        Me.gbLib.TabStop = False
        Me.gbLib.Text = "Local System Libraries"
        '
        'txtDBDLib
        '
        Me.txtDBDLib.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDBDLib.Location = New System.Drawing.Point(75, 24)
        Me.txtDBDLib.MaxLength = 255
        Me.txtDBDLib.Name = "txtDBDLib"
        Me.txtDBDLib.Size = New System.Drawing.Size(423, 20)
        Me.txtDBDLib.TabIndex = 52
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(28, 14)
        Me.Label3.TabIndex = 53
        Me.Label3.Text = "DBD"
        '
        'ctlSystem
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.gbLib)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlSystem"
        Me.Size = New System.Drawing.Size(628, 509)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.gbConn.ResumeLayout(False)
        Me.gbConn.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbLib.ResumeLayout(False)
        Me.gbLib.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtSystemName.Enabled = True '//added on 7/24 to disable object name editing

        InitControls()

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsSystem

        ClearControls(Me.Controls)

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            '// First Check Validity before Saving
            If ValidateNewName(txtSystemName.Text) = False Then
                Save = False
                Me.Cursor = Cursors.Default
                Exit Function
            End If
            If objThis.SystemName <> txtSystemName.Text Then
                objThis.IsRenamed = RenameSystem(objThis, txtSystemName.Text)
            End If
            If objThis.IsRenamed = False Then
                txtSystemName.Text = objThis.SystemName
            End If

            objThis.SystemName = txtSystemName.Text
            objThis.SystemDescription = txtSystemDesc.Text
            objThis.Port = txtPort.Text
            objThis.Host = txtHost.Text
            objThis.QueueManager = txtQMgr.Text
            objThis.OSType = cmbOSType.Text
            objThis.DBDLib = txtDBDLib.Text
            objThis.CopybookLib = txtCopybookLib.Text
            objThis.IncludeLib = txtIncludeLib.Text
            objThis.DTDLib = txtDTDLib.Text
            objThis.DDLLib = txtDDLLib.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    Save = False
                    Me.Cursor = Cursors.Default
                    Exit Function
                End If
            Else
                If objThis.Save() = False Then
                    Save = False
                    Me.Cursor = Cursors.Default
                    Exit Function
                End If
            End If
            Me.Cursor = Cursors.Default

            Save = True
            cmdSave.Enabled = False
            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If
            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex, "ctlSystem Save")
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Public Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Save()

    End Sub

    Public Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Dim ret As MsgBoxResult

        Try
            If objThis.IsModified = True Then
                ret = MsgBox("Do you want to save change(s) made to the opened form?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel, MsgTitle)
                If ret = MsgBoxResult.Yes Then
                    Save()
                ElseIf ret = MsgBoxResult.No Then
                    objThis.IsModified = False
                ElseIf ret = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Me.Visible = False
            RaiseEvent Closed(Me, objThis)


        Catch ex As Exception
            LogError(ex)
        End Try

    End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOSType.SelectedIndexChanged, txtHost.TextChanged, txtPort.TextChanged, txtQMgr.TextChanged, txtSystemDesc.TextChanged, txtSystemName.TextChanged, txtCopybookLib.TextChanged, txtDDLLib.TextChanged, txtDTDLib.TextChanged, txtIncludeLib.TextChanged, txtDBDLib.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Public Function EditObj(ByVal obj As INode) As clsSystem

        Try
            IsNewObj = False
            Me.Cursor = Cursors.WaitCursor

            StartLoad()

            objThis = obj '//Load the form env object
            objThis.LoadMe()

            UpdateFields()

            EndLoad()

            EditObj = objThis
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            LogError(ex, "ctlSystem EditObj")
            Return Nothing
            Me.Cursor = Cursors.Default
        End Try

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtSystemName.Text = objThis.SystemName
        txtPort.Text = objThis.Port
        txtHost.Text = objThis.Host
        txtQMgr.Text = objThis.QueueManager
        cmbOSType.Text = objThis.OSType
        txtSystemDesc.Text = objThis.SystemDescription
        txtDBDLib.Text = objThis.DBDLib
        txtCopybookLib.Text = objThis.CopybookLib
        txtIncludeLib.Text = objThis.IncludeLib
        txtDTDLib.Text = objThis.DTDLib
        txtDDLLib.Text = objThis.DDLLib

    End Function

    Sub InitControls()

        cmbOSType.Items.Clear()
        cmbOSType.Items.Add(New Mylist("z/OS", "z/OS"))
        cmbOSType.Items.Add(New Mylist("UNIX", "UNIX"))
        cmbOSType.Items.Add(New Mylist("Windows", "Windows"))
        cmbOSType.SelectedIndex = 0

    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Systems)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbOSType.KeyDown, txtCopybookLib.KeyDown, txtDDLLib.KeyDown, txtDTDLib.KeyDown, txtHost.KeyDown, txtIncludeLib.KeyDown, txtPort.KeyDown, txtQMgr.KeyDown, txtSystemDesc.KeyDown, txtSystemName.KeyDown, txtDBDLib.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

End Class
