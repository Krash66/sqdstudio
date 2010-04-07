Public Class frmSystem
    Inherits SQDStudio.frmBlank

    Public objThis As New clsSystem
    Dim IsNewObj As Boolean

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents txtDDLLib As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtDTDLib As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIncludeLib As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtCopybookLib As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtSystemDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmbOSType As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtQMgr As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtSystemName As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSystem))
        Me.txtDDLLib = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtDTDLib = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtIncludeLib = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtCopybookLib = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtSystemDesc = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbOSType = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtQMgr = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPort = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtHost = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtSystemName = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 361)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(232, 384)
        Me.cmdOk.TabIndex = 11
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(328, 384)
        Me.cmdCancel.TabIndex = 12
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(424, 384)
        Me.cmdHelp.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.Text = "System definition"
        '
        'PictureBox1
        '
        'Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        'Me.PictureBox1.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstu" & _
        '    "dio\images\FormTop\sq_skyblue.jpg"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(438, 44)
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'txtDDLLib
        '
        Me.txtDDLLib.Location = New System.Drawing.Point(112, 328)
        Me.txtDDLLib.MaxLength = 255
        Me.txtDDLLib.Name = "txtDDLLib"
        Me.txtDDLLib.Size = New System.Drawing.Size(384, 20)
        Me.txtDDLLib.TabIndex = 10
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 328)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 17)
        Me.Label9.TabIndex = 71
        Me.Label9.Text = "DDL Lib"
        '
        'txtDTDLib
        '
        Me.txtDTDLib.Location = New System.Drawing.Point(112, 296)
        Me.txtDTDLib.MaxLength = 255
        Me.txtDTDLib.Name = "txtDTDLib"
        Me.txtDTDLib.Size = New System.Drawing.Size(384, 20)
        Me.txtDTDLib.TabIndex = 9
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(16, 296)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 17)
        Me.Label10.TabIndex = 69
        Me.Label10.Text = "DTD Lib"
        '
        'txtIncludeLib
        '
        Me.txtIncludeLib.Location = New System.Drawing.Point(112, 264)
        Me.txtIncludeLib.MaxLength = 255
        Me.txtIncludeLib.Name = "txtIncludeLib"
        Me.txtIncludeLib.Size = New System.Drawing.Size(384, 20)
        Me.txtIncludeLib.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 264)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 17)
        Me.Label3.TabIndex = 67
        Me.Label3.Text = "Include Lib"
        '
        'txtCopybookLib
        '
        Me.txtCopybookLib.Location = New System.Drawing.Point(112, 232)
        Me.txtCopybookLib.MaxLength = 255
        Me.txtCopybookLib.Name = "txtCopybookLib"
        Me.txtCopybookLib.Size = New System.Drawing.Size(384, 20)
        Me.txtCopybookLib.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 232)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 17)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "Copybook Lib"
        '
        'txtSystemDesc
        '
        Me.txtSystemDesc.Location = New System.Drawing.Point(112, 176)
        Me.txtSystemDesc.MaxLength = 1000
        Me.txtSystemDesc.Multiline = True
        Me.txtSystemDesc.Name = "txtSystemDesc"
        Me.txtSystemDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSystemDesc.Size = New System.Drawing.Size(384, 40)
        Me.txtSystemDesc.TabIndex = 6
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(16, 176)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 17)
        Me.Label8.TabIndex = 63
        Me.Label8.Text = "Description"
        '
        'cmbOSType
        '
        Me.cmbOSType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOSType.Location = New System.Drawing.Point(360, 80)
        Me.cmbOSType.Name = "cmbOSType"
        Me.cmbOSType.Size = New System.Drawing.Size(136, 21)
        Me.cmbOSType.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(280, 80)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(80, 20)
        Me.Label7.TabIndex = 61
        Me.Label7.Text = "System Type"
        '
        'txtQMgr
        '
        Me.txtQMgr.Location = New System.Drawing.Point(112, 144)
        Me.txtQMgr.MaxLength = 128
        Me.txtQMgr.Name = "txtQMgr"
        Me.txtQMgr.Size = New System.Drawing.Size(152, 20)
        Me.txtQMgr.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 144)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 17)
        Me.Label5.TabIndex = 59
        Me.Label5.Text = "Queue Manager"
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(320, 112)
        Me.txtPort.MaxLength = 128
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(176, 20)
        Me.txtPort.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(280, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 20)
        Me.Label6.TabIndex = 57
        Me.Label6.Text = "Port"
        '
        'txtHost
        '
        Me.txtHost.Location = New System.Drawing.Point(112, 112)
        Me.txtHost.MaxLength = 128
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(152, 20)
        Me.txtHost.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(16, 112)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(96, 20)
        Me.Label11.TabIndex = 55
        Me.Label11.Text = "Host"
        '
        'txtSystemName
        '
        Me.txtSystemName.Location = New System.Drawing.Point(112, 80)
        Me.txtSystemName.MaxLength = 128
        Me.txtSystemName.Name = "txtSystemName"
        Me.txtSystemName.Size = New System.Drawing.Size(152, 20)
        Me.txtSystemName.TabIndex = 1
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(16, 80)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(96, 20)
        Me.Label12.TabIndex = 53
        Me.Label12.Text = "System Name"
        '
        'frmSystem
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 423)
        Me.Controls.Add(Me.txtDDLLib)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtDTDLib)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtIncludeLib)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtCopybookLib)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtSystemDesc)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cmbOSType)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtQMgr)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPort)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtHost)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtSystemName)
        Me.Controls.Add(Me.Label12)
        Me.Name = "frmSystem"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "System Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.Label12, 0)
        Me.Controls.SetChildIndex(Me.txtSystemName, 0)
        Me.Controls.SetChildIndex(Me.Label11, 0)
        Me.Controls.SetChildIndex(Me.txtHost, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtPort, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtQMgr, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.cmbOSType, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.txtSystemDesc, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtCopybookLib, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtIncludeLib, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.txtDTDLib, 0)
        Me.Controls.SetChildIndex(Me.Label9, 0)
        Me.Controls.SetChildIndex(Me.txtDDLLib, 0)
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOSType.SelectedIndexChanged, txtHost.TextChanged, txtPort.TextChanged, txtQMgr.TextChanged, txtSystemDesc.TextChanged, txtSystemName.TextChanged, txtCopybookLib.TextChanged, txtDDLLib.TextChanged, txtDTDLib.TextChanged, txtIncludeLib.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False

    End Sub

    Private Sub SetDefaultName()

        Try
            Dim EnvObj As clsEnvironment = objThis.Environment
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = "System" & Count.ToString '& "_" & Strings.Right(EnvObj.Text, 4)
                For Each TestSys As clsSystem In EnvObj.Systems
                    If TestSys.SystemName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtSystemName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmSys SetDefaultName")
        End Try

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

    Public Overrides Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If objThis.IsModified = False Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Exit Sub
        End If

        If MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    '//Returns new Environment object
    Public Function NewObj() As clsSystem

        IsNewObj = True

        StartLoad()

        SetDefaultName()

        InitControls()

        txtSystemName.SelectAll()

        cmbOSType.SelectedIndex = 0

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = objThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If objThis.ValidateNewObject(txtSystemName.Text) = False Or ValidateNewName(txtSystemName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.SystemName = txtSystemName.Text
            objThis.SystemDescription = txtSystemDesc.Text
            objThis.Port = txtPort.Text
            objThis.Host = txtHost.Text
            objThis.QueueManager = txtQMgr.Text
            objThis.OSType = cmbOSType.Text
            objThis.CopybookLib = txtCopybookLib.Text
            objThis.IncludeLib = txtIncludeLib.Text
            objThis.DTDLib = txtDTDLib.Text
            objThis.DDLLib = txtDDLLib.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Sub InitControls()

        cmbOSType.Items.Clear()
        cmbOSType.Items.Add(New Mylist("z/OS", "z/OS"))
        cmbOSType.Items.Add(New Mylist("UNIX", "UNIX"))
        cmbOSType.Items.Add(New Mylist("Windows", "Windows"))
        cmbOSType.SelectedIndex = 0

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Add_Systems)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) ' Handles Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click(sender, e)
            Case Keys.Enter
                cmdOk_Click(sender, e)
            Case Keys.F1
                cmdHelp_Click(sender, e)
        End Select
    End Sub

End Class
