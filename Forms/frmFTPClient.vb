Public Class frmFTPClient
    Inherits System.Windows.Forms.Form
    'Inherits SQDStudio.frmBlank

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
    Friend WithEvents lblHost As System.Windows.Forms.Label
    Friend WithEvents lblPort As System.Windows.Forms.Label
    Friend WithEvents lblUser As System.Windows.Forms.Label
    Friend WithEvents lblPass As System.Windows.Forms.Label
    Friend WithEvents txtHost As System.Windows.Forms.TextBox
    Friend WithEvents txtPort As System.Windows.Forms.TextBox
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseRemote As System.Windows.Forms.Button
    Friend WithEvents cmdUp As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents statusbar As System.Windows.Forms.Label
    Friend WithEvents cmdhelp As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseLocal As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbExtension As System.Windows.Forms.ComboBox
    Friend WithEvents lvRemote As System.Windows.Forms.ListView
    Friend WithEvents cmdDnld As System.Windows.Forms.Button
    Friend WithEvents lvLocal As System.Windows.Forms.ListView
    Friend WithEvents cmdUpld As System.Windows.Forms.Button
    Friend WithEvents txtRemotePath As System.Windows.Forms.TextBox
    Friend WithEvents txtLocalPath As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents txtFTPout As System.Windows.Forms.TextBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFTPClient))
        Me.lblHost = New System.Windows.Forms.Label()
        Me.lblPort = New System.Windows.Forms.Label()
        Me.lblUser = New System.Windows.Forms.Label()
        Me.lblPass = New System.Windows.Forms.Label()
        Me.txtHost = New System.Windows.Forms.TextBox()
        Me.txtPort = New System.Windows.Forms.TextBox()
        Me.txtUser = New System.Windows.Forms.TextBox()
        Me.txtPass = New System.Windows.Forms.TextBox()
        Me.txtRemotePath = New System.Windows.Forms.TextBox()
        Me.cmdBrowseRemote = New System.Windows.Forms.Button()
        Me.cmdUp = New System.Windows.Forms.Button()
        Me.lvRemote = New System.Windows.Forms.ListView()
        Me.cmdDnld = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.txtLocalPath = New System.Windows.Forms.TextBox()
        Me.cmdhelp = New System.Windows.Forms.Button()
        Me.statusbar = New System.Windows.Forms.Label()
        Me.cmdBrowseLocal = New System.Windows.Forms.Button()
        Me.cmbExtension = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lvLocal = New System.Windows.Forms.ListView()
        Me.cmdUpld = New System.Windows.Forms.Button()
        Me.txtFTPout = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblHost
        '
        Me.lblHost.AutoSize = True
        Me.lblHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHost.Location = New System.Drawing.Point(7, 25)
        Me.lblHost.Name = "lblHost"
        Me.lblHost.Size = New System.Drawing.Size(41, 17)
        Me.lblHost.TabIndex = 11
        Me.lblHost.Text = "Host"
        '
        'lblPort
        '
        Me.lblPort.AutoSize = True
        Me.lblPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPort.Location = New System.Drawing.Point(7, 55)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(38, 17)
        Me.lblPort.TabIndex = 14
        Me.lblPort.Text = "Port"
        '
        'lblUser
        '
        Me.lblUser.AutoSize = True
        Me.lblUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUser.Location = New System.Drawing.Point(508, 25)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(81, 17)
        Me.lblUser.TabIndex = 12
        Me.lblUser.Text = "Username"
        '
        'lblPass
        '
        Me.lblPass.AutoSize = True
        Me.lblPass.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPass.Location = New System.Drawing.Point(508, 55)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(77, 17)
        Me.lblPass.TabIndex = 15
        Me.lblPass.Text = "Password"
        '
        'txtHost
        '
        Me.txtHost.Location = New System.Drawing.Point(58, 22)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(262, 22)
        Me.txtHost.TabIndex = 1
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(58, 52)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(57, 22)
        Me.txtPort.TabIndex = 2
        Me.txtPort.Text = "21"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(598, 22)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(255, 22)
        Me.txtUser.TabIndex = 3
        '
        'txtPass
        '
        Me.txtPass.Location = New System.Drawing.Point(598, 51)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.Size = New System.Drawing.Size(255, 22)
        Me.txtPass.TabIndex = 4
        '
        'txtRemotePath
        '
        Me.txtRemotePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRemotePath.Location = New System.Drawing.Point(7, 22)
        Me.txtRemotePath.Name = "txtRemotePath"
        Me.txtRemotePath.Size = New System.Drawing.Size(304, 23)
        Me.txtRemotePath.TabIndex = 5
        '
        'cmdBrowseRemote
        '
        Me.cmdBrowseRemote.Location = New System.Drawing.Point(318, 20)
        Me.cmdBrowseRemote.Name = "cmdBrowseRemote"
        Me.cmdBrowseRemote.Size = New System.Drawing.Size(80, 25)
        Me.cmdBrowseRemote.TabIndex = 6
        Me.cmdBrowseRemote.Text = "Connect"
        '
        'cmdUp
        '
        Me.cmdUp.Image = CType(resources.GetObject("cmdUp.Image"), System.Drawing.Image)
        Me.cmdUp.Location = New System.Drawing.Point(406, 20)
        Me.cmdUp.Name = "cmdUp"
        Me.cmdUp.Size = New System.Drawing.Size(33, 25)
        Me.cmdUp.TabIndex = 7
        '
        'lvRemote
        '
        Me.lvRemote.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvRemote.FullRowSelect = True
        Me.lvRemote.Location = New System.Drawing.Point(518, 168)
        Me.lvRemote.Name = "lvRemote"
        Me.lvRemote.Size = New System.Drawing.Size(438, 268)
        Me.lvRemote.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvRemote.TabIndex = 8
        Me.lvRemote.UseCompatibleStateImageBehavior = False
        '
        'cmdDnld
        '
        Me.cmdDnld.Enabled = False
        Me.cmdDnld.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDnld.Location = New System.Drawing.Point(467, 263)
        Me.cmdDnld.Name = "cmdDnld"
        Me.cmdDnld.Size = New System.Drawing.Size(44, 27)
        Me.cmdDnld.TabIndex = 9
        Me.cmdDnld.Text = "<<"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(737, 525)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(101, 26)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "Close"
        '
        'txtLocalPath
        '
        Me.txtLocalPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtLocalPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalPath.HideSelection = False
        Me.txtLocalPath.Location = New System.Drawing.Point(7, 22)
        Me.txtLocalPath.Name = "txtLocalPath"
        Me.txtLocalPath.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtLocalPath.Size = New System.Drawing.Size(395, 23)
        Me.txtLocalPath.TabIndex = 16
        '
        'cmdhelp
        '
        Me.cmdhelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdhelp.Location = New System.Drawing.Point(845, 525)
        Me.cmdhelp.Name = "cmdhelp"
        Me.cmdhelp.Size = New System.Drawing.Size(101, 26)
        Me.cmdhelp.TabIndex = 11
        Me.cmdhelp.Text = "Help"
        '
        'statusbar
        '
        Me.statusbar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.statusbar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.statusbar.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.statusbar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.statusbar.Location = New System.Drawing.Point(7, 82)
        Me.statusbar.Name = "statusbar"
        Me.statusbar.Size = New System.Drawing.Size(718, 21)
        Me.statusbar.TabIndex = 18
        Me.statusbar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdBrowseLocal
        '
        Me.cmdBrowseLocal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseLocal.Location = New System.Drawing.Point(409, 21)
        Me.cmdBrowseLocal.Name = "cmdBrowseLocal"
        Me.cmdBrowseLocal.Size = New System.Drawing.Size(36, 24)
        Me.cmdBrowseLocal.TabIndex = 19
        Me.cmdBrowseLocal.Text = "..."
        '
        'cmbExtension
        '
        Me.cmbExtension.Location = New System.Drawing.Point(361, 51)
        Me.cmbExtension.Name = "cmbExtension"
        Me.cmbExtension.Size = New System.Drawing.Size(91, 24)
        Me.cmbExtension.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(166, 55)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(188, 17)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "File Type (file extension)"
        '
        'lvLocal
        '
        Me.lvLocal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lvLocal.FullRowSelect = True
        Me.lvLocal.Location = New System.Drawing.Point(7, 168)
        Me.lvLocal.Name = "lvLocal"
        Me.lvLocal.Size = New System.Drawing.Size(453, 268)
        Me.lvLocal.TabIndex = 22
        Me.lvLocal.UseCompatibleStateImageBehavior = False
        '
        'cmdUpld
        '
        Me.cmdUpld.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpld.Location = New System.Drawing.Point(467, 298)
        Me.cmdUpld.Name = "cmdUpld"
        Me.cmdUpld.Size = New System.Drawing.Size(44, 27)
        Me.cmdUpld.TabIndex = 23
        Me.cmdUpld.Text = ">>"
        '
        'txtFTPout
        '
        Me.txtFTPout.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFTPout.BackColor = System.Drawing.SystemColors.Window
        Me.txtFTPout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFTPout.Location = New System.Drawing.Point(7, 22)
        Me.txtFTPout.Multiline = True
        Me.txtFTPout.Name = "txtFTPout"
        Me.txtFTPout.ReadOnly = True
        Me.txtFTPout.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtFTPout.Size = New System.Drawing.Size(718, 56)
        Me.txtFTPout.TabIndex = 24
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtLocalPath)
        Me.GroupBox1.Controls.Add(Me.cmdBrowseLocal)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(7, 108)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(453, 54)
        Me.GroupBox1.TabIndex = 26
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Local Directory"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdUp)
        Me.GroupBox2.Controls.Add(Me.cmdBrowseRemote)
        Me.GroupBox2.Controls.Add(Me.txtRemotePath)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox2.Location = New System.Drawing.Point(518, 108)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(447, 54)
        Me.GroupBox2.TabIndex = 27
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Remote Directory"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblHost)
        Me.GroupBox3.Controls.Add(Me.txtHost)
        Me.GroupBox3.Controls.Add(Me.txtPort)
        Me.GroupBox3.Controls.Add(Me.lblPort)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.cmbExtension)
        Me.GroupBox3.Controls.Add(Me.lblUser)
        Me.GroupBox3.Controls.Add(Me.txtUser)
        Me.GroupBox3.Controls.Add(Me.txtPass)
        Me.GroupBox3.Controls.Add(Me.lblPass)
        Me.GroupBox3.Location = New System.Drawing.Point(7, 14)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(861, 88)
        Me.GroupBox3.TabIndex = 28
        Me.GroupBox3.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.txtFTPout)
        Me.GroupBox4.Controls.Add(Me.statusbar)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox4.Location = New System.Drawing.Point(7, 443)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(732, 108)
        Me.GroupBox4.TabIndex = 29
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Transaction Log"
        '
        'frmFTPClient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(961, 565)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdUpld)
        Me.Controls.Add(Me.lvLocal)
        Me.Controls.Add(Me.cmdhelp)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdDnld)
        Me.Controls.Add(Me.lvRemote)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(979, 612)
        Me.Name = "frmFTPClient"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "FTP File Transfer Client"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Client As FTPClient
    Private StoreDir As String
    Public GottenFiles As Collection
    Private FileType As String
    Private CalledBy As enumCalledFrom


    Public Function BrowseFile(ByVal StoreDir As String, ByVal FileType As String, ByVal CalledBy As enumCalledFrom) As String
        Me.StoreDir = StoreDir
        Me.FileType = FileType
        Me.CalledBy = CalledBy
        'Dim Destination As Collection

        InitFTPCntl()

        Me.ShowDialog()

        BrowseFile = Me.StoreDir
    End Function

    Public Function BrowseFile(ByVal StoreDir As String, ByVal FileType As String) As Collection
        Me.StoreDir = StoreDir
        Me.FileType = FileType
        Me.CalledBy = enumCalledFrom.BY_STRUCTURE
        'Dim Destination As Collection

        InitFTPCntl()

        Me.ShowDialog()
        If Not Me.GottenFiles Is Nothing Then
            BrowseFile = Me.GottenFiles
        Else
            BrowseFile = New Collection
        End If
    End Function

    Private Sub cmdBrowseRemote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseRemote.Click

        Dim LastCursor As Cursor

        LastCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If txtHost.Text <> "" And txtPass.Text <> "" And txtPort.Text <> "" And txtUser.Text <> "" Then
            Client = New FTPClient(Me.txtHost.Text, Me.txtPort.Text, _
                                   Me.txtUser.Text, Me.txtPass.Text, _
                                   Me.txtRemotePath.Text, Me)
            Client.hostNameValue = Me.txtHost.Text
            Client.portValue = Me.txtPort.Text
            Client.userNameValue = Me.txtUser.Text
            Client.dirValue = Me.txtRemotePath.Text
            lvRemote_Update()
        Else
            If (txtHost.Text = "") Then
                MsgBox("Host name is missing", MsgBoxStyle.Critical, MsgTitle)
                txtHost.Focus()
            ElseIf (txtPass.Text = "") Then
                MsgBox("Password is missing", MsgBoxStyle.Critical, MsgTitle)
                txtPass.Focus()
            ElseIf (txtPort.Text = "") Then
                MsgBox("Port number is missing", MsgBoxStyle.Critical, MsgTitle)
                txtPort.Focus()
            ElseIf (txtUser.Text = "") Then
                MsgBox("User ID is missing", MsgBoxStyle.Critical, MsgTitle)
                txtUser.Focus()
            End If
        End If
        Windows.Forms.Cursor.Current = LastCursor

    End Sub

    Private Sub lvRemote_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDnld.Click

        Dim thisFile As FTPFile
        Dim LastCursor As Cursor
        Dim Item As Windows.Forms.ListViewItem
        Dim ext As String
        Dim filename As String

        If txtLocalPath.Text = "" Then
            MsgBox("please select a local folder as the destination for your transfer", MsgBoxStyle.Information, MsgTitle)
            Exit Sub
        End If
        ext = cmbExtension.SelectedItem
        LastCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        If lvRemote.SelectedItems.Count > 0 Then
            Item = lvRemote.SelectedItems(0)
            thisFile = Item.Tag
            If thisFile.Type = FTPFileType.Directory Then
                lvRemote_Update(thisFile.Filename)
            Else
                StoreDir = txtLocalPath.Text
                Me.GottenFiles = New Collection
                Client.InitFTPScript()
                For Each Item In lvRemote.SelectedItems
                    thisFile = Item.Tag

                    filename = thisFile.Filename
                    If (Strings.Right(thisFile.Filename, 4) <> ext) Then
                        filename = thisFile.Filename & ext
                    End If

                    Client.AddGetCmd(thisFile.Filename, Me.StoreDir & "\" & filename)
                    Me.GottenFiles.Add(filename.Clone)
                Next
                Client.GETFILE()
                If (Client.TransRc() = 8) Then 'error
                    statusbar.Text = "Transfer error!!"
                    If (MsgBox("File transfer did not complete successfully.  Would you like to see the log?", MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes) Then
                        Client.ShowLog()
                    End If
                Else
                    statusbar.Text = "Transfer Successful!!"
                End If
            End If
        End If
        Windows.Forms.Cursor.Current = LastCursor
        lvLocal_Update(Me.StoreDir, cmbExtension.SelectedItem)

    End Sub

    Private Sub lvLocal_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpld.Click

        'Dim thisFile As FTPFile
        'Dim file As String
        Dim LastCursor As Cursor
        Dim Item As Windows.Forms.ListViewItem
        'Dim ext As String
        Dim filename As String

        LastCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        If lvLocal.SelectedItems.Count > 0 Then
            Item = lvLocal.SelectedItems(0)
            StoreDir = txtLocalPath.Text
            Me.GottenFiles = New Collection
            Client.InitFTPScript()
            For Each Item In lvLocal.SelectedItems
                filename = Item.Text

                Select Case Client.SystemType
                    Case FTPDType.Unix
                        Client.AddPutCmd(Me.StoreDir & "\" & filename, _
                                       txtRemotePath.Text & "/" & filename)
                    Case FTPDType.MVS
                        Client.AddPutCmd(Me.StoreDir & "\" & filename, _
                               Strings.Left(txtRemotePath.Text, Strings.Len(txtRemotePath.Text) - 1) & _
                               filename & "'")
                    Case FTPDType.MVS_PDS
                        Client.AddPutCmd(Me.StoreDir & "\" & filename, _
                               Strings.Left(txtRemotePath.Text, Strings.Len(txtRemotePath.Text) - 1) & _
                                       "(" & filename.Split(".")(0) & ")'")
                    Case FTPDType.Windows
                        Client.AddPutCmd(Me.StoreDir & "\" & filename, _
                                       txtRemotePath.Text & "\" & filename)
                End Select
                Me.GottenFiles.Add(filename.Clone)
            Next
            Client.PUTFILE()
            If (Client.TransRc() = 8) Then 'error
                statusbar.Text = "Transfer error!!"
                If (MsgBox("File transfer did not complete successfully.  Would you like to see the log?", MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes) Then
                    Client.ShowLog()
                End If
            Else
                statusbar.Text = "Transfer Successful!!"
            End If
        End If
        Windows.Forms.Cursor.Current = LastCursor
        lvRemote_Update()

    End Sub

    Private Sub lvRemote_Update(Optional ByVal NewPath As String = "")

        Dim Files As Collection
        Dim File As FTPFile
        Dim Item As New ListViewItem
        Dim LastCursor As Cursor

        Try
            LastCursor = Windows.Forms.Cursor.Current
            Windows.Forms.Cursor.Current = Cursors.WaitCursor

            If NewPath.Length > 0 Then
                Files = Client.CDLS(NewPath)
            Else
                Files = Client.LIST
            End If

            txtRemotePath.Text = Client.PWD
            lvRemote.Clear()
            lvRemote.View = View.Details
            lvRemote.Columns.Add("Name", 200, HorizontalAlignment.Left)
            lvRemote.Columns.Add("Mode", 100, HorizontalAlignment.Left)
            lvRemote.Columns.Add("Size", 100, HorizontalAlignment.Left)
            lvRemote.Columns.Add("Date", 100, HorizontalAlignment.Left)
            If Files IsNot Nothing Then
                For Each File In Files
                    Item.SubItems.Clear()
                    Item.Text = File.Filename

                    If (File.Type = FTPFileType.Directory) Then
                        Item.SubItems.Add("<DIR>")
                    Else
                        Item.SubItems.Add(File.Mode)
                    End If

                    Item.SubItems.Add(File.Size)
                    Item.SubItems.Add(File.ModDate)
                    Item.Tag = File.Clone
                    lvRemote.Items.Add(Item.Clone)
                Next
            End If

            Windows.Forms.Cursor.Current = LastCursor

        Catch ex As Exception
            LogError(ex, "frmFTPclient lvRemote_Update")
        End Try

    End Sub

    Private Sub lvLocal_Update(ByVal path As String, ByVal ext As String)

        Dim Files() As String
        Dim File As String
        Dim filenames() As String
        Dim filename As String
        Dim Item As New ListViewItem
        Dim list As Boolean

        If path <> "" Then '//added by tkarasc 9/28/06 to fix ticket#623
            Files = System.IO.Directory.GetFiles(path)

            lvLocal.Clear()
            lvLocal.View = View.Details
            lvLocal.Columns.Add("Name", 200, HorizontalAlignment.Left)
            lvLocal.Columns.Add("Date", 100, HorizontalAlignment.Left)

            For Each File In Files
                filenames = File.Split("\")
                filename = filenames(filenames.Length - 1)
                If (ext = ".*") Then
                    list = True
                ElseIf (Strings.Right(filename, 4) = ext) Then
                    list = True
                Else
                    list = False
                End If

                If (list = True) Then
                    Item.SubItems.Clear()
                    Item.Text = filename
                    Item.SubItems.Add(System.IO.File.GetCreationTime(File))
                    lvLocal.Items.Add(Item.Clone)
                End If
            Next
        End If

    End Sub

    Private Sub cmdUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUp.Click

        Dim LastCursor As Cursor
        LastCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        If txtHost.Text <> "" And txtPass.Text <> "" And txtPort.Text <> "" And txtUser.Text <> "" And txtRemotePath.Text <> "" Then
            lvRemote_Update("..")
        End If
        Windows.Forms.Cursor.Current = LastCursor

    End Sub

    Private Sub frmFTPClient_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim LastCursor As Cursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        Dim ftpobj As New FTPClient

        ftpobj.hostNameValue = modGeneral.FTPHostName
        ftpobj.portValue = modGeneral.FTPPort
        ftpobj.userNameValue = modGeneral.FTPUserID
        ftpobj.dirValue = modGeneral.FTPRemoteDir

        If (ftpobj.portValue = "") Then
            ftpobj.portValue = "21"
        End If

        txtHost.Text = ftpobj.hostNameValue
        txtPort.Text = ftpobj.portValue
        txtUser.Text = ftpobj.userNameValue
        txtRemotePath.Text = ftpobj.dirValue
        txtLocalPath.Text = StoreDir

        If (CalledBy = modDeclares.enumCalledFrom.BY_ENVIRONMENT) Then
            Select Case FileType
                Case ".dbd"
                    Me.Text = "Downloading to Local DBD Directory"
                Case ".dtd"
                    Me.Text = "Downloading to Local DTD Directory"
                Case ".ddl"
                    Me.Text = "Downloading to Local DDL Directory"
                Case ".cob"
                    Me.Text = "Downloading to Local COBOL Directory"
                Case ".h"
                    Me.Text = "Downloading to Local C Directory"
                Case ".dml"
                    Me.Text = "Downloading to Local DML Directory"
            End Select
        Else
            Me.Text = "Remote File Open"
            txtLocalPath.Enabled = True
        End If
        'Windows.Forms.Cursor.Current = LastCursor
    End Sub

    Private Sub lvRemote_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvRemote.GotFocus
        If Me.lvRemote.SelectedItems.Count = 0 Then
            cmdDnld.Enabled = False
        End If
        statusbar.Text = ""
    End Sub

    Private Sub lvLocal_GotFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvLocal.GotFocus
        If Me.lvLocal.SelectedItems.Count = 0 Then
            cmdUpld.Enabled = False
        ElseIf Client Is Nothing Then
            cmdUpld.Enabled = False
        End If
        statusbar.Text = ""
    End Sub

    Private Sub lvRemote_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvRemote.Click
        If Me.lvRemote.SelectedItems.Count > 0 Then
            cmdDnld.Enabled = True
            cmdUpld.Enabled = False
        End If
    End Sub

    Private Sub cmdBrowseLocal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseLocal.Click
        dlgBrowseFolder.SelectedPath = txtLocalPath.Text
        If dlgBrowseFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtLocalPath.Text = dlgBrowseFolder.SelectedPath
            txtLocalPath.ScrollToCaret()
            Me.StoreDir = txtLocalPath.Text
            lvLocal_Update(Me.StoreDir, cmbExtension.SelectedItem)
        End If
    End Sub

    Private Sub cmbExtension_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExtension.SelectedIndexChanged
        lvLocal_Update(Me.StoreDir, cmbExtension.SelectedItem)
    End Sub

    Private Sub lvLocal_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvLocal.Click
        If Me.lvLocal.SelectedItems.Count > 0 And Not Client Is Nothing Then
            cmdUpld.Enabled = True
            cmdDnld.Enabled = False
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub frmFTPClient_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        modGeneral.FTPPort = txtPort.Text
        modGeneral.FTPRemoteDir = txtRemotePath.Text
        modGeneral.FTPUserID = txtUser.Text
        modGeneral.FTPHostName = txtHost.Text

        'If (Not Client Is Nothing) Then
        '    Client.DelWorkFiles()
        'End If
    End Sub

    Private Sub lvLocal_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvLocal.DoubleClick
    End Sub

    Private Sub lvRemote_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvRemote.DoubleClick
        Dim files As Collection
        Dim file As New FTPFile
        Dim newdir As String
        Dim newpath As String

        If (lvRemote.Items.Count > 0) Then
            files = Client.Files
            newpath = lvRemote.SelectedItems(0).Text

            For Each file In files
                If (newpath = file.Filename And file.Type = FTPFileType.Directory) Then
                    newdir = Client.CurrentWorkingDirectory
                    If Client.SystemType = FTPDType.MVS Or Client.SystemType = FTPDType.MVS_PDS Then
                        newdir = Strings.Left(newdir, Strings.Len(newdir) - 1) & file.Filename & "'"
                    ElseIf Client.SystemType = FTPDType.Unix Then
                        newdir = newdir & "/" & file.Filename
                    Else
                        newdir = newdir & "\" & file.Filename
                    End If
                    Client.CurrentWorkingDirectory = newdir
                    lvRemote_Update()
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub InitFTPCntl()
        cmdDnld.Enabled = False
        cmdUpld.Enabled = False

        cmbExtension.Items.Add(".cob")
        cmbExtension.Items.Add(".dbd")
        cmbExtension.Items.Add(".dtd")
        cmbExtension.Items.Add(".xml")
        cmbExtension.Items.Add(".ddl")
        cmbExtension.Items.Add(".dml")
        cmbExtension.Items.Add(".h")
        cmbExtension.Items.Add(".sqd")
        cmbExtension.Items.Add(".inl")
        cmbExtension.Items.Add(".tmp")
        cmbExtension.Items.Add(".*")

        cmbExtension.SelectedItem = FileType

    End Sub

    Private Sub cmdhelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdhelp.Click

        Select Case Me.CalledBy
            Case enumCalledFrom.BY_ENVIRONMENT
                ShowHelp(modHelp.HHId.H_Add_an_Environment)
            Case enumCalledFrom.BY_SCRIPTGEN
                ShowHelp(modHelp.HHId.H_Generate_Script)
            Case enumCalledFrom.BY_STRUCTURE
                ShowHelp(modHelp.HHId.H_Add_Structures)
            Case Else

        End Select

    End Sub

    Private Sub txtLocalPath_TextChanged(ByVal sender As System.Object, ByVal e As KeyEventArgs) Handles txtLocalPath.KeyDown

        If e.KeyCode = Keys.Enter Then
            If System.IO.Directory.Exists(txtLocalPath.Text) = True Then
                Me.StoreDir = txtLocalPath.Text
                lvLocal_Update(Me.StoreDir, cmbExtension.SelectedItem)
                dlgBrowseFolder.SelectedPath = txtLocalPath.Text
            Else
                MsgBox("Folder does not exist," & Chr(13) & _
                       "Please correct Directory Path" & Chr(13) & _
                       "or use '...'button to browse for Folder.", MsgBoxStyle.Information, "Invalid Folder")
                txtLocalPath.Focus()
            End If
        End If

    End Sub

End Class