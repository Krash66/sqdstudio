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
    Friend WithEvents lblRemote As System.Windows.Forms.Label
    Friend WithEvents lblLocal As System.Windows.Forms.Label
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
    Friend WithEvents txtFTPout As System.Windows.Forms.TextBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmFTPClient))
        Me.lblHost = New System.Windows.Forms.Label
        Me.lblPort = New System.Windows.Forms.Label
        Me.lblUser = New System.Windows.Forms.Label
        Me.lblPass = New System.Windows.Forms.Label
        Me.txtHost = New System.Windows.Forms.TextBox
        Me.txtPort = New System.Windows.Forms.TextBox
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.txtPass = New System.Windows.Forms.TextBox
        Me.txtRemotePath = New System.Windows.Forms.TextBox
        Me.cmdBrowseRemote = New System.Windows.Forms.Button
        Me.cmdUp = New System.Windows.Forms.Button
        Me.lvRemote = New System.Windows.Forms.ListView
        Me.cmdDnld = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.lblRemote = New System.Windows.Forms.Label
        Me.txtLocalPath = New System.Windows.Forms.TextBox
        Me.lblLocal = New System.Windows.Forms.Label
        Me.cmdhelp = New System.Windows.Forms.Button
        Me.statusbar = New System.Windows.Forms.Label
        Me.cmdBrowseLocal = New System.Windows.Forms.Button
        Me.cmbExtension = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lvLocal = New System.Windows.Forms.ListView
        Me.cmdUpld = New System.Windows.Forms.Button
        Me.txtFTPout = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'lblHost
        '
        Me.lblHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHost.Location = New System.Drawing.Point(12, 12)
        Me.lblHost.Name = "lblHost"
        Me.lblHost.Size = New System.Drawing.Size(36, 16)
        Me.lblHost.TabIndex = 11
        Me.lblHost.Text = "Host"
        '
        'lblPort
        '
        Me.lblPort.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPort.Location = New System.Drawing.Point(12, 39)
        Me.lblPort.Name = "lblPort"
        Me.lblPort.Size = New System.Drawing.Size(36, 16)
        Me.lblPort.TabIndex = 14
        Me.lblPort.Text = "Port"
        '
        'lblUser
        '
        Me.lblUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUser.Location = New System.Drawing.Point(429, 12)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(66, 16)
        Me.lblUser.TabIndex = 12
        Me.lblUser.Text = "Username"
        '
        'lblPass
        '
        Me.lblPass.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPass.Location = New System.Drawing.Point(429, 39)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(69, 16)
        Me.lblPass.TabIndex = 15
        Me.lblPass.Text = "Password"
        '
        'txtHost
        '
        Me.txtHost.Location = New System.Drawing.Point(45, 9)
        Me.txtHost.Name = "txtHost"
        Me.txtHost.Size = New System.Drawing.Size(219, 20)
        Me.txtHost.TabIndex = 1
        '
        'txtPort
        '
        Me.txtPort.Location = New System.Drawing.Point(45, 36)
        Me.txtPort.Name = "txtPort"
        Me.txtPort.Size = New System.Drawing.Size(48, 20)
        Me.txtPort.TabIndex = 2
        Me.txtPort.Text = "21"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(501, 9)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(213, 20)
        Me.txtUser.TabIndex = 3
        '
        'txtPass
        '
        Me.txtPass.Location = New System.Drawing.Point(501, 36)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.Size = New System.Drawing.Size(213, 20)
        Me.txtPass.TabIndex = 4
        '
        'txtRemotePath
        '
        Me.txtRemotePath.Location = New System.Drawing.Point(432, 90)
        Me.txtRemotePath.Name = "txtRemotePath"
        Me.txtRemotePath.Size = New System.Drawing.Size(265, 20)
        Me.txtRemotePath.TabIndex = 5
        '
        'cmdBrowseRemote
        '
        Me.cmdBrowseRemote.Location = New System.Drawing.Point(703, 90)
        Me.cmdBrowseRemote.Name = "cmdBrowseRemote"
        Me.cmdBrowseRemote.Size = New System.Drawing.Size(67, 22)
        Me.cmdBrowseRemote.TabIndex = 6
        Me.cmdBrowseRemote.Text = "Connect"
        '
        'cmdUp
        '
        Me.cmdUp.Image = CType(resources.GetObject("cmdUp.Image"), System.Drawing.Image)
        Me.cmdUp.Location = New System.Drawing.Point(776, 90)
        Me.cmdUp.Name = "cmdUp"
        Me.cmdUp.Size = New System.Drawing.Size(28, 22)
        Me.cmdUp.TabIndex = 7
        '
        'lvRemote
        '
        Me.lvRemote.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvRemote.FullRowSelect = True
        Me.lvRemote.Location = New System.Drawing.Point(432, 118)
        Me.lvRemote.Name = "lvRemote"
        Me.lvRemote.Size = New System.Drawing.Size(372, 266)
        Me.lvRemote.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvRemote.TabIndex = 8
        Me.lvRemote.UseCompatibleStateImageBehavior = False
        '
        'cmdDnld
        '
        Me.cmdDnld.Enabled = False
        Me.cmdDnld.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdDnld.Location = New System.Drawing.Point(384, 228)
        Me.cmdDnld.Name = "cmdDnld"
        Me.cmdDnld.Size = New System.Drawing.Size(42, 23)
        Me.cmdDnld.TabIndex = 9
        Me.cmdDnld.Text = "<<"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(622, 461)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(84, 23)
        Me.cmdCancel.TabIndex = 10
        Me.cmdCancel.Text = "Close"
        '
        'lblRemote
        '
        Me.lblRemote.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRemote.Location = New System.Drawing.Point(429, 67)
        Me.lblRemote.Name = "lblRemote"
        Me.lblRemote.Size = New System.Drawing.Size(130, 16)
        Me.lblRemote.TabIndex = 13
        Me.lblRemote.Text = "Remote Directory"
        '
        'txtLocalPath
        '
        Me.txtLocalPath.Location = New System.Drawing.Point(6, 90)
        Me.txtLocalPath.Name = "txtLocalPath"
        Me.txtLocalPath.Size = New System.Drawing.Size(336, 20)
        Me.txtLocalPath.TabIndex = 16
        '
        'lblLocal
        '
        Me.lblLocal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLocal.Location = New System.Drawing.Point(12, 67)
        Me.lblLocal.Name = "lblLocal"
        Me.lblLocal.Size = New System.Drawing.Size(110, 16)
        Me.lblLocal.TabIndex = 17
        Me.lblLocal.Text = "Local Directory"
        '
        'cmdhelp
        '
        Me.cmdhelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdhelp.Location = New System.Drawing.Point(712, 461)
        Me.cmdhelp.Name = "cmdhelp"
        Me.cmdhelp.Size = New System.Drawing.Size(84, 23)
        Me.cmdhelp.TabIndex = 11
        Me.cmdhelp.Text = "Help"
        '
        'statusbar
        '
        Me.statusbar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.statusbar.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.statusbar.Font = New System.Drawing.Font("Times New Roman", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.statusbar.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.statusbar.Location = New System.Drawing.Point(6, 469)
        Me.statusbar.Name = "statusbar"
        Me.statusbar.Size = New System.Drawing.Size(586, 18)
        Me.statusbar.TabIndex = 18
        Me.statusbar.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'cmdBrowseLocal
        '
        Me.cmdBrowseLocal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseLocal.Location = New System.Drawing.Point(348, 91)
        Me.cmdBrowseLocal.Name = "cmdBrowseLocal"
        Me.cmdBrowseLocal.Size = New System.Drawing.Size(30, 21)
        Me.cmdBrowseLocal.TabIndex = 19
        Me.cmdBrowseLocal.Text = "..."
        '
        'cmbExtension
        '
        Me.cmbExtension.Location = New System.Drawing.Point(302, 64)
        Me.cmbExtension.Name = "cmbExtension"
        Me.cmbExtension.Size = New System.Drawing.Size(76, 21)
        Me.cmbExtension.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(145, 67)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 18)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "File Type (file extension)"
        '
        'lvLocal
        '
        Me.lvLocal.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lvLocal.Location = New System.Drawing.Point(6, 118)
        Me.lvLocal.Name = "lvLocal"
        Me.lvLocal.Size = New System.Drawing.Size(372, 266)
        Me.lvLocal.TabIndex = 22
        Me.lvLocal.UseCompatibleStateImageBehavior = False
        '
        'cmdUpld
        '
        Me.cmdUpld.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUpld.Location = New System.Drawing.Point(384, 258)
        Me.cmdUpld.Name = "cmdUpld"
        Me.cmdUpld.Size = New System.Drawing.Size(42, 24)
        Me.cmdUpld.TabIndex = 23
        Me.cmdUpld.Text = ">>"
        '
        'txtFTPout
        '
        Me.txtFTPout.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtFTPout.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.txtFTPout.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFTPout.Location = New System.Drawing.Point(6, 398)
        Me.txtFTPout.Multiline = True
        Me.txtFTPout.Name = "txtFTPout"
        Me.txtFTPout.ReadOnly = True
        Me.txtFTPout.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtFTPout.Size = New System.Drawing.Size(586, 61)
        Me.txtFTPout.TabIndex = 24
        '
        'frmFTPClient
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(808, 496)
        Me.Controls.Add(Me.txtFTPout)
        Me.Controls.Add(Me.cmdUpld)
        Me.Controls.Add(Me.lvLocal)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cmbExtension)
        Me.Controls.Add(Me.cmdBrowseLocal)
        Me.Controls.Add(Me.statusbar)
        Me.Controls.Add(Me.cmdhelp)
        Me.Controls.Add(Me.lblLocal)
        Me.Controls.Add(Me.txtLocalPath)
        Me.Controls.Add(Me.txtPass)
        Me.Controls.Add(Me.txtUser)
        Me.Controls.Add(Me.txtPort)
        Me.Controls.Add(Me.txtHost)
        Me.Controls.Add(Me.txtRemotePath)
        Me.Controls.Add(Me.lblRemote)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdDnld)
        Me.Controls.Add(Me.lblPass)
        Me.Controls.Add(Me.lblUser)
        Me.Controls.Add(Me.lblPort)
        Me.Controls.Add(Me.lblHost)
        Me.Controls.Add(Me.cmdUp)
        Me.Controls.Add(Me.lvRemote)
        Me.Controls.Add(Me.cmdBrowseRemote)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(816, 530)
        Me.Name = "frmFTPClient"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "File Transfer"
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
            lvRemote_Update()
            Client.hostNameValue = Me.txtHost.Text
            Client.portValue = Me.txtPort.Text
            Client.userNameValue = Me.txtUser.Text
            Client.dirValue = Me.txtRemotePath.Text
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

        LastCursor = Windows.Forms.Cursor.Current
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If NewPath.Length > 0 Then
            Files = Client.CDLS(NewPath)
        Else
            Files = Client.LIST
        End If

        txtRemotePath.Text = Client.PWD
        lvRemote.Clear()
        If (Not Files Is Nothing) Then
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
        lvRemote.View = View.Details
        lvRemote.Columns.Add("Name", 200, HorizontalAlignment.Left)
        lvRemote.Columns.Add("Mode", 100, HorizontalAlignment.Left)
        lvRemote.Columns.Add("Size", 100, HorizontalAlignment.Left)
        lvRemote.Columns.Add("Date", 100, HorizontalAlignment.Left)

        Windows.Forms.Cursor.Current = LastCursor

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
            txtLocalPath.Enabled = False
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

        If (Not Client Is Nothing) Then
            Client.DelWorkFiles()
        End If
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

    
End Class