Public Class frmProject
    Inherits SQStudio.frmBlank
    Dim objThis As New clsProject
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Friend WithEvents txtFileLocation As System.Windows.Forms.TextBox

    Dim IsNewObj As Boolean
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtProjectName As System.Windows.Forms.TextBox
    Friend WithEvents txtVersion As System.Windows.Forms.TextBox
    Friend WithEvents txtMetaDSN As System.Windows.Forms.TextBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmProject))
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtProjectName = New System.Windows.Forms.TextBox
        Me.txtVersion = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtMetaDSN = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.Label4 = New System.Windows.Forms.Label
        Me.cmdBrowseFile = New System.Windows.Forms.Button
        Me.txtFileLocation = New System.Windows.Forms.TextBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.TabIndex = 6
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.TabIndex = 7
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Text = "Project definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(432, 40)
        Me.Label2.Text = "Enter a unique project name, and a project file. Also, enter the DSN (ODBC) name " & _
            "of the SQData metadata database."
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 20)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Project Name"
        '
        'txtProjectName
        '
        Me.txtProjectName.Location = New System.Drawing.Point(136, 88)
        Me.txtProjectName.MaxLength = 255
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(351, 20)
        Me.txtProjectName.TabIndex = 1
        '
        'txtVersion
        '
        Me.txtVersion.Location = New System.Drawing.Point(136, 199)
        Me.txtVersion.MaxLength = 255
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.Size = New System.Drawing.Size(351, 20)
        Me.txtVersion.TabIndex = 5
        Me.txtVersion.Text = "1.0.0"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(24, 199)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 20)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Version"
        '
        'txtMetaDSN
        '
        Me.txtMetaDSN.Location = New System.Drawing.Point(136, 162)
        Me.txtMetaDSN.MaxLength = 255
        Me.txtMetaDSN.Name = "txtMetaDSN"
        Me.txtMetaDSN.Size = New System.Drawing.Size(351, 20)
        Me.txtMetaDSN.TabIndex = 4
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(24, 162)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(96, 20)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "Metadata DSN"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "*.sqp"
        Me.SaveFileDialog1.Filter = "SQData Studio Project Files (*.sqp)|*.sqp |All files (*.*)|*.*"
        '
        'HelpProvider1
        '
        Me.HelpProvider1.HelpNamespace = "C:\SQDStudio\SQDStudio.chm"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 125)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 20)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "File Location"
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.Location = New System.Drawing.Point(463, 125)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 3
        Me.cmdBrowseFile.Text = "..."
        '
        'txtFileLocation
        '
        Me.txtFileLocation.Location = New System.Drawing.Point(136, 125)
        Me.txtFileLocation.MaxLength = 255
        Me.txtFileLocation.Name = "txtFileLocation"
        Me.txtFileLocation.Size = New System.Drawing.Size(326, 20)
        Me.txtFileLocation.TabIndex = 2
        '
        'frmProject
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 335)
        Me.Controls.Add(Me.cmdBrowseFile)
        Me.Controls.Add(Me.txtVersion)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtMetaDSN)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtFileLocation)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtProjectName)
        Me.Controls.Add(Me.Label3)
        Me.HelpProvider1.SetHelpKeyword(Me, "project.htm")
        Me.HelpProvider1.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.Topic)
        Me.Name = "frmProject"
        Me.HelpProvider1.SetShowHelp(Me, True)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Project Properties"
        Me.TopMost = False
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtProjectName, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtFileLocation, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtMetaDSN, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtVersion, 0)
        Me.Controls.SetChildIndex(Me.cmdBrowseFile, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMetaDSN.TextChanged, txtProjectName.TextChanged, txtVersion.TextChanged, txtFileLocation.TextChanged
        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)
    End Sub

    Private Sub StartLoad()
        IsEventFromCode = True
        objThis.IsModified = False
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

        If MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If
    End Sub

    '//Returns new project object
    Friend Function NewObj() As clsProject
        IsNewObj = True
        StartLoad()
        txtMetaDSN.Text = "SQDMeta"

        'If System.IO.Directory.Exists(GetProjectPath()) = False Then
        '    System.IO.Directory.CreateDirectory(GetProjectPath())
        'End If

        'txtFileLocation.Text = GetProjectPath()
        txtProjectName.SelectAll()
        EndLoad()
doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = objThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
        End Select
    End Function

    Public Function OpenProject(ByVal sPath As String) As clsProject
        IsNewObj = False
        objThis = New clsProject

        'objThis.Load(sPath)
        UpdateFields()
        OpenProject = objThis
    End Function

    Private Function UpdateFields() As Boolean
        txtProjectName.Text = objThis.ProjectName
        txtVersion.Text = objThis.ProjectVersion
        txtMetaDSN.Text = objThis.ProjectMetaDSN
        'txtFileLocation.Text = objThis.ProjectFolder
    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidateObjectName(txtProjectName.Text) = False Or ValidateNewName(txtProjectName.Text) = False Or objThis.ValidateNewObject(txtProjectName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.ProjectName = txtProjectName.Text
            objThis.ProjectVersion = txtVersion.Text
            objThis.ProjectMetaDSN = txtMetaDSN.Text
            'objThis.ProjectFolder = txtFileLocation.Text

            If IsNewObj = True Then
                If objThis.AddNew = True Then
                    'MsgBox("Project saved successfully", MsgBoxStyle.Information)
                Else
                    'MsgBox("Error during project save operation", MsgBoxStyle.Critical)
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            Me.Close()
            LocalProjectFolderPath = txtFileLocation.Text '//new 8/15/05
            DialogResult = Windows.Forms.DialogResult.OK
        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try
    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click
        Try
            Try
                'dlgSave.InitialDirectory = System.IO.Path.GetDirectoryName(txtFileLocation.Text)
                'dlgSave.FileName = txtFileLocation.Text
            Catch ex As Exception

            End Try

            'If dlgOpen.ShowDialog = DialogResult.OK Then
            If dlgBrowseFolder.ShowDialog = Windows.Forms.DialogResult.OK Then
                'txtFileLocation.Text = dlgSave.FileName
                txtFileLocation.Text = dlgBrowseFolder.SelectedPath
            Else
            End If
        Catch ex As Exception
            LogError(ex)
        End Try
    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_Project)
    End Sub

    Private Sub frmProject_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
