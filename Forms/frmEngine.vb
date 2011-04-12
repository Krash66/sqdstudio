Public Class frmEngine
    Inherits SQDStudio.frmBlank

    Public objThis As New clsEngine
    Friend WithEvents cbForceCommit As System.Windows.Forms.CheckBox
    Friend WithEvents txtDBDLib As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label

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
    Friend WithEvents txtEngineDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtReportFile As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtCommitEvery As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtReportEvery As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtEngineName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDDLLib As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtDTDLib As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtIncludeLib As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtCopybookLib As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cbConn As System.Windows.Forms.ComboBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtEngineDesc = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtReportFile = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtCommitEvery = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtReportEvery = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtEngineName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDDLLib = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtDTDLib = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtIncludeLib = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtCopybookLib = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.cbConn = New System.Windows.Forms.ComboBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.cbForceCommit = New System.Windows.Forms.CheckBox
        Me.txtDBDLib = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(589, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 419)
        Me.GroupBox1.Size = New System.Drawing.Size(591, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(307, 442)
        Me.cmdOk.TabIndex = 10
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(403, 442)
        Me.cmdCancel.TabIndex = 11
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(499, 442)
        Me.cmdHelp.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.Text = "Engine definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(513, 39)
        Me.Label2.Text = "Enter a unique engine name within a system, and the engine attributes.  You can a" & _
            "lso override the directories (DDNAMEs) where structure files reside."
        '
        'txtEngineDesc
        '
        Me.txtEngineDesc.Location = New System.Drawing.Point(120, 205)
        Me.txtEngineDesc.MaxLength = 1000
        Me.txtEngineDesc.Multiline = True
        Me.txtEngineDesc.Name = "txtEngineDesc"
        Me.txtEngineDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtEngineDesc.Size = New System.Drawing.Size(447, 51)
        Me.txtEngineDesc.TabIndex = 5
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(24, 205)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 17)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "Description"
        '
        'txtReportFile
        '
        Me.txtReportFile.Location = New System.Drawing.Point(120, 164)
        Me.txtReportFile.MaxLength = 255
        Me.txtReportFile.Name = "txtReportFile"
        Me.txtReportFile.Size = New System.Drawing.Size(447, 20)
        Me.txtReportFile.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(24, 167)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 17)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Report File"
        '
        'txtCommitEvery
        '
        Me.txtCommitEvery.Location = New System.Drawing.Point(393, 125)
        Me.txtCommitEvery.MaxLength = 255
        Me.txtCommitEvery.Name = "txtCommitEvery"
        Me.txtCommitEvery.Size = New System.Drawing.Size(64, 20)
        Me.txtCommitEvery.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(299, 128)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 20)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "Commit Every"
        '
        'txtReportEvery
        '
        Me.txtReportEvery.Location = New System.Drawing.Point(120, 125)
        Me.txtReportEvery.MaxLength = 255
        Me.txtReportEvery.Name = "txtReportEvery"
        Me.txtReportEvery.Size = New System.Drawing.Size(66, 20)
        Me.txtReportEvery.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 20)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Report Every"
        '
        'txtEngineName
        '
        Me.txtEngineName.Location = New System.Drawing.Point(120, 85)
        Me.txtEngineName.MaxLength = 255
        Me.txtEngineName.Name = "txtEngineName"
        Me.txtEngineName.Size = New System.Drawing.Size(176, 20)
        Me.txtEngineName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 20)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Engine Name"
        '
        'txtDDLLib
        '
        Me.txtDDLLib.Location = New System.Drawing.Point(120, 368)
        Me.txtDDLLib.MaxLength = 255
        Me.txtDDLLib.Name = "txtDDLLib"
        Me.txtDDLLib.Size = New System.Drawing.Size(447, 20)
        Me.txtDDLLib.TabIndex = 9
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(24, 368)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 17)
        Me.Label9.TabIndex = 79
        Me.Label9.Text = "DDL Lib"
        '
        'txtDTDLib
        '
        Me.txtDTDLib.Location = New System.Drawing.Point(120, 342)
        Me.txtDTDLib.MaxLength = 255
        Me.txtDTDLib.Name = "txtDTDLib"
        Me.txtDTDLib.Size = New System.Drawing.Size(447, 20)
        Me.txtDTDLib.TabIndex = 8
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(24, 345)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(96, 17)
        Me.Label10.TabIndex = 77
        Me.Label10.Text = "DTD Lib"
        '
        'txtIncludeLib
        '
        Me.txtIncludeLib.Location = New System.Drawing.Point(120, 316)
        Me.txtIncludeLib.MaxLength = 255
        Me.txtIncludeLib.Name = "txtIncludeLib"
        Me.txtIncludeLib.Size = New System.Drawing.Size(447, 20)
        Me.txtIncludeLib.TabIndex = 7
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(24, 319)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(96, 17)
        Me.Label7.TabIndex = 75
        Me.Label7.Text = "Include Lib"
        '
        'txtCopybookLib
        '
        Me.txtCopybookLib.Location = New System.Drawing.Point(120, 290)
        Me.txtCopybookLib.MaxLength = 255
        Me.txtCopybookLib.Name = "txtCopybookLib"
        Me.txtCopybookLib.Size = New System.Drawing.Size(447, 20)
        Me.txtCopybookLib.TabIndex = 6
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(24, 293)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(96, 17)
        Me.Label11.TabIndex = 73
        Me.Label11.Text = "Copybook Lib"
        '
        'cbConn
        '
        Me.cbConn.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbConn.FormattingEnabled = True
        Me.cbConn.Location = New System.Drawing.Point(409, 84)
        Me.cbConn.Name = "cbConn"
        Me.cbConn.Size = New System.Drawing.Size(158, 21)
        Me.cbConn.TabIndex = 80
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(315, 88)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(71, 13)
        Me.Label12.TabIndex = 81
        Me.Label12.Text = "Connection"
        '
        'cbForceCommit
        '
        Me.cbForceCommit.AutoSize = True
        Me.cbForceCommit.CheckAlign = System.Drawing.ContentAlignment.BottomRight
        Me.cbForceCommit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbForceCommit.Location = New System.Drawing.Point(465, 127)
        Me.cbForceCommit.Name = "cbForceCommit"
        Me.cbForceCommit.Size = New System.Drawing.Size(102, 17)
        Me.cbForceCommit.TabIndex = 82
        Me.cbForceCommit.Text = "Force Commit"
        Me.cbForceCommit.UseVisualStyleBackColor = True
        '
        'txtDBDLib
        '
        Me.txtDBDLib.Location = New System.Drawing.Point(120, 264)
        Me.txtDBDLib.MaxLength = 255
        Me.txtDBDLib.Name = "txtDBDLib"
        Me.txtDBDLib.Size = New System.Drawing.Size(447, 20)
        Me.txtDBDLib.TabIndex = 83
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(24, 267)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(96, 17)
        Me.Label13.TabIndex = 84
        Me.Label13.Text = "DBD Lib"
        '
        'frmEngine
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(589, 481)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.txtDBDLib)
        Me.Controls.Add(Me.cbForceCommit)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.cbConn)
        Me.Controls.Add(Me.txtDDLLib)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtDTDLib)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.txtIncludeLib)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtCopybookLib)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.txtEngineDesc)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtReportFile)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtCommitEvery)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtReportEvery)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtEngineName)
        Me.Controls.Add(Me.Label3)
        Me.Name = "frmEngine"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Engine Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtEngineName, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtReportEvery, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtCommitEvery, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtReportFile, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.txtEngineDesc, 0)
        Me.Controls.SetChildIndex(Me.Label11, 0)
        Me.Controls.SetChildIndex(Me.txtCopybookLib, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.txtIncludeLib, 0)
        Me.Controls.SetChildIndex(Me.Label10, 0)
        Me.Controls.SetChildIndex(Me.txtDTDLib, 0)
        Me.Controls.SetChildIndex(Me.Label9, 0)
        Me.Controls.SetChildIndex(Me.txtDDLLib, 0)
        Me.Controls.SetChildIndex(Me.cbConn, 0)
        Me.Controls.SetChildIndex(Me.Label12, 0)
        Me.Controls.SetChildIndex(Me.cbForceCommit, 0)
        Me.Controls.SetChildIndex(Me.txtDBDLib, 0)
        Me.Controls.SetChildIndex(Me.Label13, 0)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCopybookLib.TextChanged, txtDDLLib.TextChanged, txtDTDLib.TextChanged, txtIncludeLib.TextChanged, txtCommitEvery.TextChanged, txtEngineDesc.TextChanged, txtEngineName.TextChanged, txtReportEvery.TextChanged, txtReportFile.TextChanged, cbConn.SelectedIndexChanged, txtDBDLib.TextChanged

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
            Dim SysObj As clsSystem = objThis.ObjSystem
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = "Engine" & Count.ToString ' & "_" & Strings.Left(SysObj.Text, 9)
                For Each TestEng As clsEngine In SysObj.Engines
                    If TestEng.EngineName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtEngineName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmEng SetDefaultName")
        End Try

    End Sub

    Sub SetComboConn()

        Try
            cbConn.Items.Clear()

            cbConn.Items.Add(New Mylist("None", "none")) ' New clsConnection))
            For Each conn As clsConnection In objThis.ObjSystem.Environment.Connections
                cbConn.Items.Add(New Mylist(conn.ConnectionName, conn))
            Next
            cbConn.SelectedIndex = 0

        Catch ex As Exception
            LogError(ex, "frmEngine setcomboconn")
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

    '//Returns new Engine object
    Public Function NewObj() As clsEngine

        IsNewObj = True
        StartLoad()

        SetComboConn()

        SetDefaultName()

        txtEngineName.SelectAll()
        txtReportEvery.Text = 0
        txtCommitEvery.Text = 0
        txtDBDLib.Text = objThis.ObjSystem.DBDLib
        txtCopybookLib.Text = objThis.ObjSystem.CopybookLib
        txtIncludeLib.Text = objThis.ObjSystem.IncludeLib
        txtDTDLib.Text = objThis.ObjSystem.DTDLib
        txtDDLLib.Text = objThis.ObjSystem.DDLLib

        txtEngineName.Focus()

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
            If objThis.ValidateNewObject(txtEngineName.Text) = False Or ValidateNewName(txtEngineName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.EngineName = txtEngineName.Text
            objThis.EngineDescription = txtEngineDesc.Text
            objThis.ReportFile = txtReportFile.Text
            objThis.ReportEvery = txtReportEvery.Text
            objThis.CommitEvery = txtCommitEvery.Text
            objThis.DBDLib = txtDBDLib.Text
            objThis.CopybookLib = txtCopybookLib.Text
            objThis.IncludeLib = txtIncludeLib.Text
            objThis.DTDLib = txtDTDLib.Text
            objThis.DDLLib = txtDDLLib.Text
            'added for RBC   3/8/11    Gives a default version
            objThis.EngVersion = "3.7.17"

            Dim temp As Mylist = cbConn.SelectedItem
            'Dim answer As MsgBoxResult

            If temp.Name = "None" Then
                'answer = MsgBox("You have not defined a connection." & Chr(13) & "If your engine requires a connection, choose it now," & Chr(13) & "you will not be allowed to add a connection to this engine later." & Chr(13) & "Do you wish to add a connection now?", MsgBoxStyle.YesNo, "Add Connection?")
                'If answer = MsgBoxResult.Yes Then
                '    DialogResult = Windows.Forms.DialogResult.Retry
                '    Exit Sub
                'Else
                objThis.Connection = Nothing
                'End If
            Else
            '/// a valid connection was chosen
            objThis.Connection = CType(temp.ItemData, clsConnection)
            End If

            'If temp Is Nothing Then
            '    '//// nothing was chosen for "connection"
            '    answer = MsgBox("You have not defined a connection." & Chr(13) & "If your engine requires a connection, choose it now," & Chr(13) & "you will not be allowed to add a connection to this engine later." & Chr(13) & "Do you wish to add a connection now?", MsgBoxStyle.YesNo, "Add Connection?")
            '    If answer = MsgBoxResult.Yes Then
            '        DialogResult = Windows.Forms.DialogResult.Retry
            '        Exit Sub
            '    End If
            'Else
            '    '/// empty "mylist" object was chosen or a name was typed in
            '    If temp.Name = "None" Then
            '        answer = MsgBox("You have not defined a connection." & Chr(13) & "If your engine requires a connection, choose it now," & Chr(13) & "you will not be allowed to add a connection to this engine later." & Chr(13) & "Do you wish to add a connection now?", MsgBoxStyle.YesNo, "Add Connection?")
            '        If answer = MsgBoxResult.Yes Then
            '            DialogResult = Windows.Forms.DialogResult.Retry
            '            Exit Sub
            '        End If
            '    Else
            '        '/// a valid connection was chosen
            '        objThis.Connection = CType(temp.ItemData, clsConnection)
            '    End If
            'End If

            'If temp IsNot Nothing Then
            '    If temp.Name = "" Then
            '        MsgBox("Please type in a new Connection Name or select one from the list", MsgBoxStyle.Information)
            '        DialogResult = Windows.Forms.DialogResult.Retry
            '        Exit Sub
            '    End If
            '    objThis.Connection = CType(temp.ItemData, clsConnection)
            'Else
            '    '//Means User typed in New connection Name into combo box
            '    If cbConn.Text = "" Then
            '        MsgBox("Please type in a new Connection Name or select one from the list", MsgBoxStyle.Information)
            '        DialogResult = Windows.Forms.DialogResult.Retry
            '        Exit Sub
            '    End If
            '    Dim NewConn As clsConnection
            '    NewConn = New clsConnection
            '    With NewConn
            '        .ConnectionName = cbConn.Text
            '        .Environment = objThis.ObjSystem.Environment
            '        .Parent = NewConn.Environment
            '    End With
            '    If NewConn.ValidateNewObject(NewConn.ConnectionName) = False Or ValidateNewName(NewConn.ConnectionName) = False Then
            '        cbConn.Focus()
            '        cbConn.SelectAll()
            '        DialogResult = Windows.Forms.DialogResult.Retry
            '        Exit Sub
            '    End If
            '    If NewConn.AddNew() = True Then
            '        objThis.Connection = NewConn
            '        objThis.Connection.ObjTreeNode = AddNode(objThis.Connection.Environment.ObjTreeNode.Nodes(0).Nodes, NODE_CONNECTION, objThis.Connection, , objThis.Connection.ConnectionName)
            '        MsgBox("You have created a new connection named, " & objThis.Connection.ConnectionName & Chr(13) & "Please make sure it is configured correctly.", MsgBoxStyle.Information)
            '    End If
            'End If

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            'If objThis.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    objThis.SaveEngProps()
            'End If

            DialogResult = Windows.Forms.DialogResult.OK

            Me.Close()

        Catch ex As Exception
            DialogResult = Windows.Forms.DialogResult.Retry
            LogError(ex, "frmEngine cmdOK_Click")
        End Try

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ShowHelp(modHelp.HHId.H_Add_Engines)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) 'Handles MyBase.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click(sender, e)
            Case Keys.Enter
                cmdOk_Click(sender, e)
            Case Keys.F1
                cmdHelp_Click(sender, e)
        End Select
    End Sub

    Private Sub cbForceCommit_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbForceCommit.CheckedChanged

        If IsEventFromCode = True Then Exit Sub
        If cbForceCommit.Checked = True Then
            objThis.ForceCommit = True
        Else
            objThis.ForceCommit = False
        End If
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

End Class
