Public Class frmTask
    Inherits SQDStudio.frmBlank

    Public objThis As New clsTask

    Dim IsNewObj As Boolean

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

#Region " Windows Form Designer generated code "
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtTaskDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtTaskName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tvSource As System.Windows.Forms.TreeView
    Friend WithEvents tvTarget As System.Windows.Forms.TreeView
    Friend WithEvents lblTaskName As System.Windows.Forms.Label

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
    End Sub

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtTaskName = New System.Windows.Forms.TextBox
        Me.lblTaskName = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tvSource = New System.Windows.Forms.TreeView
        Me.tvTarget = New System.Windows.Forms.TreeView
        Me.txtTaskDesc = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(570, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 402)
        Me.GroupBox1.Size = New System.Drawing.Size(572, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(288, 425)
        Me.cmdOk.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(384, 425)
        Me.cmdCancel.TabIndex = 6
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(480, 425)
        Me.cmdHelp.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Text = "Task definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(494, 39)
        Me.Label2.Text = "Enter a unique task name within an engine. Specify the source and target datastor" & _
            "es on which the task operates. For a lookup or join task, only select a source d" & _
            "atastore."
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(296, 120)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 16)
        Me.Label6.TabIndex = 24
        Me.Label6.Text = "Select Target(s)"
        '
        'txtTaskName
        '
        Me.txtTaskName.Location = New System.Drawing.Point(104, 88)
        Me.txtTaskName.MaxLength = 20
        Me.txtTaskName.Name = "txtTaskName"
        Me.txtTaskName.Size = New System.Drawing.Size(240, 20)
        Me.txtTaskName.TabIndex = 1
        '
        'lblTaskName
        '
        Me.lblTaskName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaskName.Location = New System.Drawing.Point(16, 88)
        Me.lblTaskName.Name = "lblTaskName"
        Me.lblTaskName.Size = New System.Drawing.Size(101, 20)
        Me.lblTaskName.TabIndex = 19
        Me.lblTaskName.Text = "Task  Name"
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 360)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(102, 16)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Description"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(102, 16)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Select Source(s)"
        '
        'tvSource
        '
        Me.tvSource.Location = New System.Drawing.Point(16, 136)
        Me.tvSource.Name = "tvSource"
        Me.tvSource.ShowPlusMinus = False
        Me.tvSource.ShowRootLines = False
        Me.tvSource.Size = New System.Drawing.Size(256, 192)
        Me.tvSource.TabIndex = 2
        '
        'tvTarget
        '
        Me.tvTarget.Location = New System.Drawing.Point(296, 136)
        Me.tvTarget.Name = "tvTarget"
        Me.tvTarget.ShowPlusMinus = False
        Me.tvTarget.ShowRootLines = False
        Me.tvTarget.Size = New System.Drawing.Size(256, 192)
        Me.tvTarget.TabIndex = 3
        '
        'txtTaskDesc
        '
        Me.txtTaskDesc.Location = New System.Drawing.Point(104, 344)
        Me.txtTaskDesc.MaxLength = 1000
        Me.txtTaskDesc.Multiline = True
        Me.txtTaskDesc.Name = "txtTaskDesc"
        Me.txtTaskDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTaskDesc.Size = New System.Drawing.Size(376, 51)
        Me.txtTaskDesc.TabIndex = 4
        '
        'frmTask
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Nothing
        Me.ClientSize = New System.Drawing.Size(570, 464)
        Me.Controls.Add(Me.txtTaskDesc)
        Me.Controls.Add(Me.tvSource)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtTaskName)
        Me.Controls.Add(Me.lblTaskName)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.tvTarget)
        Me.Name = "frmTask"
        Me.Text = "Task Properties"
        Me.Controls.SetChildIndex(Me.tvTarget, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.lblTaskName, 0)
        Me.Controls.SetChildIndex(Me.txtTaskName, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.tvSource, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.txtTaskDesc, 0)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub


#End Region

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTaskName.TextChanged, txtTaskDesc.TextChanged, tvSource.ImeModeChanged, tvTarget.ImeModeChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False

    End Sub

    Private Sub SetDefaultName(ByVal NamePrefix As String)

        Try
            Dim EngObj As clsEngine = Nothing
            Dim EnvObj As clsEnvironment = objThis.Environment
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            If objThis.Engine IsNot Nothing Then
                EngObj = objThis.Engine
            End If

            While GoodName = False
                GoodName = True
                NewName = NamePrefix & Count.ToString & "_" & Strings.Left(EngObj.Text, 9)
                If NamePrefix = "Main" Then
                    For Each TestMain As clsTask In EngObj.Mains
                        If TestMain.TaskName = NewName Then
                            GoodName = False
                            Exit For
                        End If
                    Next
                Else
                    If NamePrefix = "P_Proc" Then
                        If objThis.Engine Is Nothing Then
                            For Each TestProc As clsTask In EnvObj.Procedures
                                If TestProc.TaskName = NewName Then
                                    GoodName = False
                                    Exit For
                                End If
                            Next
                            If GoodName = True Then
                                For Each sys As clsSystem In EnvObj.Systems
                                    For Each eng As clsEngine In sys.Engines
                                        For Each proc As clsTask In eng.Procs
                                            If proc.TaskName = NewName Then
                                                GoodName = False
                                                Exit For
                                            End If
                                        Next
                                    Next
                                Next
                            End If
                        Else
                            For Each TestProc As clsTask In EngObj.Procs
                                If TestProc.TaskName = NewName Then
                                    GoodName = False
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End If


                'If objThis.Engine IsNot Nothing Then

                'End If

                'ElseIf NamePrefix = "J_Join" Then
                '    For Each TestJoin As clsTask In EngObj.Joins
                '        If TestJoin.TaskName = NewName Then
                '            GoodName = False
                '            Exit For
                '        End If
                '    Next
                'ElseIf NamePrefix = "L_LookUp" Then
                '    For Each TestLU As clsTask In EngObj.Lookups
                '        If TestLU.TaskName = NewName Then
                '            GoodName = False
                '            Exit For
                '        End If
                '    Next
                'End If
                Count += 1
            End While

            txtTaskName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmTask SetDefaultName")
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
        ElseIf MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    '//Returns new Task object
    Friend Function NewObj(ByVal cNode As TreeNode, ByVal TaskType As enumTaskType) As clsTask

        Dim NamePrefix As String = ""

        IsNewObj = True

        StartLoad()

        objThis.TaskType = TaskType

        Select Case TaskType
            Case modDeclares.enumTaskType.TASK_MAIN
                Me.Text = "Main Procedure Properties"
                Me.Label1.Text = "Main Procedure Definition"
                Me.tvTarget.Visible = False
                Me.Label6.Visible = False
                Me.tvSource.Width = 476
                NamePrefix = "Main"
            Case modDeclares.enumTaskType.TASK_PROC
                Me.Text = "Mapping Procedure Properties"
                Me.Label1.Text = "Mapping Procedure Definition"
                NamePrefix = "P_Proc"
            Case modDeclares.enumTaskType.TASK_GEN
                Me.Text = "General Procedure Properties"
                Me.Label1.Text = "General Procedure Definition"
                Me.tvTarget.Visible = False
                Me.Label6.Visible = False
                Me.tvSource.Width = 476
                NamePrefix = "P_Proc"
            Case modDeclares.enumTaskType.TASK_LOOKUP
                Me.Text = "LookUp Properties"
                Me.Label1.Text = "LookUp Definition"
                Me.tvTarget.Visible = False
                Me.Label6.Visible = False
                Me.tvSource.Width = 476
                NamePrefix = "L_LookUp"
        End Select

        SetDefaultName(NamePrefix)

        txtTaskName.SelectAll()

        FillSourceTarget(cNode)

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

    Public Function EditObj(ByVal cNode As TreeNode, ByVal obj As INode) As clsTask

        IsNewObj = False
        StartLoad()

        objThis = obj '//Load the form env object
        objThis.ObjTreeNode = cNode

        'cbGroupItems.Checked = True

        If objThis.TaskType = modDeclares.enumTaskType.TASK_GEN Or objThis.TaskType = modDeclares.enumTaskType.TASK_LOOKUP Then
            'tvTarget.Enabled = False
            tvTarget.Visible = False
            'lblTargetCaption.Visible = False
        Else
            'lblTargetCaption.Visible = True
            tvTarget.Visible = True
        End If

        UpdateFields()


        'FillMappings()

        EndLoad()

        FillSourceTarget(cNode)

        'the following if statement is to relode the source and target for the first time 
        'If firstTime = True Then
        '    FillSourceTarget(objThis.ObjTreeNode)
        '    firstTime = False
        'End If

        'Me.Visible = True

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                EditObj = objThis
                RaiseEvent Modified(Me, objThis)
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Function FillSourceTarget(ByVal cNode As TreeNode) As Boolean

        Dim nd As TreeNode
        Dim tempNode As TreeNode
        Dim srcChldNode As TreeNode
        Dim tgtChldNode As TreeNode

        tvSource.Nodes.Clear()
        tvTarget.Nodes.Clear()

        tvSource.ImageList = imgListSmall
        tvTarget.ImageList = imgListSmall

        tvSource.Font = New Font(tvSource.Font, FontStyle.Bold)
        tvTarget.Font = New Font(tvSource.Font, FontStyle.Bold)

        tvSource.CheckBoxes = True
        tvTarget.CheckBoxes = True

        For Each nd In cNode.Nodes
            If CType(nd.Tag, INode).Type = NODE_FO_SOURCEDATASTORE Then
                Dim ndSrc As TreeNode
                Dim srcTypeInode As INode
                For Each ndSrc In nd.Nodes
                    srcTypeInode = ndSrc.Tag
                    If srcTypeInode.Type = NODE_SOURCEDATASTORE Then
                        tempNode = ndSrc.Clone
                        tempNode.Nodes.Clear()
                        tvSource.Nodes.Add(tempNode)
                    End If
                Next
            ElseIf CType(nd.Tag, INode).Type = NODE_FO_TARGETDATASTORE Then
                Dim ndTgt As TreeNode
                Dim tgtTypeInode As INode
                For Each ndTgt In nd.Nodes
                    tgtTypeInode = ndTgt.Tag
                    If tgtTypeInode.Type = NODE_TARGETDATASTORE Then
                        tempNode = ndTgt.Clone
                        tempNode.Nodes.Clear()
                        tvTarget.Nodes.Add(tempNode)
                    End If
                Next
            ElseIf CType(nd.Tag, INode).Type = NODE_FO_DATASTORE Then
                Dim DSinode As INode
                For Each ndDStype As TreeNode In nd.Nodes
                    DSinode = ndDStype.Tag
                    If Strings.Left(DSinode.Type, 5) = "FO_DS" Then

                        'build source tree from Main tree
                        tempNode = ndDStype.Clone
                        tempNode.Nodes.Clear()
                        For Each ndDS As TreeNode In ndDStype.Nodes
                            srcChldNode = ndDS.Clone
                            srcChldNode.Nodes.Clear()
                            tempNode.Nodes.Add(srcChldNode)
                        Next
                        tvSource.Nodes.Add(tempNode)
                        tvSource.ExpandAll()

                        'build Target tree from Main Tree
                        tempNode = ndDStype.Clone
                        tempNode.Nodes.Clear()
                        For Each ndDS As TreeNode In ndDStype.Nodes
                            tgtChldNode = ndDS.Clone
                            tgtChldNode.Nodes.Clear()
                            tempNode.Nodes.Add(tgtChldNode)
                        Next
                        tvTarget.Nodes.Add(tempNode)
                        tvTarget.ExpandAll()
                    End If
                Next
            End If
        Next

        '//Now load selected source/target items and check if object is already 
        '//created and user is editing it
        If IsNewObj = False Then
            objThis.LoadDatastores()
            CheckSelectedDatastores()
        End If

    End Function

    Function CheckSelectedDatastores() As Boolean

        Dim i As Integer
        Dim cnt1 As Integer = tvSource.GetNodeCount(False)
        Dim cnt2 As Integer = tvTarget.GetNodeCount(False)

        If cnt1 + cnt2 <= 0 Then Exit Function

        Try
            tvSource.BeginUpdate()
            tvTarget.BeginUpdate()

            '//Looop all selected field and check that field in treeview
            For i = 0 To objThis.ObjSources.Count - 1
                If CType(objThis.ObjSources(i), INode).Type = NODE_SOURCEDATASTORE Then
                    SelectFirstMatchingNode(tvSource, "", True, CType(objThis.ObjSources(i), clsDatastore).Key)
                End If
            Next

            For i = 0 To objThis.ObjTargets.Count - 1
                If CType(objThis.ObjTargets(i), INode).Type = NODE_TARGETDATASTORE Then
                    SelectFirstMatchingNode(tvTarget, "", True, CType(objThis.ObjTargets(i), clsDatastore).Key)
                End If
            Next

        Catch ex As Exception
            LogError(ex)
        Finally
            tvSource.EndUpdate()
            tvTarget.EndUpdate()
        End Try

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtTaskName.Text = objThis.TaskName
        txtTaskDesc.Text = objThis.TaskDescription
        objThis.IsModified = False

    End Function

    Function CheckedSources() As Integer

        Dim nd As TreeNode
        Dim cnt As Integer

        For Each nd In tvSource.Nodes
            If nd.Checked Then cnt = cnt + 1
        Next
        CheckedSources = cnt

    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If IsNewObj = True Then
                If objThis.ValidateNewObject(txtTaskName.Text) = False Or ValidateNewName(txtTaskName.Text) = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            End If

            '//make sure that user dont select more than one sources
            If CheckedSources() > 1 Then
                MsgBox("You can not select more than one datastore in source", MsgBoxStyle.Critical, MsgTitle)
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            If IsNewObj = False Then
                '// check to see if renamed
                If objThis.TaskName <> txtTaskName.Text Then
                    objThis.IsRenamed = RenameTask(objThis, txtTaskName.Text)
                End If
                '// now set the task name
                If objThis.IsRenamed = False Then
                    txtTaskName.Text = objThis.TaskName
                Else
                    objThis.TaskName = txtTaskName.Text
                End If
            End If

            objThis.TaskName = txtTaskName.Text
            objThis.TaskDescription = txtTaskDesc.Text

            '//save current selected fields before we do add/save
            LoadFldArrFromTreeview(objThis.ObjSources, tvSource, True)
            For Each src As clsDatastore In objThis.ObjSources
                src.DsDirection = DS_DIRECTION_SOURCE
            Next
            LoadFldArrFromTreeview(objThis.ObjTargets, tvTarget, True)
            For Each tgt As clsDatastore In objThis.ObjTargets
                tgt.DsDirection = DS_DIRECTION_TARGET
            Next

            If IsNewObj = True Then
                If objThis.AddNew() = True Then
                    'If objThis.Engine IsNot Nothing Then
                    '    Dim objTask As clsTask
                    '    objTask = objThis.Clone(objThis.Environment)
                    '    objTask.Engine = Nothing
                    '    If objTask.AddNew = False Then
                    '        DialogResult = Windows.Forms.DialogResult.Retry
                    '        Exit Sub
                    '    End If
                    'End If
                Else
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                If objThis.Save = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            End If

            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    '/// After checking Sources
    Private Sub tv_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvSource.AfterCheck

        If e.Node.Checked = True Then
            If CType(e.Node.Tag, INode).IsFolderNode = True Then
                '//if we come here means parent node is checked. 
                '//If unchecked then fine but if checked then unchek all children which fource Rule1
                Dim nd As TreeNode
                '//uncheck children so only parent nodde remain selected
                For Each nd In e.Node.Nodes
                    nd.Checked = True
                Next
                e.Node.Checked = False
            End If
        End If

        OnChange(Me, New EventArgs)

    End Sub

    '/// After checking Targets
    Private Sub tv_AfterCheckTarget(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvTarget.AfterCheck

        If IsEventFromCode = True Then Exit Sub
        Dim NodeText As String = ""
        If e.Node.Checked = True Then
            If CType(e.Node.Tag, INode).IsFolderNode = True Then
                '//if we come here means parent node is checked. 
                '//If unchecked then fine but if checked then unchek all children which fource Rule1
                '//uncheck children so only parent nodde remain selected
                For Each nd As TreeNode In e.Node.Nodes
                    nd.Checked = True
                    NodeText = CType(nd.Tag, INode).Text
                Next
                NodeText = Strings.Right(NodeText, NodeText.Length - 2)
                txtTaskName.Text = "P_" & NodeText
                e.Node.Checked = False
            Else
                IsEventFromCode = True
                For Each nd As TreeNode In tvTarget.Nodes
                    nd.Checked = False
                    'For Each chnd As TreeNode In nd.Nodes
                    '    chnd.Checked = False
                    'Next
                Next
                e.Node.Checked = True
                NodeText = CType(e.Node.Tag, INode).Text
                NodeText = Strings.Right(NodeText, NodeText.Length - 2)
                txtTaskName.Text = "P_" & NodeText
                IsEventFromCode = False
            End If
        End If

        OnChange(Me, New EventArgs)

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Select Case Me.Text
            Case "Main Procedure Properties"
                ShowHelp(modHelp.HHId.H_Add_Main_Procedures)
            Case "Mapping Procedure Properties"
                ShowHelp(modHelp.HHId.H_Add_Mapping_Procedures)
            Case "General Procedure Properties"
                ShowHelp(modHelp.HHId.H_Add_General_Procedure)
            Case "LookUp Properties"
                ShowHelp(modHelp.HHId.H_Lookups)
            Case Else
                ShowHelp(modHelp.HHId.H_Mapping_Procedures)
        End Select

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
