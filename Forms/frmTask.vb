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
    Friend WithEvents txtTaskDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtTaskName As System.Windows.Forms.TextBox
    Friend WithEvents tvSource As System.Windows.Forms.TreeView
    Friend WithEvents tvTarget As System.Windows.Forms.TreeView
    Friend WithEvents gbTaskProp As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbSrc As System.Windows.Forms.GroupBox
    Friend WithEvents gbTgt As System.Windows.Forms.GroupBox
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
        Me.txtTaskName = New System.Windows.Forms.TextBox
        Me.lblTaskName = New System.Windows.Forms.Label
        Me.tvSource = New System.Windows.Forms.TreeView
        Me.tvTarget = New System.Windows.Forms.TreeView
        Me.txtTaskDesc = New System.Windows.Forms.TextBox
        Me.gbTaskProp = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.gbSrc = New System.Windows.Forms.GroupBox
        Me.gbTgt = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.Panel1.SuspendLayout()
        Me.gbTaskProp.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gbSrc.SuspendLayout()
        Me.gbTgt.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(626, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 454)
        Me.GroupBox1.Size = New System.Drawing.Size(628, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(344, 477)
        Me.cmdOk.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(440, 477)
        Me.cmdCancel.TabIndex = 6
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(536, 477)
        Me.cmdHelp.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Text = "Task definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(550, 39)
        Me.Label2.Text = "Enter a unique task name within an engine. Specify the source and target datastor" & _
            "es on which the task operates. For a lookup or join task, only select a source d" & _
            "atastore."
        '
        'txtTaskName
        '
        Me.txtTaskName.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.txtTaskName.Location = New System.Drawing.Point(51, 19)
        Me.txtTaskName.MaxLength = 20
        Me.txtTaskName.Name = "txtTaskName"
        Me.txtTaskName.Size = New System.Drawing.Size(311, 20)
        Me.txtTaskName.TabIndex = 1
        '
        'lblTaskName
        '
        Me.lblTaskName.AutoSize = True
        Me.lblTaskName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaskName.Location = New System.Drawing.Point(7, 22)
        Me.lblTaskName.Name = "lblTaskName"
        Me.lblTaskName.Size = New System.Drawing.Size(38, 14)
        Me.lblTaskName.TabIndex = 19
        Me.lblTaskName.Text = "Name"
        '
        'tvSource
        '
        Me.tvSource.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvSource.Location = New System.Drawing.Point(3, 16)
        Me.tvSource.Name = "tvSource"
        Me.tvSource.ShowPlusMinus = False
        Me.tvSource.ShowRootLines = False
        Me.tvSource.Size = New System.Drawing.Size(289, 197)
        Me.tvSource.TabIndex = 2
        '
        'tvTarget
        '
        Me.tvTarget.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvTarget.Location = New System.Drawing.Point(3, 16)
        Me.tvTarget.Name = "tvTarget"
        Me.tvTarget.ShowPlusMinus = False
        Me.tvTarget.ShowRootLines = False
        Me.tvTarget.Size = New System.Drawing.Size(290, 197)
        Me.tvTarget.TabIndex = 3
        '
        'txtTaskDesc
        '
        Me.txtTaskDesc.Location = New System.Drawing.Point(6, 19)
        Me.txtTaskDesc.MaxLength = 1000
        Me.txtTaskDesc.Multiline = True
        Me.txtTaskDesc.Name = "txtTaskDesc"
        Me.txtTaskDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtTaskDesc.Size = New System.Drawing.Size(346, 75)
        Me.txtTaskDesc.TabIndex = 4
        '
        'gbTaskProp
        '
        Me.gbTaskProp.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbTaskProp.Controls.Add(Me.SplitContainer1)
        Me.gbTaskProp.Controls.Add(Me.gbDesc)
        Me.gbTaskProp.Controls.Add(Me.txtTaskName)
        Me.gbTaskProp.Controls.Add(Me.lblTaskName)
        Me.gbTaskProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbTaskProp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbTaskProp.Location = New System.Drawing.Point(7, 74)
        Me.gbTaskProp.Name = "gbTaskProp"
        Me.gbTaskProp.Size = New System.Drawing.Size(607, 374)
        Me.gbTaskProp.TabIndex = 25
        Me.gbTaskProp.TabStop = False
        Me.gbTaskProp.Text = "Properties"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 152)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbSrc)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gbTgt)
        Me.SplitContainer1.Size = New System.Drawing.Size(595, 216)
        Me.SplitContainer1.SplitterDistance = 295
        Me.SplitContainer1.TabIndex = 21
        '
        'gbSrc
        '
        Me.gbSrc.Controls.Add(Me.tvSource)
        Me.gbSrc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbSrc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbSrc.Location = New System.Drawing.Point(0, 0)
        Me.gbSrc.Name = "gbSrc"
        Me.gbSrc.Size = New System.Drawing.Size(295, 216)
        Me.gbSrc.TabIndex = 21
        Me.gbSrc.TabStop = False
        Me.gbSrc.Text = "Source(s)"
        '
        'gbTgt
        '
        Me.gbTgt.Controls.Add(Me.tvTarget)
        Me.gbTgt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbTgt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbTgt.Location = New System.Drawing.Point(0, 0)
        Me.gbTgt.Name = "gbTgt"
        Me.gbTgt.Size = New System.Drawing.Size(296, 216)
        Me.gbTgt.TabIndex = 0
        Me.gbTgt.TabStop = False
        Me.gbTgt.Text = "Target(s)"
        '
        'gbDesc
        '
        Me.gbDesc.Controls.Add(Me.txtTaskDesc)
        Me.gbDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDesc.Location = New System.Drawing.Point(10, 46)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(358, 100)
        Me.gbDesc.TabIndex = 20
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'frmTask
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Nothing
        Me.ClientSize = New System.Drawing.Size(626, 516)
        Me.Controls.Add(Me.gbTaskProp)
        Me.MinimumSize = New System.Drawing.Size(401, 457)
        Me.Name = "frmTask"
        Me.Text = "Task Properties"
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.gbTaskProp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbTaskProp.ResumeLayout(False)
        Me.gbTaskProp.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.gbSrc.ResumeLayout(False)
        Me.gbTgt.ResumeLayout(False)
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.ResumeLayout(False)

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
                    If NamePrefix = "M_Proc" Or NamePrefix = "L_Proc" Then
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
                Me.gbTaskProp.Text = "Main Procedure Properties"
                Me.tvTarget.Visible = False
                'Me.Label6.Visible = False
                'Me.tvSource.Width = 476
                NamePrefix = "Main"
            Case modDeclares.enumTaskType.TASK_PROC
                Me.Text = "Mapping Procedure Properties"
                Me.Label1.Text = "Mapping Procedure Definition"
                Me.gbTaskProp.Text = "Mapping Procedure Properties"
                NamePrefix = "M_Proc"
            Case modDeclares.enumTaskType.TASK_GEN
                Me.Text = "Logic Procedure Properties"
                Me.Label1.Text = "Logic Procedure Definition"
                Me.gbTaskProp.Text = "Logic Procedure Properties"
                Me.tvTarget.Visible = False
                'Me.Label6.Visible = False
                'me.tvSource.Width = 476
                NamePrefix = "L_Proc"
            Case modDeclares.enumTaskType.TASK_LOOKUP
                Me.Text = "LookUp Properties"
                Me.Label1.Text = "LookUp Definition"
                Me.gbTaskProp.Text = "LookUp Properties"
                Me.tvTarget.Visible = False
                'Me.Label6.Visible = False
                'Me.tvSource.Width = 476
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

        If objThis.TaskType = modDeclares.enumTaskType.TASK_GEN _
        Or objThis.TaskType = modDeclares.enumTaskType.TASK_MAIN _
        Or objThis.TaskType = modDeclares.enumTaskType.TASK_LOOKUP Then
            'tvTarget.Enabled = False
            tvTarget.Visible = False
            'lblTargetCaption.Visible = False
        Else
            'lblTargetCaption.Visible = True
            tvTarget.Visible = True
        End If

        UpdateFields()


        'FillMappings()



        FillSourceTarget(cNode)
        EndLoad()
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
                txtTaskName.Text = "M_" & NodeText
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
                If objThis.TaskType = enumTaskType.TASK_GEN Then
                    txtTaskName.Text = "L_" & NodeText
                Else
                    txtTaskName.Text = "M_" & NodeText
                End If
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
            Case "Logic Procedure Properties"
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
