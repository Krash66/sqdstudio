Public Class ctlAddFlowTab

    Dim ObjEng As clsEngine
    Dim HorizSP As Integer = 20
    Dim VertSP As Integer = 20
    Private Const HorizIncr As Integer = 110
    Private Const VertIncr As Integer = 90
    Dim NodeText As String
    Dim Nodesize As Integer = 65
    Dim DSnodeSize As Integer = 55
    Dim ProcNodeSize As Integer = 55
    Dim LogicNodeSizeH As Integer = 85
    Dim LogicNodeSizeV As Integer = 55
    Private strFileName As String = ""
    Private IsEventFromCode As Boolean



#Region "Public Functions"

    Public Function RefreshAddFlow() As Boolean

        Try
            IsEventFromCode = True

            Me.tabAddFlow.Items.Clear()

            ' Add Nodes to diagram
            LoadAFNodes()
            ' Now Add the Links
            LoadAFLinks()

            IsEventFromCode = False

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab RefreshAddFlow")
        End Try

    End Function

    Public Function LoadAFtab(ByVal Eng As clsEngine) As Boolean

        Try
            Eng.ObjAddFlow = tabAddFlow
            Eng.ObjAddFlowCtl = Me

            ObjEng = Eng

            'af = tabAddFlow

            '' Create the collection of images that will be used for nodes
            'af.Images.Add("C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2008\Projects\sqdstudio\images\Copy of Can_!.ico")
            'af.Images.Add("C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2008\Projects\sqdstudio\images\Microsoft Access Macro Shortcut (.mam).ico")
            'af.Images.Add("C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2008\Projects\sqdstudio\images\Microsoft Access Macro Shortcut (.mam).ico")
            'af.Images.Add("C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2008\Projects\sqdstudio\images\Can_!.ico")

            ' Add Nodes to diagram
            'LoadAFNodes()
            '' Now Add the Links
            'LoadAFLinks()
            RefreshAddFlow()
            'LoadPalette()

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillAddFlowFromEngine")
            Return False
        End Try

    End Function

    Public Function AddSDS(ByRef Sds As clsDatastore, Optional ByVal cnt As Integer = -1, Optional ByVal Palette As Boolean = False) As Boolean

        Try
            If Palette = False Then
                HorizSP = 20
                Sds.LoadMe()
                Sds.LoadItems()
                If cnt > -1 Then
                    VertSP = 20 + (cnt * VertIncr)
                Else
                    Dim nodeCount As Integer = Sds.Engine.Sources.Count - 1
                    VertSP = 20 + (nodeCount * VertIncr)
                End If
            Else
                HorizSP = 5
                VertSP = 20
            End If

            Dim SrcNode As Node
            NodeText = Sds.DatastoreName

            '--- Get the Addflow node for each sourceDS
            SrcNode = New Node(HorizSP, VertSP, DSnodeSize, DSnodeSize, NodeText, tabAddFlow.DefNodeProp)
            'SrcNode.InLinkable = False

            If Sds.IsLookUp = False Then
                SrcNode.GradientColor = Color.LightBlue
                SrcNode.Shape.Style = ShapeStyle.DirectAccessStorage
                SrcNode.Shape.Orientation = ShapeOrientation.so_270
            Else
                SrcNode.GradientColor = Color.LightGreen
                SrcNode.Shape.Style = ShapeStyle.DirectAccessStorage
                SrcNode.Shape.Orientation = ShapeOrientation.so_270
                'SrcNode.TextColor = Color.LightGreen
            End If

            SrcNode.ImageIndex = 0
            '--- Add Object as Tag of node
            SrcNode.Tag = Sds
            '--- Add Node to DS object
            Sds.AFnode = SrcNode
            '--- Add the node to the Diagram
            If Palette = False Then
                tabAddFlow.Nodes.Add(SrcNode)
            Else
                afPalette.Nodes.Add(SrcNode)
            End If

            If cnt = -1 Then
                tabAddFlow.SelectedItem = SrcNode
            End If
            
            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddSDS")
            Return False
        End Try

    End Function

    Public Function AddMain(ByRef Main As clsTask, Optional ByVal cnt As Integer = -1, Optional ByVal Palette As Boolean = False) As Boolean

        Try
            If Palette = False Then
                Main.LoadMe()
                HorizSP = 20 + HorizIncr
                If cnt > -1 Then
                    VertSP = 20 + (cnt * VertIncr)
                Else
                    Dim nodeCount As Integer = Main.Engine.Mains.Count - 1
                    VertSP = 20 + (nodeCount * VertIncr)
                End If
            Else
                HorizSP = 5
                VertSP = 20 + VertIncr
            End If

            Dim MainNode As New Node
            NodeText = Main.TaskName

            '--- Get the Addflow node for each Main Proc
            MainNode = New Node(HorizSP, VertSP, Nodesize, Nodesize, NodeText, tabAddFlow.DefNodeProp)
            MainNode.GradientColor = Color.LightGreen
            MainNode.Shape.Style = ShapeStyle.Preparation
            MainNode.Shape.Orientation = ShapeOrientation.so_90
            'MainNode.TextColor = Color.Crimson

            MainNode.ImageIndex = 1
            '--- Add Object as Tag of node
            MainNode.Tag = Main
            '--- Add Node to Main object
            Main.AFnode = MainNode
            '--- Add the node to the Diagram
            If Palette = False Then
                tabAddFlow.Nodes.Add(MainNode)
            Else
                afPalette.Nodes.Add(MainNode)
            End If

            If cnt = -1 Then
                tabAddFlow.SelectedItem = MainNode
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddMain")
            Return False
        End Try

    End Function

    Public Function AddGen(ByRef Gen As clsTask, Optional ByVal cnt As Integer = -1, Optional ByVal Palette As Boolean = False) As Boolean

        Try
            '/// create General Nodes
            If Gen.TaskType = enumTaskType.TASK_GEN Then
                If Palette = False Then
                    Gen.LoadMe()
                    HorizSP = 20 + HorizIncr  ' * 2)
                    If cnt > -1 Then
                        VertSP = 50 + VertIncr + (cnt * VertIncr)
                    Else
                        Dim nodeCount As Integer = Gen.Engine.Gens.Count - 1
                        VertSP = 50 + VertIncr + (nodeCount * VertIncr)
                    End If
                Else
                    HorizSP = 5
                    VertSP = 20 + (2 * VertIncr)
                End If

                Dim ProcNode As New Node
                NodeText = Gen.TaskName

                '--- Get the Addflow node for each Procedure
                ProcNode = New Node(HorizSP, VertSP, LogicNodeSizeH, LogicNodeSizeV, NodeText, tabAddFlow.DefNodeProp)
                ProcNode.GradientColor = Color.LightYellow
                ProcNode.Shape.Style = ShapeStyle.Ellipse
                'ProcNode.TextColor = Color.Red

                ProcNode.ImageIndex = 2
                '--- Add Object as Tag of node
                ProcNode.Tag = Gen
                '--- Add Node to Proc object
                Gen.AFnode = ProcNode
                '--- Add the node to the Diagram
                If Palette = False Then
                    tabAddFlow.Nodes.Add(ProcNode)
                Else
                    afPalette.Nodes.Add(ProcNode)
                End If

                If cnt = -1 Then
                    tabAddFlow.SelectedItem = ProcNode
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddGen")
            Return False
        End Try

    End Function

    Public Function AddProc(ByRef Proc As clsTask, Optional ByVal cnt As Integer = -1, Optional ByVal Palette As Boolean = False) As Boolean

        Try
            If Proc.TaskType <> enumTaskType.TASK_GEN And Proc.TaskType <> enumTaskType.TASK_MAIN Then
                If Palette = False Then
                    Proc.LoadMe()
                    HorizSP = 20 + (HorizIncr * 3)
                    If cnt > -1 Then
                        VertSP = 20 + (cnt * VertIncr)
                    Else
                        Dim nodeCount As Integer = (Proc.Engine.Procs.Count - Proc.Engine.Gens.Count) - 1
                        VertSP = 20 + (nodeCount * VertIncr)
                    End If
                Else
                    HorizSP = 5
                    VertSP = 20 + (3 * VertIncr)
                End If

                Dim ProcNode As New Node
                NodeText = Proc.TaskName

                '--- Get the Addflow node for each Procedure
                ProcNode = New Node(HorizSP, VertSP, ProcNodeSize, ProcNodeSize, NodeText, tabAddFlow.DefNodeProp)
                ProcNode.GradientColor = Color.LightCyan
                ProcNode.Shape.Style = ShapeStyle.AlternateProcess
                'ProcNode.TextColor = Color.Red

                ProcNode.ImageIndex = 2
                '--- Add Object as Tag of node
                ProcNode.Tag = Proc
                '--- Add Node to Proc object
                Proc.AFnode = ProcNode
                '--- Add the node to the Diagram
                If Palette = False Then
                    tabAddFlow.Nodes.Add(ProcNode)
                Else
                    afPalette.Nodes.Add(ProcNode)
                End If

                If cnt = -1 Then
                    tabAddFlow.SelectedItem = ProcNode
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddProc")
            Return False
        End Try

    End Function

    Public Function AddTDS(ByVal Tds As clsDatastore, Optional ByVal cnt As Integer = -1, Optional ByVal Palette As Boolean = False) As Boolean

        Try
            If Palette = False Then
                Tds.LoadMe()
                HorizSP = 20 + (HorizIncr * 4)
                If cnt > -1 Then
                    VertSP = 20 + (cnt * VertIncr)
                Else
                    Dim nodeCount As Integer = Tds.Engine.Targets.Count - 1
                    VertSP = 20 + (nodeCount * VertIncr)
                End If
            Else
                HorizSP = 5
                VertSP = 20 + (4 * VertIncr)
            End If

            Dim TgtNode As New Node
            NodeText = Tds.DatastoreName

            '--- Get the Addflow node for each targetDS
            TgtNode = New Node(HorizSP, VertSP, DSnodeSize, DSnodeSize, NodeText, tabAddFlow.DefNodeProp)
            'TgtNode.OutLinkable = False

            TgtNode.GradientColor = Color.LightCoral
            TgtNode.Shape.Style = ShapeStyle.DirectAccessStorage
            TgtNode.Shape.Orientation = ShapeOrientation.so_270
            'TgtNode.TextColor = Color.LightSkyBlue
            'TgtNode.Gradient = True
            TgtNode.ImageIndex = 3
            '--- Add Object as Tag of node
            TgtNode.Tag = Tds
            '--- Add Node to Proc object
            Tds.AFnode = TgtNode
            '--- Add the node to the Diagram
            If Palette = False Then
                tabAddFlow.Nodes.Add(TgtNode)
            Else
                afPalette.Nodes.Add(TgtNode)
            End If

            If cnt = -1 Then
                tabAddFlow.SelectedItem = TgtNode
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddTDS")
            Return False
        End Try

    End Function

#End Region

#Region "Context Menu Items"

    'Private Sub mnuAddSrc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddSrc.Click

    '    Try
    '        ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
    '        My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab mnuAddSrc_Click")
    '    End Try
    'End Sub

    'Private Sub mnuAddTgt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddTgt.Click

    '    Try
    '        ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
    '        My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab mnuAddTgt_Click")
    '    End Try

    'End Sub

    Private Sub mnuAddLU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddLU.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY, , , True)
            RefreshAddFlow()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddLU_Click")
        End Try

    End Sub

    Private Sub mnuAddProcMap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddProcMap.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(3)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW)
            RefreshAddFlow()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddProcMap_Click")
        End Try

    End Sub

    Private Sub mnuAddProcGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddProcGen.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(3)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW, enumTaskType.TASK_GEN)
            RefreshAddFlow()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddProcGen_Click")
        End Try

    End Sub

    Private Sub mnuAddMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddMain.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(4)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW, enumTaskType.TASK_MAIN)
            RefreshAddFlow()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddMain_Click")
        End Try

    End Sub

    Private Sub mnuDelItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelItem.Click

        Try
            Dim itm As Flow.Item = tabAddFlow.SelectedItem 'tabAddFlow.PointedItem
            '= itm

            If itm IsNot Nothing Then
                If itm.Tag IsNot Nothing Then
                    '/// It's a Node
                    Dim NodeTag As INode = CType(itm.Tag, INode)
                    Dim tv As TreeView = NodeTag.ObjTreeNode.TreeView

                    tv.SelectedNode = NodeTag.ObjTreeNode

                    Select Case NodeTag.Type
                        Case NODE_GEN, NODE_PROC, NODE_MAIN
                            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_DELETE)

                        Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_DELETE)

                    End Select
                Else
                    '/// It's a link
                    tabAddFlow.DeleteSel()
                End If
            Else
                '/// it's the Engine
                ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode
                If MsgBox("Are You sure you want to delete the entire engine?" & Chr(13) & _
                          "If not, make sure Item you wish to delete is highlighted in diagram.", _
                          MsgBoxStyle.YesNo, "Delete Engine") = MsgBoxResult.Yes Then
                    My.Forms.frmMain.DoEngineAction(enumAction.ACTION_DELETE)
                End If
            End If

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuDelItem")
        End Try

    End Sub

    Private Sub mnuUndo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuUndo.Click
        tabAddFlow.Undo()
    End Sub

    Private Sub mnuRedo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRedo.Click
        tabAddFlow.Redo()
    End Sub

    Private Sub mnuCloseTab_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseTab.Click
        CloseTab()
    End Sub

    Private Sub mnuGenScript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGenScript.Click

        Try
            Dim envObj As clsEnvironment
            Dim strSaveDir As String

            '/// Get Environment Object
            envObj = ObjEng.ObjSystem.Environment
            envObj.LoadMe()

            '/// Get Script Directory
            If System.IO.Directory.Exists(envObj.LocalScriptDir) = True Then
                strSaveDir = envObj.LocalScriptDir
            Else
                MsgBox("Script Directory in the Environment is either not defined or cannot be accessed." & Chr(13) & _
                "Please connect to or specify a location to generate the script. ", MsgBoxStyle.OkOnly, MsgTitle)
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            '/// Open new Script Gen Form
            Dim frm As New frmScriptGen
            frm.OpenFormEng(ObjEng, strSaveDir)


        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuScriptEngine")
        End Try

    End Sub

#End Region

#Region "Mouse Actions"

    Private Sub tabAddFlow_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tabAddFlow.MouseDoubleClick

        Try
            Dim itm As Flow.Item = tabAddFlow.SelectedItem

            If itm IsNot Nothing Then
                If itm.Tag IsNot Nothing Then
                    '/// It's a Node
                    Dim NodeTag As INode = itm.Tag
                    Dim tv As TreeView = NodeTag.ObjTreeNode.TreeView

                    tv.SelectedNode = NodeTag.ObjTreeNode
                Else
                    '/// It's a link

                End If
            Else
                '/// it's the Engine
                ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode
            End If

            '/// Activate the Property Tab
            Dim tp As TabPage = Me.Parent
            Dim tc As TabControl = tp.Parent

            tc.SelectTab(0)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_MouseDoubleClick")
        End Try

    End Sub

    Private Sub tabAddFlow_MouseHoverMouseClick(ByVal sender As System.Object, ByVal e As EventArgs) Handles tabAddFlow.MouseHover, tabAddFlow.MouseUp   'System.Windows.Forms.Mouse  ', tabAddFlow.MouseClick

        Try
            tabAddFlow.SelectedItem = tabAddFlow.PointedItem

            If tabAddFlow.PointedItem IsNot Nothing Then
                If tabAddFlow.PointedItem.Tag IsNot Nothing Then
                    '/// It's a Node
                    Dim NodeTag As INode = CType(tabAddFlow.PointedItem.Tag, INode)
                    Dim tv As TreeView = NodeTag.ObjTreeNode.TreeView
                    tv.SelectedNode = NodeTag.ObjTreeNode
                    txtSelected.Text = NodeTag.Text
                    mnuDelItem.Text = "Delete Node > " & NodeTag.Text
                Else
                    '/// It's a link
                    txtSelected.Text = ""
                    mnuDelItem.Text = "Delete Link"
                End If
            Else
                '/// it's the Engine
                If ObjEng Is Nothing Then
                    ObjEng = Me.Tag
                End If
                ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode
                txtSelected.Text = ObjEng.Text
                mnuDelItem.Text = "Delete Engine > " & ObjEng.Text
            End If

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_MouseHoverMouseClick")
        End Try

    End Sub

    Private Sub tabAddFlow_MouseDownMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tabAddFlow.MouseUp  'tabAddFlow.MouseClick ',  ', tabAddFlow.MouseHover   'tabAddFlow.MouseDown, 
        Try
            Dim ItemSelected As Boolean

            tabAddFlow.SelectedItem = tabAddFlow.PointedItem

            If tabAddFlow.PointedItem Is Nothing Then
                '/// It's the Engine
                ItemSelected = False
                txtSelected.Text = ObjEng.Text
            Else
                If AddFlow.Equals(tabAddFlow.SelectedItem, tabAddFlow.PointedItem) = False Then ' Is Nothing Then
                    tabAddFlow.SelectedItem = tabAddFlow.PointedItem
                End If
                txtSelected.Text = tabAddFlow.SelectedItem.Text
                ItemSelected = True
            End If

            If e.Button = Windows.Forms.MouseButtons.Right Then
                mnuAddMain.Visible = Not ItemSelected
                mnuAddProcGen.Visible = Not ItemSelected
                mnuAddProcMap.Visible = Not ItemSelected
                mnuAddLU.Visible = Not ItemSelected
                mnuAddSrc.Visible = Not ItemSelected
                mnuAddTgt.Visible = Not ItemSelected
                mnuGenScript.Visible = True 'Not ItemSelected
                mnuRedo.Enabled = True
                mnuUndo.Enabled = True
                mnuDelItem.Enabled = True
                mnuCloseTab.Visible = False
            End If
            cms1.Refresh()


        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_MouseDownMouseClick")
        End Try

    End Sub

#End Region

#Region "Toolbar Actions"

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click

        Try
            Dim ofd As OpenFileDialog = New OpenFileDialog

            ofd.InitialDirectory = GetAppData()
            ofd.Filter = "AddFlow Diagrams(*.xml)|*.xml|All Files(*.*)|*.*"
            ofd.FileName = "*.xml"

            If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
                LoadFile(ofd.FileName)
            End If

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab OpenToolStripButton_Click")
        End Try

    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click

        Try
            Dim sfd As SaveFileDialog = New SaveFileDialog

            If (strFileName.Length > 1) Then
                sfd.FileName = strFileName
            Else
                sfd.FileName = ObjEng.EngineName & ".xml"
            End If

            sfd.Filter = "AddFlow Diagrams(*.xml)|*.xml|All Files(*.*)|*.*"
            sfd.InitialDirectory = GetAppData()

            If sfd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                strFileName = sfd.FileName
                SaveFile()
            End If

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab SaveToolStripButton_Click")
        End Try

    End Sub

    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripButton.Click

        Try
            PrnFlow1.PageSetup()
            PrnFlow1.Preview(tabAddFlow)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab PrintToolStripButton_Click")
        End Try

    End Sub

    Private Sub ZoomIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomIn.Click

        Try
            tabAddFlow.Zoom.X = tabAddFlow.Zoom.X + 0.1
            tabAddFlow.Zoom.Y = tabAddFlow.Zoom.Y + 0.1
            tabAddFlow.Refresh()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab ZoomIn_Click")
        End Try

    End Sub

    Private Sub ZoomOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomOut.Click

        Try
            If tabAddFlow.Zoom.X > 0.1 Then
                tabAddFlow.Zoom.X = tabAddFlow.Zoom.X - 0.1
                tabAddFlow.Zoom.Y = tabAddFlow.Zoom.Y - 0.1
            End If
            
            tabAddFlow.Refresh()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab ZoomOut_Click")
        End Try

    End Sub

    Private Sub ZoomNorm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ZoomNorm.Click

        Try
            tabAddFlow.Zoom.X = 1
            tabAddFlow.Zoom.Y = 1
            tabAddFlow.Refresh()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab ZoomNorm_Click")
        End Try

    End Sub

    Private Sub UpdateDiag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateDiag.Click

        Try
            RefreshAddFlow()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab UpdateDiag_Click")
        End Try

    End Sub

    Private Sub UpdateTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateTree.Click

        Try

            tabAddFlow.BackColor = Color.FromKnownColor(KnownColor.InactiveCaptionText)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab UpdateTree_Click")
        End Try

    End Sub

#End Region
   
#Region "Helper Functions"

    Private Sub CloseTab()

        Try



        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab CloseTab")
        End Try

    End Sub

    Private Sub LoadAFNodes()

        Try
            Dim count As Integer = 0
            '/// create Source node
            For Each SrcDS As clsDatastore In ObjEng.Sources
                AddSDS(SrcDS, count)
                count += 1
            Next

            count = 0

            '/// create Main Node
            For Each main As clsTask In ObjEng.Mains
                AddMain(main, count)
                count += 1
            Next

            count = 0
            '/// create General Nodes
            For Each proc As clsTask In ObjEng.Procs
                If proc.TaskType = enumTaskType.TASK_GEN Then
                    AddGen(proc, count)
                    count += 1
                End If
            Next

            count = 0
            '/// create Procedure Nodes
            For Each proc As clsTask In ObjEng.Procs
                If proc.TaskType <> enumTaskType.TASK_GEN And proc.TaskType <> enumTaskType.TASK_MAIN Then
                    AddProc(proc, count)
                    count += 1
                End If
            Next

            count = 0
            '/// create Target Nodes
            For Each TgtDS As clsDatastore In ObjEng.Targets
                AddTDS(TgtDS, count)
                count += 1
            Next

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab CreateAFNodes")
        End Try

    End Sub

    Private Sub LoadAFLinks()

        Try
            Dim PairCol As New Collection
            Dim GenPairCol As New Collection

            For Each main As clsTask In ObjEng.Mains
                '/// Link Source to Main
                For Each sDS As clsDatastore In main.ObjSources
                    Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                    sDS.AFnode.OutLinks.Add(link1, main.AFnode)
                Next

                '/// Get list of procs to map to, from main syntax
                If main.ObjMappings.Count = 1 Then
                    Dim MainMap As clsMapping = CType(main.ObjMappings(0), clsMapping)
                    Dim mainFun As clsSQFunction = CType(MainMap.MappingSource, clsSQFunction)
                    Dim FunText As String = mainFun.SQFunctionWithInnerText
                    '/// see if there is a main script or not, first
                    If FunText.Contains("CASE") Or FunText.Contains("RECNAME") Or _
                    FunText.Contains("CALLPROC") Or FunText.Contains("WHEN") Then
                        If GetSrcTablesAndProcsFromMain(PairCol, FunText) = False Then
                            MsgBox("There was an error in your Main Procedure")
                        End If
                    Else
                        'If GetTargetsFromGen(PairCol, FunText) = False Then
                        '    MsgBox("There was an error in your Logic Procedure(s)")
                        'End If
                    End If
                End If
                '/// Build Links from Main to Procs from PairCollection
                For Each map As clsMapping In PairCol
                    If map.MappingTarget IsNot Nothing Then
                        Dim link2 As Link = New Link(tabAddFlow.DefLinkProp)
                        If CType(map.MappingTarget, INode).Type = NODE_PROC Or _
                        CType(map.MappingTarget, INode).Type = NODE_GEN Then
                            main.AFnode.OutLinks.Add(link2, CType(map.MappingTarget, clsTask).AFnode)
                        End If
                    End If
                Next
            Next

            ''/// Build links from Gens to Procs
            For Each gen As clsTask In ObjEng.Gens
                '/// Get list of procs to map to, from Logic syntax
                If gen.ObjMappings.Count = 1 Then
                    Dim GenMap As clsMapping = CType(gen.ObjMappings(0), clsMapping)
                    Dim GenFun As clsSQFunction = CType(GenMap.MappingSource, clsSQFunction)
                    Dim GenText As String = GenFun.SQFunctionWithInnerText
                    '/// see if there is a main script or not, first
                    'If GenText.Contains("CASE") Or GenText.Contains("RECNAME") Or _
                    'GenText.Contains("CALLPROC") Or GenText.Contains("WHEN") Then
                    '    If GetSrcTablesAndProcsFromMain(GenPairCol, GenText) = False Then
                    '        MsgBox("There was an error in your Main Procedure")
                    '    End If
                    'Else
                    If GetTargetsFromGen(GenPairCol, GenText) = False Then
                        MsgBox("There was an error in Logic Procedure >>> " & gen.TaskName)
                    End If
                    'End If
                End If
                For Each map As clsMapping In GenPairCol
                    '/// Targets are datastores
                    If map.MappingTarget IsNot Nothing Then
                        Dim link2 As Link = New Link(tabAddFlow.DefLinkProp)
                        'If CType(map.MappingTarget, INode).Type = NODE_PROC Or _
                        'CType(map.MappingTarget, INode).Type = NODE_GEN Then
                        gen.AFnode.OutLinks.Add(link2, CType(map.MappingTarget, clsDatastore).AFnode)
                        'End If
                    End If
                    '/// SOurces are Procs
                    If map.MappingSource IsNot Nothing Then
                        Dim link3 As Link = New Link(tabAddFlow.DefLinkProp)
                        'If CType(map.MappingTarget, INode).Type = NODE_PROC Or _
                        'CType(map.MappingTarget, INode).Type = NODE_GEN Then
                        gen.AFnode.OutLinks.Add(link3, CType(map.MappingSource, clsTask).AFnode)
                        'End If
                    End If
                Next
                '/// map SrcDS to Gens not referenced in Main and where src has no link
                For Each sDS As clsDatastore In gen.ObjSources
                    If sDS.AFnode.OutLinks.Count < 1 Then
                        Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                        sDS.AFnode.OutLinks.Add(link1, gen.AFnode)
                    End If
                Next
            Next

            '/// Build links from Mapping Procs to Targets
            For Each Task As clsTask In ObjEng.Procs
                For Each Tgt As clsDatastore In Task.ObjTargets
                    Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                    Task.AFnode.OutLinks.Add(link1, Tgt.AFnode)
                Next
            Next

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab CreateAFLinks")
        End Try

    End Sub

    Function GetSrcTablesAndProcsFromMain(ByRef PairCol As Collection, ByVal Intext As String) As Boolean

        Try
            Dim NewItem As clsMapping

            For Each srcDS As clsDatastore In ObjEng.Sources
                For Each DSsel As clsDSSelection In srcDS.ObjSelections
                    If Intext.Contains(Quote(DSsel.SelectionName, "'")) Then
                        NewItem = New clsMapping
                        NewItem.MappingSource = DSsel
                        PairCol.Add(NewItem)
                        'Exit For
                    End If
                Next
            Next
            For Each proc As clsTask In ObjEng.Procs
                If Intext.Contains("CALLPROC(" & proc.TaskName & ")") Then
                    NewItem = New clsMapping
                    NewItem.MappingTarget = proc
                    PairCol.Add(NewItem)
                    'Exit For
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab GetSrcTablesAndProcsFromMain")
            Return False
        End Try

    End Function

    '    Function GetSrcTablesAndProcsFromMain(ByRef PairCol As Collection, ByVal Intext As String) As Boolean

    '        Try
    '            Dim pararray() As Char = Chr(13) & Chr(10) & Chr(12) '& Chr(9)  '" ()'" & 
    '            Dim InTextArray() As String = Intext.Split(pararray)
    '            Dim i As Integer = 0
    '            Dim tempstr As String = ""
    '            Dim sb As New System.Text.StringBuilder
    '            Dim seglist As New Collection

    '            '/// Split up the Text array into segments based on key words (i.e.When in the case of a Routing template)
    '            For i = 0 To InTextArray.Length - 1
    '                tempstr = InTextArray(i)
    '                If tempstr = "WHEN" Then
    '                    seglist.Add(sb.ToString)
    '                    sb.Remove(0, sb.Length)
    '                    sb.Append(tempstr & "~")
    '                Else
    '                    If tempstr <> "" Then
    '                        sb.Append(tempstr & "~")
    '                    End If
    '                End If
    '            Next
    '            seglist.Add(sb.ToString)

    '            For Each seg As String In seglist
    '                If seg.Contains("WHEN") = False Then
    '                    GoTo gotohere
    '                End If
    '                For Each srcDS As clsDatastore In ObjEng.Sources
    '                    Dim NewItem As clsMapping
    '                    For Each DSsel As clsDSSelection In srcDS.ObjSelections
    '                        If seg.Contains(DSsel.SelectionName & "~") Then
    '                            NewItem = New clsMapping
    '                            NewItem.MappingSource = DSsel
    '                            PairCol.Add(NewItem)
    '                            'Exit For
    '                        End If
    '                    Next
    '                    For Each proc As clsTask In ObjEng.Procs
    '                        If seg.Contains(proc.TaskName & "~") Then
    '                            NewItem = New clsMapping
    '                            NewItem.MappingTarget = proc
    '                            PairCol.Add(NewItem)
    '                            'Exit For
    '                        End If
    '                    Next
    '                    'PairCol.Add(NewItem)
    '                Next
    'gotohere:   Next

    '            Return True

    '        Catch ex As Exception
    '            LogError(ex, "ctlAddFlowTab GetSrcTablesAndProcsFromMain")
    '            Return False
    '        End Try

    '    End Function

    Function GetTargetsFromGen(ByRef PairCol As Collection, ByVal Intext As String) As Boolean

        Try
            Dim NewItem As clsMapping

            For Each tgtDS As clsDatastore In ObjEng.Targets
                For Each DSsel As clsDSSelection In tgtDS.ObjSelections
                    If Intext.Contains(Quote(DSsel.SelectionName, "'")) Then
                        NewItem = New clsMapping
                        NewItem.MappingTarget = DSsel.ObjDatastore
                        PairCol.Add(NewItem)
                        'Exit For
                    End If
                Next
            Next
            For Each proc As clsTask In ObjEng.Procs
                If Intext.Contains("CALLPROC(" & proc.TaskName & ")") Then
                    NewItem = New clsMapping
                    NewItem.MappingSource = proc
                    PairCol.Add(NewItem)
                    'Exit For
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab GetTargetsFromGen")
            Return False
        End Try

    End Function

    '    Function GetTargetsFromGen(ByRef PairCol As Collection, ByVal Intext As String) As Boolean

    '        Try
    '            Dim pararray() As Char = " ()'" & Chr(13) & Chr(10) & Chr(12) & Chr(9)
    '            Dim InTextArray() As String = Intext.Split(pararray)
    '            Dim i As Integer = 0
    '            Dim tempstr As String = ""
    '            Dim sb As New System.Text.StringBuilder
    '            Dim seglist As New Collection

    '            For i = 0 To InTextArray.Length - 1
    '                tempstr = InTextArray(i)
    '                If tempstr = "IF" Or tempstr = "WHEN" Or tempstr = "OTHERWISE" Or tempstr = "LOOK" Or tempstr = "DO" Then
    '                    seglist.Add(sb.ToString)
    '                    sb.Remove(0, sb.Length)
    '                    sb.Append(tempstr & "~")
    '                Else
    '                    If tempstr <> "" Then
    '                        sb.Append(tempstr & "~")
    '                    End If
    '                End If
    '            Next
    '            seglist.Add(sb.ToString)

    '            For Each seg As String In seglist
    '                'If seg.Contains("WHEN") = False Then
    '                '    GoTo gotohere
    '                'End If
    '                For Each srcDS As clsDatastore In ObjEng.Sources
    '                    Dim NewItem As New clsMapping
    '                    For Each DSsel As clsDSSelection In srcDS.ObjSelections
    '                        If seg.Contains(DSsel.SelectionName & "~") Then
    '                            NewItem.MappingSource = DSsel
    '                            Exit For
    '                        End If
    '                    Next
    '                    For Each proc As clsTask In ObjEng.Procs
    '                        If seg.Contains(proc.TaskName & "~") Then
    '                            NewItem.MappingTarget = proc
    '                            Exit For
    '                        End If
    '                    Next
    '                    For Each main As clsTask In ObjEng.Mains
    '                        If seg.Contains(main.TaskName) Then  '& "~"
    '                            NewItem.MappingTarget = main
    '                            Exit For
    '                        End If
    '                    Next
    '                    PairCol.Add(NewItem)
    '                Next
    'gotohere:   Next

    '            Return True

    '        Catch ex As Exception
    '            LogError(ex, "ctlAddFlowTab GetTargetsFromGen")
    '            Return False
    '        End Try

    '    End Function

    Private Sub LoadFile(ByVal strFileName As String)

        Try
            '/// Reset AddFlow
            tabAddFlow.ResetUndoRedo()
            tabAddFlow.Nodes.Clear()

            '/// Read in New Diagram
            Dim reader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(strFileName)

            reader.WhitespaceHandling = System.Xml.WhitespaceHandling.None
            tabAddFlow.ReadXml(reader)
            reader.Close()

            tabAddFlow.SetChangedFlag(False)
            Me.strFileName = strFileName

            '/// Create Engine from Nodes and Links



            tabAddFlow.BackColor = Color.FromKnownColor(KnownColor.Control)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab LoadFile")
        End Try

    End Sub

    Private Sub SaveFile()

        Try
            Cursor = Cursors.WaitCursor
            Dim writer As System.Xml.XmlTextWriter = New System.Xml.XmlTextWriter(strFileName, Nothing)
            writer.Formatting = System.Xml.Formatting.Indented
            tabAddFlow.WriteXml(writer)
            writer.Close()
            tabAddFlow.SetChangedFlag(False)
            Cursor = Cursors.Default


        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab SaveFile")
        End Try

    End Sub

#End Region

#Region "Palette"

    'Private Sub LoadPalette()

    '    Try

    '        '/// Add Source to Palette
    '        Dim NewSDS As New clsDatastore
    '        With NewSDS
    '            .DatastoreType = enumDatastore.DS_BINARY
    '            .DsDirection = DS_DIRECTION_SOURCE
    '        End With
    '        AddSDS(NewSDS, 0, True)

    '        '/// Add Main to Palette
    '        Dim NewMain As New clsTask
    '        With NewMain
    '            .TaskType = enumTaskType.TASK_MAIN
    '        End With
    '        AddMain(NewMain, 0, True)

    '        '/// Add General Proc to Palette
    '        Dim NewGen As New clsTask
    '        With NewGen
    '            .TaskType = enumTaskType.TASK_GEN
    '        End With
    '        AddGen(NewGen, 0, True)

    '        '/// Add Mapping Proc
    '        Dim NewProc As New clsTask
    '        With NewProc
    '            .TaskType = enumTaskType.TASK_PROC
    '        End With
    '        AddProc(NewProc, 0, True)

    '        '/// Add Target to Palette
    '        Dim NewTGT As New clsDatastore
    '        With NewTGT
    '            .DsDirection = DS_DIRECTION_TARGET
    '            .DatastoreType = enumDatastore.DS_BINARY
    '        End With
    '        AddTDS(NewTGT, 0, True)

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab LoadPalette")
    '    End Try

    'End Sub

    'Private Sub afPalette_AfterSelect(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles afPalette.AfterSelect

    '    Try
    '        Dim itm As Node = afPalette.SelectedItem

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab afPalette_AfterSelect")
    '    End Try

    'End Sub

    'Private Sub afPalette_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles afPalette.MouseDown

    '    Try
    '        afPalette.SelectedItem = afPalette.PointedItem
    '        afPalette.Refresh()

    '        If afPalette.PointedItem IsNot Nothing Then
    '            Windows.Forms.Clipboard.Clear()
    '            Windows.Forms.Clipboard.SetDataObject(afPalette.PointedItem)

    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab afPalette_MouseDown")
    '    End Try

    'End Sub

    'Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click

    '    Try
    '        If tabAddFlow.PointedItem Is Nothing Then
    '            Dim Node1 As New Node
    '            Node1 = CType(Windows.Forms.Clipboard.GetDataObject(), Node)
    '            tabAddFlow.Nodes.Add(Node1)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlAddFlowTab mnuPaste_Click")
    '    End Try

    'End Sub

#End Region

#Region "SrcDS Click Events"

    Private Sub mnuSrcBinary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcBinary.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcBinary_Click")
        End Try

    End Sub

    Private Sub mnuSrcText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcText.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_TEXT)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcText_Click")
        End Try

    End Sub

    Private Sub mnuSrcDelimited_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcDelimited.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DELIMITED)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcDelimited_Click")
        End Try

    End Sub

    Private Sub mnuSrcXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcXML.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_XML)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcXML_Click")
        End Try

    End Sub

    Private Sub mnuSrcRelational_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcRelational.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_RELATIONAL)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcRelational_Click")
        End Try

    End Sub

    Private Sub mnuSrcDB2LOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcDB2LOAD.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DB2LOAD)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcDB2LOAD_Click")
        End Try

    End Sub

    Private Sub mnuSrcHSSUNLOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcHSSUNLOAD.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_HSSUNLOAD)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcHSSUNLOAD_Click")
        End Try

    End Sub

    Private Sub mnuSrcIMSDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcIMSDB.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IMSDB)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcIMSDB_Click")
        End Try

    End Sub

    Private Sub mnuSrcVSAM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcVSAM.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_VSAM)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcVSAM_Click")
        End Try

    End Sub

    Private Sub mnuSrcIMSCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcIMSCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IMSCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcIMSCDC_Click")
        End Try

    End Sub

    Private Sub mnuSrcIMSCDCLE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcIMSCDCLE.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IMSCDCLE)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcIMSCDCLE_Click")
        End Try

    End Sub

    Private Sub mnuSrcDB2CDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcDB2CDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DB2CDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcDB2CDC_Click")
        End Try

    End Sub

    Private Sub mnuSrcVSAMCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcVSAMCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_VSAMCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcVSAMCDC_Click")
        End Try

    End Sub

    Private Sub mnuSrcOracleCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcOracleCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_ORACLECDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcOracleCDC_Click")
        End Try

    End Sub

    Private Sub mnuSrcUTSCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcUTSCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_UTSCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcUTSCDC_Click")
        End Try

    End Sub

    Private Sub mnuSrcIBMEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcIBMEvent.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IBMEVENT)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcIBMEvent_Click")
        End Try

    End Sub

    Private Sub mnuSrcIncludeFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSrcIncludeFile.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_INCLUDE)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuSrcIncludeFile_Click")
        End Try

    End Sub

#End Region

#Region "TgtDS Click Events"

    Private Sub mnuTgtBinary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtBinary.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtBinary_Click")
        End Try

    End Sub

    Private Sub mnuTgtText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtText.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_TEXT)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtText_Click")
        End Try

    End Sub

    Private Sub mnuTgtDelimited_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtDelimited.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DELIMITED)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtDelimited_Click")
        End Try

    End Sub

    Private Sub mnuTgtXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtXML.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_XML)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtXML_Click")
        End Try

    End Sub

    Private Sub mnuTgtRelational_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtRelational.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_RELATIONAL)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtRelational_Click")
        End Try

    End Sub

    Private Sub mnuTgtDB2LOAD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtDB2LOAD.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DB2LOAD)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtDB2LOAD_Click")
        End Try

    End Sub

    Private Sub mnuTgtIMSDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtIMSDB.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IMSDB)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtIMSDB_Click")
        End Try

    End Sub

    Private Sub mnuTgtVSAM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtVSAM.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_VSAM)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtVSAM_Click")
        End Try

    End Sub

    Private Sub mnuTgtIMSCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtIMSCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IMSCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtIMSCDC_Click")
        End Try

    End Sub

    Private Sub mnuTgtDB2CDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtDB2CDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_DB2CDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtDB2CDC_Click")
        End Try

    End Sub

    Private Sub mnuTgtVSAMCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtVSAMCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_VSAMCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtVSAMCDC_Click")
        End Try

    End Sub

    Private Sub mnuTgtOracleCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtOracleCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_ORACLECDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtOracleCDC_Click")
        End Try

    End Sub

    Private Sub mnuTgtUTSCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtUTSCDC.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_UTSCDC)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtUTSCDC_Click")
        End Try

    End Sub

    Private Sub mnuTgtIBMEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtIBMEvent.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_IBMEVENT)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtIBMEvent_Click")
        End Try

    End Sub

    Private Sub mnuTgtIncludeFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuTgtIncludeFile.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_INCLUDE)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuTgtIncludeFile_Click")
        End Try

    End Sub

#End Region

#Region "Link Event Handlers"

    Sub tabAddFlow_AfterAddLink(ByVal sender As Object, ByVal e As AfterAddLinkEventArgs) Handles tabAddFlow.AfterAddLink

        Try
            If IsEventFromCode = True Then Exit Sub

            Dim OrgNode As Node = e.Link.Org
            Dim OrgObj As INode = OrgNode.Tag
            Dim OrgType As String = OrgObj.Type
            Dim DstNode As Node = e.Link.Dst
            Dim DstObj As INode = DstNode.Tag
            Dim DstType As String = DstObj.Type

            'MsgBox("OrgNode > " & OrgNode.Text & Chr(13) & _
            '       "DstNode > " & DstNode.Text & Chr(13) & _
            '       "OrgType >> " & OrgType & Chr(13) & _
            '       "DstType >> " & DstType _
            '       , MsgBoxStyle.Information)

            Select Case OrgType
                Case NODE_GEN
                    'allow links to Tgts, procs
                    If DstType = NODE_TARGETDATASTORE Then
                        If AddLink_TaskToTarget(CType(DstObj, clsDatastore), CType(OrgObj, clsTask)) = True Then
                            CType(DstObj, clsDatastore).AFnode.DrawColor = Color.Red
                            CType(OrgObj, clsTask).AFnode.DrawColor = Color.Red
                        End If
                    ElseIf DstType = NODE_PROC Then
                        If AddLink_TaskToTask(CType(DstObj, clsTask), CType(OrgObj, clsTask)) = True Then
                            CType(DstObj, clsTask).AFnode.DrawColor = Color.Red
                            CType(OrgObj, clsTask).AFnode.DrawColor = Color.Red
                        End If
                    Else
                        MsgBox("Not Allowed to link to this node", MsgBoxStyle.Information)
                        e.Link.Remove()
                    End If

                Case NODE_LOOKUP
                    'Only allow link to Main or Gen
                    If DstType = NODE_MAIN Or DstType = NODE_GEN Then
                        'Only one src per main or Gen
                        If CType(DstObj, clsTask).ObjSources.Count > 0 Then
                            MsgBox("Only one Source Lookup allowed per Main or Logic procedure", MsgBoxStyle.Information)
                            e.Link.Remove()
                        Else
                            If AddLink_SrcToTask(CType(DstObj, clsTask), CType(OrgObj, clsDatastore)) = True Then
                                CType(DstObj, clsTask).AFnode.DrawColor = Color.Red
                                CType(OrgObj, clsDatastore).AFnode.DrawColor = Color.Red
                            End If
                        End If
                    Else
                        MsgBox("Not Allowed to draw links to this Type of Node", MsgBoxStyle.Information)
                        e.Link.Remove()
                    End If

                Case NODE_MAIN
                    'allow links to Gens, Procs, Tgts
                    If DstType = NODE_GEN Or DstType = NODE_PROC Then
                        If AddLink_TaskToTask(CType(DstObj, clsTask), CType(OrgObj, clsTask)) = True Then
                            CType(DstObj, clsTask).AFnode.DrawColor = Color.Red
                            CType(OrgObj, clsTask).AFnode.DrawColor = Color.Red
                        End If
                    ElseIf DstType = NODE_TARGETDATASTORE Then
                        If AddLink_TaskToTarget(CType(DstObj, clsDatastore), CType(OrgObj, clsTask)) = True Then
                            CType(DstObj, clsDatastore).AFnode.DrawColor = Color.Red
                            CType(OrgObj, clsTask).AFnode.DrawColor = Color.Red
                        End If
                    Else
                        MsgBox("Not Allowed to draw links to this Type of Node", MsgBoxStyle.Information)
                        e.Link.Remove()
                    End If

                Case NODE_PROC
                    'allow links to Tgts
                    If DstType = NODE_TARGETDATASTORE Then
                        If AddLink_TaskToTarget(CType(DstObj, clsDatastore), CType(OrgObj, clsTask)) = True Then
                            CType(DstObj, clsDatastore).AFnode.DrawColor = Color.Red
                            CType(OrgObj, clsTask).AFnode.DrawColor = Color.Red
                        End If
                    Else
                        MsgBox("Not Allowed to draw links to this Type of Node", MsgBoxStyle.Information)
                        e.Link.Remove()
                    End If

                Case NODE_SOURCEDATASTORE
                    'Only allow link to Main or Gen
                    If DstType = NODE_MAIN Or DstType = NODE_GEN Then
                        'Only one src per main or Gen
                        If CType(DstObj, clsTask).ObjSources.Count > 0 Then
                            MsgBox("Only one Source allowed per Main or Logic procedure", MsgBoxStyle.Information)
                            e.Link.Remove()
                        Else
                            If AddLink_SrcToTask(CType(DstObj, clsTask), CType(OrgObj, clsDatastore)) = True Then
                                CType(DstObj, clsTask).AFnode.DrawColor = Color.Red
                                CType(OrgObj, clsDatastore).AFnode.DrawColor = Color.Red
                            End If
                        End If
                    Else
                        MsgBox("Not Allowed to draw links to this Type of Node", MsgBoxStyle.Information)
                        e.Link.Remove()
                    End If

                Case NODE_TARGETDATASTORE
                    'Don't allow Target to outLink
                    MsgBox("Not Allowed to draw links out from Targets" & Chr(13) & "Draw links to Targets", MsgBoxStyle.Information)
                    e.Link.Remove()

            End Select

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_AfterAddLink")
        End Try

    End Sub

    Function AddLink_TaskToTarget(ByVal DstObj As clsDatastore, ByVal OrgObj As clsTask) As Boolean

        Try
            If OrgObj.TaskType = enumTaskType.TASK_PROC Then
                OrgObj.ObjTargets.Add(DstObj)

            ElseIf OrgObj.TaskType = enumTaskType.TASK_GEN Or OrgObj.TaskType = enumTaskType.TASK_MAIN Then
                Dim InsStr1 As String = "-- Added Target into this Procedure to use in a function"
                Dim InsStr2 As String = "       '" & DstObj.DatastoreName & "'"

                If OrgObj.ObjMappings.Count = 1 Then
                    Dim map As clsMapping = OrgObj.ObjMappings(0)
                    Dim mapsrc As clsSQFunction = map.MappingSource
                    Dim sb As New System.Text.StringBuilder
                    sb.Append(mapsrc.SQFunctionWithInnerText)
                    sb.AppendLine(InsStr1)
                    sb.AppendLine(InsStr2)
                    mapsrc.SQFunctionWithInnerText = sb.ToString
                    mapsrc.SQFunctionName = mapsrc.SQFunctionWithInnerText
                Else
                    Dim map As New clsMapping
                    Dim mapsrc As New clsSQFunction
                    Dim sb As New System.Text.StringBuilder
                    'sb.Append(mapsrc.SQFunctionWithInnerText)
                    sb.AppendLine(InsStr1)
                    sb.AppendLine(InsStr1)
                    mapsrc.SQFunctionWithInnerText = sb.ToString
                    mapsrc.SQFunctionName = mapsrc.SQFunctionWithInnerText
                    map.MappingSource = mapsrc
                    map.MappingTarget = Nothing
                    map.TargetType = enumMappingType.MAPPING_TYPE_NONE
                    OrgObj.ObjMappings.Add(map)
                End If
                'OrgObj.IsModified = True
                OrgObj.Save()

            Else
                Return False
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddLink_TaskToTarget")
            Return False
        End Try

    End Function

    Function AddLink_TaskToTask(ByVal DstObj As clsTask, ByVal OrgObj As clsTask) As Boolean

        Try
            Dim InsStr1 As String = "-- Added Procedure to use in a function"
            Dim InsStr2 As String = "       CALLPROC(" & DstObj.TaskName & ")"

            If OrgObj.ObjMappings.Count = 1 Then
                Dim map As clsMapping = OrgObj.ObjMappings(0)
                Dim mapsrc As clsSQFunction = map.MappingSource
                Dim sb As New System.Text.StringBuilder
                sb.Append(mapsrc.SQFunctionWithInnerText)
                sb.AppendLine(InsStr1)
                sb.AppendLine(InsStr2)
                mapsrc.SQFunctionWithInnerText = sb.ToString
                mapsrc.SQFunctionName = mapsrc.SQFunctionWithInnerText
            Else
                Dim map As New clsMapping
                Dim mapsrc As New clsSQFunction
                Dim sb As New System.Text.StringBuilder
                'sb.Append(mapsrc.SQFunctionWithInnerText)
                sb.AppendLine(InsStr1)
                sb.AppendLine(InsStr2)
                mapsrc.SQFunctionWithInnerText = sb.ToString
                mapsrc.SQFunctionName = mapsrc.SQFunctionWithInnerText
                map.MappingSource = mapsrc
                map.MappingTarget = Nothing
                map.TargetType = enumMappingType.MAPPING_TYPE_NONE
                OrgObj.ObjMappings.Add(map)
            End If
            OrgObj.Save()


            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddLink_TaskToTask")
            Return False
        End Try

    End Function

    Function AddLink_SrcToTask(ByVal DstObj As clsTask, ByVal OrgObj As clsDatastore) As Boolean

        Try
            DstObj.ObjTargets.Add(OrgObj)
            DstObj.IsModified = True
            DstObj.Save()

            Return True

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab AddLink_SrcToTask")
            Return False
        End Try

    End Function

#End Region

End Class
