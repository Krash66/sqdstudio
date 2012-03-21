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
    'Public af As AddFlow

#Region "Public Functions"

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
            LoadAFNodes()
            ' Now Add the Links
            LoadAFLinks()

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
            SrcNode.GradientColor = Color.LightBlue
            If Sds.IsLookUp = False Then
                SrcNode.Shape.Style = ShapeStyle.DirectAccessStorage
                SrcNode.Shape.Orientation = ShapeOrientation.so_270
            Else
                SrcNode.Shape.Style = ShapeStyle.StoredData
                SrcNode.Shape.Orientation = ShapeOrientation.so_0
            End If

            'SrcNode.TextColor = Color.LightGreen

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
            'VertSP = VertSP + VertIncr

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
            MainNode.Shape.Style = ShapeStyle.Pentagon
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
                    HorizSP = 20 + (HorizIncr * 2)
                    If cnt > -1 Then
                        VertSP = 20 + VertIncr + (cnt * VertIncr)
                    Else
                        Dim nodeCount As Integer = Gen.Engine.Gens.Count - 1
                        VertSP = 20 + VertIncr + (nodeCount * VertIncr)
                    End If
                Else
                    HorizSP = 5
                    VertSP = 20 + (2 * VertIncr)
                End If

                Dim ProcNode As New Node
                NodeText = Gen.TaskName

                '--- Get the Addflow node for each Procedure
                ProcNode = New Node(HorizSP, VertSP, ProcNodeSize, ProcNodeSize, NodeText, tabAddFlow.DefNodeProp)
                ProcNode.GradientColor = Color.LightYellow
                ProcNode.Shape.Style = ShapeStyle.Process
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
                ProcNode.Shape.Style = ShapeStyle.Merge
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

    Private Sub mnuAddSrc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddSrc.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddSrc_Click")
        End Try
    End Sub

    Private Sub mnuAddTgt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddTgt.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(1)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddTgt_Click")
        End Try

    End Sub

    Private Sub mnuAddLU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddLU.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(0)
            My.Forms.frmMain.DoDatastoreAction(enumAction.ACTION_NEW, enumDatastore.DS_BINARY, , , True)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddLU_Click")
        End Try

    End Sub

    Private Sub mnuAddProcMap_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddProcMap.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(3)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddProcMap_Click")
        End Try

    End Sub

    Private Sub mnuAddProcGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddProcGen.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(3)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW, enumTaskType.TASK_GEN)

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab mnuAddProcGen_Click")
        End Try

    End Sub

    Private Sub mnuAddMain_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddMain.Click

        Try
            ObjEng.ObjTreeNode.TreeView.SelectedNode = ObjEng.ObjTreeNode.Nodes(4)
            My.Forms.frmMain.DoTaskAction(enumAction.ACTION_NEW, enumTaskType.TASK_MAIN)

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

    Private Sub tabAddFlow_MouseHover(ByVal sender As System.Object, ByVal e As EventArgs) Handles tabAddFlow.MouseHover, tabAddFlow.MouseClick   'System.Windows.Forms.Mouse

        Try
            'Dim itm As Flow.Item = tabAddFlow.PointedItem
            tabAddFlow.SelectedItem = tabAddFlow.PointedItem
            tabAddFlow.Refresh()
            'cms1.AutoClose = True
            'cms1.Close()

            If tabAddFlow.PointedItem IsNot Nothing Then
                If tabAddFlow.PointedItem.Tag IsNot Nothing Then
                    '/// It's a Node
                    Dim NodeTag As INode = CType(tabAddFlow.PointedItem.Tag, INode)
                    Dim tv As TreeView = NodeTag.ObjTreeNode.TreeView
                    tv.SelectedNode = NodeTag.ObjTreeNode
                    txtSelected.Text = NodeTag.Text
                    mnuDelItem.Text = "Delete Node"
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
                mnuDelItem.Text = "Delete Engine"
            End If

            cms1.Refresh()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_MouseHover")
        End Try

    End Sub

    Private Sub tabAddFlow_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tabAddFlow.MouseDown ', tabAddFlow.MouseHover
        'tabAddFlow.MouseDown, 
        Try
            'Dim itm As Flow.Item = tabAddFlow.PointedItem
            Dim ItemSelected As Boolean

            cms1.Close()


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

            'tabAddFlow.Refresh()


            If e.Button = Windows.Forms.MouseButtons.Right Then
                mnuAddMain.Visible = Not ItemSelected
                mnuAddProcGen.Visible = Not ItemSelected
                mnuAddProcMap.Visible = Not ItemSelected
                mnuAddLU.Visible = Not ItemSelected
                mnuAddSrc.Visible = Not ItemSelected
                mnuAddTgt.Visible = Not ItemSelected
                mnuRedo.Enabled = True
                mnuUndo.Enabled = True
                mnuDelItem.Enabled = True
                mnuCloseTab.Visible = False
            End If

            cms1.Refresh()

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab tabAddFlow_MouseDown")
        End Try

    End Sub

#End Region

#Region "Toolbar Actions"

    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click

        Try

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab OpenToolStripButton_Click")
        End Try

    End Sub

    Private Sub SaveToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripButton.Click

        Try

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

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab UpdateDiag_Click")
        End Try

    End Sub

    Private Sub UpdateTree_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UpdateTree.Click

        Try

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
            For Each main As clsTask In ObjEng.Mains
                For Each sDS As clsDatastore In main.ObjSources
                    Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                    sDS.AFnode.OutLinks.Add(link1, main.AFnode)
                Next
            Next

            For Each Task As clsTask In ObjEng.Procs
                For Each src As clsDatastore In Task.ObjSources
                    For Each main As clsTask In ObjEng.Mains
                        If main.ObjSources.Contains(src) Then
                            Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                            main.AFnode.OutLinks.Add(link1, Task.AFnode)
                        End If
                    Next
                Next
                For Each Tgt As clsDatastore In Task.ObjTargets
                    Dim link1 As Link = New Link(tabAddFlow.DefLinkProp)
                    Task.AFnode.OutLinks.Add(link1, Tgt.AFnode)
                Next
            Next

        Catch ex As Exception
            LogError(ex, "ctlAddFlowTab CreateAFLinks")
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

    
End Class
