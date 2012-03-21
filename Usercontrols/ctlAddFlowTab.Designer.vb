<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlAddFlowTab
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctlAddFlowTab))
        Me.tabAddFlow = New Lassalle.Flow.AddFlow
        Me.cms1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuAddSrc = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddTgt = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddLU = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddProcMap = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddProcGen = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuAddMain = New System.Windows.Forms.ToolStripMenuItem
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.mnuDelItem = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuUndo = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuRedo = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuCloseTab = New System.Windows.Forms.ToolStripMenuItem
        Me.PrnFlow1 = New Lassalle.PrnFlow.PrnFlow
        Me.afPalette = New Lassalle.Flow.AddFlow
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.toolStripSeparator = New System.Windows.Forms.ToolStripSeparator
        Me.UpdateDiag = New System.Windows.Forms.ToolStripButton
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator
        Me.UpdateTree = New System.Windows.Forms.ToolStripButton
        Me.toolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel
        Me.txtSelected = New System.Windows.Forms.ToolStripTextBox
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator
        Me.ZoomOut = New System.Windows.Forms.ToolStripButton
        Me.ZoomIn = New System.Windows.Forms.ToolStripButton
        Me.ZoomNorm = New System.Windows.Forms.ToolStripButton
        Me.cms1.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabAddFlow
        '
        Me.tabAddFlow.AllowDrop = True
        Me.tabAddFlow.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabAddFlow.AntiAliasing = False
        Me.tabAddFlow.AutoScroll = True
        Me.tabAddFlow.AutoScrollMinSize = New System.Drawing.Size(594, 455)
        Me.tabAddFlow.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.tabAddFlow.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.tabAddFlow.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.tabAddFlow.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tabAddFlow.CanDrawNode = False
        Me.tabAddFlow.CanReflexLink = False
        Me.tabAddFlow.CanSizeNode = False
        Me.tabAddFlow.ContextMenuStrip = Me.cms1
        Me.tabAddFlow.CursorSetting = Lassalle.Flow.CursorSetting.ResizeAndDrag
        Me.tabAddFlow.CycleMode = Lassalle.Flow.CycleMode.NoCycle
        Me.tabAddFlow.DefLinkProp.AdjustDst = True
        Me.tabAddFlow.DefLinkProp.ArrowDst = New Lassalle.Flow.Arrow(Lassalle.Flow.ArrowStyle.Arrow, Lassalle.Flow.ArrowSize.Large, Lassalle.Flow.ArrowAngle.deg15, False)
        Me.tabAddFlow.DefLinkProp.ConnectionStyleDst = Lassalle.Flow.ConnectionStyle.Inside
        Me.tabAddFlow.DefLinkProp.ConnectionStyleOrg = Lassalle.Flow.ConnectionStyle.Inside
        Me.tabAddFlow.DefLinkProp.CustomEndCap = Nothing
        Me.tabAddFlow.DefLinkProp.CustomStartCap = Nothing
        Me.tabAddFlow.DefLinkProp.DrawColor = System.Drawing.Color.DarkBlue
        Me.tabAddFlow.DefLinkProp.DrawWidth = 2
        Me.tabAddFlow.DefLinkProp.EndCap = System.Drawing.Drawing2D.LineCap.ArrowAnchor
        Me.tabAddFlow.DefLinkProp.Jump = Lassalle.Flow.Jump.Arc
        Me.tabAddFlow.DefLinkProp.Line = New Lassalle.Flow.Line(Lassalle.Flow.LineStyle.Spline, True, True, False)
        Me.tabAddFlow.DefLinkProp.StartCap = System.Drawing.Drawing2D.LineCap.Round
        Me.tabAddFlow.DefLinkProp.Tag = Nothing
        Me.tabAddFlow.DefLinkProp.ZOrder = -1
        Me.tabAddFlow.DefNodeProp.Alignment = Lassalle.Flow.Alignment.CenterTOP
        Me.tabAddFlow.DefNodeProp.BackMode = Lassalle.Flow.BackMode.Opaque
        Me.tabAddFlow.DefNodeProp.DrawColor = System.Drawing.SystemColors.InactiveCaption
        Me.tabAddFlow.DefNodeProp.DrawWidth = 5
        Me.tabAddFlow.DefNodeProp.FillColor = System.Drawing.SystemColors.MenuHighlight
        Me.tabAddFlow.DefNodeProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tabAddFlow.DefNodeProp.Gradient = True
        Me.tabAddFlow.DefNodeProp.GradientColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.tabAddFlow.DefNodeProp.GradientMode = System.Drawing.Drawing2D.LinearGradientMode.Vertical
        Me.tabAddFlow.DefNodeProp.ImageIndex = 13
        Me.tabAddFlow.DefNodeProp.LabelEdit = False
        Me.tabAddFlow.DefNodeProp.Location = CType(resources.GetObject("resource.Location"), System.Drawing.PointF)
        Me.tabAddFlow.DefNodeProp.Rect = CType(resources.GetObject("resource.Rect"), System.Drawing.RectangleF)
        Me.tabAddFlow.DefNodeProp.Shape = New Lassalle.Flow.Shape(Lassalle.Flow.ShapeStyle.Octogon, Lassalle.Flow.ShapeOrientation.so_0)
        Me.tabAddFlow.DefNodeProp.Tag = Nothing
        Me.tabAddFlow.DefNodeProp.TextColor = System.Drawing.SystemColors.Info
        Me.tabAddFlow.DefNodeProp.ZOrder = -1
        Me.tabAddFlow.ForeColor = System.Drawing.SystemColors.ControlText
        Me.tabAddFlow.LinkHandleSize = Lassalle.Flow.HandleSize.Medium
        Me.tabAddFlow.Location = New System.Drawing.Point(3, 28)
        Me.tabAddFlow.MultiSel = False
        Me.tabAddFlow.Name = "tabAddFlow"
        Me.tabAddFlow.SelectionHandleSize = Lassalle.Flow.HandleSize.Medium
        Me.tabAddFlow.Size = New System.Drawing.Size(598, 459)
        Me.tabAddFlow.TabIndex = 3
        Me.tabAddFlow.UndoSize = 50
        '
        'cms1
        '
        Me.cms1.AllowMerge = False
        Me.cms1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuAddSrc, Me.mnuAddTgt, Me.mnuAddLU, Me.mnuAddProcMap, Me.mnuAddProcGen, Me.mnuAddMain, Me.ToolStripSeparator1, Me.mnuDelItem, Me.mnuUndo, Me.mnuRedo, Me.mnuCloseTab})
        Me.cms1.Name = "cms1"
        Me.cms1.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional
        Me.cms1.ShowImageMargin = False
        Me.cms1.Size = New System.Drawing.Size(175, 230)
        '
        'mnuAddSrc
        '
        Me.mnuAddSrc.Name = "mnuAddSrc"
        Me.mnuAddSrc.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddSrc.Text = "Add Source"
        '
        'mnuAddTgt
        '
        Me.mnuAddTgt.Name = "mnuAddTgt"
        Me.mnuAddTgt.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddTgt.Text = "Add Target"
        '
        'mnuAddLU
        '
        Me.mnuAddLU.Name = "mnuAddLU"
        Me.mnuAddLU.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddLU.Text = "Add Lookup"
        '
        'mnuAddProcMap
        '
        Me.mnuAddProcMap.Name = "mnuAddProcMap"
        Me.mnuAddProcMap.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddProcMap.Text = "Add Mapping Procedure"
        '
        'mnuAddProcGen
        '
        Me.mnuAddProcGen.Name = "mnuAddProcGen"
        Me.mnuAddProcGen.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddProcGen.Text = "Add Logic Procedure"
        '
        'mnuAddMain
        '
        Me.mnuAddMain.Name = "mnuAddMain"
        Me.mnuAddMain.Size = New System.Drawing.Size(174, 22)
        Me.mnuAddMain.Text = "Add Main"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(171, 6)
        '
        'mnuDelItem
        '
        Me.mnuDelItem.Name = "mnuDelItem"
        Me.mnuDelItem.Size = New System.Drawing.Size(174, 22)
        Me.mnuDelItem.Text = "Delete Item"
        '
        'mnuUndo
        '
        Me.mnuUndo.Name = "mnuUndo"
        Me.mnuUndo.Size = New System.Drawing.Size(174, 22)
        Me.mnuUndo.Text = "Undo"
        '
        'mnuRedo
        '
        Me.mnuRedo.Name = "mnuRedo"
        Me.mnuRedo.Size = New System.Drawing.Size(174, 22)
        Me.mnuRedo.Text = "Redo"
        '
        'mnuCloseTab
        '
        Me.mnuCloseTab.Name = "mnuCloseTab"
        Me.mnuCloseTab.Size = New System.Drawing.Size(174, 22)
        Me.mnuCloseTab.Text = "Close Tab"
        '
        'PrnFlow1
        '
        Me.PrnFlow1.FitToPage = True
        '
        'afPalette
        '
        Me.afPalette.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.afPalette.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.afPalette.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.afPalette.CanDragScroll = False
        Me.afPalette.CanDrawLink = False
        Me.afPalette.CanDrawNode = False
        Me.afPalette.CanLabelEdit = False
        Me.afPalette.CanMoveNode = False
        Me.afPalette.CanMultiLink = False
        Me.afPalette.CanReflexLink = False
        Me.afPalette.CanShowJumps = False
        Me.afPalette.CanSizeNode = False
        Me.afPalette.CanStretchLink = False
        Me.afPalette.DefLinkProp.ArrowDst = New Lassalle.Flow.Arrow(Lassalle.Flow.ArrowStyle.Arrow, Lassalle.Flow.ArrowSize.Small, Lassalle.Flow.ArrowAngle.deg15, False)
        Me.afPalette.DefLinkProp.CustomEndCap = Nothing
        Me.afPalette.DefLinkProp.CustomStartCap = Nothing
        Me.afPalette.DefLinkProp.OwnerDraw = True
        Me.afPalette.DefLinkProp.Tag = Nothing
        Me.afPalette.DefLinkProp.ZOrder = -1
        Me.afPalette.Location = New System.Drawing.Point(3, 28)
        Me.afPalette.MultiSel = False
        Me.afPalette.Name = "afPalette"
        Me.afPalette.Size = New System.Drawing.Size(78, 459)
        Me.afPalette.TabIndex = 4
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 4)
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 4)
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.ToolStripSeparator4, Me.ZoomIn, Me.ZoomOut, Me.ZoomNorm, Me.toolStripSeparator, Me.UpdateDiag, Me.ToolStripSeparator3, Me.UpdateTree, Me.toolStripSeparator2, Me.ToolStripLabel1, Me.txtSelected})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(704, 25)
        Me.ToolStrip1.TabIndex = 5
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "&New"
        Me.NewToolStripButton.Visible = False
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.OpenToolStripButton.Text = "&Open"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "&Save"
        Me.SaveToolStripButton.ToolTipText = "Save Diagram as XML"
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "&Print"
        '
        'toolStripSeparator
        '
        Me.toolStripSeparator.Name = "toolStripSeparator"
        Me.toolStripSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'UpdateDiag
        '
        Me.UpdateDiag.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.UpdateDiag.Image = CType(resources.GetObject("UpdateDiag.Image"), System.Drawing.Image)
        Me.UpdateDiag.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.UpdateDiag.Name = "UpdateDiag"
        Me.UpdateDiag.Size = New System.Drawing.Size(138, 22)
        Me.UpdateDiag.Text = "Update Diagram from Tree"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'UpdateTree
        '
        Me.UpdateTree.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.UpdateTree.Image = CType(resources.GetObject("UpdateTree.Image"), System.Drawing.Image)
        Me.UpdateTree.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.UpdateTree.Name = "UpdateTree"
        Me.UpdateTree.Size = New System.Drawing.Size(138, 22)
        Me.UpdateTree.Text = "Update Tree from Diagram"
        '
        'toolStripSeparator2
        '
        Me.toolStripSeparator2.Name = "toolStripSeparator2"
        Me.toolStripSeparator2.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ToolStripLabel1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(90, 22)
        Me.ToolStripLabel1.Text = "Selected Item:"
        '
        'txtSelected
        '
        Me.txtSelected.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.txtSelected.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelected.Name = "txtSelected"
        Me.txtSelected.ReadOnly = True
        Me.txtSelected.Size = New System.Drawing.Size(100, 25)
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'ZoomOut
        '
        Me.ZoomOut.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ZoomOut.Image = CType(resources.GetObject("ZoomOut.Image"), System.Drawing.Image)
        Me.ZoomOut.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ZoomOut.Name = "ZoomOut"
        Me.ZoomOut.Size = New System.Drawing.Size(23, 22)
        Me.ZoomOut.Text = "-"
        Me.ZoomOut.ToolTipText = "Zoom Out"
        '
        'ZoomIn
        '
        Me.ZoomIn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ZoomIn.Image = CType(resources.GetObject("ZoomIn.Image"), System.Drawing.Image)
        Me.ZoomIn.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ZoomIn.Name = "ZoomIn"
        Me.ZoomIn.Size = New System.Drawing.Size(23, 22)
        Me.ZoomIn.Text = "+"
        Me.ZoomIn.ToolTipText = "Zoom In"
        '
        'ZoomNorm
        '
        Me.ZoomNorm.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.ZoomNorm.Image = CType(resources.GetObject("ZoomNorm.Image"), System.Drawing.Image)
        Me.ZoomNorm.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.ZoomNorm.Name = "ZoomNorm"
        Me.ZoomNorm.Size = New System.Drawing.Size(56, 22)
        Me.ZoomNorm.Text = "Zoom 1:1"
        '
        'ctlAddFlowTab
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Controls.Add(Me.ToolStrip1)
        Me.Controls.Add(Me.tabAddFlow)
        Me.Controls.Add(Me.afPalette)
        Me.Name = "ctlAddFlowTab"
        Me.Size = New System.Drawing.Size(704, 490)
        Me.cms1.ResumeLayout(False)
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tabAddFlow As Lassalle.Flow.AddFlow
    Friend WithEvents cms1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuAddSrc As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddTgt As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddLU As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddProcMap As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddProcGen As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuAddMain As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuDelItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuUndo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuRedo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuCloseTab As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrnFlow1 As Lassalle.PrnFlow.PrnFlow
    Friend WithEvents afPalette As Lassalle.Flow.AddFlow
    Friend WithEvents ToolStripButton1 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripButton2 As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents NewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents OpenToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents PrintToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents UpdateDiag As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ToolStripLabel1 As System.Windows.Forms.ToolStripLabel
    Friend WithEvents txtSelected As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents UpdateTree As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator4 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents ZoomIn As System.Windows.Forms.ToolStripButton
    Friend WithEvents ZoomOut As System.Windows.Forms.ToolStripButton
    Friend WithEvents ZoomNorm As System.Windows.Forms.ToolStripButton

End Class
