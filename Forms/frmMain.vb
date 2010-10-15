Public Class frmMain
    Inherits System.Windows.Forms.Form

    Dim prevNode As TreeNode
    Dim IsEventFromCode As Boolean
    Dim IsFocusOnExplorer As Boolean
    Dim CurLoadedProject As clsProject = Nothing
    Dim IsRename As Boolean


    '//This collection holds Inode objects copied into clipboard
    Dim m_ClipObjects As New ArrayList

    Enum enumToolBarButtons

        TB_PROJECT = 0
        TB_S1 = TB_PROJECT + 1
        TB_OPEN = TB_S1 + 1
        TB_SAVE = TB_OPEN + 1
        TB_S2 = TB_SAVE + 1
        TB_CUT = TB_S2 + 1
        TB_COPY = TB_CUT + 1
        TB_PASTE = TB_COPY + 1
        TB_DEL = TB_PASTE + 1
        TB_S3 = TB_DEL + 1
        TB_SCRIPT = TB_S3 + 1
        TB_S4 = TB_SCRIPT + 1
        TB_ENV = TB_S4 + 1
        TB_STRUCT = TB_ENV + 1
        TB_STRUCTSEL = TB_STRUCT + 1
        TB_SYS = TB_STRUCTSEL + 1
        TB_CONN = TB_SYS + 1
        TB_ENGINE = TB_CONN + 1
        TB_DATASTORE = TB_ENGINE + 1
        TB_VAR = TB_DATASTORE + 1
        TB_PROC = TB_VAR + 1
        TB_JOIN = TB_PROC + 1
        TB_LOOKUP = TB_JOIN + 1
        TB_MAIN = TB_LOOKUP + 1
        TB_HELP = TB_MAIN + 1
        TB_LOG = TB_HELP + 1
        TB_S5 = TB_LOG + 1
        TB_QUERY = TB_S5 + 1
        TB_S6 = TB_QUERY + 1
        TB_MAP = TB_S6 + 1
        TB_S7 = TB_MAP + 1
        TB_TOGGLE = TB_S7 + 1

    End Enum

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        InitUI()
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents ImageListSmall As System.Windows.Forms.ImageList
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents mnuMain As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem45 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupEnv As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupSystem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupStruct As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupConnection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupEngine As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupVariable As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupTask As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddEnv As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddSys As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStruct As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructXML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructCStruct As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddConn As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddEngine As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelEnv As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelSystem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelStruct As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelConn As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructDDL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructDML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructCOBOL_IMS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupStructSelection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructSelection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelStructSelection As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelEngine As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopup As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPopupDS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSBinary As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSDelimited As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSXML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSIMS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSRelational As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSIMSCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSDB2CDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSVSAMCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelDS As System.Windows.Forms.MenuItem
    'Trigger based cdc
    Friend WithEvents mnuAddDSDB2LOAD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSHSSUNLOAD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddDSVSAM As System.Windows.Forms.MenuItem
    Friend WithEvents tvExplorer As System.Windows.Forms.TreeView
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents pnlProp As System.Windows.Forms.Panel
    Friend WithEvents ctPrj As SQDStudio.ctlProject
    Friend WithEvents ctEnv As SQDStudio.ctlEnvironment
    Friend WithEvents ctStr As SQDStudio.ctlStructure
    Friend WithEvents ctDs As SQDStudio.ctlDatastore
    Friend WithEvents ctConn As SQDStudio.ctlConnection
    Friend WithEvents ctSys As SQDStudio.ctlSystem
    Friend WithEvents ctStrSel As SQDStudio.ctlStructureSelection
    Friend WithEvents ctEng As SQDStudio.ctlEngine
    Friend WithEvents ctInc As SQDStudio.ctlInclude
    Friend WithEvents ctVar As SQDStudio.ctlVariable
    Friend WithEvents ctTask As SQDStudio.ctlTask
    Friend WithEvents ctMain As SQDStudio.ctlMain
    Friend WithEvents ctFolder As SQDStudio.ctlFolderNode
    Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
    Friend WithEvents tbtnNew As System.Windows.Forms.ToolBarButton
    Friend WithEvents s1 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnOpen As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnSave As System.Windows.Forms.ToolBarButton
    Friend WithEvents s2 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnCut As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnCopy As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnPaste As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnDel As System.Windows.Forms.ToolBarButton
    Friend WithEvents s3 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnScript As System.Windows.Forms.ToolBarButton
    Friend WithEvents s4 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnEnv As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnStructure As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnStructSel As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnSystem As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnConnection As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnEngine As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnDS As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnVar As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnProc As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnJoin As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnLookup As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnMain As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuDelVar As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelTask As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddTask As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddVar As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructSubset As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructDDL As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructDTD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructLOD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructSQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructMSSQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructDB2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructSelectionDDL As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructSelectionDTD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScriptEngine As System.Windows.Forms.MenuItem
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents mnuAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuNewProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuOpenProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseProject As System.Windows.Forms.MenuItem
    Friend WithEvents mnuSave As System.Windows.Forms.MenuItem
    Friend WithEvents mnuExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopy As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPaste As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelete As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCloseAllProjects As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyEnv As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteEnv As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopySystem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteSystem As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyStruct As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteStruct As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyConn As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteConn As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyEngine As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteEngine As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyDS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteDS As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem39 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyVar As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteVar As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem42 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyTask As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteTask As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuCopyStructSel As System.Windows.Forms.MenuItem
    Friend WithEvents mnuPasteStructSel As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructHeader As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelStructSelectionHeader As System.Windows.Forms.MenuItem
    Friend WithEvents tbtHelp As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuHelpIdx As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents tbLog As System.Windows.Forms.ToolBarButton
    Friend WithEvents lblStatusMsg As System.Windows.Forms.Label
    Friend WithEvents mnuAddDSIBMEvent As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsBinary As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsText As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsDelimited As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsXML As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsRelational As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsVSAM As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsIMS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsDB2LOAD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsHSSUNLOAD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsIBMEvent As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsIMSCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsDB2CDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsVSAMCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsXMLCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsTriggerCDC As System.Windows.Forms.MenuItem
    Friend WithEvents s5 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbtnQuery As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuQuery As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChgStruct As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuChangeTask As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents mnuDelAllDS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelAllTask As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDelAllStruct As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDMLChoose As System.Windows.Forms.MenuItem
    Friend WithEvents mnuDMLUser As System.Windows.Forms.MenuItem
    Friend WithEvents s6 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbMap As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuDMLexisting As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelSSLOD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelSSSQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelSSMSSQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModelSSDB2 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddStructCOBOL As System.Windows.Forms.MenuItem
    Friend WithEvents SCmain As System.Windows.Forms.SplitContainer
    Friend WithEvents s7 As System.Windows.Forms.ToolBarButton
    Friend WithEvents tbToggle As System.Windows.Forms.ToolBarButton
    Friend WithEvents mnuAddOraCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddSQDCDC As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddIMSLE As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsOraCDC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsSQDCDC As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsIMSLE As System.Windows.Forms.MenuItem
    Friend WithEvents mnuMapAsIMSLEBat As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEngMerge As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuIncludeDS As System.Windows.Forms.MenuItem
    Friend WithEvents mnuIncProc As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScriptDS As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuScriptProc As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuGenRepStr As System.Windows.Forms.MenuItem
    Friend WithEvents mnuLU As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents mnuAddGen As System.Windows.Forms.MenuItem
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents mnuModDSSelDTD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModDSSelDDL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModDSSelC As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModDSSelLOD As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModDSSelSQL As System.Windows.Forms.MenuItem
    Friend WithEvents mnuModDSSelMSSQL As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.lblStatusMsg = New System.Windows.Forms.Label
        Me.ToolBar1 = New System.Windows.Forms.ToolBar
        Me.tbtnNew = New System.Windows.Forms.ToolBarButton
        Me.s1 = New System.Windows.Forms.ToolBarButton
        Me.tbtnOpen = New System.Windows.Forms.ToolBarButton
        Me.tbtnSave = New System.Windows.Forms.ToolBarButton
        Me.s2 = New System.Windows.Forms.ToolBarButton
        Me.tbtnCut = New System.Windows.Forms.ToolBarButton
        Me.tbtnCopy = New System.Windows.Forms.ToolBarButton
        Me.tbtnPaste = New System.Windows.Forms.ToolBarButton
        Me.tbtnDel = New System.Windows.Forms.ToolBarButton
        Me.s3 = New System.Windows.Forms.ToolBarButton
        Me.tbtnScript = New System.Windows.Forms.ToolBarButton
        Me.s4 = New System.Windows.Forms.ToolBarButton
        Me.tbtnEnv = New System.Windows.Forms.ToolBarButton
        Me.tbtnStructure = New System.Windows.Forms.ToolBarButton
        Me.tbtnStructSel = New System.Windows.Forms.ToolBarButton
        Me.tbtnSystem = New System.Windows.Forms.ToolBarButton
        Me.tbtnConnection = New System.Windows.Forms.ToolBarButton
        Me.tbtnEngine = New System.Windows.Forms.ToolBarButton
        Me.tbtnDS = New System.Windows.Forms.ToolBarButton
        Me.tbtnVar = New System.Windows.Forms.ToolBarButton
        Me.tbtnProc = New System.Windows.Forms.ToolBarButton
        Me.tbtnJoin = New System.Windows.Forms.ToolBarButton
        Me.tbtnLookup = New System.Windows.Forms.ToolBarButton
        Me.tbtnMain = New System.Windows.Forms.ToolBarButton
        Me.tbtHelp = New System.Windows.Forms.ToolBarButton
        Me.tbLog = New System.Windows.Forms.ToolBarButton
        Me.s5 = New System.Windows.Forms.ToolBarButton
        Me.tbtnQuery = New System.Windows.Forms.ToolBarButton
        Me.s6 = New System.Windows.Forms.ToolBarButton
        Me.tbMap = New System.Windows.Forms.ToolBarButton
        Me.s7 = New System.Windows.Forms.ToolBarButton
        Me.tbToggle = New System.Windows.Forms.ToolBarButton
        Me.ImageListSmall = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.tvExplorer = New System.Windows.Forms.TreeView
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.mnuMain = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuNewProject = New System.Windows.Forms.MenuItem
        Me.mnuOpenProject = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.mnuCloseProject = New System.Windows.Forms.MenuItem
        Me.mnuCloseAllProjects = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.mnuSave = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.mnuExit = New System.Windows.Forms.MenuItem
        Me.mnuEdit = New System.Windows.Forms.MenuItem
        Me.mnuCopy = New System.Windows.Forms.MenuItem
        Me.mnuPaste = New System.Windows.Forms.MenuItem
        Me.mnuDelete = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuHelpIdx = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.mnuAbout = New System.Windows.Forms.MenuItem
        Me.mnuPopup = New System.Windows.Forms.MenuItem
        Me.mnuPopupEnv = New System.Windows.Forms.MenuItem
        Me.mnuAddEnv = New System.Windows.Forms.MenuItem
        Me.mnuDelEnv = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.mnuCopyEnv = New System.Windows.Forms.MenuItem
        Me.mnuPasteEnv = New System.Windows.Forms.MenuItem
        Me.mnuPopupSystem = New System.Windows.Forms.MenuItem
        Me.mnuAddSys = New System.Windows.Forms.MenuItem
        Me.mnuDelSystem = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.mnuCopySystem = New System.Windows.Forms.MenuItem
        Me.mnuPasteSystem = New System.Windows.Forms.MenuItem
        Me.mnuPopupStruct = New System.Windows.Forms.MenuItem
        Me.mnuAddStruct = New System.Windows.Forms.MenuItem
        Me.mnuAddStructCOBOL = New System.Windows.Forms.MenuItem
        Me.mnuAddStructCOBOL_IMS = New System.Windows.Forms.MenuItem
        Me.mnuAddStructXML = New System.Windows.Forms.MenuItem
        Me.mnuAddStructCStruct = New System.Windows.Forms.MenuItem
        Me.mnuAddStructDDL = New System.Windows.Forms.MenuItem
        Me.mnuAddStructDML = New System.Windows.Forms.MenuItem
        Me.mnuDMLChoose = New System.Windows.Forms.MenuItem
        Me.mnuDMLUser = New System.Windows.Forms.MenuItem
        Me.mnuDMLexisting = New System.Windows.Forms.MenuItem
        Me.mnuAddStructSubset = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.mnuModelStructDTD = New System.Windows.Forms.MenuItem
        Me.mnuModelStructDDL = New System.Windows.Forms.MenuItem
        Me.mnuModelStructHeader = New System.Windows.Forms.MenuItem
        Me.mnuModelStructLOD = New System.Windows.Forms.MenuItem
        Me.mnuModelStructSQL = New System.Windows.Forms.MenuItem
        Me.mnuModelStructMSSQL = New System.Windows.Forms.MenuItem
        Me.mnuModelStructDB2 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.mnuDelStruct = New System.Windows.Forms.MenuItem
        Me.mnuChgStruct = New System.Windows.Forms.MenuItem
        Me.mnuDelAllStruct = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.mnuCopyStruct = New System.Windows.Forms.MenuItem
        Me.mnuPasteStruct = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.mnuGenRepStr = New System.Windows.Forms.MenuItem
        Me.mnuPopupConnection = New System.Windows.Forms.MenuItem
        Me.mnuAddConn = New System.Windows.Forms.MenuItem
        Me.mnuDelConn = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.mnuCopyConn = New System.Windows.Forms.MenuItem
        Me.mnuPasteConn = New System.Windows.Forms.MenuItem
        Me.mnuPopupEngine = New System.Windows.Forms.MenuItem
        Me.mnuAddEngine = New System.Windows.Forms.MenuItem
        Me.mnuDelEngine = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.mnuScriptEngine = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.mnuCopyEngine = New System.Windows.Forms.MenuItem
        Me.mnuPasteEngine = New System.Windows.Forms.MenuItem
        Me.mnuEngMerge = New System.Windows.Forms.MenuItem
        Me.mnuPopupDS = New System.Windows.Forms.MenuItem
        Me.mnuAddDS = New System.Windows.Forms.MenuItem
        Me.mnuAddDSBinary = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.mnuAddDSDelimited = New System.Windows.Forms.MenuItem
        Me.mnuAddDSXML = New System.Windows.Forms.MenuItem
        Me.mnuAddDSRelational = New System.Windows.Forms.MenuItem
        Me.mnuAddDSDB2LOAD = New System.Windows.Forms.MenuItem
        Me.mnuAddDSHSSUNLOAD = New System.Windows.Forms.MenuItem
        Me.mnuAddDSIMS = New System.Windows.Forms.MenuItem
        Me.mnuAddDSVSAM = New System.Windows.Forms.MenuItem
        Me.MenuItem45 = New System.Windows.Forms.MenuItem
        Me.mnuAddDSIMSCDC = New System.Windows.Forms.MenuItem
        Me.mnuAddIMSLE = New System.Windows.Forms.MenuItem
        Me.mnuAddDSDB2CDC = New System.Windows.Forms.MenuItem
        Me.mnuAddDSVSAMCDC = New System.Windows.Forms.MenuItem
        Me.mnuAddOraCDC = New System.Windows.Forms.MenuItem
        Me.mnuAddSQDCDC = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.mnuAddDSIBMEvent = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.mnuIncludeDS = New System.Windows.Forms.MenuItem
        Me.mnuLU = New System.Windows.Forms.MenuItem
        Me.mnuDelDS = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.mnuCopyDS = New System.Windows.Forms.MenuItem
        Me.mnuPasteDS = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.mnuMapAsBinary = New System.Windows.Forms.MenuItem
        Me.mnuMapAsText = New System.Windows.Forms.MenuItem
        Me.mnuMapAsDelimited = New System.Windows.Forms.MenuItem
        Me.mnuMapAsXML = New System.Windows.Forms.MenuItem
        Me.mnuMapAsRelational = New System.Windows.Forms.MenuItem
        Me.mnuMapAsVSAM = New System.Windows.Forms.MenuItem
        Me.mnuMapAsIMS = New System.Windows.Forms.MenuItem
        Me.mnuMapAsDB2LOAD = New System.Windows.Forms.MenuItem
        Me.mnuMapAsHSSUNLOAD = New System.Windows.Forms.MenuItem
        Me.mnuMapAsIBMEvent = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.mnuMapAsIMSCDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsDB2CDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsVSAMCDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsXMLCDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsTriggerCDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsOraCDC = New System.Windows.Forms.MenuItem
        Me.mnuMapAsSQDCDC = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.mnuMapAsIMSLE = New System.Windows.Forms.MenuItem
        Me.mnuMapAsIMSLEBat = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelDTD = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelDDL = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelC = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelLOD = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelSQL = New System.Windows.Forms.MenuItem
        Me.mnuModDSSelMSSQL = New System.Windows.Forms.MenuItem
        Me.MenuItem31 = New System.Windows.Forms.MenuItem
        Me.mnuQuery = New System.Windows.Forms.MenuItem
        Me.mnuDelAllDS = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.mnuScriptDS = New System.Windows.Forms.MenuItem
        Me.mnuPopupVariable = New System.Windows.Forms.MenuItem
        Me.mnuAddVar = New System.Windows.Forms.MenuItem
        Me.mnuDelVar = New System.Windows.Forms.MenuItem
        Me.MenuItem39 = New System.Windows.Forms.MenuItem
        Me.mnuCopyVar = New System.Windows.Forms.MenuItem
        Me.mnuPasteVar = New System.Windows.Forms.MenuItem
        Me.mnuPopupTask = New System.Windows.Forms.MenuItem
        Me.mnuAddTask = New System.Windows.Forms.MenuItem
        Me.mnuAddGen = New System.Windows.Forms.MenuItem
        Me.mnuDelTask = New System.Windows.Forms.MenuItem
        Me.mnuIncProc = New System.Windows.Forms.MenuItem
        Me.MenuItem42 = New System.Windows.Forms.MenuItem
        Me.mnuCopyTask = New System.Windows.Forms.MenuItem
        Me.mnuPasteTask = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.mnuChangeTask = New System.Windows.Forms.MenuItem
        Me.mnuDelAllTask = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.mnuScriptProc = New System.Windows.Forms.MenuItem
        Me.mnuPopupProject = New System.Windows.Forms.MenuItem
        Me.mnuDelProject = New System.Windows.Forms.MenuItem
        Me.mnuPopupStructSelection = New System.Windows.Forms.MenuItem
        Me.mnuAddStructSelection = New System.Windows.Forms.MenuItem
        Me.mnuDelStructSelection = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.mnuModelStructSelectionDTD = New System.Windows.Forms.MenuItem
        Me.mnuModelStructSelectionDDL = New System.Windows.Forms.MenuItem
        Me.mnuModelStructSelectionHeader = New System.Windows.Forms.MenuItem
        Me.mnuModelSSLOD = New System.Windows.Forms.MenuItem
        Me.mnuModelSSSQL = New System.Windows.Forms.MenuItem
        Me.mnuModelSSMSSQL = New System.Windows.Forms.MenuItem
        Me.mnuModelSSDB2 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.mnuCopyStructSel = New System.Windows.Forms.MenuItem
        Me.mnuPasteStructSel = New System.Windows.Forms.MenuItem
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.pnlProp = New System.Windows.Forms.Panel
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SCmain = New System.Windows.Forms.SplitContainer
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.ctMain = New SQDStudio.ctlMain
        Me.ctInc = New SQDStudio.ctlInclude
        Me.ctFolder = New SQDStudio.ctlFolderNode
        Me.ctTask = New SQDStudio.ctlTask
        Me.ctVar = New SQDStudio.ctlVariable
        Me.ctEng = New SQDStudio.ctlEngine
        Me.ctStrSel = New SQDStudio.ctlStructureSelection
        Me.ctSys = New SQDStudio.ctlSystem
        Me.ctConn = New SQDStudio.ctlConnection
        Me.ctStr = New SQDStudio.ctlStructure
        Me.ctEnv = New SQDStudio.ctlEnvironment
        Me.ctPrj = New SQDStudio.ctlProject
        Me.ctDs = New SQDStudio.ctlDatastore
        Me.Panel1.SuspendLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlProp.SuspendLayout()
        Me.SCmain.Panel1.SuspendLayout()
        Me.SCmain.Panel2.SuspendLayout()
        Me.SCmain.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Control
        Me.Panel1.Controls.Add(Me.lblStatusMsg)
        Me.Panel1.Controls.Add(Me.ToolBar1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1016, 29)
        Me.Panel1.TabIndex = 2
        '
        'lblStatusMsg
        '
        Me.lblStatusMsg.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatusMsg.ForeColor = System.Drawing.Color.Blue
        Me.lblStatusMsg.Location = New System.Drawing.Point(321, 5)
        Me.lblStatusMsg.Name = "lblStatusMsg"
        Me.lblStatusMsg.Size = New System.Drawing.Size(625, 21)
        Me.lblStatusMsg.TabIndex = 7
        Me.lblStatusMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ToolBar1
        '
        Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.tbtnNew, Me.s1, Me.tbtnOpen, Me.tbtnSave, Me.s2, Me.tbtnCut, Me.tbtnCopy, Me.tbtnPaste, Me.tbtnDel, Me.s3, Me.tbtnScript, Me.s4, Me.tbtnEnv, Me.tbtnStructure, Me.tbtnStructSel, Me.tbtnSystem, Me.tbtnConnection, Me.tbtnEngine, Me.tbtnDS, Me.tbtnVar, Me.tbtnProc, Me.tbtnJoin, Me.tbtnLookup, Me.tbtnMain, Me.tbtHelp, Me.tbLog, Me.s5, Me.tbtnQuery, Me.s6, Me.tbMap, Me.s7, Me.tbToggle})
        Me.ToolBar1.Dock = System.Windows.Forms.DockStyle.None
        Me.ToolBar1.DropDownArrows = True
        Me.ToolBar1.ImageList = Me.ImageListSmall
        Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
        Me.ToolBar1.Name = "ToolBar1"
        Me.ToolBar1.ShowToolTips = True
        Me.ToolBar1.Size = New System.Drawing.Size(1016, 28)
        Me.ToolBar1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.ToolBar1, "Custom Field Mapping")
        '
        'tbtnNew
        '
        Me.tbtnNew.ImageIndex = 2
        Me.tbtnNew.Name = "tbtnNew"
        Me.tbtnNew.Tag = "New"
        Me.tbtnNew.ToolTipText = "Create new project"
        '
        's1
        '
        Me.s1.Name = "s1"
        Me.s1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbtnOpen
        '
        Me.tbtnOpen.ImageIndex = 24
        Me.tbtnOpen.Name = "tbtnOpen"
        Me.tbtnOpen.Tag = "Open"
        Me.tbtnOpen.ToolTipText = "Open existing project"
        '
        'tbtnSave
        '
        Me.tbtnSave.Enabled = False
        Me.tbtnSave.ImageIndex = 23
        Me.tbtnSave.Name = "tbtnSave"
        Me.tbtnSave.Tag = "Save"
        Me.tbtnSave.ToolTipText = "Save"
        '
        's2
        '
        Me.s2.Name = "s2"
        Me.s2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbtnCut
        '
        Me.tbtnCut.Enabled = False
        Me.tbtnCut.ImageIndex = 21
        Me.tbtnCut.Name = "tbtnCut"
        Me.tbtnCut.Tag = "Cut"
        Me.tbtnCut.ToolTipText = "Cut"
        Me.tbtnCut.Visible = False
        '
        'tbtnCopy
        '
        Me.tbtnCopy.Enabled = False
        Me.tbtnCopy.ImageIndex = 20
        Me.tbtnCopy.Name = "tbtnCopy"
        Me.tbtnCopy.Tag = "Copy"
        Me.tbtnCopy.ToolTipText = "Copy"
        '
        'tbtnPaste
        '
        Me.tbtnPaste.Enabled = False
        Me.tbtnPaste.ImageIndex = 19
        Me.tbtnPaste.Name = "tbtnPaste"
        Me.tbtnPaste.Tag = "Paste"
        Me.tbtnPaste.ToolTipText = "Paste"
        '
        'tbtnDel
        '
        Me.tbtnDel.Enabled = False
        Me.tbtnDel.ImageIndex = 22
        Me.tbtnDel.Name = "tbtnDel"
        Me.tbtnDel.Tag = "Delete"
        Me.tbtnDel.ToolTipText = "Delete"
        '
        's3
        '
        Me.s3.Name = "s3"
        Me.s3.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbtnScript
        '
        Me.tbtnScript.Enabled = False
        Me.tbtnScript.ImageIndex = 28
        Me.tbtnScript.Name = "tbtnScript"
        Me.tbtnScript.Tag = "Script"
        Me.tbtnScript.ToolTipText = "Generate Script"
        '
        's4
        '
        Me.s4.Name = "s4"
        Me.s4.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        Me.s4.Visible = False
        '
        'tbtnEnv
        '
        Me.tbtnEnv.Enabled = False
        Me.tbtnEnv.ImageIndex = 3
        Me.tbtnEnv.Name = "tbtnEnv"
        Me.tbtnEnv.ToolTipText = "Create new environment"
        Me.tbtnEnv.Visible = False
        '
        'tbtnStructure
        '
        Me.tbtnStructure.Enabled = False
        Me.tbtnStructure.ImageIndex = 4
        Me.tbtnStructure.Name = "tbtnStructure"
        Me.tbtnStructure.ToolTipText = "Create new structure"
        Me.tbtnStructure.Visible = False
        '
        'tbtnStructSel
        '
        Me.tbtnStructSel.Enabled = False
        Me.tbtnStructSel.ImageIndex = 29
        Me.tbtnStructSel.Name = "tbtnStructSel"
        Me.tbtnStructSel.ToolTipText = "Create new selection"
        Me.tbtnStructSel.Visible = False
        '
        'tbtnSystem
        '
        Me.tbtnSystem.Enabled = False
        Me.tbtnSystem.ImageIndex = 5
        Me.tbtnSystem.Name = "tbtnSystem"
        Me.tbtnSystem.ToolTipText = "Create new system"
        Me.tbtnSystem.Visible = False
        '
        'tbtnConnection
        '
        Me.tbtnConnection.Enabled = False
        Me.tbtnConnection.ImageIndex = 7
        Me.tbtnConnection.Name = "tbtnConnection"
        Me.tbtnConnection.ToolTipText = "Create new connection"
        Me.tbtnConnection.Visible = False
        '
        'tbtnEngine
        '
        Me.tbtnEngine.Enabled = False
        Me.tbtnEngine.ImageIndex = 6
        Me.tbtnEngine.Name = "tbtnEngine"
        Me.tbtnEngine.ToolTipText = "Create new engine"
        Me.tbtnEngine.Visible = False
        '
        'tbtnDS
        '
        Me.tbtnDS.Enabled = False
        Me.tbtnDS.ImageIndex = 16
        Me.tbtnDS.Name = "tbtnDS"
        Me.tbtnDS.ToolTipText = "Create new datastore"
        Me.tbtnDS.Visible = False
        '
        'tbtnVar
        '
        Me.tbtnVar.Enabled = False
        Me.tbtnVar.ImageIndex = 14
        Me.tbtnVar.Name = "tbtnVar"
        Me.tbtnVar.ToolTipText = "Create new variable"
        Me.tbtnVar.Visible = False
        '
        'tbtnProc
        '
        Me.tbtnProc.Enabled = False
        Me.tbtnProc.ImageIndex = 11
        Me.tbtnProc.Name = "tbtnProc"
        Me.tbtnProc.ToolTipText = "Create new procedure"
        Me.tbtnProc.Visible = False
        '
        'tbtnJoin
        '
        Me.tbtnJoin.Enabled = False
        Me.tbtnJoin.ImageIndex = 8
        Me.tbtnJoin.Name = "tbtnJoin"
        Me.tbtnJoin.ToolTipText = "Create new join"
        Me.tbtnJoin.Visible = False
        '
        'tbtnLookup
        '
        Me.tbtnLookup.Enabled = False
        Me.tbtnLookup.ImageIndex = 9
        Me.tbtnLookup.Name = "tbtnLookup"
        Me.tbtnLookup.ToolTipText = "Create new lookup"
        Me.tbtnLookup.Visible = False
        '
        'tbtnMain
        '
        Me.tbtnMain.Enabled = False
        Me.tbtnMain.ImageIndex = 10
        Me.tbtnMain.Name = "tbtnMain"
        Me.tbtnMain.ToolTipText = "Create new main procedure"
        Me.tbtnMain.Visible = False
        '
        'tbtHelp
        '
        Me.tbtHelp.ImageIndex = 37
        Me.tbtHelp.Name = "tbtHelp"
        Me.tbtHelp.Tag = "Help"
        Me.tbtHelp.ToolTipText = "Help"
        '
        'tbLog
        '
        Me.tbLog.ImageIndex = 38
        Me.tbLog.Name = "tbLog"
        Me.tbLog.Tag = "Log"
        Me.tbLog.ToolTipText = "Enable/View Log"
        '
        's5
        '
        Me.s5.Name = "s5"
        Me.s5.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbtnQuery
        '
        Me.tbtnQuery.Enabled = False
        Me.tbtnQuery.ImageIndex = 18
        Me.tbtnQuery.Name = "tbtnQuery"
        Me.tbtnQuery.Tag = "Query"
        Me.tbtnQuery.ToolTipText = "Datastore Query Tool"
        Me.tbtnQuery.Visible = False
        '
        's6
        '
        Me.s6.Name = "s6"
        Me.s6.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbMap
        '
        Me.tbMap.ImageIndex = 9
        Me.tbMap.Name = "tbMap"
        Me.tbMap.Tag = "Map"
        Me.tbMap.ToolTipText = "Custom Field Mapping"
        '
        's7
        '
        Me.s7.Name = "s7"
        Me.s7.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
        '
        'tbToggle
        '
        Me.tbToggle.ImageIndex = 3
        Me.tbToggle.Name = "tbToggle"
        Me.tbToggle.Pushed = True
        Me.tbToggle.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
        Me.tbToggle.Tag = "Toggle"
        Me.tbToggle.ToolTipText = "Show/Hide Main Tree"
        '
        'ImageListSmall
        '
        Me.ImageListSmall.ImageStream = CType(resources.GetObject("ImageListSmall.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageListSmall.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageListSmall.Images.SetKeyName(0, "")
        Me.ImageListSmall.Images.SetKeyName(1, "")
        Me.ImageListSmall.Images.SetKeyName(2, "")
        Me.ImageListSmall.Images.SetKeyName(3, "")
        Me.ImageListSmall.Images.SetKeyName(4, "")
        Me.ImageListSmall.Images.SetKeyName(5, "")
        Me.ImageListSmall.Images.SetKeyName(6, "")
        Me.ImageListSmall.Images.SetKeyName(7, "")
        Me.ImageListSmall.Images.SetKeyName(8, "")
        Me.ImageListSmall.Images.SetKeyName(9, "")
        Me.ImageListSmall.Images.SetKeyName(10, "")
        Me.ImageListSmall.Images.SetKeyName(11, "")
        Me.ImageListSmall.Images.SetKeyName(12, "")
        Me.ImageListSmall.Images.SetKeyName(13, "")
        Me.ImageListSmall.Images.SetKeyName(14, "")
        Me.ImageListSmall.Images.SetKeyName(15, "")
        Me.ImageListSmall.Images.SetKeyName(16, "")
        Me.ImageListSmall.Images.SetKeyName(17, "")
        Me.ImageListSmall.Images.SetKeyName(18, "")
        Me.ImageListSmall.Images.SetKeyName(19, "")
        Me.ImageListSmall.Images.SetKeyName(20, "")
        Me.ImageListSmall.Images.SetKeyName(21, "")
        Me.ImageListSmall.Images.SetKeyName(22, "")
        Me.ImageListSmall.Images.SetKeyName(23, "")
        Me.ImageListSmall.Images.SetKeyName(24, "")
        Me.ImageListSmall.Images.SetKeyName(25, "")
        Me.ImageListSmall.Images.SetKeyName(26, "")
        Me.ImageListSmall.Images.SetKeyName(27, "")
        Me.ImageListSmall.Images.SetKeyName(28, "")
        Me.ImageListSmall.Images.SetKeyName(29, "")
        Me.ImageListSmall.Images.SetKeyName(30, "")
        Me.ImageListSmall.Images.SetKeyName(31, "")
        Me.ImageListSmall.Images.SetKeyName(32, "")
        Me.ImageListSmall.Images.SetKeyName(33, "")
        Me.ImageListSmall.Images.SetKeyName(34, "")
        Me.ImageListSmall.Images.SetKeyName(35, "")
        Me.ImageListSmall.Images.SetKeyName(36, "")
        Me.ImageListSmall.Images.SetKeyName(37, "")
        Me.ImageListSmall.Images.SetKeyName(38, "")
        Me.ImageListSmall.Images.SetKeyName(39, "")
        Me.ImageListSmall.Images.SetKeyName(40, "")
        Me.ImageListSmall.Images.SetKeyName(41, "")
        Me.ImageListSmall.Images.SetKeyName(42, "")
        Me.ImageListSmall.Images.SetKeyName(43, "")
        Me.ImageListSmall.Images.SetKeyName(44, "")
        Me.ImageListSmall.Images.SetKeyName(45, "")
        Me.ImageListSmall.Images.SetKeyName(46, "")
        Me.ImageListSmall.Images.SetKeyName(47, "")
        Me.ImageListSmall.Images.SetKeyName(48, "")
        '
        'StatusBar1
        '
        Me.StatusBar1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusBar1.Location = New System.Drawing.Point(0, 542)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(1016, 24)
        Me.StatusBar1.TabIndex = 5
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.Icon = CType(resources.GetObject("StatusBarPanel1.Icon"), System.Drawing.Icon)
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 600
        '
        'tvExplorer
        '
        Me.tvExplorer.AllowDrop = True
        Me.tvExplorer.BackColor = System.Drawing.Color.White
        Me.tvExplorer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvExplorer.HideSelection = False
        Me.tvExplorer.HotTracking = True
        Me.tvExplorer.ImageIndex = 0
        Me.tvExplorer.ImageList = Me.ImageListSmall
        Me.tvExplorer.Indent = 19
        Me.tvExplorer.Location = New System.Drawing.Point(0, 0)
        Me.tvExplorer.Name = "tvExplorer"
        Me.tvExplorer.SelectedImageIndex = 7
        Me.tvExplorer.ShowNodeToolTips = True
        Me.tvExplorer.Size = New System.Drawing.Size(346, 513)
        Me.tvExplorer.TabIndex = 2
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "SQData Studio Project Files (*.sqp)|*.sqp|All files (*.*)|*.*"
        Me.OpenFileDialog1.FilterIndex = 0
        '
        'mnuMain
        '
        Me.mnuMain.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuHelp, Me.mnuPopup})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuNewProject, Me.mnuOpenProject, Me.MenuItem2, Me.mnuCloseProject, Me.mnuCloseAllProjects, Me.MenuItem7, Me.mnuSave, Me.MenuItem10, Me.mnuExit})
        Me.mnuFile.Shortcut = System.Windows.Forms.Shortcut.CtrlF
        Me.mnuFile.Text = "&File"
        '
        'mnuNewProject
        '
        Me.mnuNewProject.Index = 0
        Me.mnuNewProject.Shortcut = System.Windows.Forms.Shortcut.CtrlN
        Me.mnuNewProject.Text = "New Project"
        '
        'mnuOpenProject
        '
        Me.mnuOpenProject.Index = 1
        Me.mnuOpenProject.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.mnuOpenProject.Text = "Open Project"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 2
        Me.MenuItem2.Text = "-"
        '
        'mnuCloseProject
        '
        Me.mnuCloseProject.Index = 3
        Me.mnuCloseProject.Text = "Close Project"
        '
        'mnuCloseAllProjects
        '
        Me.mnuCloseAllProjects.Index = 4
        Me.mnuCloseAllProjects.Text = "Close All PROJECTS"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 5
        Me.MenuItem7.Text = "-"
        '
        'mnuSave
        '
        Me.mnuSave.Index = 6
        Me.mnuSave.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.mnuSave.Text = "Save"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 7
        Me.MenuItem10.Text = "-"
        '
        'mnuExit
        '
        Me.mnuExit.Index = 8
        Me.mnuExit.Text = "Exit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuCopy, Me.mnuPaste, Me.mnuDelete})
        Me.mnuEdit.Shortcut = System.Windows.Forms.Shortcut.CtrlE
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuCopy
        '
        Me.mnuCopy.Enabled = False
        Me.mnuCopy.Index = 0
        Me.mnuCopy.Shortcut = System.Windows.Forms.Shortcut.CtrlC
        Me.mnuCopy.Text = "Copy"
        '
        'mnuPaste
        '
        Me.mnuPaste.Enabled = False
        Me.mnuPaste.Index = 1
        Me.mnuPaste.Shortcut = System.Windows.Forms.Shortcut.CtrlV
        Me.mnuPaste.Text = "Paste"
        '
        'mnuDelete
        '
        Me.mnuDelete.Enabled = False
        Me.mnuDelete.Index = 2
        Me.mnuDelete.Shortcut = System.Windows.Forms.Shortcut.CtrlD
        Me.mnuDelete.Text = "Delete"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 2
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpIdx, Me.MenuItem5, Me.mnuAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpIdx
        '
        Me.mnuHelpIdx.Index = 0
        Me.mnuHelpIdx.Text = "&Index"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 1
        Me.MenuItem5.Text = "-"
        '
        'mnuAbout
        '
        Me.mnuAbout.Index = 2
        Me.mnuAbout.Text = ""
        '
        'mnuPopup
        '
        Me.mnuPopup.Index = 3
        Me.mnuPopup.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuPopupEnv, Me.mnuPopupSystem, Me.mnuPopupStruct, Me.mnuPopupConnection, Me.mnuPopupEngine, Me.mnuPopupDS, Me.mnuPopupVariable, Me.mnuPopupTask, Me.mnuPopupProject, Me.mnuPopupStructSelection})
        Me.mnuPopup.Text = "mnuPopup"
        Me.mnuPopup.Visible = False
        '
        'mnuPopupEnv
        '
        Me.mnuPopupEnv.Index = 0
        Me.mnuPopupEnv.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddEnv, Me.mnuDelEnv, Me.MenuItem1, Me.mnuCopyEnv, Me.mnuPasteEnv})
        Me.mnuPopupEnv.Text = "mnuPopupEnv"
        '
        'mnuAddEnv
        '
        Me.mnuAddEnv.Index = 0
        Me.mnuAddEnv.Text = "Add Environment"
        '
        'mnuDelEnv
        '
        Me.mnuDelEnv.Index = 1
        Me.mnuDelEnv.Text = "Delete Environment"
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 2
        Me.MenuItem1.Text = "-"
        '
        'mnuCopyEnv
        '
        Me.mnuCopyEnv.Index = 3
        Me.mnuCopyEnv.Text = "Copy"
        '
        'mnuPasteEnv
        '
        Me.mnuPasteEnv.Index = 4
        Me.mnuPasteEnv.Text = "Paste"
        '
        'mnuPopupSystem
        '
        Me.mnuPopupSystem.Index = 1
        Me.mnuPopupSystem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddSys, Me.mnuDelSystem, Me.MenuItem8, Me.mnuCopySystem, Me.mnuPasteSystem})
        Me.mnuPopupSystem.Text = "mnuPopupSystem"
        '
        'mnuAddSys
        '
        Me.mnuAddSys.Index = 0
        Me.mnuAddSys.Text = "Add System"
        '
        'mnuDelSystem
        '
        Me.mnuDelSystem.Index = 1
        Me.mnuDelSystem.Text = "Delete System"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.Text = "-"
        '
        'mnuCopySystem
        '
        Me.mnuCopySystem.Index = 3
        Me.mnuCopySystem.Text = "Copy"
        '
        'mnuPasteSystem
        '
        Me.mnuPasteSystem.Index = 4
        Me.mnuPasteSystem.Text = "Paste"
        '
        'mnuPopupStruct
        '
        Me.mnuPopupStruct.Index = 2
        Me.mnuPopupStruct.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddStruct, Me.mnuAddStructSubset, Me.MenuItem13, Me.mnuModelStructDTD, Me.mnuModelStructDDL, Me.mnuModelStructHeader, Me.mnuModelStructLOD, Me.mnuModelStructSQL, Me.mnuModelStructMSSQL, Me.mnuModelStructDB2, Me.MenuItem15, Me.mnuDelStruct, Me.mnuChgStruct, Me.mnuDelAllStruct, Me.MenuItem6, Me.mnuCopyStruct, Me.mnuPasteStruct, Me.MenuItem21, Me.mnuGenRepStr})
        Me.mnuPopupStruct.Text = "mnuPopupStruct"
        '
        'mnuAddStruct
        '
        Me.mnuAddStruct.Index = 0
        Me.mnuAddStruct.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddStructCOBOL, Me.mnuAddStructCOBOL_IMS, Me.mnuAddStructXML, Me.mnuAddStructCStruct, Me.mnuAddStructDDL, Me.mnuAddStructDML})
        Me.mnuAddStruct.Text = "Add Description"
        '
        'mnuAddStructCOBOL
        '
        Me.mnuAddStructCOBOL.Index = 0
        Me.mnuAddStructCOBOL.Text = "Cobol Copybook"
        '
        'mnuAddStructCOBOL_IMS
        '
        Me.mnuAddStructCOBOL_IMS.Index = 1
        Me.mnuAddStructCOBOL_IMS.Text = "Cobol Copybook/IMS DBD"
        '
        'mnuAddStructXML
        '
        Me.mnuAddStructXML.Index = 2
        Me.mnuAddStructXML.Text = "XML DTD"
        '
        'mnuAddStructCStruct
        '
        Me.mnuAddStructCStruct.Index = 3
        Me.mnuAddStructCStruct.Text = "C Structure"
        '
        'mnuAddStructDDL
        '
        Me.mnuAddStructDDL.Index = 4
        Me.mnuAddStructDDL.Text = "Relational DDL"
        '
        'mnuAddStructDML
        '
        Me.mnuAddStructDML.Index = 5
        Me.mnuAddStructDML.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDMLChoose, Me.mnuDMLUser, Me.mnuDMLexisting})
        Me.mnuAddStructDML.Text = "Relational DML"
        '
        'mnuDMLChoose
        '
        Me.mnuDMLChoose.Index = 0
        Me.mnuDMLChoose.Text = "Browse Database"
        '
        'mnuDMLUser
        '
        Me.mnuDMLUser.Index = 1
        Me.mnuDMLUser.Text = "User Defined SQL"
        '
        'mnuDMLexisting
        '
        Me.mnuDMLexisting.Index = 2
        Me.mnuDMLexisting.Text = "Existing DML File"
        '
        'mnuAddStructSubset
        '
        Me.mnuAddStructSubset.Index = 1
        Me.mnuAddStructSubset.Text = "Add Subset"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 2
        Me.MenuItem13.Text = "-"
        '
        'mnuModelStructDTD
        '
        Me.mnuModelStructDTD.Index = 3
        Me.mnuModelStructDTD.Text = "Model DTD"
        '
        'mnuModelStructDDL
        '
        Me.mnuModelStructDDL.Index = 4
        Me.mnuModelStructDDL.Text = "Model DDL"
        '
        'mnuModelStructHeader
        '
        Me.mnuModelStructHeader.Index = 5
        Me.mnuModelStructHeader.Text = "Model C"
        '
        'mnuModelStructLOD
        '
        Me.mnuModelStructLOD.Index = 6
        Me.mnuModelStructLOD.Text = "Oracle Load"
        '
        'mnuModelStructSQL
        '
        Me.mnuModelStructSQL.Index = 7
        Me.mnuModelStructSQL.Text = "Oracle Trigger"
        '
        'mnuModelStructMSSQL
        '
        Me.mnuModelStructMSSQL.Index = 8
        Me.mnuModelStructMSSQL.Text = "SQL Server Trigger"
        '
        'mnuModelStructDB2
        '
        Me.mnuModelStructDB2.Enabled = False
        Me.mnuModelStructDB2.Index = 9
        Me.mnuModelStructDB2.Text = "DB2 Load"
        Me.mnuModelStructDB2.Visible = False
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 10
        Me.MenuItem15.Text = "-"
        '
        'mnuDelStruct
        '
        Me.mnuDelStruct.Index = 11
        Me.mnuDelStruct.Text = "Delete Description"
        '
        'mnuChgStruct
        '
        Me.mnuChgStruct.Index = 12
        Me.mnuChgStruct.Text = "Change Description File"
        '
        'mnuDelAllStruct
        '
        Me.mnuDelAllStruct.Index = 13
        Me.mnuDelAllStruct.Text = "Delete All"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 14
        Me.MenuItem6.Text = "-"
        '
        'mnuCopyStruct
        '
        Me.mnuCopyStruct.Index = 15
        Me.mnuCopyStruct.Text = "Copy"
        '
        'mnuPasteStruct
        '
        Me.mnuPasteStruct.Index = 16
        Me.mnuPasteStruct.Text = "Paste"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 17
        Me.MenuItem21.Text = "-"
        '
        'mnuGenRepStr
        '
        Me.mnuGenRepStr.Index = 18
        Me.mnuGenRepStr.Text = "Generate Report"
        '
        'mnuPopupConnection
        '
        Me.mnuPopupConnection.Index = 3
        Me.mnuPopupConnection.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddConn, Me.mnuDelConn, Me.MenuItem22, Me.mnuCopyConn, Me.mnuPasteConn})
        Me.mnuPopupConnection.Text = "mnuPopupConnection"
        '
        'mnuAddConn
        '
        Me.mnuAddConn.Index = 0
        Me.mnuAddConn.Text = "Add Connection"
        '
        'mnuDelConn
        '
        Me.mnuDelConn.Index = 1
        Me.mnuDelConn.Text = "Delete Connection"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 2
        Me.MenuItem22.Text = "-"
        '
        'mnuCopyConn
        '
        Me.mnuCopyConn.Index = 3
        Me.mnuCopyConn.Text = "Copy"
        '
        'mnuPasteConn
        '
        Me.mnuPasteConn.Index = 4
        Me.mnuPasteConn.Text = "Paste"
        '
        'mnuPopupEngine
        '
        Me.mnuPopupEngine.Index = 4
        Me.mnuPopupEngine.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddEngine, Me.mnuDelEngine, Me.MenuItem30, Me.mnuScriptEngine, Me.MenuItem3, Me.mnuCopyEngine, Me.mnuPasteEngine, Me.mnuEngMerge})
        Me.mnuPopupEngine.Text = "mnuPopupEngine"
        '
        'mnuAddEngine
        '
        Me.mnuAddEngine.Index = 0
        Me.mnuAddEngine.Text = "Add Engine"
        '
        'mnuDelEngine
        '
        Me.mnuDelEngine.Index = 1
        Me.mnuDelEngine.Text = "Delete Engine"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 2
        Me.MenuItem30.Text = "-"
        '
        'mnuScriptEngine
        '
        Me.mnuScriptEngine.Index = 3
        Me.mnuScriptEngine.Text = "Generate Script"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 4
        Me.MenuItem3.Text = "-"
        '
        'mnuCopyEngine
        '
        Me.mnuCopyEngine.Index = 5
        Me.mnuCopyEngine.Text = "Copy"
        '
        'mnuPasteEngine
        '
        Me.mnuPasteEngine.Index = 6
        Me.mnuPasteEngine.Text = "Paste"
        '
        'mnuEngMerge
        '
        Me.mnuEngMerge.Index = 7
        Me.mnuEngMerge.Text = "Merge"
        '
        'mnuPopupDS
        '
        Me.mnuPopupDS.Index = 5
        Me.mnuPopupDS.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddDS, Me.mnuLU, Me.mnuDelDS, Me.MenuItem29, Me.mnuCopyDS, Me.mnuPasteDS, Me.MenuItem11, Me.MenuItem9, Me.MenuItem33, Me.mnuModDSSelDTD, Me.mnuModDSSelDDL, Me.mnuModDSSelC, Me.mnuModDSSelLOD, Me.mnuModDSSelSQL, Me.mnuModDSSelMSSQL, Me.MenuItem31, Me.mnuQuery, Me.mnuDelAllDS, Me.MenuItem14, Me.mnuScriptDS})
        Me.mnuPopupDS.Text = "mnuPopupDatastore"
        '
        'mnuAddDS
        '
        Me.mnuAddDS.DefaultItem = True
        Me.mnuAddDS.Index = 0
        Me.mnuAddDS.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddDSBinary, Me.MenuItem23, Me.mnuAddDSDelimited, Me.mnuAddDSXML, Me.mnuAddDSRelational, Me.mnuAddDSDB2LOAD, Me.mnuAddDSHSSUNLOAD, Me.mnuAddDSIMS, Me.mnuAddDSVSAM, Me.MenuItem45, Me.mnuAddDSIMSCDC, Me.mnuAddIMSLE, Me.mnuAddDSDB2CDC, Me.mnuAddDSVSAMCDC, Me.mnuAddOraCDC, Me.mnuAddSQDCDC, Me.MenuItem16, Me.mnuAddDSIBMEvent, Me.MenuItem17, Me.mnuIncludeDS})
        Me.mnuAddDS.Text = "Add Datastore"
        '
        'mnuAddDSBinary
        '
        Me.mnuAddDSBinary.Index = 0
        Me.mnuAddDSBinary.Text = "Binary"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 1
        Me.MenuItem23.Text = "Text"
        '
        'mnuAddDSDelimited
        '
        Me.mnuAddDSDelimited.Index = 2
        Me.mnuAddDSDelimited.Text = "Delimited"
        '
        'mnuAddDSXML
        '
        Me.mnuAddDSXML.Index = 3
        Me.mnuAddDSXML.Text = "XML"
        '
        'mnuAddDSRelational
        '
        Me.mnuAddDSRelational.Index = 4
        Me.mnuAddDSRelational.Text = "Relational"
        '
        'mnuAddDSDB2LOAD
        '
        Me.mnuAddDSDB2LOAD.Index = 5
        Me.mnuAddDSDB2LOAD.Text = "DB2LOAD"
        '
        'mnuAddDSHSSUNLOAD
        '
        Me.mnuAddDSHSSUNLOAD.Index = 6
        Me.mnuAddDSHSSUNLOAD.Text = "HSSUNLOAD"
        '
        'mnuAddDSIMS
        '
        Me.mnuAddDSIMS.Index = 7
        Me.mnuAddDSIMS.Text = "IMS DB"
        '
        'mnuAddDSVSAM
        '
        Me.mnuAddDSVSAM.Index = 8
        Me.mnuAddDSVSAM.Text = "VSAM"
        '
        'MenuItem45
        '
        Me.MenuItem45.Index = 9
        Me.MenuItem45.Text = "-"
        '
        'mnuAddDSIMSCDC
        '
        Me.mnuAddDSIMSCDC.Index = 10
        Me.mnuAddDSIMSCDC.Text = "IMS CDC"
        '
        'mnuAddIMSLE
        '
        Me.mnuAddIMSLE.Index = 11
        Me.mnuAddIMSLE.Text = "IMS CDC LE"
        '
        'mnuAddDSDB2CDC
        '
        Me.mnuAddDSDB2CDC.Index = 12
        Me.mnuAddDSDB2CDC.Text = "DB2 CDC"
        '
        'mnuAddDSVSAMCDC
        '
        Me.mnuAddDSVSAMCDC.Index = 13
        Me.mnuAddDSVSAMCDC.Text = "VSAM CDC"
        '
        'mnuAddOraCDC
        '
        Me.mnuAddOraCDC.Index = 14
        Me.mnuAddOraCDC.Text = "Oracle CDC"
        '
        'mnuAddSQDCDC
        '
        Me.mnuAddSQDCDC.Index = 15
        Me.mnuAddSQDCDC.Text = "Generic CDC"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 16
        Me.MenuItem16.Text = "-"
        '
        'mnuAddDSIBMEvent
        '
        Me.mnuAddDSIBMEvent.Index = 17
        Me.mnuAddDSIBMEvent.Text = "IBM Event"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 18
        Me.MenuItem17.Text = "-"
        '
        'mnuIncludeDS
        '
        Me.mnuIncludeDS.Index = 19
        Me.mnuIncludeDS.Text = "Include File"
        '
        'mnuLU
        '
        Me.mnuLU.Index = 1
        Me.mnuLU.Text = "Add Look Up"
        '
        'mnuDelDS
        '
        Me.mnuDelDS.Index = 2
        Me.mnuDelDS.Text = "Delete Datastore"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 3
        Me.MenuItem29.Text = "-"
        '
        'mnuCopyDS
        '
        Me.mnuCopyDS.Index = 4
        Me.mnuCopyDS.Text = "Copy"
        '
        'mnuPasteDS
        '
        Me.mnuPasteDS.Index = 5
        Me.mnuPasteDS.Text = "Paste"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 6
        Me.MenuItem11.Text = "-"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 7
        Me.MenuItem9.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuMapAsBinary, Me.mnuMapAsText, Me.mnuMapAsDelimited, Me.mnuMapAsXML, Me.mnuMapAsRelational, Me.mnuMapAsVSAM, Me.mnuMapAsIMS, Me.mnuMapAsDB2LOAD, Me.mnuMapAsHSSUNLOAD, Me.mnuMapAsIBMEvent, Me.MenuItem36, Me.mnuMapAsIMSCDC, Me.mnuMapAsDB2CDC, Me.mnuMapAsVSAMCDC, Me.mnuMapAsXMLCDC, Me.mnuMapAsTriggerCDC, Me.mnuMapAsOraCDC, Me.mnuMapAsSQDCDC, Me.MenuItem4, Me.mnuMapAsIMSLE, Me.mnuMapAsIMSLEBat})
        Me.MenuItem9.Text = "Map As"
        '
        'mnuMapAsBinary
        '
        Me.mnuMapAsBinary.Index = 0
        Me.mnuMapAsBinary.Text = "Binary"
        '
        'mnuMapAsText
        '
        Me.mnuMapAsText.Index = 1
        Me.mnuMapAsText.Text = "Text"
        '
        'mnuMapAsDelimited
        '
        Me.mnuMapAsDelimited.Index = 2
        Me.mnuMapAsDelimited.Text = "Delimited"
        '
        'mnuMapAsXML
        '
        Me.mnuMapAsXML.Index = 3
        Me.mnuMapAsXML.Text = "XML"
        '
        'mnuMapAsRelational
        '
        Me.mnuMapAsRelational.Index = 4
        Me.mnuMapAsRelational.Text = "Relational"
        '
        'mnuMapAsVSAM
        '
        Me.mnuMapAsVSAM.Index = 5
        Me.mnuMapAsVSAM.Text = "VSAM"
        '
        'mnuMapAsIMS
        '
        Me.mnuMapAsIMS.Index = 6
        Me.mnuMapAsIMS.Text = "IMS"
        '
        'mnuMapAsDB2LOAD
        '
        Me.mnuMapAsDB2LOAD.Index = 7
        Me.mnuMapAsDB2LOAD.Text = "DB2LOAD"
        '
        'mnuMapAsHSSUNLOAD
        '
        Me.mnuMapAsHSSUNLOAD.Index = 8
        Me.mnuMapAsHSSUNLOAD.Text = "HSSUNLOAD"
        '
        'mnuMapAsIBMEvent
        '
        Me.mnuMapAsIBMEvent.Index = 9
        Me.mnuMapAsIBMEvent.Text = "IBM EVENT"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 10
        Me.MenuItem36.Text = "-"
        '
        'mnuMapAsIMSCDC
        '
        Me.mnuMapAsIMSCDC.Index = 11
        Me.mnuMapAsIMSCDC.Text = "IMS CDC"
        '
        'mnuMapAsDB2CDC
        '
        Me.mnuMapAsDB2CDC.Index = 12
        Me.mnuMapAsDB2CDC.Text = "DB2 CDC"
        '
        'mnuMapAsVSAMCDC
        '
        Me.mnuMapAsVSAMCDC.Index = 13
        Me.mnuMapAsVSAMCDC.Text = "VSAM CDC"
        '
        'mnuMapAsXMLCDC
        '
        Me.mnuMapAsXMLCDC.Index = 14
        Me.mnuMapAsXMLCDC.Text = "XML CDC"
        '
        'mnuMapAsTriggerCDC
        '
        Me.mnuMapAsTriggerCDC.Index = 15
        Me.mnuMapAsTriggerCDC.Text = "Trigger Based CDC"
        '
        'mnuMapAsOraCDC
        '
        Me.mnuMapAsOraCDC.Index = 16
        Me.mnuMapAsOraCDC.Text = "Oracle CDC"
        '
        'mnuMapAsSQDCDC
        '
        Me.mnuMapAsSQDCDC.Index = 17
        Me.mnuMapAsSQDCDC.Text = "Generic CDC"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 18
        Me.MenuItem4.Text = "-"
        '
        'mnuMapAsIMSLE
        '
        Me.mnuMapAsIMSLE.Index = 19
        Me.mnuMapAsIMSLE.Text = "IMS LE"
        '
        'mnuMapAsIMSLEBat
        '
        Me.mnuMapAsIMSLEBat.Index = 20
        Me.mnuMapAsIMSLEBat.Text = "IMS LE Batch"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 8
        Me.MenuItem33.Text = "-"
        '
        'mnuModDSSelDTD
        '
        Me.mnuModDSSelDTD.Enabled = False
        Me.mnuModDSSelDTD.Index = 9
        Me.mnuModDSSelDTD.Text = "Model DTD"
        '
        'mnuModDSSelDDL
        '
        Me.mnuModDSSelDDL.Enabled = False
        Me.mnuModDSSelDDL.Index = 10
        Me.mnuModDSSelDDL.Text = "Model DDL"
        '
        'mnuModDSSelC
        '
        Me.mnuModDSSelC.Enabled = False
        Me.mnuModDSSelC.Index = 11
        Me.mnuModDSSelC.Text = "Model C"
        '
        'mnuModDSSelLOD
        '
        Me.mnuModDSSelLOD.Enabled = False
        Me.mnuModDSSelLOD.Index = 12
        Me.mnuModDSSelLOD.Text = "Oracle LOD"
        '
        'mnuModDSSelSQL
        '
        Me.mnuModDSSelSQL.Enabled = False
        Me.mnuModDSSelSQL.Index = 13
        Me.mnuModDSSelSQL.Text = "Oracle Trigger"
        '
        'mnuModDSSelMSSQL
        '
        Me.mnuModDSSelMSSQL.Enabled = False
        Me.mnuModDSSelMSSQL.Index = 14
        Me.mnuModDSSelMSSQL.Text = "SQL Server Trigger"
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 15
        Me.MenuItem31.Text = "-"
        '
        'mnuQuery
        '
        Me.mnuQuery.Index = 16
        Me.mnuQuery.Text = "Query"
        Me.mnuQuery.Visible = False
        '
        'mnuDelAllDS
        '
        Me.mnuDelAllDS.Index = 17
        Me.mnuDelAllDS.Text = "Delete all"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 18
        Me.MenuItem14.Text = "-"
        '
        'mnuScriptDS
        '
        Me.mnuScriptDS.Index = 19
        Me.mnuScriptDS.Text = "Script"
        '
        'mnuPopupVariable
        '
        Me.mnuPopupVariable.Index = 6
        Me.mnuPopupVariable.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddVar, Me.mnuDelVar, Me.MenuItem39, Me.mnuCopyVar, Me.mnuPasteVar})
        Me.mnuPopupVariable.Text = "mnuPopupVar"
        '
        'mnuAddVar
        '
        Me.mnuAddVar.Index = 0
        Me.mnuAddVar.Text = "Add Variable"
        '
        'mnuDelVar
        '
        Me.mnuDelVar.Index = 1
        Me.mnuDelVar.Text = "Delete Variable"
        '
        'MenuItem39
        '
        Me.MenuItem39.Index = 2
        Me.MenuItem39.Text = "-"
        '
        'mnuCopyVar
        '
        Me.mnuCopyVar.Index = 3
        Me.mnuCopyVar.Text = "Copy"
        '
        'mnuPasteVar
        '
        Me.mnuPasteVar.Index = 4
        Me.mnuPasteVar.Text = "Paste"
        '
        'mnuPopupTask
        '
        Me.mnuPopupTask.Index = 7
        Me.mnuPopupTask.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddTask, Me.mnuAddGen, Me.mnuDelTask, Me.mnuIncProc, Me.MenuItem42, Me.mnuCopyTask, Me.mnuPasteTask, Me.MenuItem12, Me.mnuChangeTask, Me.mnuDelAllTask, Me.MenuItem20, Me.mnuScriptProc})
        Me.mnuPopupTask.Text = "mnuPopupTask"
        '
        'mnuAddTask
        '
        Me.mnuAddTask.Index = 0
        Me.mnuAddTask.Text = "Add XXXX"
        '
        'mnuAddGen
        '
        Me.mnuAddGen.Index = 1
        Me.mnuAddGen.Text = "Add General Procedure"
        '
        'mnuDelTask
        '
        Me.mnuDelTask.Index = 2
        Me.mnuDelTask.Text = "Delete XXX"
        '
        'mnuIncProc
        '
        Me.mnuIncProc.Index = 3
        Me.mnuIncProc.Text = "Add Include Proc"
        '
        'MenuItem42
        '
        Me.MenuItem42.Index = 4
        Me.MenuItem42.Text = "-"
        '
        'mnuCopyTask
        '
        Me.mnuCopyTask.Index = 5
        Me.mnuCopyTask.Text = "Copy"
        '
        'mnuPasteTask
        '
        Me.mnuPasteTask.Index = 6
        Me.mnuPasteTask.Text = "Paste"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 7
        Me.MenuItem12.Text = "-"
        '
        'mnuChangeTask
        '
        Me.mnuChangeTask.Index = 8
        Me.mnuChangeTask.Text = "Add/Remove Datastores"
        '
        'mnuDelAllTask
        '
        Me.mnuDelAllTask.Index = 9
        Me.mnuDelAllTask.Text = "Delete All"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 10
        Me.MenuItem20.Text = "-"
        '
        'mnuScriptProc
        '
        Me.mnuScriptProc.Index = 11
        Me.mnuScriptProc.Text = "Script Procedure"
        '
        'mnuPopupProject
        '
        Me.mnuPopupProject.Index = 8
        Me.mnuPopupProject.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuDelProject})
        Me.mnuPopupProject.Text = "mnuPopupProject"
        '
        'mnuDelProject
        '
        Me.mnuDelProject.Index = 0
        Me.mnuDelProject.Text = "Delete Entire Project"
        '
        'mnuPopupStructSelection
        '
        Me.mnuPopupStructSelection.Index = 9
        Me.mnuPopupStructSelection.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuAddStructSelection, Me.mnuDelStructSelection, Me.MenuItem18, Me.mnuModelStructSelectionDTD, Me.mnuModelStructSelectionDDL, Me.mnuModelStructSelectionHeader, Me.mnuModelSSLOD, Me.mnuModelSSSQL, Me.mnuModelSSMSSQL, Me.mnuModelSSDB2, Me.MenuItem19, Me.mnuCopyStructSel, Me.mnuPasteStructSel})
        Me.mnuPopupStructSelection.Text = "mnuPopupStructSelection"
        '
        'mnuAddStructSelection
        '
        Me.mnuAddStructSelection.Index = 0
        Me.mnuAddStructSelection.Text = "Add Subset"
        '
        'mnuDelStructSelection
        '
        Me.mnuDelStructSelection.Index = 1
        Me.mnuDelStructSelection.Text = "Delete Subset"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 2
        Me.MenuItem18.Text = "-"
        '
        'mnuModelStructSelectionDTD
        '
        Me.mnuModelStructSelectionDTD.Index = 3
        Me.mnuModelStructSelectionDTD.Text = "Model DTD"
        '
        'mnuModelStructSelectionDDL
        '
        Me.mnuModelStructSelectionDDL.Index = 4
        Me.mnuModelStructSelectionDDL.Text = "Model DDL"
        '
        'mnuModelStructSelectionHeader
        '
        Me.mnuModelStructSelectionHeader.Index = 5
        Me.mnuModelStructSelectionHeader.Text = "Model C"
        '
        'mnuModelSSLOD
        '
        Me.mnuModelSSLOD.Index = 6
        Me.mnuModelSSLOD.Text = "Oracle Load"
        '
        'mnuModelSSSQL
        '
        Me.mnuModelSSSQL.Index = 7
        Me.mnuModelSSSQL.Text = "Oracle Trigger"
        '
        'mnuModelSSMSSQL
        '
        Me.mnuModelSSMSSQL.Index = 8
        Me.mnuModelSSMSSQL.Text = "SQLServer Trigger"
        '
        'mnuModelSSDB2
        '
        Me.mnuModelSSDB2.Enabled = False
        Me.mnuModelSSDB2.Index = 9
        Me.mnuModelSSDB2.Text = "DB2 Load"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 10
        Me.MenuItem19.Text = "-"
        '
        'mnuCopyStructSel
        '
        Me.mnuCopyStructSel.Index = 11
        Me.mnuCopyStructSel.Text = "Copy"
        '
        'mnuPasteStructSel
        '
        Me.mnuPasteStructSel.Index = 12
        Me.mnuPasteStructSel.Text = "Paste"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "*.sqp"
        Me.SaveFileDialog1.Filter = "SQData Studio Project Files (*.sqp)|*.sqp |All files (*.*)|*.*"
        '
        'pnlProp
        '
        Me.pnlProp.AllowDrop = True
        Me.pnlProp.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.pnlProp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlProp.Controls.Add(Me.ctMain)
        Me.pnlProp.Controls.Add(Me.ctInc)
        Me.pnlProp.Controls.Add(Me.ctFolder)
        Me.pnlProp.Controls.Add(Me.ctTask)
        Me.pnlProp.Controls.Add(Me.ctVar)
        Me.pnlProp.Controls.Add(Me.ctEng)
        Me.pnlProp.Controls.Add(Me.ctStrSel)
        Me.pnlProp.Controls.Add(Me.ctSys)
        Me.pnlProp.Controls.Add(Me.ctConn)
        Me.pnlProp.Controls.Add(Me.ctStr)
        Me.pnlProp.Controls.Add(Me.ctEnv)
        Me.pnlProp.Controls.Add(Me.ctPrj)
        Me.pnlProp.Controls.Add(Me.ctDs)
        Me.pnlProp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlProp.Location = New System.Drawing.Point(0, 0)
        Me.pnlProp.Name = "pnlProp"
        Me.pnlProp.Padding = New System.Windows.Forms.Padding(5, 0, 0, 0)
        Me.pnlProp.Size = New System.Drawing.Size(665, 513)
        Me.pnlProp.TabIndex = 3
        '
        'SCmain
        '
        Me.SCmain.AllowDrop = True
        Me.SCmain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SCmain.Location = New System.Drawing.Point(0, 29)
        Me.SCmain.Name = "SCmain"
        '
        'SCmain.Panel1
        '
        Me.SCmain.Panel1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.SCmain.Panel1.Controls.Add(Me.tvExplorer)
        '
        'SCmain.Panel2
        '
        Me.SCmain.Panel2.Controls.Add(Me.pnlProp)
        Me.SCmain.Size = New System.Drawing.Size(1016, 513)
        Me.SCmain.SplitterDistance = 346
        Me.SCmain.SplitterWidth = 5
        Me.SCmain.TabIndex = 9
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(700, 545)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(234, 19)
        Me.ProgressBar1.TabIndex = 10
        '
        'ctMain
        '
        Me.ctMain.AllowDrop = True
        Me.ctMain.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.ctMain.Location = New System.Drawing.Point(48, 48)
        Me.ctMain.Name = "ctMain"
        Me.ctMain.Size = New System.Drawing.Size(874, 611)
        Me.ctMain.TabIndex = 12
        '
        'ctInc
        '
        Me.ctInc.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctInc.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.ctInc.Location = New System.Drawing.Point(155, 22)
        Me.ctInc.Name = "ctInc"
        Me.ctInc.Size = New System.Drawing.Size(653, 530)
        Me.ctInc.TabIndex = 11
        '
        'ctFolder
        '
        Me.ctFolder.BackColor = System.Drawing.SystemColors.Desktop
        Me.ctFolder.ForeColor = System.Drawing.Color.White
        Me.ctFolder.Location = New System.Drawing.Point(7, 2)
        Me.ctFolder.Name = "ctFolder"
        Me.ctFolder.Size = New System.Drawing.Size(488, 344)
        Me.ctFolder.TabIndex = 10
        Me.ctFolder.Visible = False
        '
        'ctTask
        '
        Me.ctTask.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctTask.ForeColor = System.Drawing.Color.White
        Me.ctTask.Location = New System.Drawing.Point(21, 22)
        Me.ctTask.Name = "ctTask"
        Me.ctTask.Size = New System.Drawing.Size(672, 472)
        Me.ctTask.TabIndex = 9
        '
        'ctVar
        '
        Me.ctVar.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctVar.ForeColor = System.Drawing.Color.White
        Me.ctVar.Location = New System.Drawing.Point(31, 48)
        Me.ctVar.Name = "ctVar"
        Me.ctVar.Size = New System.Drawing.Size(488, 344)
        Me.ctVar.TabIndex = 8
        '
        'ctEng
        '
        Me.ctEng.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctEng.ForeColor = System.Drawing.Color.White
        Me.ctEng.Location = New System.Drawing.Point(24, 192)
        Me.ctEng.Name = "ctEng"
        Me.ctEng.Size = New System.Drawing.Size(496, 344)
        Me.ctEng.TabIndex = 7
        '
        'ctStrSel
        '
        Me.ctStrSel.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctStrSel.ForeColor = System.Drawing.Color.White
        Me.ctStrSel.Location = New System.Drawing.Point(232, 112)
        Me.ctStrSel.Name = "ctStrSel"
        Me.ctStrSel.Size = New System.Drawing.Size(608, 448)
        Me.ctStrSel.TabIndex = 6
        '
        'ctSys
        '
        Me.ctSys.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctSys.ForeColor = System.Drawing.Color.White
        Me.ctSys.Location = New System.Drawing.Point(48, 32)
        Me.ctSys.Name = "ctSys"
        Me.ctSys.Size = New System.Drawing.Size(504, 352)
        Me.ctSys.TabIndex = 5
        '
        'ctConn
        '
        Me.ctConn.AutoValidate = System.Windows.Forms.AutoValidate.Disable
        Me.ctConn.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctConn.ForeColor = System.Drawing.Color.White
        Me.ctConn.ImeMode = System.Windows.Forms.ImeMode.[On]
        Me.ctConn.Location = New System.Drawing.Point(272, 104)
        Me.ctConn.Name = "ctConn"
        Me.ctConn.Size = New System.Drawing.Size(560, 352)
        Me.ctConn.TabIndex = 4
        '
        'ctStr
        '
        Me.ctStr.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctStr.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ctStr.ForeColor = System.Drawing.Color.White
        Me.ctStr.Location = New System.Drawing.Point(328, 48)
        Me.ctStr.Name = "ctStr"
        Me.ctStr.Size = New System.Drawing.Size(608, 456)
        Me.ctStr.TabIndex = 2
        '
        'ctEnv
        '
        Me.ctEnv.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctEnv.ForeColor = System.Drawing.Color.White
        Me.ctEnv.Location = New System.Drawing.Point(144, 152)
        Me.ctEnv.Name = "ctEnv"
        Me.ctEnv.Size = New System.Drawing.Size(488, 344)
        Me.ctEnv.TabIndex = 1
        '
        'ctPrj
        '
        Me.ctPrj.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctPrj.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ctPrj.ForeColor = System.Drawing.Color.White
        Me.ctPrj.Location = New System.Drawing.Point(272, 168)
        Me.ctPrj.Name = "ctPrj"
        Me.ctPrj.Size = New System.Drawing.Size(512, 328)
        Me.ctPrj.TabIndex = 0
        '
        'ctDs
        '
        Me.ctDs.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ctDs.ForeColor = System.Drawing.Color.White
        Me.ctDs.Location = New System.Drawing.Point(44, 29)
        Me.ctDs.Name = "ctDs"
        Me.ctDs.Size = New System.Drawing.Size(568, 648)
        Me.ctDs.TabIndex = 3
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 566)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.SCmain)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.mnuMain
        Me.MinimumSize = New System.Drawing.Size(800, 600)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlProp.ResumeLayout(False)
        Me.SCmain.Panel1.ResumeLayout(False)
        Me.SCmain.Panel2.ResumeLayout(False)
        Me.SCmain.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    '//////// Events for the form itself ////////
#Region "Main Form Events"

    Function InitUI() As Boolean
        Try
            tvExplorer.Font = New Font("Arial", 8, FontStyle.Bold)

            imgListSmall = Me.ImageListSmall
            dlgOpen = Me.OpenFileDialog1
            dlgSave = Me.SaveFileDialog1
            dlgBrowseFolder = Me.FolderBrowserDialog1

            EnableTreeActionButton(False)

            HideAllUC()

            ctPrj.Location = New Point(0, 0)
            ctEnv.Location = ctPrj.Location
            ctStr.Location = ctPrj.Location
            ctDs.Location = ctPrj.Location
            ctSys.Location = ctPrj.Location
            ctConn.Location = ctPrj.Location
            ctEng.Location = ctPrj.Location
            ctStrSel.Location = ctPrj.Location
            ctVar.Location = ctPrj.Location
            ctTask.Location = ctPrj.Location
            ctMain.Location = ctPrj.Location
            ctFolder.Location = ctPrj.Location
            ctInc.Location = ctPrj.Location


        Catch ex As Exception
            LogError(ex, "frmMain InitUI")
        End Try
        '//TODO
    End Function

    Private Sub pnlProp_SizeChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlProp.SizeChanged

        Try
            ctPrj.Size = pnlProp.Size
            ctEnv.Size = pnlProp.Size
            ctStr.Size = pnlProp.Size
            ctDs.Size = pnlProp.Size
            ctSys.Size = pnlProp.Size
            ctConn.Size = pnlProp.Size
            ctEng.Size = pnlProp.Size
            ctStrSel.Size = pnlProp.Size
            ctVar.Size = pnlProp.Size
            ctTask.Size = pnlProp.Size
            ctMain.Size = pnlProp.Size
            ctFolder.Size = pnlProp.Size
            ctInc.Size = pnlProp.Size

        Catch ex As Exception
            LogError(ex, "pnlProp_SizeChanged")
        End Try

    End Sub

    Sub HideAllUC()

        ctPrj.Visible = False
        ctEnv.Visible = False
        ctStr.Visible = False
        ctDs.Visible = False
        ctConn.Visible = False
        ctSys.Visible = False
        ctEng.Visible = False
        ctStrSel.Visible = False
        ctVar.Visible = False
        ctTask.Visible = False
        ctMain.Visible = False
        ctFolder.Visible = False
        ctInc.Visible = False

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Application.CommonAppDataRegistry.GetValue("WinState") IsNot Nothing Then
            Dim winStr As String = Application.CommonAppDataRegistry.GetValue("WinState")
            Select Case winStr
                Case "Maximized"
                    Me.WindowState = FormWindowState.Maximized
                Case "Normal"
                    Me.WindowState = FormWindowState.Normal
                Case Else
                    Me.WindowState = FormWindowState.Normal
            End Select

        End If
        InitMain()
        LoadGlobalValues()
        ToolBar1.Buttons(enumToolBarButtons.TB_MAP).Enabled = False
        System.Environment.SetEnvironmentVariable("PATH", GetAppPath() & ";" & System.Environment.GetEnvironmentVariable("PATH"))

    End Sub

    Sub InitMain()

        '#If CONFIG = "ETI" Then
        '        Me.Text = "ETI CDC Studio" ' V3 " & Application.ProductVersion
        '        Me.mnuAbout.Text = "About ETI CDC Studio"
        '        'Me.imgLogo.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstudio\images\ETI_Logo_Installation_Graphic\ETI-installer-right.bmp"
        '        Me.Name = "ETICDCStudio"
        '#Else
        Me.Text = "Design Studio V3 " & Application.ProductVersion
        Me.mnuAbout.Text = "About Design Studio"
        'Me.imgLogo.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstudio\images\FormTop\sq_skyblue.jpg"
        'Me.Icon = Drawing.Icon.ExtractAssociatedIcon("C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstudio\images\icons\SQ.ico")
        '#End If

        mainwindow = Me

    End Sub

    Private Sub Splitter1_SplitterMoved(ByVal sender As System.Object, ByVal e As System.Windows.Forms.SplitterEventArgs) Handles SCmain.SplitterMoved

        If CurLoadedProject Is Nothing Then Exit Sub

        If IsEventFromCode = False Then
            CurLoadedProject.MainSeparatorX = e.X
        End If

    End Sub

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        '/// Save Window State
        Application.CommonAppDataRegistry.SetValue("WinState", Me.WindowState)
        '//save project settings
        If Not CurLoadedProject Is Nothing Then
            CurLoadedProject.Save(True)
            CurLoadedProject.SaveToRegistry()
            CurLoadedProject = Nothing
        End If
        If cnn IsNot Nothing Then
            cnn.Close()
        End If

    End Sub

    '/// from user control event
    Private Sub OnSave(ByVal sender As System.Object, ByVal e As INode) Handles ctEnv.Saved, ctPrj.Saved, ctStr.Saved, ctConn.Saved, ctSys.Saved, ctEng.Saved, ctStrSel.Saved, ctVar.Saved, ctTask.Saved, ctDs.Saved, ctInc.Saved, ctMain.Saved

        Try
            '// logical stupidity fixed by tk 8/16/2007
            '/// (logical stupidity created by tk 1/07  |8-O )
            lblStatusMsg.Text = "Save Successful"
            lblStatusMsg.Show()

            ToolBar1.Buttons(enumToolBarButtons.TB_SAVE).Enabled = False

            If Not prevNode Is Nothing Then
                UpdateNode(prevNode, e)
            End If

            Dim NodeText As String = tvExplorer.SelectedNode.Text
            'tvExplorer.Refresh()
            'FillProject(CurLoadedProject, , True)
            tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, NodeText)
            ShowUsercontrol(tvExplorer.SelectedNode)

        Catch ex As Exception
            LogError(ex, "frmMain OnSave")
        End Try

    End Sub

    '/// from user control event
    Private Sub OnModify(ByVal sender As System.Object, ByVal e As INode) Handles ctEnv.Modified, ctPrj.Modified, ctStr.Modified, ctDs.Modified, ctConn.Modified, ctSys.Modified, ctEng.Modified, ctStrSel.Modified, ctVar.Modified, ctTask.Modified, ctInc.Modified, ctMain.Modified

        ToolBar1.Buttons(enumToolBarButtons.TB_SAVE).Enabled = True

    End Sub

    '/// from user control event
    Private Sub OnRename(ByVal sender As System.Object, ByVal e As INode) Handles ctEnv.Renamed, ctPrj.Renamed, ctStr.Renamed, ctDs.Renamed, ctConn.Renamed, ctSys.Renamed, ctEng.Renamed, ctStrSel.Renamed, ctVar.Renamed, ctTask.Renamed, ctInc.Renamed, ctMain.Renamed

        lblStatusMsg.Text = "Rename and Save Successful"
        lblStatusMsg.Show()
        IsRename = True
        '// Created for Renaming by Tom Karasch
        ToolBar1.Buttons(enumToolBarButtons.TB_SAVE).Enabled = False
        FillProject(e.Project, True)
        If e.Type <> NODE_SOURCEDATASTORE And e.Type <> NODE_TARGETDATASTORE Then
            e.ObjTreeNode.Text = e.Text
        End If
        'If PrevObjTreeNode IsNot Nothing Then
        tvExplorer.SelectedNode = e.ObjTreeNode
        'Else
        'tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, NameOfNodeBefore)
        'End If
        ShowUsercontrol(tvExplorer.SelectedNode)
        IsRename = False

    End Sub

    '/// from user control event
    Private Sub OnClose(ByVal sender As System.Object, ByVal e As INode) Handles ctPrj.Closed, ctEnv.Closed, ctConn.Closed, ctDs.Closed, ctEng.Closed, ctInc.Closed, ctStr.Closed, ctStrSel.Closed, ctSys.Closed, ctTask.Closed, ctVar.Closed, ctMain.Closed

        If SCmain.Panel1Collapsed = True Then
            Call SCmain_DoubleClick(Me, New EventArgs)
        End If

    End Sub

#End Region

    '//////// TVexplorer tree events .... etc... ////////
#Region "tvexplorer functions"

    Private Sub tvExplorer_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tvExplorer.Leave

        IsFocusOnExplorer = False
        mnuEdit.Enabled = False

        EnableTreeActionButton(False)

    End Sub

    Private Sub tvExplorer_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tvExplorer.DoubleClick

        '// Modified 1/07 by Tom Karasch

        If Not (tvExplorer.SelectedNode Is Nothing) Then
            '//if user pressed cancel then dont open popup form
            If SavePreviousScreen(tvExplorer.SelectedNode) = False Then
                Exit Sub
            End If
            ShowUsercontrol(tvExplorer.SelectedNode, True)
        End If

    End Sub

    Private Sub tvExplorer_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvExplorer.MouseClick

        ' See if this is the right button.
        If e.Button = Windows.Forms.MouseButtons.Right Then
            IsEventFromCode = True
            ' Select this node.
            Dim node_here As TreeNode = tvExplorer.GetNodeAt(e.X, e.Y)
            tvExplorer.SelectedNode = node_here

            ' See if we got a node.
            If node_here Is Nothing Then Exit Sub

            ShowPopupMenu(tvExplorer.SelectedNode, New Point(e.X, e.Y))
            IsEventFromCode = False
        Else
            Dim node_here As TreeNode = tvExplorer.GetNodeAt(e.X, e.Y)
            tvExplorer.SelectedNode = node_here
        End If

    End Sub

    Private Sub tvExplorer_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvExplorer.AfterSelect
        '// Modified 11/06 by Tom Karasch

        Me.Cursor = Cursors.WaitCursor
        IsFocusOnExplorer = True
        mnuEdit.Enabled = True

        If Not (tvExplorer.SelectedNode Is Nothing) Then
            EnableTreeActionButton(True)
            ShowStatusMessage("Selected node : [" & CType(tvExplorer.SelectedNode.Tag, INode).Key.Replace("-", "->") & "]")
        End If

        If Not objTopMost Is Nothing Then
            Dim obj As INode
            obj = tvExplorer.SelectedNode.Tag

            '//This will prevent firing save event if user select the same node which he
            '// is editing. Will allow editing of DS selections

            '/// changed 6/5/07 by tk 
            If objTopMost Is obj And tvExplorer.SelectedNode Is prevNode Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
        End If

        If SavePreviousScreen(e.Node) = True Then
            If IsRename = False Then
                ShowUsercontrol(e.Node, True)
            Else
                ShowUsercontrol(e.Node)
            End If

        End If

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub tvExplorer_AfterExpandAfterCollapse(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvExplorer.AfterExpand, tvExplorer.AfterCollapse

        
        If tvExplorer.SelectedNode IsNot Nothing Then
            Dim Nodetype As String = CType(tvExplorer.SelectedNode.Tag, INode).Type
            Dim NodeText As String = tvExplorer.SelectedNode.Text
            Dim NodeState As Boolean = tvExplorer.SelectedNode.IsExpanded

            Select Case Nodetype
                Case NODE_FO_ENVIRONMENT
                    CurLoadedProject.ENV_FOexpanded = NodeState

                Case NODE_ENVIRONMENT
                    CurLoadedProject.ENVexpanded = NodeState

                Case NODE_FO_CONNECTION
                    CurLoadedProject.CONNexpanded = NodeState

                Case NODE_FO_STRUCT
                    If NodeText.Contains("Descriptions") = True Then
                        CurLoadedProject.STRexpanded = NodeState
                    ElseIf NodeText.EndsWith("COBOL") = True Then
                        CurLoadedProject.COBOLexpanded = NodeState
                    ElseIf NodeText.EndsWith("COBOLIMS") = True Then
                        CurLoadedProject.COBOLIMSexpanded = NodeState
                    ElseIf NodeText.EndsWith("CHeader") = True Then
                        CurLoadedProject.Cexpanded = NodeState
                    ElseIf NodeText.EndsWith("XMLDTD") = True Then
                        CurLoadedProject.XMLDTDexpanded = NodeState
                    ElseIf NodeText.EndsWith("DDL") = True Then
                        CurLoadedProject.DDLexpanded = NodeState
                    ElseIf NodeText.EndsWith("DML") = True Then
                        CurLoadedProject.DMLexpanded = NodeState
                    End If

                Case NODE_FO_SYSTEM
                    CurLoadedProject.SYS_FOexpanded = NodeState

                Case NODE_SYSTEM
                    CurLoadedProject.SYSexpanded = NodeState

                Case NODE_FO_ENGINE
                    CurLoadedProject.ENG_FOexpanded = NodeState

                Case NODE_ENGINE
                    CurLoadedProject.ENGexpanded = NodeState

                Case NODE_FO_SOURCEDATASTORE
                    CurLoadedProject.SRCexpanded = NodeState

                Case NODE_SOURCEDATASTORE
                    CurLoadedProject.SRCselExpanded = NodeState

                Case NODE_FO_TARGETDATASTORE
                    CurLoadedProject.TGTexpanded = NodeState

                Case NODE_FO_VARIABLE
                    CurLoadedProject.VARexpanded = NodeState

                Case NODE_FO_PROC
                    CurLoadedProject.PROCexpanded = NodeState

                Case NODE_FO_MAIN
                    CurLoadedProject.MAINexpanded = NodeState

                Case Else

            End Select
        End If

    End Sub

    'Private Sub tvExplorer_AfterCollapse(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)

    '    If tvExplorer.SelectedNode.IsExpanded = True Then Exit Sub

    '    If tvExplorer.SelectedNode IsNot Nothing Then
    '        Dim Nodetype As String = CType(tvExplorer.SelectedNode.Tag, INode).Type
    '        Dim NodeText As String = tvExplorer.SelectedNode.Text
    '        Dim NodeState As Boolean = tvExplorer.SelectedNode.IsExpanded

    '        Select Case Nodetype
    '            Case NODE_FO_ENVIRONMENT
    '                CurLoadedProject.ENV_FOexpanded = NodeState

    '            Case NODE_ENVIRONMENT
    '                CurLoadedProject.ENVexpanded = NodeState

    '            Case NODE_FO_CONNECTION
    '                CurLoadedProject.CONNexpanded = NodeState

    '            Case NODE_FO_STRUCT
    '                If NodeText.Contains("Descriptions") = True Then
    '                    CurLoadedProject.STRexpanded = NodeState
    '                ElseIf NodeText.EndsWith("COBOL") = True Then
    '                    CurLoadedProject.COBOLexpanded = NodeState
    '                ElseIf NodeText.EndsWith("COBOLIMS") = True Then
    '                    CurLoadedProject.COBOLIMSexpanded = NodeState
    '                ElseIf NodeText.EndsWith("CHeader") = True Then
    '                    CurLoadedProject.Cexpanded = NodeState
    '                ElseIf NodeText.EndsWith("XMLDTD") = True Then
    '                    CurLoadedProject.XMLDTDexpanded = NodeState
    '                ElseIf NodeText.EndsWith("DDL") = True Then
    '                    CurLoadedProject.DDLexpanded = NodeState
    '                ElseIf NodeText.EndsWith("DML") = True Then
    '                    CurLoadedProject.DMLexpanded = NodeState
    '                End If

    '            Case NODE_FO_SYSTEM
    '                CurLoadedProject.SYS_FOexpanded = NodeState

    '            Case NODE_SYSTEM
    '                CurLoadedProject.SYSexpanded = NodeState

    '            Case NODE_FO_ENGINE
    '                CurLoadedProject.ENG_FOexpanded = NodeState

    '            Case NODE_ENGINE
    '                CurLoadedProject.ENGexpanded = NodeState

    '            Case NODE_FO_SOURCEDATASTORE
    '                CurLoadedProject.SRCexpanded = NodeState

    '            Case NODE_SOURCEDATASTORE
    '                CurLoadedProject.SRCselExpanded = NodeState

    '            Case NODE_FO_TARGETDATASTORE
    '                CurLoadedProject.TGTexpanded = NodeState

    '            Case NODE_FO_VARIABLE
    '                CurLoadedProject.VARexpanded = NodeState

    '            Case NODE_FO_PROC
    '                CurLoadedProject.PROCexpanded = NodeState

    '            Case NODE_FO_MAIN
    '                CurLoadedProject.MAINexpanded = NodeState

    '        End Select
    '    End If

    'End Sub

    Private Sub tvExplorer_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvExplorer.KeyDown

        '///OverHauled by Tom Karasch 9/06-1/07

        If tvExplorer.Focused = False Then Exit Sub

        If e.KeyCode = Keys.Enter Then
            tvExplorer_DoubleClick(Me, New EventArgs)
        End If

        If e.KeyCode = Keys.F1 Then
            Dim obj As INode
            Dim HelpObj As INode = Nothing
            If tvExplorer.SelectedNode Is Nothing Then
                ShowHelp(modHelp.HHId.H_Object_Tree)
            Else
                obj = tvExplorer.SelectedNode.Tag
                If Not tvExplorer.SelectedNode.Parent Is Nothing Then
                    HelpObj = tvExplorer.SelectedNode.Parent.Tag
                End If

                '//Global procedure for all type node F1 HELP 

                If Not (obj Is Nothing) Then
                    Select Case obj.Type
                        Case NODE_PROJECT
                            ShowHelp(modHelp.HHId.H_Project)
                        Case NODE_ENVIRONMENT, NODE_FO_ENVIRONMENT
                            ShowHelp(modHelp.HHId.H_Environment)
                        Case NODE_STRUCT, NODE_FO_STRUCT
                            Select Case HelpObj.Text
                                Case "COBOL"
                                    ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                Case "CHeader"
                                    ShowHelp(modHelp.HHId.H_C_Structure)
                                Case "COBOLIMS"
                                    ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                Case "IMS"
                                    ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                                Case "DDL"
                                    ShowHelp(modHelp.HHId.H_Relational_DDL)
                                Case "DML"
                                    ShowHelp(modHelp.HHId.H_Relational_DML)
                                Case "XMLDTD"
                                    ShowHelp(modHelp.HHId.H_XML_DTD)
                                Case "Unknown"
                                    ShowHelp(modHelp.HHId.H_Structures)
                                Case Else
                                    ShowHelp(modHelp.HHId.H_Structures)
                            End Select
                        Case NODE_STRUCT_FLD
                            ShowHelp(modHelp.HHId.H_Structure_Field_Attributes)
                        Case NODE_STRUCT_SEL, NODE_FO_STRUCT_SEL
                            Select Case CType(HelpObj, clsStructure).StructureType
                                Case enumStructure.STRUCT_COBOL
                                    ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                Case enumStructure.STRUCT_C
                                    ShowHelp(modHelp.HHId.H_C_Structure)
                                Case enumStructure.STRUCT_COBOL_IMS
                                    ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                Case enumStructure.STRUCT_IMS
                                    ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                                Case enumStructure.STRUCT_REL_DDL
                                    ShowHelp(modHelp.HHId.H_Relational_DDL)
                                Case enumStructure.STRUCT_REL_DML
                                    ShowHelp(modHelp.HHId.H_Relational_DML)
                                Case enumStructure.STRUCT_XMLDTD
                                    ShowHelp(modHelp.HHId.H_XML_DTD)
                                Case enumStructure.STRUCT_UNKNOWN
                                    ShowHelp(modHelp.HHId.H_Structures)
                                Case Else
                                    ShowHelp(modHelp.HHId.H_Structures)
                            End Select
                        Case NODE_SYSTEM, NODE_FO_SYSTEM
                            ShowHelp(modHelp.HHId.H_Systems)
                        Case NODE_CONNECTION, NODE_FO_CONNECTION
                            ShowHelp(modHelp.HHId.H_Connections)
                        Case NODE_ENGINE, NODE_FO_ENGINE
                            ShowHelp(modHelp.HHId.H_Engines)
                        Case NODE_SOURCEDATASTORE, NODE_FO_SOURCEDATASTORE, NODE_SOURCEDSSEL
                            ShowHelp(modHelp.HHId.H_Sources)
                        Case NODE_TARGETDATASTORE, NODE_FO_TARGETDATASTORE, NODE_TARGETDSSEL
                            ShowHelp(modHelp.HHId.H_Targets)
                        Case NODE_VARIABLE, NODE_FO_VARIABLE
                            ShowHelp(modHelp.HHId.H_Variables)
                        Case NODE_PROC, NODE_FO_PROC
                            ShowHelp(modHelp.HHId.H_Mapping_Procedures)
                        Case NODE_GEN, NODE_FO_JOIN
                            ShowHelp(modHelp.HHId.H_Joins)
                        Case NODE_LOOKUP, NODE_FO_LOOKUP
                            ShowHelp(modHelp.HHId.H_Lookups)
                        Case NODE_MAIN, NODE_FO_MAIN
                            ShowHelp(modHelp.HHId.H_Main_Procedure)
                        Case NODE_MAPPING, NODE_FO_MAPPING
                            ShowHelp(modHelp.HHId.H_MAP)
                        Case NODE_FUN, NODE_FO_FUNCTION, NODE_FO_FUNCTION_RECENT
                            ShowHelp(modHelp.HHId.H_Function_Reference)
                        Case NODE_TEMPLATE, NODE_FO_TEMPLATE
                            ShowHelp(modHelp.HHId.H_Reference)
                        Case Else
                            ShowHelp(modHelp.HHId.H_DATETIME)
                    End Select
                End If
            End If
        End If

    End Sub

    Private Sub tvExplorer_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles tvExplorer.ItemDrag

        Select Case CType(CType(e.Item, TreeNode).Tag, INode).Type
            Case NODE_VARIABLE, NODE_GEN, NODE_PROC, NODE_LOOKUP, NODE_MAIN
                'Set the drag node and initiate the DragDrop
                DoDragDrop(e.Item, DragDropEffects.Copy)
        End Select

    End Sub

    Private Sub tvExplorer_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvExplorer.DragOver

        Dim sourceNode As TreeNode
        Dim obj As Object, objI As INode

        'Get the TreeView raising the event (incase multiple on form)

        '//dont allow drag and drop from other treeview
        Dim sourceTreeview As TreeView = CType(sender, TreeView)
        If sourceTreeview.Name <> "tvExplorer" Then
            e.Effect = DragDropEffects.None
            Exit Sub
        End If

        'As the mouse moves over nodes, provide feedback to the user
        'by highlighting the node that is the current drop target
        Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
        Dim targetNode As TreeNode = sourceTreeview.GetNodeAt(pt)

        'See if there is a TreeNode being dragged
        If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
            'TreeNode found allow move effect

            sourceNode = e.Data.GetData("System.Windows.Forms.TreeNode")
            obj = sourceNode.Tag
            If Not (obj Is Nothing) Then
                If Not (obj.GetType.GetInterface("INode") Is Nothing) Then
                    objI = obj
                    If objI.Type = CType(targetNode.Tag, INode).Type Then
                        e.Effect = DragDropEffects.Copy
                    Else
                        e.Effect = DragDropEffects.None
                    End If
                End If
            End If
        Else
            'No TreeNode found, prevent move
            e.Effect = DragDropEffects.None
        End If

    End Sub

    Private Sub tvExplorer_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvExplorer.DragDrop

        '// Modified 3/07 by Tom Karasch

        Dim sourceNode As TreeNode
        Dim NewPos As Integer
        Dim SourceIndex, TargetIndex As Integer
        'Get the TreeView raising the event (incase multiple on form)

        '//dont allow drag and drop from other treeview
        Dim sourceTreeview As TreeView = CType(sender, TreeView)

        Try

            If sourceTreeview.Name <> "tvExplorer" Then
                e.Effect = DragDropEffects.None
                Exit Sub
            End If

            tvExplorer.BeginUpdate()

            'As the mouse moves over nodes, provide feedback to the user
            'by highlighting the node that is the current drop target
            Dim pt As Point = CType(sender, TreeView).PointToClient(New Point(e.X, e.Y))
            Dim targetNode As TreeNode = sourceTreeview.GetNodeAt(pt)

            sourceNode = e.Data.GetData("System.Windows.Forms.TreeNode")

            SourceIndex = sourceNode.Index
            TargetIndex = sourceNode.Index

            NewPos = MoveNode(sourceNode, targetNode.Index)

            '//Since we reshuffled nodes we must store New SeqNo
            Dim nd As TreeNode

            For Each nd In sourceNode.Parent.Nodes
                If nd.Tag.SeqNo <> nd.Index Then
                    nd.Tag.SeqNo = nd.Index
                    nd.Tag.SaveSeqNo()
                End If
            Next

            tvExplorer.SelectedNode = sourceNode

        Catch ex As Exception
            LogError(ex)
        Finally
            tvExplorer.EndUpdate()
        End Try

    End Sub

#End Region

    '//////// actions for specific treenode types ////////
#Region "Node Actions"

    Private Function DoProjectAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode '//cNode = collection of treenodes as an array
        Dim obj As clsProject '//Actual/Modified object stored in node.tag
        Dim folderObj As clsFolderNode
        Dim frm1 As frmProjOpen
        Dim frm2 As frmNewProj

        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    If CurLoadedProject IsNot Nothing Then
                        cnn.Close()
                    End If
                    '//If this is a new project build a new tree as collection 

                    frm2 = New frmNewProj
                    obj = frm2.NewProj()

                    If Not (obj Is Nothing) Then
                        cNode = AddTreeNode(tvExplorer, NODE_PROJECT, obj)
                        obj.SeqNo = cNode.Index '//store position in array of nodes
                        If Not (cNode Is Nothing) Then
                            '// set the selected node "project name" to be a 
                            '//collection of it's nodes
                            tvExplorer.SelectedNode = cNode
                        End If
                        '//When we create new project by default add Environment folder node
                        folderObj = New clsFolderNode("Environments", NODE_FO_ENVIRONMENT)
                        folderObj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, folderObj.Type, folderObj)
                    End If

                Case modDeclares.enumAction.ACTION_OPEN
                    If CurLoadedProject IsNot Nothing Then
                        cnn.Close()
                    End If
                    Me.Cursor = Cursors.WaitCursor

                    frm1 = New frmProjOpen
                    obj = frm1.OpenProj()
                    

                    If obj IsNot Nothing Then
                        '// set the current loaded project as the object containing 
                        '//the array of treenode objects
                        obj.LoadMe()
                        CurLoadedProject = obj
                        '// fill the object array with it's child treenodes
                        FillProject(obj)
                    Else
                        Exit Try
                    End If

                    '/// This loads the Mapping pattern list if one exists
                    If Application.UserAppDataRegistry.GetValue(CurLoadedProject.ProjectName & "MapListPath") IsNot Nothing Then
                        CurLoadedProject.MapListPath = Application.UserAppDataRegistry.GetValue(CurLoadedProject.ProjectName & "MapListPath").ToString
                    Else
                        CurLoadedProject.MapListPath = ""
                    End If

                    Me.Cursor = Cursors.Default
                    ToolBar1.Buttons(enumToolBarButtons.TB_MAP).Enabled = True

                Case modDeclares.enumAction.ACTION_DELETE
                    '// if delete is chosen delete the whole array of tree nodes 
                    '//contained in the current project node (array of nodes)
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain DoProjectAction")
            Return False
        Finally
            Me.Cursor = Cursors.Default
        End Try


    End Function

    Private Function DoEnvironmentAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmEnvironment

        cNode = tvExplorer.SelectedNode '// current node in the tree (this case environment)
        Try

            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up if node is not folder 
                    '//i.e. if an existing environment is right clicked on, 
                    '//a new environment should be built under the environments folder, 
                    '//not the existing environment
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmEnvironment
                        '//Go one step up to Project

                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                        obj = frm.NewObj() '//Show the form for new environment entry
                    End If

                    If obj IsNot Nothing Then
                        '// add a node to the parent (project, not the environments folder)
                        cNode = AddNode(cNode, NODE_ENVIRONMENT, obj)
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            '// if it's a valid environment then make it the selected 
                            '// node so that folders can be built under it.
                            tvExplorer.SelectedNode = cNode
                        End If

                        '//Now add two folders under this new node 
                        '//(Connections folder, structure folder, system folder)
                        '//struct folder
                        obj = New clsFolderNode("Connections", NODE_FO_CONNECTION)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//struct folder
                        obj = New clsFolderNode("Descriptions", NODE_FO_STRUCT)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//Datastores folder
                        'obj = New clsFolderNode("Datastores", NODE_FO_DATASTORE)
                        'obj.Parent = CType(cNode.Tag, INode)
                        'AddNode(cNode.Nodes, obj.Type, obj)

                        '//targets folder
                        'obj = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
                        'obj.Parent = CType(cNode.Tag, INode)
                        'AddNode(cNode.Nodes, obj.Type, obj)

                        '//variables folder
                        obj = New clsFolderNode("Variables", NODE_FO_VARIABLE)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//tasks folder
                        'obj = New clsFolderNode("Procedures", NODE_FO_PROC)
                        'obj.Parent = CType(cNode.Tag, INode)
                        'AddNode(cNode.Nodes, obj.Type, obj)

                        '//sys folder
                        obj = New clsFolderNode("Systems", NODE_FO_SYSTEM)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// do delete of ENV and recursive delete of all child objects of ENV
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain DoEnvironmentAction")
            Return False
        End Try

    End Function

    Private Function DoFolderAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        'Dim obj As INode '//Actual/Modified object stored in node tag
        cNode = tvExplorer.SelectedNode
        '// no other actions available to folders except to make it the currently selected node

        Select Case ActionType
            Case modDeclares.enumAction.ACTION_NEW

        End Select

        Return True

    End Function

    Private Function DoVarAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmVariable
        Dim NodeText As String

        cNode = tvExplorer.SelectedNode
        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up to the "variables" 
                    '//folder if mouse was right clicked on an already defined variable
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmVariable
                        '//Go one step up to Engine and make the engine it's parent object
                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)

                        obj = frm.NewObj() '//Show the form for new entry
                        '// open the variables form and create a new variable object
                    End If

                    If Not (obj Is Nothing) Then
                        cNode = AddNode(cNode, NODE_VARIABLE, obj)
                        '// add the new variable object to the tree as a "variable" node
                        obj.SeqNo = cNode.Index '//index it's store position

                        If Not (cNode Is Nothing) Then
                            '// make this node the currently selected node on the tree
                            tvExplorer.SelectedNode = cNode

                        End If
                        NodeText = tvExplorer.SelectedNode.Text
                        'tvExplorer.Refresh()
                        'FillProject(CurLoadedProject, , True)
                        tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, NodeText)
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the variable object, it will have no children
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain DoVarAction")
            Return False
        End Try

    End Function

    Private Function DoTaskAction(ByVal ActionType As enumAction, Optional ByVal TaskType As enumTaskType = modDeclares.enumTaskType.TASK_PROC, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim tempobj As INode
        Dim tempnode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmTask
        Dim typeOfTask As enumTaskType = TaskType
        Dim frmInc As frmInclude
        Dim Parent As INode
        Dim NodeText As String

        Try
            '// make the current node the selected node on the tree
            cNode = tvExplorer.SelectedNode
            tempnode = cNode
            tempobj = cNode.Tag

            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up if node is currently 
                    '//existing procedure and not a main, procedure, join, or lookup folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If
                    '/// now see what folder the user has selected to define a new 
                    '//task and set the task type accordingly
                    Select Case tempobj.Type
                        Case NODE_FO_MAIN
                            TaskType = modDeclares.enumTaskType.TASK_MAIN
                        Case NODE_FO_JOIN
                            TaskType = modDeclares.enumTaskType.TASK_GEN
                        Case NODE_FO_LOOKUP
                            TaskType = modDeclares.enumTaskType.TASK_LOOKUP
                        Case NODE_FO_PROC
                            If typeOfTask <> enumTaskType.TASK_IncProc And typeOfTask <> enumTaskType.TASK_GEN Then
                                TaskType = modDeclares.enumTaskType.TASK_PROC
                            End If
                    End Select

                    If IsClipboardAction = False Then
                        If TaskType = enumTaskType.TASK_IncProc Then
                            frmInc = New frmInclude
                            Parent = CType(cNode.Parent.Tag, INode)
                            obj = frmInc.NewObj(NODE_PROC, Parent)
                        Else
                            frm = New frmTask
                            '//Go one step up to engine and make this a child of engine 
                            '//and not the folder it will reside in, in the tree
                            frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                            '//Show the form for new task entry and make obj the new task object
                            obj = frm.NewObj(cNode.Parent, TaskType)
                        End If
                    End If


                    If Not (obj Is Nothing) Then
                        '// if the task is a valid task then add a node to the tree 
                        '//under the proper folder
                        cNode = AddNode(cNode, obj.Type, obj)
                        '// make the new task object the current object and 
                        '//save the sequence number so it will be displayed 
                        '//properly in the treeview
                        obj.SeqNo = cNode.Index '//store position
                        CType(obj, clsTask).SaveSeqNo()
                        If Not (cNode Is Nothing) Then
                            '// if it's a valid object, 
                            '//then make it the currently selected node in the tree
                            tvExplorer.SelectedNode = cNode
                            NodeText = tvExplorer.SelectedNode.Text
                            'tvExplorer.Refresh()
                            'FillProject(CurLoadedProject, , True)
                            tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, NodeText)
                        End If
                    End If

                Case modDeclares.enumAction.ACTION_CHANGE
                    '//Make parent node to one step up if node is currently 
                    '//existing procedure and not a main, procedure, join, or lookup folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    Else
                        Exit Try
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmTask
                        '//Go one step up to engine and make this a child of engine 
                        '//and not the folder it will reside in, in the tree
                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                        '//Show the form for new task entry and make obj the new task object
                        obj = frm.EditObj(cNode.Parent, CType(tempobj, INode))
                    End If

                    If Not (obj Is Nothing) Then
                        '// if the task is a valid task then update node on the tree 
                        '//under the proper folder
                        UpdateNode(tempnode, obj)
                        '// make the new task object the current object and 
                        '//save the sequence number so it will be displayed 
                        '//properly in the treeview
                        CType(obj, clsTask).SaveSeqNo()

                        If Not (cNode Is Nothing) Then
                            '// if it's a valid object, 
                            '//then make it the currently selected node in the tree
                            tvExplorer.SelectedNode = tempnode
                            ShowUsercontrol(tempnode, True)

                        End If
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the treenode and the object it contains
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoTaskAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoSystemAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmSystem
        '// make the current node the node on tree the user selected
        cNode = tvExplorer.SelectedNode
        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent tree node to one step up if node is a current 
                    '//system and not the systems folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmSystem
                        '// make this new system a child of the environment is resides in
                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                        obj = frm.NewObj() '//Show the form for new entry
                    End If

                    If Not (obj Is Nothing) Then
                        '// if the new system object is valid add the node to 
                        '// the tree as a system node and index it's position
                        cNode = AddNode(cNode, NODE_SYSTEM, obj)
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            '// if it's a valid node then make it the currently 
                            '// selected node on the tree
                            tvExplorer.SelectedNode = cNode
                        End If


                        '// Add engines folder
                        '// make a new engines folder tree node and make it's parent 
                        '// the(New System)
                        obj = New clsFolderNode("Engines", NODE_FO_ENGINE)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoSystemAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoStructureAction(ByVal ActionType As enumAction, Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0, Optional ByVal ByUser As Boolean = False, Optional ByVal fromFile As Boolean = False) As Boolean

        '// Overhauled for DMLs by Tom Karasch  6/07
        Dim cNode As TreeNode
        Dim nd As TreeNode = Nothing
        Dim obj As INode '//Actual/Modified "structure" object stored in node tag
        Dim frm As frmStructure
        Dim frmDMLInf As frmDMLInfo
        Dim i As Integer
        Dim ChgStrObj As clsStructure
        Dim objArr As ArrayList = Nothing
        Dim frmDMLfrm As frmDML
        Dim DMLRetStr As String = ""

        cNode = tvExplorer.SelectedNode
        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '// first make the current tree node the structures folder
                    '//Make parent tree node two steps up if node is an existing structure
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent.Parent '[Structure]->[type folder]->Struct
                        '// make the current tree node one step up if it's a 
                        '//structure "type" folder
                    ElseIf cNode.Text.Contains("Descriptions") = False Then
                        cNode = cNode.Parent '[Structure]->[type folder]
                    End If
                    '/// get the environment to pass into New Structure Form
                    Dim ObjEnv As clsEnvironment = CType(cNode.Parent.Tag, clsEnvironment)

                    If fromFile = True Or (StructType <> enumStructure.STRUCT_REL_DML And _
                                           StructType <> enumStructure.STRUCT_REL_DML_FILE) Then
                        If IsClipboardAction = False Then
                            frm = New frmStructure
                            '// make the new structure a child of the environment
                            '//Go step up to Environment
                            frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                            '// make a new structure object in the structures object array 
                            '//and put it in the proper structures "type" folder
                            objArr = frm.NewObj(StructType, , ObjEnv)  '//Show the form for new entry
                        End If

                        '// build proper structure and type folder nodes as needed to 
                        '//store the new structure and put the new structure in the 
                        '//proper place in the tree node structure
                        If objArr IsNot Nothing Then
                            For i = 0 To objArr.Count - 1
                                obj = objArr(i)
                                If obj IsNot Nothing Then
                                    nd = AddStructNode(obj, cNode)
                                    obj.SeqNo = nd.Index '//store position
                                End If
                            Next
                            '// make the new structure the currently selected node in the tree
                            If nd IsNot Nothing Then
                                tvExplorer.SelectedNode = nd
                            End If
                        End If
                        '********************
                    Else '****  DML  ********
                        '********************

                        '/// DML Modeling here
                        If ByUser = False Then
                            '/// Get the environment so we can see if previous DML were created so 
                            '/// Repeated ODBC login is not required if the Source DSN doesn't change
                            If IsClipboardAction = False Then
                                frmDMLInf = New frmDMLInfo
                                Dim ObjDMLArray As New Collection

                                '/// Get DML Object from form 
                                ObjDMLArray = frmDMLInf.GetInfo(ObjEnv)
                                If ObjDMLArray Is Nothing Then Exit Select

                                If ObjDMLArray.Count > 0 Then
                                    '/// if the form data is valid, add the new structure
                                    For Each NewDMLinfo As clsDMLinfo In ObjDMLArray
                                        Dim objstr As New clsStructure
                                        '/// Create New Structure Object from DML Object Data, and read in Field Attributes
                                        objstr = MakeDMLStructure(NewDMLinfo, ObjEnv)
                                        '/// Validate the new Structure and add it to the project
                                        If objstr IsNot Nothing Then
tryAgain:                                   If objstr.ValidateNewObject() = False Then
                                                objstr.StructureName = InputBox("Enter different Description Name ", "Duplicate Description Name", objstr.Text)
                                                If objstr.StructureName <> "" Then
                                                    GoTo tryAgain
                                                End If
                                            Else
                                                '/// Add the structure and add it to the tree
                                                If objstr.AddNew() = True Then
                                                    nd = AddStructNode(objstr, cNode)
                                                    objstr.SeqNo = nd.Index '//store position
                                                    '// Add DMLobj to Environment
                                                    ObjEnv.OldDMLobj = NewDMLinfo
                                                    tvExplorer.SelectedNode = nd
                                                    ShowUsercontrol(tvExplorer.SelectedNode, True)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Else
                            '/// New User Defined DML Here
                            If IsClipboardAction = False Then
                                '///Open DML form and get User's Filepath that has their SQL Statement
                                '///defining the new DML structure
                                frmDMLfrm = New frmDML
                                DMLRetStr = frmDMLfrm.NewObj(ObjEnv)
                                '/// send this file path into the new structure form so form 
                                '///will generate new XML structure file to be read in by SQDuiimp.exe
                                If DMLRetStr IsNot Nothing Then
                                    '// open the new structure form
                                    frm = New frmStructure
                                    '// make the new structure a child of the environment
                                    '//Go step up to Environment
                                    frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                                    '// make a new structure object in the structures object array 
                                    '//and put it in the proper structures "type" folder
                                    objArr = frm.NewObj(StructType, DMLRetStr, ObjEnv)  '//Show the form for new entry
                                End If
                            End If

                            '// build proper structure and type folder nodes as needed to 
                            '//store the new structure and put the new structure in the 
                            '//proper place in the tree node structure
                            If objArr IsNot Nothing Then
                                For i = 0 To objArr.Count - 1
                                    obj = objArr(i)
                                    If obj IsNot Nothing Then
                                        nd = AddStructNode(obj, cNode)
                                        obj.SeqNo = nd.Index '//store position
                                    End If
                                Next
                                '// make the new structure the currently selected node in the tree
                                If nd IsNot Nothing Then
                                    tvExplorer.SelectedNode = nd
                                End If
                            End If
                        End If
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the selected structure from the tree
                    DoAction(modDeclares.enumAction.ACTION_DELETE)

                Case enumAction.ACTION_CHANGE
                    If CType(cNode.Tag, INode).Type IsNot NODE_STRUCT Then
                        Exit Function
                    Else
                        frm = New frmStructure
                        '// make a new structure object in the structures object array 
                        '//and put it in the proper structures "type" folder
                        ChgStrObj = frm.ChangeObj(CType(cNode.Tag, clsStructure))  '//Show the form for new entry
                        '// attach new structure pointer to Node Tag
                        If ChgStrObj IsNot Nothing Then
                            cNode.Tag = ChgStrObj
                            '// Reload Project to reflect MetaData Changes
                            FillProject(ChgStrObj.Project, False, True)
                        End If
                    End If
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoStructureAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoStructureSelectionAction(ByVal ActionType As enumAction, Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmStructSelection

        '// Make the current node the user selectexd node in the tree
        cNode = tvExplorer.SelectedNode
        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up if node is selection. 
                    '//We add selection under the structure that is the parent of the selection
                    If CType(cNode.Tag, INode).Type = NODE_STRUCT_SEL Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmStructSelection
                        '//Make the new object's parent, the current node 
                        '//(the base structure for the selection)
                        frm.objThis.Parent = CType(cNode.Tag, INode)
                        '// open new StrSel Form and take user entry and set it to obj
                        obj = frm.NewObj()
                    End If

                    If Not (obj Is Nothing) Then
                        '// If the user entered a valid str sel, then make the new 
                        '//tree node the current node and put object "obj" into it. 
                        '//Then index it in the tree, and make it the currently 
                        '//selected tree node
                        cNode = AddNode(cNode, NODE_STRUCT_SEL, obj)
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            tvExplorer.SelectedNode = cNode
                        End If
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the structure selection and the tree node and update the tree
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoStructureSelectionAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoConnectionAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmConnection
        '// make the current node the user selected connection node in the tree
        cNode = tvExplorer.SelectedNode

        Try

            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '// make the current node the system object 
                    '//Make parent node to one step up if node is a current "connection" object
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmConnection
                        '//Go one step up to system
                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                        '// open the connections form and make obj the new connection object
                        obj = frm.NewObj() '//Show the form for new entry
                    End If

                    '// if it's a valid connection object, then give it the proper 
                    '//index in the tree structure and add the node to the tree, 
                    '//then make it the selected node in the tree
                    If obj IsNot Nothing Then
                        cNode = AddNode(cNode, NODE_CONNECTION, obj)
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            '// make the new engine the currently selected tree node
                            tvExplorer.SelectedNode = cNode
                        End If
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the current connection object and it's node in the tree
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoConnectionAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoEngineAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim cNode As TreeNode
        Dim obj As INode = Nothing '//Actual/Modified object stored in node tag
        Dim frm As frmEngine

        '// make the current node the user selected engine node in the tree
        cNode = tvExplorer.SelectedNode
        Try
            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up if node is an existing engine tree node
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then
                        frm = New frmEngine
                        '// make a new engine object and make it a child of the system 
                        '//it's created under
                        'objTopMost = ctEng.newobj()

                        frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                        obj = frm.NewObj() '//Show the form for new entry
                    End If

                    If Not (obj Is Nothing) Then
                        '// add the new engine node to the tree and index it's position
                        cNode = AddNode(cNode, NODE_ENGINE, obj)
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            '// make the new engine the currently selected tree node
                            tvExplorer.SelectedNode = cNode
                        End If

                        '//Now add 6 folders under this new node 
                        '//(sources, targets, variables, main, procedures, joins, 
                        '//and lookups folders)

                        '// make these new folders tree node objects and make 
                        '//them all children of the engine in the tree

                        '//sources folder
                        obj = New clsFolderNode("Sources", NODE_FO_SOURCEDATASTORE)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//targets folder
                        obj = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//variables folder
                        obj = New clsFolderNode("Variables", NODE_FO_VARIABLE)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//tasks folder
                        'obj = New clsFolderNode("Joins", NODE_FO_JOIN)
                        'obj.Parent = CType(cNode.Tag, INode)
                        'AddNode(cNode.Nodes, obj.Type, obj)

                        '//tasks folder
                        obj = New clsFolderNode("Procedures", NODE_FO_PROC)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                        '//tasks folder
                        obj = New clsFolderNode("Main", NODE_FO_MAIN)
                        obj.Parent = CType(cNode.Tag, INode)
                        AddNode(cNode, obj.Type, obj)

                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the engine object and 
                    '// recursively delete all tree nodes and objects under it
                    DoAction(modDeclares.enumAction.ACTION_DELETE)

                Case enumAction.ACTION_MERGE
                    '// Merge two engines together
                    If m_ClipObjects.Count > 0 Then

                    End If
                    MsgBox("Merge with this engine", MsgBoxStyle.OkOnly)



            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "DoEngineAction=>" & ex.Message)
            Return False
        End Try

    End Function

    Private Function DoDatastoreAction(ByVal ActionType As enumAction, Optional ByVal DatastoreType As enumDatastore = modDeclares.enumDatastore.DS_UNKNOWN, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0, Optional ByVal IsLU As Boolean = False) As Boolean

        '/// Modified numerous times by Tom Karasch

        Dim cNode As TreeNode
        Dim obj As clsDatastore = Nothing '//Actual/Modified object stored in node tag
        Dim selobj As clsDSSelection
        Dim frm As frmDatastore
        Dim frmInc As frmInclude
        Dim Parent As INode
        Dim NodeText As String

        '// make the current node the user selected node in the tree
        cNode = tvExplorer.SelectedNode

        '// FIRST see if the selected treenode is a Datastore Selection
        '// if it is, edit the datastore it belongs to
        If CType(cNode.Tag, INode).Type = NODE_SOURCEDSSEL Or CType(cNode.Tag, INode).Type = NODE_TARGETDSSEL Then
            selobj = cNode.Tag
            cNode = selobj.ObjDatastore.ObjTreeNode
        End If

        '// see if it is a source or target datastore 
        Dim NodeType As String
        Dim Direction As String
        Try
            If CType(cNode.Tag, INode).Type = NODE_FO_SOURCEDATASTORE Or _
               CType(cNode.Tag, INode).Type = NODE_SOURCEDATASTORE Then

                NodeType = NODE_SOURCEDATASTORE
                Direction = DS_DIRECTION_SOURCE

            Else
                'CType(cNode.Tag, INode).Type = NODE_FO_TARGETDATASTORE Or _
                '   CType(cNode.Tag, INode).Type = NODE_TARGETDATASTORE Then

                Direction = DS_DIRECTION_TARGET
                NodeType = NODE_TARGETDATASTORE

                'Else
                '    Direction = ""
                '    NodeType = GetFolType(DatastoreType)
            End If

            Select Case ActionType
                Case modDeclares.enumAction.ACTION_NEW
                    '//Make parent node to one step up if node is not folder
                    If CType(cNode.Tag, INode).IsFolderNode = True Then
                        If cNode.Parent.Text = "Datastores" Then
                            cNode = cNode.Parent
                        End If
                    Else
                        cNode = cNode.Parent
                    End If

                    If IsClipboardAction = False Then

                        If DatastoreType = enumDatastore.DS_INCLUDE Then
                            frmInc = New frmInclude
                            Parent = CType(cNode.Parent.Tag, INode)
                            obj = frmInc.NewObj(NodeType, Parent)
                        Else
                            frm = New frmDatastore
                            frm.objThis.DsDirection = Direction

                            '// add direction as source or target to this datastore object
                            '// make this node a child of the engine it it created in
                            '//Go one step up to engine
                            frm.objThis.Parent = CType(cNode.Parent.Tag, INode)
                            '// make new object open a new datastore form and get 
                            '//user input and store it as the new object
                            obj = frm.NewObj(cNode, DatastoreType, IsLU) '//Show the form for new entry
                        End If

                    End If

                    If Not (obj Is Nothing) Then
                        '//If the new datastore is valid, add it as a child 
                        '//of the engine and index it properly for it's position in the tree
                        'If obj.Engine IsNot Nothing Then
                        cNode = AddNode(cNode, NodeType, obj, , obj.DsPhysicalSource)
                        'Else
                        '    cNode = AddDSNode(obj, cNode)
                        'End If

                        obj.ObjTreeNode = cNode
                        obj.SeqNo = cNode.Index '//store position
                        If Not (cNode Is Nothing) Then
                            tvExplorer.SelectedNode = cNode
                        End If
                        '// Now add it's Selections to the tree
                        AddDSstructuresToTree(cNode, obj)
                        'NodeText = tvExplorer.SelectedNode.Text
                        NodeText = obj.DsPhysicalSource
                        'tvExplorer.Refresh()
                        'FillProject(CurLoadedProject, , True)
                        tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, NodeText)
                    End If

                Case modDeclares.enumAction.ACTION_DELETE
                    '// delete the currently selected datastore and remove it's node from the tree
                    DoAction(modDeclares.enumAction.ACTION_DELETE)
            End Select

            Return True

        Catch ex As Exception
            Log("DoDatastoreAction=>" & ex.Message)
            Return False
        End Try

    End Function

#End Region

    '//////// Filling the treenodes from MetaDataDB by node type ////////
#Region "Fill Tree"

    Function FillProject(ByVal obj As clsProject, Optional ByVal Renamed As Boolean = False, Optional ByVal Updated As Boolean = False) As Boolean
        '// Modified 2/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp")
        Dim i As Integer
        Dim sql As String
        Dim MsgClear As Boolean

        Try
            cnn = New System.Data.Odbc.OdbcConnection(obj.Project.MetaConnectionString)
            cnn.Open()
            Dim cNode As TreeNode

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            'obj.LoadMe()

            '// Optional Passed Vars added By Tom Karasch for change(i.e. delete, etc...) and rename
            '/// to reload project according to what occured in the object tree.
            If Renamed = False Then
                If Updated = False Then  '// opening new project
                    '//Add project node
                    cNode = AddTreeNode(tvExplorer, NODE_PROJECT, obj)
                    obj.SeqNo = cNode.Index '//store position
                    IsEventFromCode = True
                    SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    'Me.WindowState = CType(obj, clsProject).WinState
                    IsEventFromCode = False
                    MsgClear = True
                Else   '// Reloading project and tree to reflect MetaData Changes
                    cNode = obj.ObjTreeNode
                    cNode.Nodes.Clear()
                    IsEventFromCode = True
                    SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    IsEventFromCode = False
                    MsgClear = False
                End If
            Else
                '// Obj is existing project with renamed objects so refill with name changes
                cNode = obj.ObjTreeNode
                cNode.Nodes.Clear()
                cNode.Text = obj.Project.ProjectName
                'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                CType(obj, clsProject).Save()
                CType(obj, clsProject).SaveToRegistry()
                MsgClear = False
            End If

            ShowUsercontrol(cNode, MsgClear)
            tvExplorer.SelectedNode = cNode

            '//Now add and process each environment 
            'If obj.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblEnvironments & " where ProjectName=" & obj.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,ENVIRONMENTDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblEnvironments & " where ProjectName=" & obj.GetQuotedText
            Log(sql)
            'End If

            'Log(obj.Project.MetaConnectionString)    //// commented so Password is not in the log
            'Log("cnn.connectionstring >> " & cnn.ConnectionString)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            '//Add Environment folder node
            Dim objF As clsFolderNode
            objF = New clsFolderNode("Environments", NODE_FO_ENVIRONMENT)
            objF.Parent = CType(cNode.Tag, INode)
            cNode = AddNode(cNode, objF.Type, objF)
            objF.SeqNo = cNode.Index '//store position

            For i = 0 To dt.Rows.Count - 1
                '//Process this env
                FillEnv(cNode, dt.Rows(i), cnn)
            Next

            If obj.ENV_FOexpanded = True Then
                cNode.Expand()
            Else
                cNode.Collapse()
            End If
            tvExplorer.Refresh()

            cNode = cNode.Parent

        Catch ex As Exception
            LogError(ex, "fillProject")
        Finally
            ShowStatusMessage("")
        End Try

    End Function

    Function FillEnv(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim i As Integer

        Dim obj As New clsEnvironment
        Dim sql As String = ""
        Dim Dir As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent '//Project->[Env Folder]->Env
            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
            '    obj.LocalDTDDir = GetVal(dr.Item("LocalDTDDir"))
            '    obj.LocalDDLDir = GetVal(dr.Item("LocalDDLDir"))
            '    obj.LocalCobolDir = GetVal(dr.Item("LocalCobolDir"))
            '    obj.LocalCDir = GetVal(dr.Item("LocalCDir"))
            '    obj.LocalScriptDir = GetVal(dr.Item("LocalScriptDir"))
            '    obj.LocalDBDDir = GetVal(dr.Item("LocalModelDir"))

            '    obj.EnvironmentDescription = GetVal(dr.Item("Description"))

            'Else
            obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
            obj.EnvironmentDescription = GetVal(dr.Item("ENVIRONMENTDESCRIPTION"))
            'End If


            AddToCollection(obj.Project.Environments, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            '/// Added 6/4/07 by TK
            obj.GetStructureDir()

            If CurLoadedProject.ENVexpanded = True Then
                cNode.Expand()
            Else
                cNode.Collapse()
            End If

            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "fillEnv")
        End Try
        '///////////////////////////////////////////////
        '//Now add connections
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under Env
            Dim cNodeCnn As TreeNode
            Dim objCnn As INode

            objCnn = New clsFolderNode("Connections", NODE_FO_CONNECTION)
            objCnn.Parent = CType(cNode.Tag, INode)
            cNodeCnn = AddNode(cNode, objCnn.Type, objCnn)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & objCnn.Project.tblConnections & " where EnvironmentName=" & obj.GetQuotedText & " AND ProjectName=" & obj.Project.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,CONNECTIONNAME,CONNECTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objCnn.Project.tblConnections & _
                        " where ProjectName=" & obj.Project.GetQuotedText & _
                        " AND EnvironmentName =" & obj.GetQuotedText

            Log(sql)
            'End If


            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillConn(cNodeCnn, dt.Rows(i), cnn)
            Next

            If CurLoadedProject.CONNexpanded = True Then
                cNodeCnn.Expand()
            Else
                cNodeCnn.Collapse()
            End If

            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "fillConn")
        End Try
        '///////////////////////////////////////////////
        '//Now add structures
        '///////////////////////////////////////////////
        Try
            '//Add Structure folder node under env
            Dim cNodeStruct As TreeNode
            Dim objStruct As INode
            objStruct = New clsFolderNode("Descriptions", NODE_FO_STRUCT)
            objStruct.Parent = CType(cNode.Tag, INode)
            cNodeStruct = AddNode(cNode, objStruct.Type, objStruct)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblStructures & _
            '    " where EnvironmentName=" & obj.GetQuotedText & _
            '    " AND ProjectName=" & obj.Project.GetQuotedText & _
            '    " order by structuretype, structureName"
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONTYPE,DESCRIPTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblDescriptions & _
            " where ProjectName=" & obj.Project.GetQuotedText & _
            " AND EnvironmentName = " & obj.GetQuotedText & _
            " order by DESCRIPTIONTYPE, DescriptionName"
            Log(sql)
            'End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New System.Data.DataTable("temp2")

            da.Fill(dt)
            da.Dispose()

            'cNodeStruct.Text = "(" & dt.Rows.Count.ToString & ")" & " Descriptions"

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillStruct(cNodeStruct, dt.Rows(i), cnn)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.STRexpanded = True Then
                cNodeStruct.Expand()
            Else
                cNodeStruct.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "fillEnv-fillStr")
        End Try

        '///////////////////////////////////////////////
        '//Now All datastores put into folders by type
        '///////////////////////////////////////////////
        'Try
        '    Dim cNode1 As TreeNode
        '    Dim obj1 As INode

        '    obj1 = New clsFolderNode("Datastores", NODE_FO_DATASTORE)
        '    obj1.Parent = CType(cNode.Tag, INode)
        '    cNode1 = AddNode(cNode.Nodes, obj1.Type, obj1, False)

        '    If obj1.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
        '        sql = "Select * from " & obj.Project.tblDatastores & _
        '        " where EnvironmentName=" & obj.GetQuotedText & _
        '        " AND ProjectName=" & obj.Project.GetQuotedText & _
        '        " AND SYSTEMNAME=" & Quote(DBNULL) & _
        '        " AND ENGINENAME=" & Quote(DBNULL) & _
        '        " order by DATASTORETYPE"
        '    Else
        '        sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
        '        " where projectname=" & Quote(obj.Project.ProjectName) & _
        '        " and environmentname=" & Quote(obj.EnvironmentName) & _
        '        " AND SYSTEMNAME=" & Quote(DBNULL) & _
        '        " AND ENGINENAME=" & Quote(DBNULL) & _
        '        " order by dstype"
        '    End If

        '    Log(sql)

        '    da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
        '    dt = New DataTable("temp")
        '    da.Fill(dt)
        '    da.Dispose()

        '    For i = 0 To dt.Rows.Count - 1
        '        ''//Process this (cNode1 is root node under which we add other structures)
        '        FillDSbyType(cNode1, dt.Rows(i), cnn, obj1.Project.ProjectMetaVersion)
        '    Next

        'Catch ex As Exception
        '    LogError(ex, "frmMain FillEnv_DS")
        'End Try

        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add variables folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode, objVar.Type, objVar, False)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from variables where EngineName=" & obj.GetQuotedText & _
            '    " AND ENGINENAME=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & obj.GetQuotedText & _
            '    " AND ProjectName=" & obj.Project.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables " & _
                        "where EnvironmentName=" & obj.GetQuotedText & _
                        " AND ProjectName=" & obj.Project.GetQuotedText & _
                        " AND SYSTEMNAME=" & Quote(DBNULL) & _
                        " AND ENGINENAME=" & Quote(DBNULL) & _
                        " ORDER BY VARIABLENAME"
            Log(sql)
            'End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVar(cNodeVar, dt.Rows(i), cnn, True)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.VARtopExpanded = True Then
                cNodeVar.Expand()
            Else
                cNodeVar.Collapse()
            End If

            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>VariableLoad")
        End Try

        '///////////////////////////////////////////////
        '//Now add procs
        '///////////////////////////////////////////////
        'Try
        '    'Dim cNodeMain As TreeNode
        '    'Dim cNodeLookup As TreeNode
        '    'Dim cNodeJoin As TreeNode
        '    Dim cNodeProc As TreeNode
        '    'Dim objMain As INode
        '    'Dim objLook As INode
        '    'Dim objJoin As INode
        '    Dim objProc As INode
        '    'Dim ttype As enumTaskType

        '    '//Add Join folder
        '    'objJoin = New clsFolderNode("Join", NODE_FO_JOIN)
        '    'objJoin.Parent = CType(cNode.Tag, INode)
        '    'cNodeJoin = AddNode(cNode.Nodes, objJoin.Type, objJoin, False)

        '    '//Add Proc folder
        '    objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
        '    objProc.Parent = CType(cNode.Tag, INode)
        '    cNodeProc = AddNode(cNode.Nodes, objProc.Type, objProc, False)

        '    '//Add Lookup folder
        '    'objLook = New clsFolderNode("Lookup", NODE_FO_LOOKUP)
        '    'objLook.Parent = CType(cNode.Tag, INode)
        '    'cNodeLookup = AddNode(cNode.Nodes, objLook.Type, objLook, False)

        '    '//Add Main folder
        '    'objMain = New clsFolderNode("Main", NODE_FO_MAIN)
        '    'objMain.Parent = CType(cNode.Tag, INode)
        '    'cNodeMain = AddNode(cNode.Nodes, objMain.Type, objMain, False)
        '    dt = New DataTable("temp")

        '    If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
        '        sql = "Select * from " & obj.Project.tblTasks & _
        '        " where EnvironmentName = " & obj.GetQuotedText & _
        '        " AND ProjectName=" & obj.Project.GetQuotedText & _
        '        " AND systemName=" & Quote(DBNULL) & _
        '        " AND enginename=" & Quote(DBNULL) & _
        '        " Order by taskname"
        '    Else
        '        sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,TASKNAME,TASKTYPE,TASKSEQNO,TASKDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblTasks & _
        '        " where EnvironmentName = " & obj.GetQuotedText & _
        '        " AND ProjectName=" & obj.Project.GetQuotedText & _
        '        " AND systemName=" & Quote(DBNULL) & _
        '        " AND enginename=" & Quote(DBNULL) & _
        '        " Order by taskname"
        '    End If

        '    Log(sql)
        '    da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

        '    da.Fill(dt)
        '    da.Dispose()

        '    For i = 0 To dt.Rows.Count - 1
        '        ''//Process this (cNode is root node under which we add other nodes)
        '        'ttype = GetVal(dt.Rows(i).Item("TaskType"))
        '        'Select Case ttype
        '        '    'Case enumTaskType.TASK_MAIN
        '        '    '    FillTasks(cNodeMain, dt.Rows(i), cnn)
        '        '    Case enumTaskType.TASK_JOIN, enumTaskType.TASK_LOOKUP, enumTaskType.TASK_PROC, _
        '        '    enumTaskType.TASK_IncProc, enumTaskType.TASK_MAIN
        '        FillTasks(cNodeProc, dt.Rows(i), cnn)
        '        'Case enumTaskType.TASK_JOIN, enumTaskType.TASK_LOOKUP
        '        '    FillTasks(cNodeJoin, dt.Rows(i), cnn)
        '        'Case 
        '        '    FillTasks(cNodeLookup, dt.Rows(i), cnn)
        '        'End Select
        '    Next

        'Catch ex As Exception
        '    LogError(ex, "FillEnv-FillProc")
        'End Try


        '///////////////////////////////////////////////
        '//Now add systems
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Systems", NODE_FO_SYSTEM)
            objSys.Parent = CType(cNode.Tag, INode)
            cNodeSys = AddNode(cNode, objSys.Type, objSys)

            dt = New DataTable("temp")

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblSystems & _
            '    " where EnvironmentName=" & obj.GetQuotedText & _
            '    " AND ProjectName=" & obj.Project.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,SYSTEMDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblSystems & " where ProjectName=" & obj.Project.GetQuotedText & _
                        " AND EnvironmentName=" & obj.GetQuotedText
            Log(sql)
            'End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillSys(cNodeSys, dt.Rows(i), cnn)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.SYS_FOexpanded = True Then
                cNodeSys.Expand()
            Else
                cNodeSys.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "FillEnv-FillSys")
        End Try

    End Function

    Function FillStruct(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsStructure
        Dim sql As String = ""
        '///////////////////////////////////////////////
        '//Construct structure object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent
            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.StructureName = GetVal(dr.Item("StructureName"))
            '    obj.StructureType = GetVal(dr.Item("StructureType"))
            '    obj.fPath1 = GetVal(dr.Item("fPath1"))
            '    obj.fPath2 = GetVal(dr.Item("fPath2"))
            '    obj.IMSDBName = GetVal(dr.Item("IMSDBName"))
            '    obj.SegmentName = GetVal(dr.Item("SegmentName"))
            '    obj.StructureDescription = GetVal(dr.Item("Description"))
            '    obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->StructFolder->Struct
            '    Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
            '    If ConnName <> "" Then
            '        For Each conn As clsConnection In obj.Environment.Connections
            '            If conn.Text = ConnName Then
            '                obj.Connection = conn
            '            End If
            '        Next
            '    End If
            'Else
            obj.StructureName = GetVal(dr.Item("DescriptionName"))
            obj.StructureDescription = GetStr(GetVal(dr.Item("DescriptionDescription")))
            obj.StructureType = GetVal(dr.Item("DescriptionTYPE"))
            'End If


            AddToCollection(obj.Environment.Structures, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddStructNode(obj, cNode)

        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

        '///////////////////////////////////////////////
        '//Now add fieldselections
        '///////////////////////////////////////////////
        Try
            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblStructSel & " where StructureName=" & obj.GetQuotedText & _
            '    " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & _
            '    obj.Environment.Project.GetQuotedText & " order by selectionname"
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,SELECTIONNAME,ISSYSTEMSEL,SELECTDESCRIPTION from " & obj.Project.tblDescriptionSelect & _
            " where ProjectName=" & obj.Project.GetQuotedText & _
            " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
            " AND DescriptionName = " & obj.GetQuotedText & _
            " order by selectionname"
            'End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode is root node under which we add other nodes)
                FillStructSel(cNode, dt.Rows(i), cnn)
                tvExplorer.Refresh()
            Next

        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

    End Function

    Function FillStructSel(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsStructureSelection

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.ObjStructure = CType(cNode.Tag, INode)
            'If obj.ObjStructure.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSELECT"))
            '    obj.SelectionName = GetVal(dr.Item("SelectionName"))
            '    obj.SelectionDescription = GetVal(dr.Item("Description"))
            'Else
            obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSEL"))
            obj.SelectionName = GetVal(dr.Item("SelectionName"))
            obj.SelectionDescription = GetVal(dr.Item("SELECTDescription"))
            'End If

            AddToCollection(obj.ObjStructure.StructureSelections, obj, obj.GUID)

            If obj.IsSystemSelection = "1" Then
                obj.ObjStructure.SysAllSelection = obj
                Exit Function '//dont add this node if system selection
            End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            cNode.Collapse()

        Catch ex As Exception
            LogError(ex, "fillstrSel")
        End Try

    End Function

    Function FillSys(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsSystem
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.SystemName = GetVal(dr.Item("SystemName"))
            '    obj.Port = GetVal(dr.Item("Port"))
            '    obj.Host = GetVal(dr.Item("Host"))
            '    obj.QueueManager = GetVal(dr.Item("QMgr"))
            '    obj.OSType = GetVal(dr.Item("OSType"))

            '    obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
            '    obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
            '    obj.DTDLib = GetVal(dr.Item("DTDLib"))
            '    obj.DDLLib = GetVal(dr.Item("DDLLib"))

            '    obj.SystemDescription = GetVal(dr.Item("Description"))
            'Else
            obj.SystemName = GetVal(dr.Item("SystemName"))
            obj.SystemDescription = GetVal(dr.Item("SystemDescription"))
            'End If

            AddToCollection(obj.Environment.Systems, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "fillSys")
        End Try

        '///////////////////////////////////////////////
        '//Now add engines
        '///////////////////////////////////////////////
        Try
            '//Add engines folder node under env
            Dim cNodeFOENG As TreeNode
            Dim objENG As INode
            objENG = New clsFolderNode("Engines", NODE_FO_ENGINE)
            objENG.Parent = CType(cNode.Tag, INode)
            cNodeFOENG = AddNode(cNode, objENG.Type, objENG)
            tvExplorer.Refresh()

            dt = New DataTable("temp")

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & objSys.Project.tblEngines & " where SystemName=" & obj.GetQuotedText & " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & obj.Environment.Project.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,ENGINEDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objENG.Project.tblEngines & _
                        " where ProjectName=" & obj.Project.GetQuotedText & _
                        " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
                        " AND SystemName=" & obj.GetQuotedText
            Log(sql)
            'End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillEngine(cNodeFOENG, dt.Rows(i), cnn)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.ENG_FOexpanded = True Then
                cNodeFOENG.Expand()
            Else
                cNodeFOENG.Collapse()
            End If
            tvExplorer.Refresh()

            If CurLoadedProject.SYSexpanded = True Then
                cNode.Expand()
            Else
                cNode.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "fillEng")
        End Try

    End Function

    Function FillEngine(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer
        Dim Dir As String = ""
        Dim obj As New clsEngine
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.EngineName = GetVal(dr.Item("EngineName"))
            obj.ObjSystem = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.ReportEvery = GetVal(dr.Item("ReportEvery"))
            '    obj.CommitEvery = GetVal(dr.Item("CommitEvery"))
            '    obj.ReportFile = GetVal(dr.Item("ReportFile"))
            '    obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
            '    obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
            '    obj.DTDLib = GetVal(dr.Item("DTDLib"))
            '    obj.DDLLib = GetVal(dr.Item("DDLLib"))
            '    obj.EngineDescription = GetVal(dr.Item("Description"))
            '    Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
            '    If ConnName <> "" Then
            '        For Each conn As clsConnection In obj.ObjSystem.Environment.Connections
            '            If conn.Text = ConnName Then
            '                obj.Connection = conn
            '            End If
            '        Next
            '    End If
            '    obj.LoadEngProps()
            'Else
            obj.EngineDescription = GetVal(dr.Item("EngineDescription"))
            'End If

            AddToCollection(obj.ObjSystem.Engines, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now source datastore
        '///////////////////////////////////////////////
        Try
            Dim cNode1 As TreeNode
            Dim obj1 As INode

            obj1 = New clsFolderNode("Sources", NODE_FO_SOURCEDATASTORE)
            obj1.Parent = CType(cNode.Tag, INode)
            cNode1 = AddNode(cNode, obj1.Type, obj1, False)


            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='S' AND EngineName=" & _
            '    obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
            '    obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            'Else
            sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
            " where projectname=" & Quote(obj.Project.ProjectName) & _
            " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
            " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
            " and enginename=" & Quote(obj.EngineName) & _
            " and DSDIRECTION='S'" & _
            " order by datastorename"

            Dir = "S"
            'End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            '/// Add Count to Folder Title
            cNode1.Text = "(" & dt.Rows.Count.ToString & ")" & " Sources"
            tvExplorer.Refresh()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStore(cNode1, dt.Rows(i), cnn, Dir)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.SRCexpanded = True Then
                cNode1.Expand()
            Else
                cNode1.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add targets
        '///////////////////////////////////////////////
        Try
            Dim cNode2 As TreeNode
            Dim obj2 As INode


            obj2 = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
            obj2.Parent = CType(cNode.Tag, INode)
            cNode2 = AddNode(cNode, obj2.Type, obj2, False)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='T' AND EngineName=" & _
            '    obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
            '    obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            'Else
            sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
            " where projectname=" & Quote(obj.Project.ProjectName) & _
            " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
            " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
            " and enginename=" & Quote(obj.EngineName) & _
            " and DSDIRECTION='T'" & _
            " order by datastorename"

            Dir = "T"
            'End If
            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            '/// Add Count to Folder Title
            cNode2.Text = "(" & dt.Rows.Count.ToString & ")" & " Targets"
            tvExplorer.Refresh()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStore(cNode2, dt.Rows(i), cnn, Dir)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.TGTexpanded = True Then
                cNode2.Expand()
            Else
                cNode2.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, sql)
        End Try

        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode, objVar.Type, objVar, False)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.Project.GetQuotedText
            Log(sql)
            'End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            cNodeVar.Text = "(" & dt.Rows.Count.ToString & ")" & " Variables"
            tvExplorer.Refresh()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVar(cNodeVar, dt.Rows(i), cnn)
                tvExplorer.Refresh()
            Next

            If CurLoadedProject.VARexpanded = True Then
                cNodeVar.Expand()
            Else
                cNodeVar.Collapse()
            End If
            tvExplorer.Refresh()

        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>VariableLoad")
        End Try

        '///////////////////////////////////////////////
        '//Now add tasks (main, join, lookup, proc)
        '///////////////////////////////////////////////
        Try
            Dim cNodeMain As TreeNode
            'Dim cNodeLookup As TreeNode
            'Dim cNodeJoin As TreeNode
            Dim cNodeProc As TreeNode
            Dim objMain As INode
            'Dim objLook As INode
            'Dim objGen As INode
            Dim objProc As INode
            Dim ttype As enumTaskType

            '//Add Join folder
            'objJoin = New clsFolderNode("Join", NODE_FO_JOIN)
            'objJoin.Parent = CType(cNode.Tag, INode)
            'cNodeJoin = AddNode(cNode.Nodes, objJoin.Type, objJoin, False)

            '//Add Proc folder
            objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
            objProc.Parent = CType(cNode.Tag, INode)
            cNodeProc = AddNode(cNode, objProc.Type, objProc, False)
            tvExplorer.Refresh()

            '//Add Lookup folder
            'objLook = New clsFolderNode("Lookup", NODE_FO_LOOKUP)
            'objLook.Parent = CType(cNode.Tag, INode)
            'cNodeLookup = AddNode(cNode.Nodes, objLook.Type, objLook, False)

            '//Add Main folder
            objMain = New clsFolderNode("Main Procedure(s)", NODE_FO_MAIN)
            objMain.Parent = CType(cNode.Tag, INode)
            cNodeMain = AddNode(cNode, objMain.Type, objMain, False)
            tvExplorer.Refresh()


            dt = New DataTable("temp")

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    If obj.ObjSystem IsNot Nothing Then
            '        sql = "Select * from " & obj.Project.tblTasks & _
            '        " where EngineName=" & obj.GetQuotedText & _
            '        " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
            '        " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
            '        " AND ProjectName=" & obj.Project.GetQuotedText & _
            '        " Order by SeqNo"
            '    Else
            '        sql = "Select * from " & obj.Project.tblTasks & _
            '        " where EnvironmentName = " & obj.ObjSystem.Environment.GetQuotedText & _
            '        " AND ProjectName=" & obj.Project.GetQuotedText & _
            '        " Order by SeqNo"
            '    End If
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,TASKNAME,TASKTYPE,TASKSEQNO,TASKDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblTasks & _
            " where EngineName=" & obj.GetQuotedText & _
            " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
            " AND ProjectName=" & obj.Project.GetQuotedText & _
            " Order by TASKSEQNO"
            'End If

            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            Dim ProcCount As Integer = 0

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode is root node under which we add other nodes)
                ttype = GetVal(dt.Rows(i).Item("TaskType"))
                Select Case ttype
                    Case enumTaskType.TASK_MAIN
                        FillTasks(cNodeMain, dt.Rows(i), cnn)
                    Case enumTaskType.TASK_GEN, enumTaskType.TASK_LOOKUP, enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc
                        FillTasks(cNodeProc, dt.Rows(i), cnn)
                        ProcCount += 1
                        'Case enumTaskType.TASK_JOIN, enumTaskType.TASK_LOOKUP
                        '    FillTasks(cNodeJoin, dt.Rows(i), cnn)
                        'Case 
                        '    FillTasks(cNodeLookup, dt.Rows(i), cnn)
                End Select
                tvExplorer.Refresh()
            Next

            '/// Add Count to Folder Title
            cNodeProc.Text = "(" & ProcCount.ToString & ")" & " Procedures"

            If CurLoadedProject.PROCexpanded = True Then
                cNodeProc.Expand()
            Else
                cNodeProc.Collapse()
            End If
            tvExplorer.Refresh()

            If CurLoadedProject.MAINexpanded = True Then
                cNodeMain.Expand()
            Else
                cNodeMain.Collapse()
            End If
            tvExplorer.Refresh()

            If CurLoadedProject.ENGexpanded = True Then
                cNode.Expand()
            Else
                cNode.Collapse()
            End If
            tvExplorer.Refresh()

            For Each ds As clsDatastore In obj.Sources
                ds.SetIsMapped(False, True)
            Next
            For Each ds As clsDatastore In obj.Targets
                ds.SetIsMapped(False, True)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>fillTasks", sql)
            Return False
        End Try

    End Function

    Function FillTasks(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsTask
        Dim Otype As String

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try

            Otype = CType(cNode.Parent.Tag, INode).Type
            If Otype = NODE_ENVIRONMENT Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->TaskFolder->Any Task
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.TaskName = GetVal(dr.Item("TaskName"))
            obj.TaskType = GetVal(dr.Item("TaskType"))
            obj.TaskDescription = GetStr(GetVal(dr("TASKDESCRIPTION")))

            Select Case obj.TaskType
                'Case modDeclares.enumTaskType.TASK_JOIN, modDeclares.enumTaskType.TASK_LOOKUP
                '    If obj.Engine IsNot Nothing Then
                '        AddToCollection(obj.Engine.Joins, obj, obj.GUID)
                '    Else
                '        AddToCollection(obj.Environment.Joins, obj, obj.GUID)
                '    End If
                Case modDeclares.enumTaskType.TASK_MAIN
                    'If obj.Engine IsNot Nothing Then
                    AddToCollection(obj.Engine.Mains, obj, obj.GUID)
                    'Else
                    'AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    'End If
                Case modDeclares.enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc, _
                modDeclares.enumTaskType.TASK_GEN, modDeclares.enumTaskType.TASK_LOOKUP
                    'If obj.Engine IsNot Nothing Then
                    AddToCollection(obj.Engine.Procs, obj, obj.GUID)
                    'Else
                    'AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    'End If
            End Select

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.LoadDatastores()
            '    obj.LoadMappings()
            'End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store treeview node index
            If Procs.Contains(obj.TaskName) Then
                cNode.ForeColor = Color.Red
            End If

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillVar(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal EnvLoad As Boolean = False) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsVariable
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            If EnvLoad = True Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->VarFolder->Var
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.VariableName = GetVal(dr.Item("VariableName"))


            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.VarInitVal = GetVal(dr.Item("VarInitVal"))
            '    obj.VarSize = GetVal(dr.Item("VarSize"))
            '    obj.VariableDescription = GetVal(dr.Item("Description"))
            'Else
            obj.VariableDescription = GetVal(dr.Item("VariableDescription"))
            'End If

            If EnvLoad = True Then
                AddToCollection(obj.Environment.Variables, obj, obj.GUID)
            Else
                AddToCollection(obj.Engine.Variables, obj, obj.GUID)
            End If


            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex, "frmMain fillVar")
        End Try

    End Function

    Function FillConn(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsConnection
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'If CType(cNode.Tag, INode).Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
            '    obj.ConnectionType = GetVal(dr.Item("ConnectionType"))
            '    obj.ConnectionDescription = GetVal(dr.Item("Description"))
            '    obj.Database = GetVal(dr.Item("DBName"))
            '    obj.UserId = GetVal(dr.Item("UserId"))
            '    obj.Password = GetVal(dr.Item("Password"))
            '    obj.DateFormat = GetVal(dr.Item("Dateformat"))
            'Else
            obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
            obj.ConnectionDescription = GetVal(dr.Item("ConnectionDescription"))
            'End If

            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->ConnFolder->Conn
            AddToCollection(obj.Environment.Connections, obj, obj.GUID)

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillDataStore(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal DSD As String = "") As Boolean

        '/// Modified by Tom Karasch 12/06
        Dim obj As New clsDatastore

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Parent.Tag, INode) 'Engine->Folder(source/target)->ds
            obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

            'If DSD = "" Then
            '    obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
            '    obj.DatastoreDescription = GetVal(dr.Item("Description"))
            '    obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
            '    obj.DsDirection = GetVal(dr.Item("DsDirection"))
            '    obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
            '    obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
            '    'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
            '    obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
            '    obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
            '    'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
            '    obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

            '    obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
            '    obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
            '    obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
            '    obj.DatastoreDescription = GetVal(dr.Item("Description"))
            '    obj.OperationType = GetVal(dr.Item("OperationType"))
            '    obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
            '    obj.DsPort = GetVal(dr.Item("Port"))
            '    obj.DsUOW = GetVal(dr.Item("UOW"))
            '    obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
            '    'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
            '    '/// Load Globals From File
            '    'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    '    obj.LoadGlobals()
            '    'End If

            '    '//If AnyTree Object is renamed, then reload all datastore items to 
            '    '// Make sure renaming is propagated to all nodes
            '    obj.LoadItems()
            'Else
            obj.DatastoreDescription = GetVal(dr.Item("DATASTOREDESCRIPTION"))

            If obj.DatastoreDescription Is Nothing Then
                obj.DatastoreDescription = ""
            End If
            If DSD = "S" Then
                obj.DsDirection = DS_DIRECTION_SOURCE
            Else
                obj.DsDirection = DS_DIRECTION_TARGET
            End If

            obj.LoadItems(, True)

            obj.LoadMe()

            'End If

            Select Case obj.DsDirection
                Case DS_DIRECTION_SOURCE
                    'If obj.Engine Is Nothing Then
                    '    AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    'Else
                    AddToCollection(obj.Engine.Sources, obj, obj.GUID)
                    'End If
                Case DS_DIRECTION_TARGET
                    'If obj.Engine Is Nothing Then
                    '    AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    'Else
                    AddToCollection(obj.Engine.Targets, obj, obj.GUID)
                    'End If
            End Select

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False, obj.Text)
            obj.SeqNo = cNode.Index '//store position

            '// now add Datastore selections to tree
            AddDSstructuresToTree(cNode, obj)

            If obj.DsDirection = DS_DIRECTION_SOURCE Then
                cNode.Text = "(" & obj.ObjSelections.Count.ToString & ")" & obj.DsPhysicalSource
                If CurLoadedProject.SRCselExpanded = True Then
                    cNode.Expand()
                Else
                    cNode.Collapse()
                End If
                tvExplorer.Refresh()
            Else
                cNode.Text = obj.DsPhysicalSource
                cNode.Collapse()
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillDatastore")
            Return False
        End Try

    End Function

    'Function FillDSbyType(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal Ver As enumMetaVersion = enumMetaVersion.V2) As Boolean

    '    Try
    '        Dim obj As New clsDatastore

    '        obj.Parent = CType(cNode.Tag, INode).Parent
    '        obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

    '        If Ver = enumMetaVersion.V2 Then
    '            obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
    '            obj.DatastoreDescription = GetVal(dr.Item("Description"))
    '            obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
    '            obj.DsDirection = GetVal(dr.Item("DsDirection"))
    '            obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
    '            obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
    '            'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
    '            obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
    '            obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
    '            'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
    '            obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

    '            obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
    '            obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
    '            obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
    '            obj.DatastoreDescription = GetVal(dr.Item("Description"))
    '            obj.OperationType = GetVal(dr.Item("OperationType"))
    '            obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
    '            obj.DsPort = GetVal(dr.Item("Port"))
    '            obj.DsUOW = GetVal(dr.Item("UOW"))
    '            obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
    '            'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
    '            '/// Load Globals From File
    '            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
    '                obj.LoadGlobals()
    '            End If

    '            '//If AnyTree Object is renamed, then reload all datastore items to 
    '            '// Make sure renaming is propagated to all nodes
    '            obj.LoadItems()
    '        Else
    '            obj.DatastoreDescription = GetStr(GetVal(dr.Item("DATASTOREDESCRIPTION")))
    '            obj.DatastoreType = GetVal(dr.Item("DSTYPE"))

    '            obj.LoadItems(, True)

    '            obj.LoadAttr()
    '        End If


    '        AddToCollection(obj.Environment.Datastores, obj, obj.GUID)

    '        ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

    '        cNode = AddDSNode(obj, cNode)
    '        obj.SeqNo = cNode.Index '//store position

    '        cNode.Text = obj.DsPhysicalSource

    '        '// now add Datastore selections to tree
    '        AddDSstructuresToTree(cNode, obj)
    '        cNode.Collapse()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "frmMain FillDSbyType")
    '        Return False
    '    End Try

    'End Function

    '//This will add Datastore node under proper Category Folder
    '//cNode : is "Datastores" folder node
    '//obj : is object to represent struct node
    '// [Datastores]
    '//     |
    '//    [+]-[<DSType>]
    '//         |
    '//         [+]-DS1
    '//         [+]-DS2
    'Function AddDSNode(ByVal obj As clsDatastore, ByVal cNode As TreeNode) As TreeNode
    '    '//Now look for Datastore type and insert it in proper folder
    '    Dim nd As TreeNode = Nothing
    '    Dim IsFound As Boolean

    '    Try
    '        For Each nd In cNode.Nodes
    '            If nd.Text = GetDSFolderText(obj.DatastoreType) Then
    '                IsFound = True
    '                Exit For
    '            End If
    '        Next

    '        '//if folder for structure category is not found then add new folder 
    '        '//and then add structure node under it
    '        If IsFound = True Then
    '            '//Add struct node under struct category
    '            cNode = AddNode(nd.Nodes, obj.Type, obj, False)
    '            obj.SeqNo = cNode.Index '//store position
    '        Else
    '            Dim objFol As INode
    '            Dim objFolType As String = GetFolType(obj.DatastoreType)

    '            objFol = New clsFolderNode(GetDSFolderText(obj.DatastoreType), objFolType)

    '            objFol.Parent = CType(cNode.Parent.Tag, INode)
    '            '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
    '            nd = AddNode(cNode.Nodes, objFol.Type, objFol, False)
    '            '//Add struct node under struct category
    '            cNode = AddNode(nd.Nodes, obj.Type, obj, False, obj.DatastoreName)
    '            obj.SeqNo = cNode.Index '//store position
    '        End If
    '        Return cNode

    '    Catch ex As Exception
    '        LogError(ex, "frmMain AddDSNode")
    '        Return Nothing
    '    End Try

    'End Function

    'Function GetDSFolderText(ByVal DSType As enumDatastore) As String

    '    Select Case DSType
    '        Case modDeclares.enumDatastore.DS_BINARY
    '            Return "BINARY"
    '        Case modDeclares.enumDatastore.DS_DB2CDC
    '            Return "DB2CDC"
    '        Case modDeclares.enumDatastore.DS_DB2LOAD
    '            Return "DB2LOAD"
    '        Case modDeclares.enumDatastore.DS_DELIMITED
    '            Return "DELIMITED"
    '        Case modDeclares.enumDatastore.DS_GENERICCDC
    '            Return "GENERICCDC"
    '        Case modDeclares.enumDatastore.DS_HSSUNLOAD
    '            Return "HSSUNLOAD"
    '        Case modDeclares.enumDatastore.DS_IBMEVENT
    '            Return "IBMEVENT"
    '            'Case modDeclares.enumDatastore.DS_IMSCDC
    '            '    Return "IMSCDC"
    '        Case modDeclares.enumDatastore.DS_IMSDB
    '            Return "IMSDB"
    '            'Case modDeclares.enumDatastore.DS_IMSLE
    '            '    Return "IMSLE"
    '            'Case modDeclares.enumDatastore.DS_IMSLEBATCH
    '            '    Return "IMSLEBATCH"
    '        Case modDeclares.enumDatastore.DS_INCLUDE
    '            Return "INCLUDE"
    '        Case modDeclares.enumDatastore.DS_ORACLECDC
    '            Return "ORACLECDC"
    '        Case modDeclares.enumDatastore.DS_RELATIONAL
    '            Return "RELATIONAL"
    '        Case modDeclares.enumDatastore.DS_SUBVAR
    '            Return "SUBVAR"
    '        Case modDeclares.enumDatastore.DS_TEXT
    '            Return "TEXT"
    '            'Case modDeclares.enumDatastore.DS_TRBCDC
    '            '    Return "TRBCDC"
    '        Case modDeclares.enumDatastore.DS_UNKNOWN
    '            Return "UNKNOWN"
    '        Case modDeclares.enumDatastore.DS_VSAM
    '            Return "VSAM"
    '        Case modDeclares.enumDatastore.DS_VSAMCDC
    '            Return "VSAMCDC"
    '        Case modDeclares.enumDatastore.DS_XML
    '            Return "XML"
    '            'Case modDeclares.enumDatastore.DS_XMLCDC
    '            '    Return "XMLCDC"
    '        Case Else
    '            Return "Unknown"
    '    End Select

    'End Function

    'Function GetFolType(ByVal DStype As enumDatastore) As String

    '    Select Case DStype
    '        Case enumDatastore.DS_BINARY
    '            Return DS_BINARY
    '        Case enumDatastore.DS_DB2CDC
    '            Return DS_DB2CDC
    '        Case enumDatastore.DS_DB2LOAD
    '            Return DS_DB2LOAD
    '        Case enumDatastore.DS_DELIMITED
    '            Return DS_DELIMITED
    '        Case enumDatastore.DS_GENERICCDC
    '            Return DS_GENERICCDC
    '        Case enumDatastore.DS_HSSUNLOAD
    '            Return DS_HSSUNLOAD
    '        Case enumDatastore.DS_IBMEVENT
    '            Return DS_IBMEVENT
    '        Case enumDatastore.DS_IMSCDCLE
    '            Return DS_IMSCDCLE
    '        Case enumDatastore.DS_IMSDB
    '            Return DS_IMSDB
    '            'Case enumDatastore.DS_IMSLE
    '            '    Return DS_IMSLE
    '            'Case enumDatastore.DS_IMSLEBATCH
    '            '    Return DS_IMSLEBATCH
    '        Case enumDatastore.DS_INCLUDE
    '            Return DS_INCLUDE
    '        Case enumDatastore.DS_ORACLECDC
    '            Return DS_ORACLECDC
    '        Case enumDatastore.DS_RELATIONAL
    '            Return DS_RELATIONAL
    '        Case enumDatastore.DS_SUBVAR
    '            Return DS_SUBVAR
    '        Case enumDatastore.DS_TEXT
    '            Return DS_TEXT
    '            'Case enumDatastore.DS_TRBCDC
    '            '    Return DS_TRBCDC
    '        Case enumDatastore.DS_UNKNOWN
    '            Return DS_UNKNOWN
    '        Case enumDatastore.DS_VSAM
    '            Return DS_VSAM
    '        Case enumDatastore.DS_VSAMCDC
    '            Return DS_VSAMCDC
    '        Case enumDatastore.DS_XML
    '            Return DS_XML
    '            'Case enumDatastore.DS_XMLCDC
    '            '    Return DS_XMLCDC
    '        Case Else
    '            Return NODE_FO_DATASTORE
    '    End Select

    'End Function

    Function AddDSstructuresToTree(ByVal pNode As TreeNode, ByVal obj As clsDatastore, Optional ByVal MapAs As Boolean = False) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            pNode.Nodes.Clear()
            obj.LoadItems(, True)

            For i = 0 To obj.ObjSelections.Count - 1
                Dim DSselobj As clsDSSelection = obj.ObjSelections(i)

                DSselobj.LoadItems()

                If Not DSselobj.IsChildDSSelection = True Or MapAs = True Then
                    cnode = AddNode(pNode, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    tvExplorer.Refresh()
                    Call addDSSelChildrenToTree(cnode, DSselobj)
                End If
                tvExplorer.Refresh()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "main AddDSstructuresToTree")
            Return False
        End Try

    End Function

    Private Function addDSSelChildrenToTree(ByRef pNode As TreeNode, ByRef obj As clsDSSelection) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            For i = 0 To obj.ObjDSSelections.Count - 1
                Dim DSselobj As clsDSSelection
                DSselobj = obj.ObjDSSelections(i)
                DSselobj.LoadItems(, True)
                If DSselobj.Parent Is obj Then
                    cnode = AddNode(pNode, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    Call addDSSelChildrenToTree(cnode, DSselobj) '// recurse for all children
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "main addDSselChildrenToTree")
            Return False
        End Try

    End Function

    '//This will add Structure node under proper Category Folder
    '//cNode : is "Structures" folder node
    '//obj : is object to represent struct node
    '// [Structures]
    '//     |
    '//    [+]-[<category>]
    '//         |
    '//         [+]-Struct1
    '//         [+]-struct2
    Function AddStructNode(ByVal obj As clsStructure, ByVal cNode As TreeNode) As TreeNode
        '//Now look for Structure type and insert it in proper folder
        Dim nd As TreeNode = Nothing
        Dim IsFound As Boolean

        Try
            For Each nd In cNode.Nodes
                If nd.Text.EndsWith(GetStructureFolderText(obj.StructureType)) = True Then
                    IsFound = True
                    Exit For
                End If
            Next

            '//if folder for structure category is not found then add new folder 
            '//and then add structure node under it
            If IsFound = True Then
                '//Add struct node under struct category
                cNode = AddNode(nd, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            Else
                Dim objFol As INode
                objFol = New clsFolderNode(GetStructureFolderText(obj.StructureType), NODE_FO_STRUCT)
                objFol.Parent = CType(cNode.Parent.Tag, INode)
                '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
                nd = AddNode(cNode, objFol.Type, objFol, True)

                '//Add struct node under struct category
                cNode = AddNode(nd, obj.Type, obj, True)
                obj.SeqNo = cNode.Index '//store position
            End If

            nd.Text = "(" & nd.Nodes.Count.ToString & ")" & GetStructureFolderText(obj.StructureType)

            Dim nodeState As Boolean
            If nd.Text.EndsWith("COBOL") = True Then
                nodeState = CurLoadedProject.COBOLexpanded
            ElseIf nd.Text.EndsWith("COBOLIMS") = True Then
                nodeState = CurLoadedProject.COBOLIMSexpanded
            ElseIf nd.Text.EndsWith("CHeader") = True Then
                nodeState = CurLoadedProject.Cexpanded
            ElseIf nd.Text.EndsWith("XMLDTD") = True Then
                nodeState = CurLoadedProject.XMLDTDexpanded
            ElseIf nd.Text.EndsWith("DDL") = True Then
                nodeState = CurLoadedProject.DDLexpanded
            ElseIf nd.Text.EndsWith("DML") = True Then
                nodeState = CurLoadedProject.DMLexpanded
            End If

            If nodeState = True Then
                nd.Expand()
            Else
                nd.Collapse()
            End If
            tvExplorer.Refresh()

            Return cNode

        Catch ex As Exception
            Log("AddStructNode=>" & ex.Message)
            Return Nothing
        End Try

    End Function

    Function GetStructureFolderText(ByVal stType As enumStructure) As String

        Select Case stType
            Case modDeclares.enumStructure.STRUCT_C
                Return "CHeader"
            Case modDeclares.enumStructure.STRUCT_COBOL
                Return "COBOL"
            Case modDeclares.enumStructure.STRUCT_COBOL_IMS
                Return "COBOLIMS"
            Case modDeclares.enumStructure.STRUCT_IMS
                Return "IMS"
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                Return "DDL"
            Case modDeclares.enumStructure.STRUCT_REL_DML, enumStructure.STRUCT_REL_DML_FILE
                Return "DML"
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                Return "XMLDTD"
            Case Else
                Return "Unknown"
        End Select

    End Function

    '//////////////////////////////////////////////////////////////////
    '//// Now the tree is built from loaded data in the MetaDataDB ////
    '//////////////////////////////////////////////////////////////////
#End Region

    '//////// Right-Click on Tree Nodes ////////
#Region "Right Click Menu Event Handlers"

    Sub ShowPopupMenu(ByVal cNode As TreeNode, ByVal P As Point)

        '// Modified Many times by Tom Karasch
        Dim obj As INode
        Dim ctxmenu As New ContextMenu
        Dim mnuPop As MenuItem = Nothing
        Dim bFolderClick As Boolean
        Dim bAllowPaste As Boolean
        Dim bStructFolder As Boolean
        Dim bEngine As Boolean

        Try
            obj = CType(cNode.Tag, INode)

            bFolderClick = CType(cNode.Tag, INode).IsFolderNode
            EnableCopy(Not bFolderClick)
            If CType(cNode.Tag, INode).Text = "Descriptions" Then
                bStructFolder = True
            Else
                bStructFolder = False
            End If
            If CType(cNode.Tag, INode).Type = NODE_ENGINE Then
                bEngine = True
            Else
                bEngine = False
            End If


            bAllowPaste = IsClipboardAvailableforThisNodeType(obj.Type)

            Select Case obj.Type
                Case NODE_PROJECT
                    mnuPop = mnuPopupProject.CloneMenu
                Case NODE_FO_ENVIRONMENT, NODE_ENVIRONMENT
                    '//If clicked on folder then disable del and edit options
                    mnuDelEnv.Enabled = Not bFolderClick
                    'mnuEditEnv.Enabled = Not bFolderClick
                    'mnuCopyEnv.Enabled = Not bFolderClick
                    mnuPasteEnv.Enabled = bAllowPaste

                    mnuPop = mnuPopupEnv.CloneMenu
                Case NODE_FO_STRUCT, NODE_STRUCT
                    '//If clicked on folder then disable del and edit options
                    mnuDelStruct.Enabled = Not bFolderClick
                    'mnuEditStruct.Enabled = Not bFolderClick
                    mnuModelStructDDL.Enabled = Not bStructFolder
                    mnuModelStructDTD.Enabled = Not bStructFolder
                    mnuModelStructHeader.Enabled = Not bStructFolder '//8/15/05
                    mnuModelStructLOD.Enabled = Not bStructFolder
                    mnuModelStructSQL.Enabled = Not bStructFolder
                    mnuModelStructMSSQL.Enabled = Not bStructFolder
                    'mnuModelStructDB2.Enabled = Not bStructFolder
                    mnuAddStructSubset.Enabled = Not bFolderClick
                    'mnuCopyStruct.Enabled = Not bFolderClick
                    mnuPasteStruct.Enabled = bAllowPaste
                    mnuChgStruct.Enabled = Not bFolderClick
                    If obj.Text <> "Descriptions" Then
                        mnuDelAllStruct.Enabled = bFolderClick
                        Select Case obj.Text
                            Case "COBOL"
                                mnuDelAllStruct.Text = "Delete All COBOL Descriptions"
                            Case "DML"
                                mnuDelAllStruct.Text = "Delete All DML Descriptions"
                            Case "DDL"
                                mnuDelAllStruct.Text = "Delete All DDL Descriptions"
                            Case "CHeader"
                                mnuDelAllStruct.Text = "Delete All C Descriptions"
                            Case "XMLDTD"
                                mnuDelAllStruct.Text = "Delete All XMLDTD Descriptions"
                            Case "IMS"
                                mnuDelAllStruct.Text = "Delete All IMS Descriptions"
                            Case "COBOLIMS"
                                mnuDelAllStruct.Text = "Delete All COBOLIMS Descriptions"
                        End Select
                    Else
                        mnuDelAllStruct.Enabled = Not bFolderClick
                    End If

                    mnuPop = mnuPopupStruct.CloneMenu
                Case NODE_FO_STRUCT_SEL, NODE_STRUCT_SEL
                    '//If clicked on folder then disable del and edit options
                    mnuDelStructSelection.Enabled = Not bFolderClick
                    'mnuEditStructSelection.Enabled = Not bFolderClick
                    mnuCopyStructSel.Enabled = Not bFolderClick
                    mnuPasteStructSel.Enabled = bAllowPaste

                    mnuPop = mnuPopupStructSelection.CloneMenu
                Case NODE_FO_SYSTEM, NODE_SYSTEM
                    '//If clicked on folder then disable del and edit options
                    mnuDelSystem.Enabled = Not bFolderClick
                    'mnuEditSystem.Enabled = Not bFolderClick
                    'mnuCopySystem.Enabled = Not bFolderClick
                    mnuPasteSystem.Enabled = bAllowPaste

                    mnuPop = mnuPopupSystem.CloneMenu
                Case NODE_FO_CONNECTION, NODE_CONNECTION
                    '//If clicked on folder then disable del and edit options
                    mnuDelConn.Enabled = Not bFolderClick
                    'mnuEditConn.Enabled = Not bFolderClick
                    'mnuCopyConn.Enabled = Not bFolderClick
                    mnuPasteConn.Enabled = bAllowPaste

                    mnuPop = mnuPopupConnection.CloneMenu
                Case NODE_FO_ENGINE, NODE_ENGINE
                    '//If clicked on folder then disable del and edit options
                    mnuDelEngine.Enabled = Not bFolderClick
                    'mnuEditEngine.Enabled = Not bFolderClick
                    mnuScriptEngine.Enabled = Not bFolderClick
                    'mnuCopyEngine.Enabled = Not bFolderClick
                    mnuPasteEngine.Enabled = bAllowPaste
                    'mnuGenDebug.Enabled = Not bFolderClick
                    'mnuParseDebug.Enabled = Not bFolderClick
                    'mnuParse.Enabled = Not bFolderClick
                    mnuEngMerge.Enabled = bEngine And bAllowPaste

                    mnuPop = mnuPopupEngine.CloneMenu
                    '//TODO
                Case NODE_FO_DATASTORE, DS_UNKNOWN, DS_BINARY, DS_TEXT, DS_DELIMITED, DS_XML, DS_RELATIONAL, DS_DB2LOAD, DS_HSSUNLOAD, DS_IMSDB, DS_VSAM, DS_IMSCDCLE, DS_DB2CDC, DS_VSAMCDC, DS_IBMEVENT, DS_ORACLECDC, DS_GENERICCDC, DS_SUBVAR, DS_INCLUDE
                    'DS_IMSLE, DS_IMSLEBATCH, DS_XMLCDC, DS_TRBCDC,
                    '/// Modeling
                    mnuModDSSelC.Enabled = False
                    mnuModDSSelDDL.Enabled = False
                    mnuModDSSelDTD.Enabled = False
                    mnuModDSSelLOD.Enabled = False
                    mnuModDSSelSQL.Enabled = False
                    mnuModDSSelMSSQL.Enabled = False
                    mnuModDSSelC.Visible = False
                    mnuModDSSelDDL.Visible = False
                    mnuModDSSelDTD.Visible = False
                    mnuModDSSelLOD.Visible = False
                    mnuModDSSelSQL.Visible = False
                    mnuModDSSelMSSQL.Visible = False
                    '//If clicked on folder then disable del and edit options
                    mnuDelDS.Enabled = Not bFolderClick
                    mnuAddDSHSSUNLOAD.Enabled = False
                    mnuAddIMSLE.Enabled = False
                    'mnuEditDS.Enabled = Not bFolderClick
                    'If obj.Type = NODE_FO_TARGETDATASTORE And m_ClipObjects.Count > 0 Then
                    '    MenuItem9.Enabled = True
                    'Else
                    '    MenuItem9.Enabled = False
                    'End If
                    'If obj.Type = NODE_TARGETDATASTORE Or obj.Type = NODE_FO_TARGETDATASTORE Then
                    '    mnuAddDS.Text = "Add Target"
                    '    mnuDelDS.Text = "Delete Target"
                    'End If
                    'If obj.Type = NODE_SOURCEDATASTORE Or obj.Type = NODE_FO_SOURCEDATASTORE Then
                    '    mnuAddDS.Text = "Add Source"
                    '    mnuDelDS.Text = "Delete Source"
                    'End If
                    mnuDelAllDS.Text = "Delete All Datastores"
                    If obj.Type = NODE_TARGETDATASTORE Or obj.Type = NODE_SOURCEDATASTORE Or _
                    obj.Type = NODE_FO_DATASTORE Then
                        mnuDelAllDS.Enabled = False

                    Else
                        mnuDelAllDS.Enabled = True
                        Select Case obj.Type
                            Case NODE_FO_TARGETDATASTORE
                                mnuDelAllDS.Text = "Delete All Targets"
                            Case NODE_FO_SOURCEDATASTORE
                                mnuDelAllDS.Text = "Delete All Sources"
                            Case DS_UNKNOWN
                                mnuDelAllDS.Text = "Delete All UNKNOWN"
                            Case DS_BINARY
                                mnuDelAllDS.Text = "Delete All BINARY"
                            Case DS_TEXT
                                mnuDelAllDS.Text = "Delete All TEXT"
                            Case DS_DELIMITED
                                mnuDelAllDS.Text = "Delete All DELIMITED"
                            Case DS_XML
                                mnuDelAllDS.Text = "Delete All XML"
                            Case DS_RELATIONAL
                                mnuDelAllDS.Text = "Delete All RELATIONAL"
                            Case DS_DB2LOAD
                                mnuDelAllDS.Text = "Delete All DB2LOAD"
                            Case DS_HSSUNLOAD
                                mnuDelAllDS.Text = "Delete All HSSUNLOAD"
                            Case DS_IMSDB
                                mnuDelAllDS.Text = "Delete All IMSDB"
                            Case DS_VSAM
                                mnuDelAllDS.Text = "Delete All VSAM"
                            Case DS_IMSCDCLE
                                mnuDelAllDS.Text = "Delete All IMSCDCLE"
                            Case DS_DB2CDC
                                mnuDelAllDS.Text = "Delete All DB2CDC"
                            Case DS_VSAMCDC
                                mnuDelAllDS.Text = "Delete All VSAMCDC"
                                'Case DS_XMLCDC
                                '    mnuDelAllDS.Text = "Delete All XMLCDC"
                            Case DS_IBMEVENT
                                mnuDelAllDS.Text = "Delete All IBMEVENT"
                                'Case DS_TRBCDC
                                '    mnuDelAllDS.Text = "Delete All TRBCDC"
                            Case DS_ORACLECDC
                                mnuDelAllDS.Text = "Delete All ORACLECDC"
                            Case DS_GENERICCDC
                                mnuDelAllDS.Text = "Delete All GENERICCDC"
                            Case DS_IMSCDCLE
                                mnuDelAllDS.Text = "Delete All IMSCDCLE"
                                'Case DS_IMSLEBATCH
                                '    mnuDelAllDS.Text = "Delete All IMSLEBATCH"
                            Case DS_SUBVAR
                                mnuDelAllDS.Text = "Delete All SUBVAR"
                            Case DS_INCLUDE
                                mnuDelAllDS.Text = "Delete All INCLUDE"
                        End Select
                    End If
                    mnuCopyDS.Enabled = True 'Not bFolderClick
                    mnuPasteDS.Enabled = True 'bAllowPaste
                    mnuQuery.Enabled = False
                    mnuAddDS.Enabled = True
                    mnuLU.Enabled = False

                    mnuAddDS.Text = "Add Datastore"
                    mnuDelDS.Text = "Delete Datastore"


                    mnuPop = mnuPopupDS.CloneMenu

                Case NODE_FO_SOURCEDATASTORE, NODE_SOURCEDATASTORE, NODE_FO_TARGETDATASTORE, NODE_TARGETDATASTORE
                    '/// Modeling
                    mnuModDSSelC.Enabled = False
                    mnuModDSSelDDL.Enabled = False
                    mnuModDSSelDTD.Enabled = False
                    mnuModDSSelLOD.Enabled = False
                    mnuModDSSelSQL.Enabled = False
                    mnuModDSSelMSSQL.Enabled = False
                    mnuModDSSelC.Visible = False
                    mnuModDSSelDDL.Visible = False
                    mnuModDSSelDTD.Visible = False
                    mnuModDSSelLOD.Visible = False
                    mnuModDSSelSQL.Visible = False
                    mnuModDSSelMSSQL.Visible = False
                    '//If clicked on folder then disable del and edit options
                    mnuDelDS.Enabled = Not bFolderClick
                    'mnuEditDS.Enabled = Not bFolderClick
                    If obj.Type = NODE_FO_TARGETDATASTORE And m_ClipObjects.Count > 0 Then
                        MenuItem9.Enabled = True
                    Else
                        MenuItem9.Enabled = False
                    End If
                    If obj.Type = NODE_TARGETDATASTORE Or obj.Type = NODE_SOURCEDATASTORE Then
                        If CType(obj, clsDatastore).Engine Is Nothing Then
                            mnuAddDS.Text = "Add Datastore"
                            mnuDelDS.Text = "Delete Datastore"
                            mnuDelAllDS.Text = "Delete All Datastores"
                            mnuLU.Enabled = False
                        Else
                            If obj.Type = NODE_TARGETDATASTORE Then
                                mnuAddDS.Text = "Add Target"
                                mnuDelDS.Text = "Delete Target"
                                mnuDelAllDS.Text = "Delete All Targets"
                                mnuLU.Enabled = False
                            End If

                            If obj.Type = NODE_SOURCEDATASTORE Then
                                mnuAddDS.Text = "Add Source"
                                mnuDelDS.Text = "Delete Source"
                                mnuDelAllDS.Text = "Delete All Sources"
                                mnuLU.Enabled = True
                            End If
                        End If
                    End If

                    If obj.Type = NODE_FO_TARGETDATASTORE Then
                        mnuAddDS.Text = "Add Target"
                        mnuDelDS.Text = "Delete Target"
                        mnuDelAllDS.Text = "Delete All Targets"
                        mnuLU.Enabled = False
                    End If

                    If obj.Type = NODE_TARGETDATASTORE Or obj.Type = NODE_FO_TARGETDATASTORE Then
                        mnuAddDSHSSUNLOAD.Enabled = False
                        mnuAddIMSLE.Enabled = False
                    Else
                        mnuAddDSHSSUNLOAD.Enabled = True
                        mnuAddIMSLE.Enabled = True
                    End If

                    If obj.Type = NODE_FO_SOURCEDATASTORE Then
                        mnuAddDS.Text = "Add Source"
                        mnuDelDS.Text = "Delete Source"
                        mnuDelAllDS.Text = "Delete All Sources"
                        mnuLU.Enabled = True
                    End If

                    If obj.Type = NODE_TARGETDATASTORE Or obj.Type = NODE_SOURCEDATASTORE Then
                        mnuDelAllDS.Enabled = False
                    Else
                        mnuDelAllDS.Enabled = True
                    End If

                    mnuCopyDS.Enabled = True 'Not bFolderClick
                    mnuPasteDS.Enabled = True 'bAllowPaste
                    mnuQuery.Enabled = False
                    mnuAddDS.Enabled = True

                    mnuPop = mnuPopupDS.CloneMenu


                Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                    mnuAddDS.Enabled = True
                    '/// enable Modeling
                    mnuModDSSelC.Enabled = True
                    mnuModDSSelDDL.Enabled = True
                    mnuModDSSelDTD.Enabled = True
                    mnuModDSSelLOD.Enabled = True
                    mnuModDSSelSQL.Enabled = True
                    mnuModDSSelMSSQL.Enabled = True
                    '/// make sure modeling options are visible
                    mnuModDSSelC.Visible = False
                    mnuModDSSelDDL.Visible = False
                    mnuModDSSelDTD.Visible = False
                    mnuModDSSelLOD.Visible = False
                    mnuModDSSelSQL.Visible = False
                    mnuModDSSelMSSQL.Visible = False

                    If CType(obj, clsDSSelection).ObjDatastore.Engine Is Nothing Then
                        mnuAddDS.Text = "Add Datastore"
                        mnuDelDS.Text = "Delete Datastore"
                        mnuDelAllDS.Text = "Delete All Datastores"
                        mnuLU.Enabled = False
                    Else
                        If obj.Type = NODE_TARGETDSSEL Then
                            mnuAddDS.Text = "Add Target"
                            mnuDelDS.Text = "Delete Target"
                            mnuDelAllDS.Text = "Delete All Targets"
                            mnuLU.Enabled = False

                        End If
                        If obj.Type = NODE_SOURCEDSSEL Then
                            mnuAddDS.Text = "Add Source"
                            mnuDelDS.Text = "Delete Source"
                            mnuDelAllDS.Text = "Delete All Sources"
                            mnuLU.Enabled = True

                        End If
                    End If

                    If obj.Type = NODE_TARGETDSSEL Then
                        mnuAddDSHSSUNLOAD.Enabled = False
                        mnuAddIMSLE.Enabled = False
                    Else
                        mnuAddDSHSSUNLOAD.Enabled = True
                        mnuAddIMSLE.Enabled = True
                    End If

                    mnuDelDS.Enabled = False
                    'mnuEditDS.Enabled = True
                    mnuCopyDS.Enabled = False
                    mnuPasteDS.Enabled = False
                    MenuItem9.Enabled = False
                    '*********************************************************
                    '******* debug false --- to use set to true *********************
                    '*******************************************************
                    mnuQuery.Enabled = False
                    mnuPop = mnuPopupDS.CloneMenu

                Case NODE_FO_VARIABLE, NODE_VARIABLE
                    '//If clicked on folder then disable del and edit options
                    mnuDelVar.Enabled = Not bFolderClick
                    'mnuEditVar.Enabled = Not bFolderClick
                    'mnuCopyVar.Enabled = Not bFolderClick
                    mnuPasteVar.Enabled = bAllowPaste

                    mnuPop = mnuPopupVariable.CloneMenu

                Case NODE_FO_JOIN, NODE_FO_MAIN, NODE_FO_LOOKUP, NODE_FO_PROC, NODE_GEN, NODE_MAIN, NODE_LOOKUP, NODE_PROC
                    '//If clicked on folder then disable del and edit options
                    mnuDelTask.Enabled = Not bFolderClick
                    'mnuEditTask.Enabled = Not bFolderClick
                    'mnuCopyTask.Enabled = Not bFolderClick
                    mnuPasteTask.Enabled = bAllowPaste
                    mnuDelAllTask.Enabled = bFolderClick
                    mnuChangeTask.Enabled = Not bFolderClick

                    Select Case obj.Type
                        Case NODE_FO_JOIN
                            mnuAddTask.Text = "Add Join"
                            mnuDelTask.Text = "Delete Join"
                            mnuDelAllTask.Text = "Delete All Joins"
                            mnuAddGen.Enabled = False
                        Case NODE_MAIN, NODE_FO_MAIN
                            mnuAddTask.Text = "Add Main Procedure"
                            mnuDelTask.Text = "Delete Main Procedure"
                            mnuDelAllTask.Text = "Delete all Main Procedures"
                            mnuAddGen.Enabled = False
                        Case NODE_LOOKUP, NODE_FO_LOOKUP
                            mnuAddTask.Text = "Add Lookup"
                            mnuDelTask.Text = "Delete Lookup"
                            mnuDelAllTask.Text = "Delete All Lookups"
                            mnuAddGen.Enabled = False
                        Case NODE_GEN, NODE_PROC, NODE_FO_PROC
                            mnuAddTask.Text = "Add Mapping Procedure"
                            mnuDelTask.Text = "Delete Procedure"
                            mnuDelAllTask.Text = "Delete All Procedures"
                            mnuAddGen.Enabled = True
                    End Select

                    mnuPop = mnuPopupTask.CloneMenu

                    '//TODO
            End Select

            '//////////////////////////////////////////////////////
            'Create empty array to store MenuItem objects.
            Dim myItems(mnuPop.MenuItems.Count - 1) As MenuItem

            mnuPop.MenuItems.CopyTo(myItems, 0)
            ctxmenu.MenuItems.AddRange(myItems)
            ctxmenu.Show(tvExplorer, P)
            '//////////////////////////////////////////////////////
        Catch ex As Exception
            LogError(ex, "ShowPopupMenu=>" & ex.Message)
        End Try

    End Sub

    Private Sub OnEnvAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddEnv.Click

        DoEnvironmentAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub OnEnvDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelEnv.Click

        DoAction(modDeclares.enumAction.ACTION_DELETE)

    End Sub

    '/// added by TK for query tool
    Private Sub OnQueryClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuQuery.Click

        StartQueryTool()

    End Sub

    Private Sub mnuAddStructCOBOL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructCOBOL.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_COBOL)

    End Sub

    Private Sub OnStructAddClickCobolIMS(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructCOBOL_IMS.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_COBOL_IMS)

    End Sub

    Private Sub OnStructAddClickXML(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructXML.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_XMLDTD)

    End Sub

    Private Sub OnStructAddClickDDL(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructDDL.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_REL_DDL)

    End Sub

    Private Sub mnuDMLUser_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDMLUser.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_REL_DML_FILE, , , True)

    End Sub

    Private Sub mnuDMLChoose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDMLChoose.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_REL_DML)

    End Sub

    Private Sub mnuDMLexisting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDMLexisting.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, enumStructure.STRUCT_REL_DML_FILE, , , , True)

    End Sub

    Private Sub OnStructAddClickCHeader(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructCStruct.Click

        DoStructureAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumStructure.STRUCT_C)

    End Sub

    Private Sub OnStructModelDDLClickCHeader(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructDDL.Click

        DoModelAction("STR", "DDL")

    End Sub

    Private Sub OnStructModelDTDClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructDTD.Click
        '//Generate DTD
        DoModelAction("STR", "DTD")

    End Sub

    Private Sub OnStructModelLODClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructLOD.Click
        '//Generate Oracle Load
        DoModelAction("STR", "LOD")

    End Sub

    Private Sub OnStructModelSQLClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructSQL.Click
        '//Generate Oracle Trigger
        DoModelAction("STR", "SQL")

    End Sub

    Private Sub OnStructModelMSSQLClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructMSSQL.Click
        '//Generate MS SQLServer Trigger
        DoModelAction("STR", "MSSQL")

    End Sub

    Private Sub OnStructModelDB2Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructDB2.Click
        '//Generate DTD
        DoModelAction("STR", "DB2")

    End Sub

    Private Sub OnStructModelHeaderClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructHeader.Click
        '//Generate C Header
        DoModelAction("STR", "H")

    End Sub

    Private Sub OnStructSelectionModelDDLClickCHeader(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructSelectionDDL.Click
        '//Generate DDL
        DoModelAction("SEL", "DDL")

    End Sub

    Private Sub OnStructSelectionModelDTDClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructSelectionDTD.Click
        '//Generate DTD
        DoModelAction("SEL", "DTD")

    End Sub

    Private Sub OnStructSelectionModelHeaderClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelStructSelectionHeader.Click
        '//Generate C Header file
        DoModelAction("SEL", "H")

    End Sub

    Private Sub mnuModelSSLOD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelSSLOD.Click

        DoModelAction("SEL", "LOD")

    End Sub

    Private Sub mnuModelSSSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelSSSQL.Click

        DoModelAction("SEL", "SQL")

    End Sub

    Private Sub mnuModelSSMSSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelSSMSSQL.Click

        DoModelAction("SEL", "MSSQL")

    End Sub

    Private Sub mnuModelSSDB2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModelSSDB2.Click

        DoModelAction("SEL", "DB2")

    End Sub

    Private Sub mnuModDSSelDTD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelDTD.Click

        DoModelAction(, "DTD")

    End Sub

    Private Sub mnuModDSSelDDL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelDDL.Click

        DoModelAction(, "DDL")

    End Sub

    Private Sub mnuModDSSelC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelC.Click

        DoModelAction(, "H")

    End Sub

    Private Sub mnuModDSSelLOD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelLOD.Click

        DoModelAction(, "LOD")

    End Sub

    Private Sub mnuModDSSelSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelSQL.Click

        DoModelAction(, "SQL")

    End Sub

    Private Sub mnuModDSSelMSSQL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuModDSSelMSSQL.Click

        DoModelAction(, "MSSQL")

    End Sub

    Private Sub OnStructDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelStruct.Click

        DoAction(modDeclares.enumAction.ACTION_DELETE)

    End Sub

    '// Added by TK to allow structure file to be changed
    Private Sub OnChangeStructureClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChgStruct.Click

        Dim ret As MsgBoxResult
        ret = MsgBox("Changing a Description file where field names do not match will result in Description fields" & (Chr(13)) & "being removed from datastores and mappings where they are used. Do you still want to continue?", MsgBoxStyle.YesNo, "Replace Description File")
        If ret = MsgBoxResult.Yes Then
            DoStructureAction(modDeclares.enumAction.ACTION_CHANGE)
        End If

    End Sub

    Private Sub OnStructSelectionAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddStructSelection.Click, mnuAddStructSubset.Click

        DoStructureSelectionAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub OnStructSelectionDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelStructSelection.Click

        DoAction(modDeclares.enumAction.ACTION_DELETE)

    End Sub

    Private Sub OnSystemAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddSys.Click

        DoSystemAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub OnSystemDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelSystem.Click

        DoAction(modDeclares.enumAction.ACTION_DELETE)

    End Sub

    Private Sub OnConnectionAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddConn.Click

        DoConnectionAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub OnConnectionDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelConn.Click

        DoAction(modDeclares.enumAction.ACTION_DELETE)

    End Sub

    Private Sub OnEngineAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddEngine.Click

        DoEngineAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub OnEngineDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelEngine.Click
        DoEngineAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub mnuEngMerge_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEngMerge.Click
        DoEngineAction(enumAction.ACTION_MERGE, True)
    End Sub

    Private Sub OnDSAddClickBinary(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSBinary.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_BINARY)
    End Sub

    'Private Sub OnDSAddClickText(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_TEXT)
    'End Sub

    Private Sub OnDSAddClickDelimited(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSDelimited.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_DELIMITED)
    End Sub

    Private Sub OnDSAddClickXML(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSXML.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_XML)
    End Sub

    Private Sub OnDSAddClickRelational(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSRelational.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_RELATIONAL)
    End Sub

    Private Sub OnDSAddClickDB2LOAD(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSDB2LOAD.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_DB2LOAD)
    End Sub

    Private Sub OnDSAddClickHSSUNLOAD(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSHSSUNLOAD.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_HSSUNLOAD)
    End Sub

    Private Sub OnDSAddClickIMSDB(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSIMS.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IMSDB)
    End Sub

    Private Sub OnDSAddClickVSAM(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSVSAM.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_VSAM)
    End Sub

    Private Sub OnDSAddClickIMSCDC(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSIMSCDC.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IMSCDC)
    End Sub

    Private Sub OnDSAddClickIMSLE(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddIMSLE.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IMSCDCLE)
    End Sub

    Private Sub OnDSAddClickDB2CDC(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSDB2CDC.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_DB2CDC)
    End Sub

    Private Sub OnDSAddClickVSAMCDC(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSVSAMCDC.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_VSAMCDC)
    End Sub

    'Private Sub OnDSAddClickXMLCDC(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_XMLCDC)
    'End Sub

    ''Trigger based cdc
    'Private Sub OnDSAddClickTRBCDC(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_TRBCDC)
    'End Sub

    '------ New V3 DS Types
    Private Sub mnuAddOraCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddOraCDC.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_ORACLECDC)
    End Sub

    Private Sub mnuAddSQDCDC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddSQDCDC.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_GENERICCDC)
    End Sub

    'Private Sub mnuAddIMSLE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddIMSLE.Click
    '    DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IMSLE)
    'End Sub

    'Private Sub mnuAddIMSLEBat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IMSLEBATCH)
    'End Sub

    Private Sub OnDSDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelDS.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    '//New: 8/10/05 by Npatel
    Private Sub OnDSAddClickIBMEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddDSIBMEvent.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_IBMEVENT)
    End Sub

    Private Sub mnuIncludeDS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIncludeDS.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_INCLUDE)
    End Sub

    Private Sub OnProjectDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelProject.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub OnVarAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddVar.Click
        DoVarAction(modDeclares.enumAction.ACTION_NEW)
    End Sub

    Private Sub OnVarDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelVar.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub OnTaskAddClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddTask.Click
        DoTaskAction(modDeclares.enumAction.ACTION_NEW, enumTaskType.TASK_PROC)
    End Sub

    Private Sub mnuAddGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAddGen.Click
        DoTaskAction(modDeclares.enumAction.ACTION_NEW, enumTaskType.TASK_GEN)
    End Sub

    Private Sub OnTaskDelClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelTask.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub mnuIncProc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIncProc.Click
        DoTaskAction(enumAction.ACTION_NEW, enumTaskType.TASK_IncProc)
    End Sub

    Private Sub mnuChangeTask_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuChangeTask.Click
        DoTaskAction(enumAction.ACTION_CHANGE)
    End Sub

    Private Sub mnuDelAllDS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelAllDS.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub mnuDelAllTask_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelAllTask.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub mnuDelAllStruct_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelAllStruct.Click
        DoAction(modDeclares.enumAction.ACTION_DELETE)
    End Sub

    Private Sub mnuLU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuLU.Click
        DoDatastoreAction(modDeclares.enumAction.ACTION_NEW, modDeclares.enumDatastore.DS_RELATIONAL, , , True)
    End Sub

#End Region

    '//////// Script Generation /////////
#Region "Generate Scripts"

    Private Sub OnEngineGenerateScriptClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScriptEngine.Click

        Dim obj As clsEngine
        'Dim sysobj As clsSystem
        Dim envObj As clsEnvironment
        Dim strSaveDir As String

        'If System.IO.Directory.Exists(GetAppPath() & "Scripts") = False Then
        '    System.IO.Directory.CreateDirectory(GetAppPath() & "Scripts")
        'End If

        'Me.Cursor = Cursors.WaitCursor
        '/// Get Engine Object
        obj = CType(tvExplorer.SelectedNode.Tag, clsEngine)
        obj.LoadMe()
        'sysobj = obj.ObjSystem
        'sysobj.LoadItems()

        '/// Get Environment Object
        envObj = obj.ObjSystem.Environment
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
        '******************************* New Script Generation v3 ************************************
        'Dim RetCode As clsRcode
        ''/// Generate Script ... Returns a return code object
        'RetCode = GenerateEngScriptV3(obj, strSaveDir, True)

        Dim frm As New frmScriptGen
        frm.OpenFormEng(obj, strSaveDir)

        'Log("********* Script Generation End ***********")

        'Me.Cursor = Cursors.Default

    End Sub

    Private Sub mnuScriptDS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScriptDS.Click

        Dim obj As clsDatastore
        Dim envObj As clsEnvironment
        Dim dsSelobj As clsDSSelection
        Dim strSaveDir As String

        '/// Get Engine Object
        If CType(tvExplorer.SelectedNode.Tag, INode).Type = NODE_SOURCEDSSEL Or _
        CType(tvExplorer.SelectedNode.Tag, INode).Type = NODE_TARGETDSSEL Then
            dsSelobj = tvExplorer.SelectedNode.Tag
            obj = dsSelobj.ObjDatastore
        Else
            obj = tvExplorer.SelectedNode.Tag
        End If

        '/// Get Environment Object
        envObj = obj.Engine.ObjSystem.Environment
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
        '******************************* New Script Generation v3 ************************************
        'Dim RetCode As clsRcode
        ''/// Generate Script ... Returns a return code object
        'RetCode = GenerateEngScriptV3(obj, strSaveDir, True)

        Dim frm As New frmScriptGen
        frm.OpenFormDS(obj, strSaveDir)

        'Log("********* Script Generation End ***********")

        'Me.Cursor = Cursors.Default
    End Sub

    Private Sub mnuScriptProc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuScriptProc.Click

        Dim obj As clsTask
        Dim envObj As clsEnvironment
        Dim strSaveDir As String

        'If System.IO.Directory.Exists(GetAppPath() & "Scripts") = False Then
        '    System.IO.Directory.CreateDirectory(GetAppPath() & "Scripts")
        'End If

        'Me.Cursor = Cursors.WaitCursor
        '/// Get Engine Object
        obj = tvExplorer.SelectedNode.Tag
        '/// Get Environment Object
        envObj = obj.Engine.ObjSystem.Environment
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
        '******************************* New Script Generation v3 ************************************
        'Dim RetCode As clsRcode
        ''/// Generate Script ... Returns a return code object
        'RetCode = GenerateEngScriptV3(obj, strSaveDir, True)

        Dim frm As New frmScriptGen
        frm.OpenFormProc(obj, strSaveDir)

        'Log("********* Script Generation End ***********")

        'Me.Cursor = Cursors.Default

    End Sub

    'Private Sub mnuParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuParse.Click

    '    Dim obj As clsEngine
    '    Dim envObj As clsEnvironment
    '    Dim strSaveDir As String

    '    If System.IO.Directory.Exists(GetAppPath() & "Scripts") = False Then
    '        System.IO.Directory.CreateDirectory(GetAppPath() & "Scripts")
    '    End If

    '    Me.Cursor = Cursors.WaitCursor
    '    '/// Get Engine Object
    '    obj = tvExplorer.SelectedNode.Tag
    '    '/// Get Environment Object
    '    envObj = obj.ObjSystem.Environment
    '    '/// Get Script Directory
    '    If envObj.LocalScriptDir <> "" Then
    '        strSaveDir = envObj.LocalScriptDir
    '    Else
    '        MsgBox("LocalScriptDir in the environment is not filled, please specify a location to generate the script. ", MsgBoxStyle.OkOnly, MsgTitle)
    '        Me.Cursor = Cursors.Default
    '        Exit Sub
    '    End If
    '    '******************************* New Script Generation v3 ************************************
    '    Dim RetCode As clsRcode
    '    '/// Generate Script ... Returns a return code object
    '    RetCode = GenerateEngScriptV3(obj, strSaveDir)

    '    Dim frm As New frmScriptGen
    '    frm.OpenScripts(RetCode, strSaveDir)

    '    'If RetCode.HasError = False Then
    '    '    '/// Good Script
    '    '    If MsgBox("Script Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode, MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
    '    '        Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
    '    '    End If
    '    'Else
    '    '    If RetCode.ErrorLocation = enumErrorLocation.SQDParse Then
    '    '        '/// Error while Parsing Script
    '    '        If MsgBox("Script Caused an Error while Parsing:" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode & Chr(13) & Chr(13) & _
    '    '        "Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script Report?", MsgBoxStyle.YesNo, _
    '    '        "Script Generation Parse Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    Else
    '    '        '/// Error in Windows (e.g. Object problems)
    '    '        If MsgBox("Script Generation Returned The Following Error:" & Chr(13) & _
    '    '        RetCode.GetGUIErrorMsg() & Chr(13) & Chr(13) & _
    '    '        "Error Occured at:" & Chr(13) & _
    '    '        "SQD file line Number: " & RetCode.SQDline & Chr(13) & _
    '    '        "INL file Line Number: " & RetCode.INLline & Chr(13) & _
    '    '        "TMP file Line Number: " & RetCode.TMPline & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script?", MsgBoxStyle.YesNo, "Script Generation Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ErrorPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    End If
    '    'End If
    '    '*****************************************************************************************
    '    Log("********* Script Generation End ***********")

    '    Me.Cursor = Cursors.Default

    'End Sub

    'Private Sub mnuGenDebug_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGenDebug.Click

    '    Dim obj As clsEngine
    '    Dim envObj As clsEnvironment
    '    Dim strSaveDir As String

    '    If System.IO.Directory.Exists(GetAppPath() & "Scripts") = False Then
    '        System.IO.Directory.CreateDirectory(GetAppPath() & "Scripts")
    '    End If

    '    Me.Cursor = Cursors.WaitCursor
    '    '/// Get Engine Object
    '    obj = tvExplorer.SelectedNode.Tag
    '    '/// Get Environment Object
    '    envObj = obj.ObjSystem.Environment
    '    '/// Get Script Directory
    '    If envObj.LocalScriptDir <> "" Then
    '        strSaveDir = envObj.LocalScriptDir
    '    Else
    '        MsgBox("LocalScriptDir in the environment is not filled, please specify a location to generate the script. ", MsgBoxStyle.OkOnly, MsgTitle)
    '        Me.Cursor = Cursors.Default
    '        Exit Sub
    '    End If
    '    '******************************* New Script Generation v3 ************************************
    '    Dim RetCode As clsRcode
    '    '/// Generate Script ... Returns a return code object
    '    RetCode = GenerateEngScriptV3(obj, strSaveDir, True, True)

    '    Dim frm As New frmScriptGen
    '    frm.OpenScripts(RetCode, strSaveDir)

    '    'If RetCode.HasError = False Then
    '    '    '/// Good Script
    '    '    If MsgBox("Script Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode, MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
    '    '        Shell("notepad.exe " & Quote(RetCode.Path, """"), AppWinStyle.NormalFocus)
    '    '    End If
    '    'Else
    '    '    If RetCode.ErrorLocation = enumErrorLocation.SQDParse Then
    '    '        '/// Error while Parsing Script
    '    '        If MsgBox("Script Caused an Error while Parsing:" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode & Chr(13) & Chr(13) & _
    '    '        "Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script Report?", MsgBoxStyle.YesNo, _
    '    '        "Script Generation Parse Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    Else
    '    '        '/// Error in Windows (e.g. Object problems)
    '    '        If MsgBox("Script Generation Returned The Following Error:" & Chr(13) & _
    '    '        RetCode.GetGUIErrorMsg() & Chr(13) & Chr(13) & _
    '    '        "Error Occured at:" & Chr(13) & _
    '    '        "SQD file line Number: " & RetCode.SQDline & Chr(13) & _
    '    '        "INL file Line Number: " & RetCode.INLline & Chr(13) & _
    '    '        "TMP file Line Number: " & RetCode.TMPline & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script?", MsgBoxStyle.YesNo, "Script Generation Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ErrorPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    End If
    '    'End If
    '    '*****************************************************************************************
    '    Log("********* Script Generation End ***********")

    '    Me.Cursor = Cursors.Default

    'End Sub

    'Private Sub mnuParseDebug_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuParseDebug.Click

    '    Dim obj As clsEngine
    '    Dim envObj As clsEnvironment
    '    Dim strSaveDir As String

    '    If System.IO.Directory.Exists(GetAppPath() & "Scripts") = False Then
    '        System.IO.Directory.CreateDirectory(GetAppPath() & "Scripts")
    '    End If

    '    Me.Cursor = Cursors.WaitCursor
    '    '/// Get Engine Object
    '    obj = tvExplorer.SelectedNode.Tag
    '    '/// Get Environment Object
    '    envObj = obj.ObjSystem.Environment
    '    '/// Get Script Directory
    '    If envObj.LocalScriptDir <> "" Then
    '        strSaveDir = envObj.LocalScriptDir
    '    Else
    '        MsgBox("LocalScriptDir in the environment is not filled, please specify a location to generate the script. ", MsgBoxStyle.OkOnly, MsgTitle)
    '        Me.Cursor = Cursors.Default
    '        Exit Sub
    '    End If
    '    '******************************* New Script Generation v3 ************************************
    '    Dim RetCode As clsRcode
    '    '/// Generate Script ... Returns a return code object
    '    RetCode = GenerateEngScriptV3(obj, strSaveDir, , True)

    '    Dim frm As New frmScriptGen
    '    frm.OpenScripts(RetCode, strSaveDir)

    '    'If RetCode.HasError = False Then
    '    '    '/// Good Script
    '    '    If MsgBox("Script Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode, MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
    '    '        Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
    '    '    End If
    '    'Else
    '    '    If RetCode.ErrorLocation = enumErrorLocation.SQDParse Then
    '    '        '/// Error while Parsing Script
    '    '        If MsgBox("Script Caused an Error while Parsing:" & Chr(13) & Chr(13) & _
    '    '        RetCode.ReturnCode & Chr(13) & Chr(13) & _
    '    '        "Generator wrote:" & Chr(13) & _
    '    '        RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
    '    '        RetCode.INLline & " lines to .INL file and" & Chr(13) & _
    '    '        RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script Report?", MsgBoxStyle.YesNo, _
    '    '        "Script Generation Parse Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    Else
    '    '        '/// Error in Windows (e.g. Object problems)
    '    '        If MsgBox("Script Generation Returned The Following Error:" & Chr(13) & _
    '    '        RetCode.GetGUIErrorMsg() & Chr(13) & Chr(13) & _
    '    '        "Error Occured at:" & Chr(13) & _
    '    '        "SQD file line Number: " & RetCode.SQDline & Chr(13) & _
    '    '        "INL file Line Number: " & RetCode.INLline & Chr(13) & _
    '    '        "TMP file Line Number: " & RetCode.TMPline & Chr(13) & Chr(13) & _
    '    '        "Would You Like to view the Script?", MsgBoxStyle.YesNo, "Script Generation Error") = MsgBoxResult.Yes Then
    '    '            Shell("notepad.exe " & Quote(RetCode.ErrorPath, """"), AppWinStyle.NormalFocus)
    '    '        End If
    '    '    End If
    '    'End If
    '    '*****************************************************************************************
    '    Log("********* Script Generation End ***********")

    '    Me.Cursor = Cursors.Default

    'End Sub

#End Region

    '//////// Toolbar Click actions ////////
#Region "Toolbar events"

    Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick
        '//
        Dim buttonTXT As String = e.Button.Tag.ToString

        Try
            Select Case buttonTXT
                Case "New" '//new project
                    Call DoProjectAction(modDeclares.enumAction.ACTION_NEW)
                Case "Save" '//save
                    DoSave(False)
                Case "Open" '//open project
                    Call DoProjectAction(modDeclares.enumAction.ACTION_OPEN)
                Case "Script"
                    Call OnEngineGenerateScriptClick(sender, e)
                Case "Copy"
                    CopyClick(sender, e)
                Case "Paste"
                    PasteClick(sender, e)
                Case "Delete"
                    DoDeleteAction()
                Case "Query"  '//added 11/21/2006 by TK for new Query Tool
                    StartQueryTool()
                Case "Map"    '//added 7/18/2007 for source to target field name associations
                    Call OpenMapList()
                Case "Toggle"
                    Call SCmain_DoubleClick(Me, New EventArgs)

                    '
                    '//TODO More buttons may be added here
                    '
                Case "Help"
                    ShowHelp(modHelp.HHId.H_Welcome_to_SQData_Studio)
                Case "Log"
                    ShowLog()
                Case Else
            End Select

        Catch ex As Exception
            Log("ToolBar1_ButtonClick=>" & ex.Message)
        End Try

    End Sub

    Private Sub mnuFile_Popup(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuFile.Popup

        If objTopMost IsNot Nothing Then
            mnuSave.Enabled = objTopMost.IsModified
        Else
            mnuSave.Enabled = False
        End If

        mnuCloseProject.Enabled = Not (tvExplorer.SelectedNode Is Nothing)
    End Sub

    Private Sub mnuCloseAllProjects_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseAllProjects.Click

        If tvExplorer.GetNodeCount(False) = 1 Then '//if one project loaded then just close it
            If DoSave() = MsgBoxResult.Cancel Then Exit Sub
            If Not CurLoadedProject Is Nothing Then
                CurLoadedProject.Save()
                CurLoadedProject.SaveToRegistry()
                CurLoadedProject = Nothing
            End If
            tvExplorer.Nodes(0).Remove()
            HideAllUC()
            ctFolder.Clear()
        ElseIf tvExplorer.GetNodeCount(False) > 1 Then '//if mulitple projects loaded then force user to select one
            If DoSave() = MsgBoxResult.Cancel Then Exit Sub
            If Not CurLoadedProject Is Nothing Then
                CurLoadedProject.Save()
                CurLoadedProject.SaveToRegistry()
                CurLoadedProject = Nothing
            End If
            Dim nd As TreeNode
            For Each nd In tvExplorer.Nodes
                nd = tvExplorer.Nodes(0) 'with out this it wont remove the nodes in an order and causes an exception. 
                nd.Remove()
            Next
            HideAllUC()
            ctFolder.Clear()
        End If
        cnn.Close()

        HideAllUC()

    End Sub

    Private Sub mnuCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopy.Click
        CopyClick(sender, e)

    End Sub

    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click
        PasteClick(sender, e)

    End Sub

    Private Sub mnuDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuDelete.Click
        DoDeleteAction()

    End Sub

    Private Sub mnuOpenProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenProject.Click

        If DoSave() = MsgBoxResult.Cancel Then Exit Sub
        Call DoProjectAction(modDeclares.enumAction.ACTION_OPEN)

    End Sub

    Private Sub mnuHelpIdx_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpIdx.Click
        ShowHelp(modHelp.HHId.H_Welcome_to_SQData_Studio)
    End Sub

    Private Sub mnuAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuAbout.Click

#If CONFIG = "ETI" Then
        Dim f As New frmAboutETI
        f.ShowDialog()
#Else
        Dim f As New frmAbout
        f.ShowDialog()
#End If


    End Sub

    Private Sub mnuExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuExit.Click

        If DoSave() = MsgBoxResult.Cancel Then Exit Sub
        If Not CurLoadedProject Is Nothing Then
            CurLoadedProject.SaveToRegistry()
            CurLoadedProject = Nothing
        End If
        If cnn IsNot Nothing Then
            cnn.Close()
        End If
        Application.Exit()

    End Sub

    Private Sub mnuNewProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuNewProject.Click

        If DoSave() = MsgBoxResult.Cancel Then Exit Sub
        Call DoProjectAction(modDeclares.enumAction.ACTION_NEW)

    End Sub

    Private Sub mnuSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuSave.Click
        If DoSave() = MsgBoxResult.Cancel Then Exit Sub

    End Sub

    Private Sub mnuCloseProject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCloseProject.Click

        If tvExplorer.GetNodeCount(False) = 1 Then '//if one project loaded then just close it
            If DoSave() = MsgBoxResult.Cancel Then Exit Sub
            tvExplorer.Nodes(0).Remove()

            ctFolder.Clear()
        ElseIf tvExplorer.GetNodeCount(False) > 1 Then '//if mulitple projects loaded then force user to select one
            If tvExplorer.SelectedNode Is Nothing Then
                MsgBox("Please select project from treeview", MsgBoxStyle.OkOnly, MsgTitle)
                Exit Sub
            Else
                If DoSave() = MsgBoxResult.Cancel Then Exit Sub
                GetTopLevelNode(tvExplorer.SelectedNode).Remove()
            End If
        End If

        If Not CurLoadedProject Is Nothing Then
            CurLoadedProject.Save()
            CurLoadedProject.SaveToRegistry()
            CurLoadedProject = Nothing
        End If
        cnn.Close()

        HideAllUC()

    End Sub

    Public Sub ToolBar1_Keydown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ToolBar1.KeyDown

        If ToolBar1.Focused = True Then
            If e.KeyCode = Keys.F1 Then ShowHelp(modHelp.HHId.H_Menu_and_Toolbar)
        End If

    End Sub

#End Region

    '//////// Copy/Paste functionality ////////
#Region "Clipboard Functions"

    '//This function takes treenode as input and if treenode is INode it checks if 
    '//foldernode then all child nodes will be copied in to clipboard else if INode
    '// but not folder then single node is copied in to clipboard buffer
    Function SetClipobjects(ByVal cNode As TreeNode) As Boolean
        Try
            'TODO : In future we will aaddsupport for copy entire folder
            If CType(cNode.Tag, INode).IsFolderNode = True Then
                ClearClipboard()
                Select Case CType(cNode.Tag, INode).Type
                    Case NODE_FO_STRUCT
                        If cNode.Text = "Descriptions" Then
                            Dim parentNode As TreeNode
                            Dim childNode As TreeNode
                            For Each parentNode In cNode.Nodes
                                For Each childNode In parentNode.Nodes
                                    m_ClipObjects.Add(childNode.Tag)
                                Next
                            Next
                        Else
                            Dim childNode As TreeNode
                            For Each childNode In cNode.Nodes
                                m_ClipObjects.Add(childNode.Tag)
                            Next
                        End If
                    Case NODE_FO_DATASTORE, DS_UNKNOWN, DS_BINARY, DS_TEXT, DS_DELIMITED, DS_XML, DS_RELATIONAL, DS_DB2LOAD, DS_HSSUNLOAD, DS_IMSDB, DS_VSAM, DS_IMSCDC, DS_DB2CDC, DS_VSAMCDC, DS_IBMEVENT, DS_ORACLECDC, DS_GENERICCDC, DS_SUBVAR, DS_INCLUDE
                        'DS_IMSLE, DS_IMSLEBATCH,  DS_XMLCDC, DS_TRBCDC, 
                        If cNode.Text = "Datastores" Then
                            Dim parentNode As TreeNode
                            Dim childNode As TreeNode
                            For Each parentNode In cNode.Nodes
                                For Each childNode In parentNode.Nodes
                                    CType(childNode.Tag, clsDatastore).LoadMe()
                                    m_ClipObjects.Add(childNode.Tag)
                                Next
                            Next
                        Else
                            Dim childNode As TreeNode
                            For Each childNode In cNode.Nodes
                                CType(childNode.Tag, clsDatastore).LoadMe()
                                m_ClipObjects.Add(childNode.Tag)
                            Next
                        End If
                    Case NODE_FO_SOURCEDATASTORE, NODE_FO_TARGETDATASTORE, NODE_FO_VARIABLE, NODE_FO_MAIN, NODE_FO_PROC, NODE_FO_JOIN, NODE_FO_LOOKUP, NODE_FO_ENGINE, NODE_FO_CONNECTION, NODE_FO_SYSTEM, NODE_FO_ENVIRONMENT
                        Dim childNode As TreeNode
                        For Each childNode In cNode.Nodes
                            m_ClipObjects.Add(childNode.Tag)
                        Next
                End Select
                EnablePaste(True)
            Else
                ClearClipboard()
                m_ClipObjects.Add(cNode.Tag)
                EnablePaste(True)
            End If

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    '//Rerurns copy of specified clipboard object. If index is not passed then always 
    '//first clipboard object is returned index is zero based
    Function GetClipboardObject(ByVal NewParent As INode, Optional ByVal index As Integer = 0) As INode

        If index + 1 <= m_ClipObjects.Count Then
            Return CType(m_ClipObjects(index), INode).Clone(NewParent) '//return copy of object
        Else
            Return Nothing
        End If

    End Function

    Function GetClipboardObjectType(Optional ByVal index As Integer = 0) As String

        If index + 1 <= m_ClipObjects.Count Then
            Return CType(m_ClipObjects(index), INode).Type '//return object type
        Else
            '//No object in clipboard
            Return ""
        End If

    End Function

    Function ClearClipboard() As Boolean

        m_ClipObjects = Nothing
        m_ClipObjects = New ArrayList
        EnablePaste(False)

    End Function

    Function IsClipboardAvailable() As Boolean

        Return m_ClipObjects.Count

    End Function

    Function EnableCopy(ByVal bEnable As Boolean) As Boolean

        ToolBar1.Buttons(enumToolBarButtons.TB_COPY).Enabled = bEnable
        mnuCopy.Enabled = bEnable

    End Function

    Function EnablePaste(ByVal bEnable As Boolean) As Boolean

        ToolBar1.Buttons(enumToolBarButtons.TB_PASTE).Enabled = bEnable
        mnuPaste.Enabled = bEnable

    End Function

    Function EnableDelete(ByVal bEnable As Boolean) As Boolean

        ToolBar1.Buttons(enumToolBarButtons.TB_DEL).Enabled = bEnable
        mnuDelete.Enabled = bEnable

    End Function

    Private Sub CopyClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopyConn.Click, mnuCopyDS.Click, mnuCopyEngine.Click, mnuCopyEnv.Click, mnuCopyStruct.Click, mnuCopyStructSel.Click, mnuCopySystem.Click, mnuCopyTask.Click, mnuCopyVar.Click

        SetClipobjects(tvExplorer.SelectedNode)

    End Sub

    Private Sub PasteClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPasteConn.Click, mnuPasteDS.Click, mnuPasteEngine.Click, mnuPasteEnv.Click, mnuPasteStruct.Click, mnuPasteStructSel.Click, mnuPasteSystem.Click, mnuPasteTask.Click, mnuPasteVar.Click


        Dim cNode As TreeNode
        Dim newObjParent As INode = Nothing '//this is parent of new cloned object

        '//only fire paste event if comes from toolbar paste click or when explorer has focus
        If Not (sender.GetType.Name = "Toolbar" Or IsFocusOnExplorer = True) Then
            Exit Sub
        End If
        Me.Cursor = Cursors.WaitCursor

        Try
            cNode = tvExplorer.SelectedNode

            Select Case GetClipboardObjectType()
                Case NODE_STRUCT
                    '//Make parent node to one step up if node is not folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent.Parent '[Structure]->[type folder]->Struct
                        newObjParent = cNode.Parent.Tag '//environment will become the parent for new cloned obj 
                    ElseIf cNode.Text <> "Descriptions" Then
                        '//clicked on type folder
                        cNode = cNode.Parent '[Structure]->[type folder]
                        newObjParent = cNode.Parent.Tag '//environment will become the parent for new cloned obj 
                    Else '//clicked on structure folder
                        'cNode=cNode
                        newObjParent = cNode.Parent.Tag
                    End If
                Case NODE_STRUCT_SEL
                    '//Make parent node to one step up if node is selection. We add selection under Structure
                    If CType(cNode.Tag, INode).Type = NODE_STRUCT_SEL Then
                        cNode = cNode.Parent
                        newObjParent = cNode.Tag '//structure will become the parent for new cloned obj 
                    End If
                Case Else
                    '//Make parent node to one step up if node is not folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                        newObjParent = cNode.Parent.Tag '//parent for new cloned obj 
                    Else
                        newObjParent = cNode.Parent.Tag '//parent for new cloned obj 
                    End If
            End Select

            For Each obj As INode In m_ClipObjects
                Dim cmd As New System.Data.Odbc.OdbcCommand
                Dim conn As New Odbc.OdbcConnection(obj.Project.MetaConnectionString)
                Dim destObj As INode
                conn.Open()
                cmd.Connection = conn

                destObj = obj.Clone(newObjParent, True, cmd)
                conn.Close()

                If tvExplorer.SelectedNode.Text = "Sources" Then
                    CType(destObj, clsDatastore).DsDirection = "S"
                    'CType(destObj, clsDatastore).LoadAttr()

                ElseIf tvExplorer.SelectedNode.Text = "Targets" Then
                    CType(destObj, clsDatastore).DsDirection = "T"
                    'CType(destObj, clsDatastore).LoadAttr()
                End If
                '/*************************************************
                '/*** Copies Structure file to new location
                '/*************************************************
                'If obj.Type = NODE_STRUCT Then
                '    If copyBackGroundFiles(obj, CType(destObj, clsStructure)) = False Then
                '        GoTo [continue]
                '    End If
                'End If

                If HasMissingDependency(destObj, cNode, newObjParent) = True Then
                    MsgBox("One or more dependant objects are missing" & vbCrLf & _
                    "Note: Please make sure that you copy all Descriptions before you copy any object dependent on that Description.", MsgBoxStyle.Critical, MsgTitle)
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If

                Windows.Forms.Cursor.Current = Cursors.WaitCursor
                '/// modified by Tom Karasch 6/07
                '// if cancel button clicked then destobj.text = "" and will fall through to next in loop
                If destObj.Text <> "" Then

                    If ValidateNewName(destObj.Text) = True Then
                        If destObj.ValidateNewObject() = False Then
                            destObj.Text = InputBox("Please Enter Object Name", "", obj.Text)
                        End If
                    End If



                    If destObj.AddNew(True) = True Then '//8/15/05
                        'tran.Commit()

                        Select Case GetClipboardObjectType()
                            Case NODE_PROJECT
                                FillProjectFromClipboard(destObj)

                            Case NODE_ENVIRONMENT
                                FillEnvFromClipboard(cNode, destObj)

                            Case NODE_SYSTEM
                                FillSysFromClipboard(cNode, destObj)

                            Case NODE_ENGINE
                                FillEngineFromClipboard(cNode, destObj)

                            Case NODE_CONNECTION
                                FillConnFromClipboard(cNode, destObj)

                            Case NODE_STRUCT
                                FillStructFromClipboard(cNode, destObj)

                            Case NODE_STRUCT_SEL
                                FillStructSelFromClipboard(cNode, destObj)

                            Case NODE_SOURCEDATASTORE
                                FillDataStoreFromClipboard(cNode, destObj)

                            Case NODE_TARGETDATASTORE
                                FillDataStoreFromClipboard(cNode, destObj)

                            Case NODE_PROC
                                FillTasksFromClipboard(cNode, destObj)

                            Case NODE_GEN
                                FillTasksFromClipboard(cNode, destObj)

                            Case NODE_LOOKUP
                                FillTasksFromClipboard(cNode, destObj)

                            Case NODE_MAIN
                                FillTasksFromClipboard(cNode, destObj)

                            Case NODE_VARIABLE
                                FillVarFromClipboard(cNode, destObj)

                            Case Else
                                Me.Cursor = Cursors.Default
                                Exit Sub
                        End Select
                        'Else
                        '    tran.Rollback()
                    End If
                End If
[continue]: Next

            ShowUsercontrol(cNode, True)

        Catch ex As Exception
            LogError(ex)
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Function copyBackGroundFiles(ByVal ClipObj As INode, ByVal obj As clsStructure) As Boolean

        Dim returnVal As Boolean = True
        Dim oldParent As clsEnvironment
        Dim newParent As clsEnvironment
        Dim FileToCopy As String = ""
        Dim NewCopy As String = ""
        Dim DBDFileToCopy As String = ""
        Dim DBDNewCopy As String = ""

        oldParent = ClipObj.Parent()
        newParent = obj.Parent()

        Select Case CType(ClipObj, clsStructure).StructureType
            Case modDeclares.enumStructure.STRUCT_C
                FileToCopy = oldParent.LocalCDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalCDir & GetFileName(obj.fPath1)
            Case modDeclares.enumStructure.STRUCT_COBOL
                FileToCopy = oldParent.LocalCobolDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalCobolDir & GetFileName(obj.fPath1)
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                FileToCopy = oldParent.LocalDDLDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalDDLDir & GetFileName(obj.fPath1)
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                FileToCopy = oldParent.LocalDTDDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalDTDDir & GetFileName(obj.fPath1)
            Case modDeclares.enumStructure.STRUCT_COBOL_IMS
                FileToCopy = oldParent.LocalCobolDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalCobolDir & GetFileName(obj.fPath1)
                DBDFileToCopy = oldParent.LocalCobolDir & GetFileName(obj.fPath2)
                DBDNewCopy = newParent.LocalCobolDir & GetFileName(obj.fPath2)
            Case modDeclares.enumStructure.STRUCT_REL_DML, enumStructure.STRUCT_REL_DML_FILE
                FileToCopy = oldParent.LocalDMLDir & GetFileName(obj.fPath1)
                NewCopy = newParent.LocalDMLDir & GetFileName(obj.fPath1)
        End Select
        Try

            If FileToCopy <> Nothing & NewCopy <> Nothing Then
                If System.IO.File.Exists(NewCopy) = True Then
                    If MsgBox(NewCopy & "   already exists." & vbNewLine & vbNewLine & " Do you want to overwrite?" & vbNewLine & vbNewLine & "Respond YES to overwrite the file or NO to cancel copying this Structure", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, MsgTitle) = MsgBoxResult.Yes Then
                        System.IO.File.Delete(NewCopy)
                        System.IO.File.Copy(FileToCopy, NewCopy)
                    Else
                        returnVal = False
                        Return returnVal
                    End If
                Else
                    System.IO.File.Copy(FileToCopy, NewCopy)
                End If
            Else
                MsgBox("Source OR Destination path is EMPTY!!", MsgBoxStyle.Critical, MsgTitle)
                returnVal = False
            End If
            If DBDFileToCopy <> Nothing & DBDNewCopy <> Nothing Then
                If System.IO.File.Exists(DBDNewCopy) = True Then
                    If MsgBox(DBDNewCopy & "   already exists." & vbNewLine & vbNewLine & " Do you want to overwrite?" & vbNewLine & vbNewLine & "Respond YES to overwrite the file or NO to cancel copying this Structure", MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, MsgTitle) = MsgBoxResult.Yes Then
                        System.IO.File.Delete(DBDNewCopy)
                        System.IO.File.Copy(DBDFileToCopy, DBDNewCopy)
                    Else
                        returnVal = False
                        Return returnVal
                    End If
                Else
                    System.IO.File.Copy(DBDFileToCopy, DBDNewCopy)
                End If
            End If

        Catch ex As Exception
            LogError(ex)
            returnVal = False
        End Try
        Return returnVal

    End Function

    Function GetFileName(ByVal filePath As String) As String
        Dim retFilePath As String
        retFilePath = filePath.Substring(filePath.LastIndexOf("\"))
        Return (retFilePath)
    End Function

    '/obj is object being pasted. and cNode is target node underwhich object is being pasted
    'newObjParent is new parent of object being pasted. 
    Public Function ChangeLinkedObjects(ByRef obj As INode, ByVal cNode As TreeNode, ByVal newObjParent As INode) As Boolean
        Return HasMissingDependency(obj, cNode, newObjParent, True)
    End Function

    '//obj is object being pasted. and cNode is target node underwhich object is being pasted
    '//newObjParent is new parent of object being pasted. 
    '//if ChangeReference=true then obj being searched for match will be get reference of 
    '//new matched object if match found. This option will only swap reference of Structure and Selection 
    Public Function HasMissingDependency(ByRef obj As INode, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean
        Select Case obj.Type
            Case NODE_VARIABLE, NODE_CONNECTION, NODE_STRUCT
                Return False
            Case NODE_STRUCT_SEL
                Return HasMissingDependencyForStructSel(obj, cNode, newObjParent, ChangeReference)
            Case NODE_ENVIRONMENT
                Return HasMissingDependencyForEnvironment(obj, cNode, newObjParent, ChangeReference)
            Case NODE_SYSTEM
                Return HasMissingDependencyForSystem(obj, cNode, newObjParent, ChangeReference)
            Case NODE_ENGINE
                Return HasMissingDependencyForEngine(obj, cNode, newObjParent, ChangeReference)
            Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                Return HasMissingDependencyForDatastore(obj, cNode, newObjParent, ChangeReference)
            Case NODE_MAIN, NODE_PROC, NODE_LOOKUP, NODE_GEN
                Return HasMissingDependencyForTask(obj, cNode, newObjParent, ChangeReference)
            Case Else
                Return False
        End Select
    End Function

    '//If we can not find obj in current environment that means 
    Public Function HasMissingDependencyForStructSel(ByRef obj As clsStructureSelection, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean
        Dim sel As clsStructureSelection
        Dim str As clsStructure
        Dim env As clsEnvironment

        env = obj.ObjStructure.Environment

        Try
            '/// Modified by Tom Karasch 4/07
            '//Try to find Selection under target Environment, If found then make sure that it has same defination and same number of fields.
            For Each str In env.Structures
                '//System Selection represents the structure so handle it differently
                If obj.IsSystemSelection = "1" Then
                    '//Match names
                    If str.Text = obj.Text Then
                        str.LoadItems()
                        obj.LoadMe()
                        '//Match field counts
                        If str.ObjFields.Count = obj.ObjSelectionFields.Count Then
                            If ChangeReference = True Then
                                obj = str.SysAllSelection
                            End If
                            Return False
                            Exit Try
                        End If
                    End If
                Else
                    For Each sel In str.StructureSelections
                        '//Match names
                        If sel.Text = obj.Text Then
                            sel.LoadMe()
                            obj.LoadMe()
                            '//Match field counts
                            If sel.ObjSelectionFields.Count = obj.ObjSelectionFields.Count Then
                                If ChangeReference = True Then
                                    obj = sel
                                End If
                                Return False
                                Exit Try
                            End If
                        End If
                    Next
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain HasMissingDepForStrSel")
        End Try

    End Function

    Public Function HasMissingDependencyForDSSel(ByRef obj As clsDSSelection, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        Dim sel As clsStructureSelection
        Dim str As clsStructure
        Dim env As clsEnvironment

        Try
            env = obj.ObjStructure.Environment

            '/// Modified by Tom Karasch 4/07
            '//Try to find Selection under target Environment, If found then 
            '//make sure that it has same defination and same number of fields.
            For Each str In env.Structures
                '//System Selection represents the structure so handle it differently
                If obj.ObjSelection.IsSystemSelection = "1" Then
                    '//Match names
                    If str.Text = obj.Text Then
                        str.LoadItems()
                        obj.LoadMe()
                        '//Match field counts
                        If str.ObjFields.Count = obj.DSSelectionFields.Count Then
                            Return False
                            Exit Try
                        End If
                    End If
                Else
                    For Each sel In str.StructureSelections
                        '//Match names
                        If sel.Text = obj.Text Then
                            sel.LoadMe()
                            obj.LoadMe()
                            '//Match field counts
                            If sel.ObjSelectionFields.Count = obj.DSSelectionFields.Count Then
                                Return False
                                Exit Try
                            End If
                        End If
                    Next
                End If
            Next
            Return True

        Catch ex As Exception
            LogError(ex, "frmMain HasMissingDepForDSSel")
        End Try

    End Function

    Public Function HasMissingDependencyForEnvironment(ByRef obj As clsEnvironment, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        '/// Modified by Tom Karasch 6/07
        Try
            For Each o As INode In obj.Systems
                If HasMissingDependencyForSystem(o, cNode, newObjParent, ChangeReference) = True Then
                    Return True
                    Exit Try
                End If
            Next
            Return False

        Catch ex As Exception
            LogError(ex, "frmMain HasMissingDepForEnv")
        End Try

    End Function

    Public Function HasMissingDependencyForSystem(ByRef obj As clsSystem, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        '/// Modified by Tom Karasch 6/07
        Try
            For Each o As INode In obj.Engines
                If HasMissingDependencyForEngine(o, cNode, newObjParent, ChangeReference) = True Then
                    Return True
                    Exit Try
                End If
            Next
            Return False

        Catch ex As Exception
            LogError(ex, "frmMain HasMissingDepForSys")
        End Try

    End Function

    Public Function HasMissingDependencyForEngine(ByRef obj As clsEngine, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        Dim o As INode

        For Each o In obj.Sources
            If HasMissingDependencyForDatastore(o, cNode, newObjParent, ChangeReference) = True Then
                Return True
                Exit Function
            End If
        Next

        For Each o In obj.Targets
            If HasMissingDependencyForDatastore(o, cNode, newObjParent, ChangeReference) = True Then
                Return True
                Exit Function
            End If
        Next

        'For Each o In obj.Mains
        '    If HasMissingDependencyForTask(o, cNode, newObjParent, ChangeReference) = True Then
        '        Return True
        '        Exit Function
        '    End If
        'Next

        For Each o In obj.Procs
            If HasMissingDependencyForTask(o, cNode, newObjParent, ChangeReference) = True Then
                Return True
                Exit Function
            End If
        Next

        'For Each o In obj.Lookups
        '    If HasMissingDependencyForTask(o, cNode, newObjParent, ChangeReference) = True Then
        '        Return True
        '        Exit Function
        '    End If
        'Next

        'For Each o In obj.Joins
        '    If HasMissingDependencyForTask(o, cNode, newObjParent, ChangeReference) = True Then
        '        Return True
        '        Exit Function
        '    End If
        'Next

        Return False

    End Function

    Public Function HasMissingDependencyForDatastore(ByRef obj As clsDatastore, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        Dim o As INode
        Dim i As Integer
        Dim oTemp As INode

        Try
            If ChangeReference = True Then
                'obj.LoadItems() '//loads selections
                For i = 0 To obj.ObjSelections.Count - 1
                    oTemp = SearchDSSel(obj.Engine.ObjSystem.Environment, obj.ObjSelections(i))
                    If oTemp Is Nothing Then
                        clsLogging.LogEvent("Can not locate " & obj.ObjSelections(i).text & " in target environment")
                    Else
                        obj.ObjSelections(i) = oTemp
                    End If
                    Return True
                Next
                Exit Try
            End If

            For Each o In obj.ObjSelections
                If HasMissingDependencyForDSSel(o, cNode, newObjParent, ChangeReference) = True Then
                    Return True
                    Exit Try
                End If
            Next

            Return False

        Catch ex As Exception
            LogError(ex, "frmMain HasmissingDepForDS")
        End Try

    End Function

    Public Function HasMissingDependencyForTask(ByRef obj As clsTask, ByVal cNode As TreeNode, ByVal newObjParent As INode, Optional ByVal ChangeReference As Boolean = False) As Boolean

        Dim o As INode

        '// modified 4/07 and 6/07 by Tom Karasch
        Try
            '//This block will try to find matching datastores on the target environment
            If ChangeReference = True Then
                Dim i As Integer
                Dim oTemp As Object
                Dim eng As clsEngine

                eng = obj.Engine
                obj.LoadDatastores()
                obj.LoadMappings()

                For i = 0 To obj.ObjSources.Count - 1
                    oTemp = SearchDatastore(eng, obj.ObjSources(i))
                    If oTemp Is Nothing Then
                        clsLogging.LogEvent("Can not locate " & obj.TaskName & "'s datastore " & obj.ObjSources(i).text & " in target environment")
                    Else
                        obj.ObjSources(i) = oTemp
                    End If
                Next

                For i = 0 To obj.ObjTargets.Count - 1
                    oTemp = SearchDatastore(eng, obj.ObjTargets(i))
                    If oTemp Is Nothing Then
                        clsLogging.LogEvent("Can not locate " & obj.TaskName & "'s datastore " & obj.ObjSources(i).text & " in target environment")
                    Else
                        obj.ObjTargets(i) = oTemp
                    End If
                Next

                Dim map As clsMapping
                For i = 0 To obj.ObjMappings.Count - 1
                    map = obj.ObjMappings(i)

                    If map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                        '//do nothing
                    Else
                        If map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
                            map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Or _
                            map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                            '//already referencing to new copy
                        ElseIf map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_VAR Then
                            '//need to change object reference
                            oTemp = SearchVariable(eng, map.MappingSource)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Source " & map.MappingSource.text & " in target environment")
                            Else
                                obj.ObjMappings(i).MappingSource = oTemp
                            End If
                        ElseIf map.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                            '//need to change object reference
                            oTemp = SearchField(eng, map.MappingSource, modDeclares.enumDirection.DI_SOURCE)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Source " & map.MappingSource.text & " in target environment")
                            Else
                                obj.ObjMappings(i).MappingSource = oTemp
                            End If
                        Else
                            '//need to change object reference
                            oTemp = SearchTask(eng, map.MappingSource)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Source " & map.MappingSource.text & " in target environment")
                            Else
                                obj.ObjMappings(i).MappingSource = oTemp
                            End If
                        End If
                    End If

                    If map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                        '//do nothing
                    Else
                        If map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
                            map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Or _
                            map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                            '//already referencing to new copy
                        ElseIf map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_VAR Then
                            '//need to change object reference
                            oTemp = SearchTask(eng, map.MappingTarget)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Target " & map.MappingTarget.text & " in target environment")
                            Else
                                obj.ObjMappings(i).MappingTarget = oTemp
                            End If
                        ElseIf map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                            '//need to change object reference
                            oTemp = SearchField(eng, map.MappingTarget, modDeclares.enumDirection.DI_TARGET)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Target " & map.MappingTarget.text & " in target environment")
                            Else

                                obj.ObjMappings(i).MappingTarget = oTemp
                            End If
                        Else
                            '//need to change object reference
                            oTemp = SearchTask(eng, map.MappingTarget)
                            If oTemp Is Nothing Then
                                clsLogging.LogEvent("Can not locate " & map.Text & "'s Target " & map.MappingTarget.text & " in target environment")
                            Else
                                obj.ObjMappings(i).MappingTarget = oTemp
                            End If
                        End If
                    End If
                Next
                Return False
            End If

            For Each o In obj.ObjSources
                If HasMissingDependencyForDatastore(o, cNode, newObjParent, ChangeReference) = True Then
                    Return True
                    Exit Function
                End If
            Next

            For Each o In obj.ObjTargets
                If HasMissingDependencyForDatastore(o, cNode, newObjParent, ChangeReference) = True Then
                    Return True
                    Exit Function
                End If
            Next

            Return False

        Catch ex As Exception
            LogError(ex, "frmMain HasMissingDepForTask")
        End Try

    End Function

    '//Returns INode object's environment object
    Function GetEnvironment(ByRef obj As INode) As clsEnvironment

        Select Case obj.Type
            Case NODE_ENVIRONMENT
                Return obj
            Case NODE_STRUCT
                Return CType(obj, clsStructure).Environment
            Case NODE_STRUCT_SEL
                Return CType(obj, clsStructureSelection).ObjStructure.Environment
            Case NODE_SYSTEM
                Return CType(obj, clsSystem).Environment
            Case NODE_STRUCT
                Return CType(obj, clsStructure).Environment
            Case NODE_ENGINE
                Return CType(obj, clsEngine).ObjSystem.Environment
            Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                Return CType(obj, clsDatastore).Engine.ObjSystem.Environment
            Case NODE_MAIN, NODE_PROC, NODE_GEN, NODE_LOOKUP, NODE_VARIABLE
                Return CType(obj, clsTask).Engine.ObjSystem.Environment
            Case Else
                Return Nothing
        End Select

    End Function

    '//Returns INode object's engine object
    Function GetEngine(ByRef obj As INode) As clsEngine

        Select Case obj.Type
            Case NODE_ENGINE
                Return obj
            Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                Return CType(obj, clsDatastore).Engine
            Case NODE_MAIN, NODE_PROC, NODE_GEN, NODE_LOOKUP, NODE_VARIABLE
                Return CType(obj, clsTask).Engine
            Case Else
                Return Nothing
        End Select

    End Function

    Function IsClipboardAvailableforThisNodeType(ByVal ndType As String) As Boolean

        If IsClipboardAvailable() = False Then Exit Function

        If ndType = GetClipboardObjectType() Then
            Return True
            Exit Function
        End If

        Dim ClipType As String

        ClipType = GetClipboardObjectType()

        If ClipType = NODE_ENVIRONMENT And ndType = NODE_FO_ENVIRONMENT Then
            Return True
        ElseIf ClipType = NODE_SYSTEM And ndType = NODE_FO_SYSTEM Then
            Return True
        ElseIf ClipType = NODE_ENGINE And ndType = NODE_FO_ENGINE Then
            Return True
        ElseIf ClipType = NODE_CONNECTION And ndType = NODE_FO_CONNECTION Then
            Return True
        ElseIf ClipType = NODE_STRUCT And ndType = NODE_FO_STRUCT Then
            Return True
        ElseIf ClipType = NODE_SOURCEDATASTORE And ndType = NODE_FO_SOURCEDATASTORE Then
            Return True
        ElseIf ClipType = NODE_TARGETDATASTORE And ndType = NODE_FO_TARGETDATASTORE Then
            Return True
        ElseIf ClipType = NODE_PROC And ndType = NODE_FO_PROC Then
            Return True
        ElseIf ClipType = NODE_GEN And ndType = NODE_FO_JOIN Then
            Return True
        ElseIf ClipType = NODE_LOOKUP And ndType = NODE_FO_LOOKUP Then
            Return True
        ElseIf ClipType = NODE_MAIN And ndType = NODE_FO_MAIN Then
            Return True
        ElseIf ClipType = NODE_VARIABLE And ndType = NODE_FO_VARIABLE Then
            Return True
        Else
            Return False
        End If

    End Function

    Function FillProjectFromClipboard(ByVal obj As clsProject) As Boolean

        Dim i As Integer
        Dim objClip As clsProject
        Dim foldObj As clsFolderNode

        objClip = obj

        Try
            Dim cNode As TreeNode

            If obj.AddNew(True) = True Then
                '//Add project node
                cNode = AddTreeNode(tvExplorer, NODE_PROJECT, obj)
                obj.SeqNo = cNode.Index '//store position

                '//Now add and process each environment 
                '//Add Environment folder node
                foldObj = New clsFolderNode("Environments", NODE_FO_ENVIRONMENT)
                foldObj.Parent = CType(cNode.Tag, INode)
                cNode = AddNode(cNode, foldObj.Type, foldObj)
                obj.SeqNo = cNode.Index '//store position

                For i = 0 To objClip.Environments.Count - 1
                    '//Process this env
                    FillEnvFromClipboard(cNode, objClip.Environments(i + 1))
                Next
            End If


        Catch ex As Exception
            LogError(ex, "frmMain FillProjForClipboard")
        End Try

    End Function

    Function FillEnvFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsEnvironment) As Boolean

        Dim i As Integer
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            obj.LoadMe()

            obj.Parent = CType(cNode.Tag, INode).Parent '//Project->[Env Folder]->Env
            'obj.Project = CType(cNode.Tag, INode).Project
            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add connections
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under sys
            Dim cNodeCnn As TreeNode
            Dim objCnn As INode
            objCnn = New clsFolderNode("Connections", NODE_FO_CONNECTION)
            objCnn.Parent = CType(cNode.Tag, INode)
            'objCnn.Project = CType(cNode.Tag, INode).Project
            cNodeCnn = AddNode(cNode, objCnn.Type, objCnn)

            For i = 0 To obj.Connections.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillConnFromClipboard(cNodeCnn, obj.Connections(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add structures
        '///////////////////////////////////////////////
        Try
            '//Add Structure folder node under env
            Dim cNodeStruct As TreeNode
            Dim objStruct As INode
            objStruct = New clsFolderNode("Descriptions", NODE_FO_STRUCT)
            objStruct.Parent = CType(cNode.Tag, INode)
            'objStruct.Project = CType(cNode.Tag, INode).Project
            cNodeStruct = AddNode(cNode, objStruct.Type, objStruct)

            For i = 0 To obj.Structures.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillStructFromClipboard(cNodeStruct, obj.Structures(i + 1))
            Next
        Catch ex As Exception
            LogError(ex)
        End Try

        '//////////////////////////////////////////////
        '//Now add Datastores
        '//////////////////////////////////////////////
        'Try
        '    Dim cNode1 As TreeNode
        '    Dim obj1 As INode

        '    obj1 = New clsFolderNode("Datastores", NODE_FO_DATASTORE)
        '    obj1.Parent = CType(cNode.Tag, INode)
        '    cNode1 = AddNode(cNode.Nodes, obj1.Type, obj1)

        '    For i = 0 To obj.Datastores.Count - 1
        '        ''//Process this (cNode1 is root node under which we add other structures)
        '        FillDSbyTypeFromClipboard(cNode1, obj.Datastores(i + 1))
        '    Next


        'Catch ex As Exception
        '    LogError(ex)
        'End Try

        '///////////////////////////////////////////////
        '//Now add Variables
        '///////////////////////////////////////////////
        Try
            '//Add variables folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = obj
            cNodeVar = AddNode(cNode, objVar.Type, objVar)

            For i = 0 To obj.Variables.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVarFromClipboard(cNodeVar, obj.Variables(i + 1))
            Next

        Catch ex As Exception

        End Try

        '///////////////////////////////////////////////
        '//Now add Procedures
        '///////////////////////////////////////////////
        'Try
        '    Dim cNodeProc As TreeNode
        '    Dim objProc As INode

        '    '//Add Proc folder
        '    objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
        '    objProc.Parent = CType(cNode.Tag, INode)
        '    cNodeProc = AddNode(cNode.Nodes, objProc.Type, objProc, False)

        '    For i = 0 To obj.Procedures.Count - 1
        '        ''//Process this (cNode is root node under which we add other nodes)
        '        FillTasksFromClipboard(cNodeProc, obj.Procedures(i + 1))
        '    Next

        'Catch ex As Exception

        'End Try

        '///////////////////////////////////////////////
        '//Now add systems
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Systems", NODE_FO_SYSTEM)
            objSys.Parent = CType(cNode.Tag, INode)
            'objSys.Project = CType(cNode.Tag, INode).Project
            cNodeSys = AddNode(cNode, objSys.Type, objSys)

            For i = 0 To obj.Systems.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillSysFromClipboard(cNodeSys, obj.Systems(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillStructFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsStructure) As Boolean

        Dim i As Integer

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Parent = CType(cNode.Parent.Tag, INode) 'Env->StructFolder->Struct
            cNode = AddStructNode(obj, cNode)

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add fieldselections
        '///////////////////////////////////////////////
        Try
            For i = 0 To obj.StructureSelections.Count - 1
                '//Process this (cNode is root node under which we add other nodes)
                FillStructSelFromClipboard(cNode, obj.StructureSelections(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillStructSelFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsStructureSelection) As Boolean

        If obj.IsSystemSelection = "1" Then
            Exit Function '//dont add this node if system selection
        End If

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Parent = CType(cNode.Tag, INode)
            'obj.Project = CType(cNode.Tag, INode).Project

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    'Function FillDSbyTypeFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsDatastore) As Boolean

    '    Try
    '        obj.Parent = CType(cNode.Parent.Tag, INode)

    '        'obj.LoadItems(, True)

    '        'obj.LoadAttr()

    '        'AddToCollection(obj.Environment.Datastores, obj, obj.GUID)

    '        'ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

    '        cNode = AddDSNode(obj, cNode)
    '        obj.SeqNo = cNode.Index '//store position

    '        cNode.Text = obj.DsPhysicalSource

    '        '// now add Datastore selections to tree
    '        AddDSstructuresToTree(cNode, obj)
    '        cNode.Collapse()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "frmMain FillDSbyTypeFromClipboard")
    '        Return False
    '    End Try

    'End Function

    Function FillSysFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsSystem) As Boolean

        Dim i As Integer
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys
            'obj.Project = CType(cNode.Tag, INode).Project

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

        '//TODO : Add Engines, Connections under this system...

        '///////////////////////////////////////////////
        '//Now add engines
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Engines", NODE_FO_ENGINE)
            objSys.Parent = CType(cNode.Tag, INode)
            'objSys.Project = CType(cNode.Tag, INode).Project
            cNodeSys = AddNode(cNode, objSys.Type, objSys)

            For i = 0 To obj.Engines.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillEngineFromClipboard(cNodeSys, obj.Engines(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillEngineFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsEngine) As Boolean

        Dim i As Integer
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Parent = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys
            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now source datastore
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under sys
            Dim cNode1 As TreeNode
            Dim obj1 As clsFolderNode

            obj1 = New clsFolderNode("Sources", NODE_FO_SOURCEDATASTORE)
            obj1.Parent = obj
            cNode1 = AddNode(cNode, obj1.Type, obj1)

            cNode1.Text = "(" & obj.Sources.Count.ToString & ")" & " Sources"

            For i = 0 To obj.Sources.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStoreFromClipboard(cNode1, obj.Sources(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add targets
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under sys
            Dim cNode2 As TreeNode
            Dim obj2 As clsFolderNode
            obj2 = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
            obj2.Parent = obj
            cNode2 = AddNode(cNode, obj2.Type, obj2)

            cNode2.Text = "(" & obj.Targets.Count.ToString & ")" & " Targets"

            For i = 0 To obj.Targets.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStoreFromClipboard(cNode2, obj.Targets(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = obj
            cNodeVar = AddNode(cNode, objVar.Type, objVar)

            cNodeVar.Text = "(" & obj.Variables.Count.ToString & ")" & " Variables"

            For i = 0 To obj.Variables.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVarFromClipboard(cNodeVar, obj.Variables(i + 1))
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add tasks (main, join, lookup, proc)
        '///////////////////////////////////////////////
        Try
            Dim cNodeProc As TreeNode
            Dim cNodeMain As TreeNode
            Dim objProc As INode
            Dim objMain As INode

            ''//Add Join folder
            'objJoin = New clsFolderNode("Join", NODE_FO_JOIN)
            'objJoin.Parent = obj
            'cNodeJoin = AddNode(cNode.Nodes, objJoin.Type, objJoin)

            'For i = 0 To obj.Joins.Count - 1
            '    FillTasksFromClipboard(cNodeJoin, obj.Joins(i + 1))
            'Next
            'For i = 0 To obj.Lookups.Count - 1
            '    FillTasksFromClipboard(cNodeJoin, obj.Lookups(i + 1))
            'Next

            '//Add Proc folder
            objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
            objProc.Parent = obj
            cNodeProc = AddNode(cNode, objProc.Type, objProc)

            cNodeProc.Text = "(" & obj.Procs.Count.ToString & ")" & " Procedures"

            For i = 0 To obj.Procs.Count - 1
                FillTasksFromClipboard(cNodeProc, obj.Procs(i + 1))
            Next

            'For i = 0 To obj.Joins.Count - 1
            '    FillTasksFromClipboard(cNodeJoin, obj.Joins(i + 1))
            'Next
            'For i = 0 To obj.Lookups.Count - 1
            '    FillTasksFromClipboard(cNodeJoin, obj.Lookups(i + 1))
            'Next
            '//Add Main folder
            objMain = New clsFolderNode("Main Procedure(s)", NODE_FO_MAIN)
            objMain.Parent = obj
            cNodeMain = AddNode(cNode, objMain.Type, objMain)

            For i = 0 To obj.Mains.Count - 1
                FillTasksFromClipboard(cNodeMain, obj.Mains(i + 1))
            Next

            ''//Add Lookup folder
            'objLook = New clsFolderNode("Lookup", NODE_FO_LOOKUP)
            'objLook.Parent = obj
            'cNodeLookup = AddNode(cNode.Nodes, objLook.Type, objLook)

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillConnFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsConnection) As Boolean

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Parent = CType(cNode.Parent.Tag, INode) 'Env->ConnFolder->Conn
            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillDataStoreFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsDatastore, Optional ByRef cmd As Odbc.OdbcCommand = Nothing) As Boolean

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'If CType(cNode.Parent.Tag, INode).Type = NODE_ENGINE Then
            obj.Parent = CType(cNode.Parent.Tag, clsEngine) 'Engine->Folder(source/target)->ds
            'Else
            ''If CType(cNode.Parent.Tag, INode).Type = NODE_ENVIRONMENT Then
            ''    obj.Parent = CType(cNode.Parent.Tag, clsEnvironment)
            ''End If
            'FillDSbyTypeFromClipboard(cNode, obj)
            'End If

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    obj.LoadAttr()
            'End If
            'obj.LoadItems()   'make sure it is loaded

            'obj.SetDSselParents()   'set all the DSselection parents, based on Fkey
            cNode = AddNode(cNode, obj.Type, obj)   ' add node to the tree
            obj.SeqNo = cNode.Index '//store position



            If cmd Is Nothing Then ' save datastore so that meta is updated correctly
                obj.Save()
                cNode.Text = obj.DsPhysicalSource
                AddDSstructuresToTree(cNode, obj)  'add it's DSselections under it
            Else
                obj.MapAsSave(cmd)
                cNode.Text = obj.Text
                AddDSstructuresToTree(cNode, obj, True)  'add it's DSselections under it
            End If


            Return True

        Catch ex As Exception
            LogError(ex)
            Return False
        End Try

    End Function

    Function FillTasksFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsTask) As Boolean

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            obj.LoadMe()
            'obj.LoadDatastores()
            'obj.LoadMappings(True)

            obj.Parent = CType(cNode.Parent.Tag, INode) 'Engine->TaskFolder->Any Task
            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillVarFromClipboard(ByVal cNode As TreeNode, ByVal obj As clsVariable) As Boolean

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            'obj.LoadItems()

            obj.Parent = CType(cNode.Parent.Tag, INode) 'Engine->VarFolder->Var
            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

#End Region

    '//////// Generalized Functions for the Main Form ////////
#Region "Non-Node Specific Main Form Functions"

    '//Purpose: Moves tree node in the same level
    Function MoveNode(ByVal nd As TreeNode, ByVal Pos As Integer) As Integer

        Dim cNode As TreeNodeCollection

        '//if new position is same as current then dont do anything
        If Pos = nd.Index Then
            Exit Function
        End If

        If nd.Parent Is Nothing Then
            cNode = nd.TreeView.Nodes
            cNode.RemoveAt(nd.Index)
            cNode.Insert(Pos, nd)
        Else
            cNode = nd.Parent.Nodes
            cNode.RemoveAt(nd.Index)
            cNode.Insert(Pos, nd)
        End If

        Return nd.Index

    End Function

    Sub ShowLog()
        Dim frm As New frmLog ': 8/15/05
        frm.ShowLog()
    End Sub

    Private Function DoAction(ByVal ActionType As enumAction, Optional ByVal IsClipboardAction As Boolean = False, Optional ByVal ClipboardObjIndex As Integer = 0) As Boolean

        Dim obj As INode

        Try
            '//If no node present in the treeview then we execute project Addnew action 
            If tvExplorer.GetNodeCount(True) = 0 Then
                Select Case ActionType
                    Case modDeclares.enumAction.ACTION_NEW, modDeclares.enumAction.ACTION_OPEN
                        DoAction = DoProjectAction(ActionType)
                End Select
                Exit Function
            End If

            obj = tvExplorer.SelectedNode.Tag

            '//Global procedure for all type node delete 
            If ActionType = modDeclares.enumAction.ACTION_DELETE Then
                DoAction = DoDeleteAction()
                Exit Function
            End If

            If Not (obj Is Nothing) Then
                '// determine what type of treenode object is selected and 
                '//do the corresponding action type to the selected treenode object
                Select Case obj.Type
                    Case NODE_PROJECT
                        DoAction = DoProjectAction(ActionType)
                    Case NODE_ENVIRONMENT, NODE_FO_ENVIRONMENT
                        DoAction = DoEnvironmentAction(ActionType)
                    Case NODE_STRUCT, NODE_FO_STRUCT
                        DoAction = DoStructureAction(ActionType)
                    Case NODE_STRUCT_SEL, NODE_FO_STRUCT_SEL
                        DoAction = DoStructureSelectionAction(ActionType)
                    Case NODE_SYSTEM, NODE_FO_SYSTEM
                        DoAction = DoSystemAction(ActionType)
                    Case NODE_CONNECTION, NODE_FO_CONNECTION
                        DoAction = DoConnectionAction(ActionType)
                    Case NODE_ENGINE, NODE_FO_ENGINE
                        DoAction = DoEngineAction(ActionType)
                    Case NODE_SOURCEDATASTORE, NODE_SOURCEDSSEL
                        DoAction = DoDatastoreAction(ActionType)
                    Case NODE_TARGETDATASTORE, NODE_TARGETDSSEL
                        DoAction = DoDatastoreAction(ActionType)
                    Case NODE_VARIABLE
                        DoAction = DoVarAction(ActionType)
                    Case NODE_PROC, NODE_FO_PROC
                        DoAction = DoTaskAction(ActionType, modDeclares.enumTaskType.TASK_PROC)
                    Case NODE_GEN, NODE_FO_JOIN
                        DoAction = DoTaskAction(ActionType, modDeclares.enumTaskType.TASK_GEN)
                    Case NODE_LOOKUP, NODE_FO_LOOKUP
                        DoAction = DoTaskAction(ActionType, modDeclares.enumTaskType.TASK_LOOKUP)
                    Case NODE_MAIN, NODE_FO_MAIN
                        DoAction = DoTaskAction(ActionType, modDeclares.enumTaskType.TASK_MAIN)
                    Case Else
                        '//TODO
                End Select
            End If

        Catch ex As Exception
            LogError(ex, "frmMain DoAction")
            DoAction = False
        End Try

    End Function

    Function DoDeleteAction() As Boolean
        '/// Function Overhauled by Tom Karasch 3/07

        Try
            Dim inodetype As String = CType(tvExplorer.SelectedNode.Tag, INode).Type
            Dim pNode As TreeNode = tvExplorer.SelectedNode.Parent
            Dim selNode As TreeNode = tvExplorer.SelectedNode

            If CType(tvExplorer.SelectedNode.Tag, INode).IsFolderNode = True Then
                If MsgBox("Are you sure you want to delete everything in this folder?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
                    If Not (tvExplorer.SelectedNode Is Nothing) Then
                        Dim cnode As TreeNode = tvExplorer.SelectedNode
                        Dim Flag As Boolean = True
                        Dim i As Int32
                        Dim range As Int32 = cnode.Nodes.Count - 1
                        Dim ChildNode As TreeNode
                        '/// Nodes must be deleted in reverse order (don't ask me why)
                        '/// also this cannot be done with a "for each ..." (don't ask me why again)
                        For i = range To 0 Step -1
                            ChildNode = cnode.Nodes.Item(i)
                            '// until the current node (cnode) is empty .. do a recursive delete
                            Flag = DoRecursiveDelete(ChildNode, inodetype)
                            If Flag = False Then
                                Exit For
                            End If
                        Next
                        If cnode.Parent.Text.Contains("Descriptions") Then
                            cnode.Remove()
                        End If
                        DoDeleteAction = Flag
                    End If
                End If
                If tvExplorer.SelectedNode IsNot Nothing Then
                    UpdateParentNodeCount(selNode)
                    tvExplorer.SelectedNode.Expand()
                End If
            Else
                If inodetype = NODE_STRUCT Or inodetype = NODE_STRUCT_SEL Then
                    If MsgBox("*** Warning ***" & (Chr(13)) & "All references to this Description, throughout the project," & (Chr(13)) & " will be deleted to maintain referential integrity of MetaData." & (Chr(13)) & "(This will delete Description from Datastores and Mappings)", MsgBoxStyle.OkCancel, MsgTitle) = MsgBoxResult.Ok Then
                        If Not (tvExplorer.SelectedNode Is Nothing) Then
                            '// until the current node is empty .. do a recursive delete
                            DoDeleteAction = DoRecursiveDelete(tvExplorer.SelectedNode, inodetype)
                        End If
                    End If
                Else
                    If MsgBox("Are you sure you want to delete selected item and sub items?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
                        If Not (tvExplorer.SelectedNode Is Nothing) Then
                            '// until the current node is empty .. do a recursive delete
                            DoDeleteAction = DoRecursiveDelete(tvExplorer.SelectedNode, inodetype)

                        End If
                    End If
                End If
                If tvExplorer.SelectedNode IsNot Nothing Then
                    UpdateParentNodeCount(pNode)
                    tvExplorer.SelectedNode.Expand()
                End If
            End If

            DoDeleteAction = True

        Catch ex As Exception
            LogError(ex, "frmMain DoDeleteAction")
            DoDeleteAction = False
        End Try

    End Function

    '//Delete selected node and all children nodes
    Function DoRecursiveDelete(ByVal cNode As TreeNode, Optional ByVal InodeType As String = "") As Boolean
        '/// OverHauled By Tom Karasch 4/07

        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim Successful As Boolean
        Dim tempnode As TreeNode = tvExplorer.SelectedNode.PrevNode
        Dim nodetext As String = cNode.Text

        Try
            If tempnode Is Nothing Then
                tempnode = CurLoadedProject.ObjTreeNode
            End If

            If cNode Is Nothing Then
                '// if the current node is empty, the exit this function
                Exit Function
            End If


            cmd.Connection = cnn
            '//We need to put in transaction because we will add structure and 
            '//fields in two steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            Successful = CType(cNode.Tag, INode).Delete(cmd, cnn)

            '//Delete all children under this node
            If Successful = True Then
                tran.Commit()

                cNode.Nodes.Clear()
                cNode.Remove()
                ctFolder.Clear()
                DoRecursiveDelete = True
                HideAllUC()
                If InodeType = NODE_STRUCT Or InodeType = NODE_STRUCT_SEL Then
                    If CurLoadedProject IsNot Nothing Then
                        FillProject(CurLoadedProject, False, True)
                    End If
                    '' vv Was here vv
                End If
                tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, tempnode.Text)
            Else
                tran.Rollback()
                Return False
                Exit Try
            End If

            Return True

        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "frmMain DoRecursiveDelete")
            Return False
        End Try

    End Function

    Sub DoModelAction(Optional ByVal InType As String = "", Optional ByVal OutType As String = "DDL")

        '//Generate DDL, etc....
        Dim obj As INode
        Dim retPath As String
        Dim strSaveDir As String
        Dim NewName As String
        Dim StrObj As clsStructure
        Dim envobj As clsEnvironment

        Try

            obj = CType(tvExplorer.SelectedNode.Tag, INode)

            If obj.IsFolderNode = False Then
                If obj.Type = NODE_STRUCT Then
                    envobj = CType(obj, clsStructure).Environment
                ElseIf obj.Type = NODE_STRUCT_SEL Then
                    envobj = CType(obj, clsStructureSelection).ObjStructure.Environment
                ElseIf obj.Type = NODE_SOURCEDSSEL Or obj.Type = NODE_TARGETDSSEL Then
                    envobj = CType(obj, clsDSSelection).ObjStructure.Environment
                Else
                    envobj = Nothing
                End If

                envobj.LoadMe()

                Select Case OutType
                    Case "DTD"
                        strSaveDir = envobj.LocalDTDDir
                    Case "DDL"
                        strSaveDir = envobj.LocalDDLDir
                    Case "H"
                        strSaveDir = envobj.LocalCDir
                    Case "LOD"
                        strSaveDir = envobj.LocalDMLDir
                    Case "SQL"
                        strSaveDir = envobj.LocalDMLDir
                    Case "MSSQL"
                        strSaveDir = envobj.LocalDMLDir
                    Case Else
                        strSaveDir = OutType
                End Select
                FolderBrowserDialog1.SelectedPath = strSaveDir

                If FolderBrowserDialog1.ShowDialog(Me) <> Windows.Forms.DialogResult.OK Then
                    Exit Sub
                Else
                    strSaveDir = FolderBrowserDialog1.SelectedPath
                End If

                NewName = InputBox("Please Name Your Model", "Model Name", "m" & obj.Text)


                'retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                'strSaveDir, obj.Text, obj.Project.TablePrefix)



                '//////////////////// NEW MODELER /////////////////////
                'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                '//////////////////////////////////////////////////////
                Select Case OutType
                    Case "DTD"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case "DDL"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case "H"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case "LOD"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case "SQL"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case "MSSQL"
                        'retPath = ModelStructure(obj, OutType, NewName, strSaveDir)
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                    Case Else
                        retPath = ModelScript(obj.Project.MetaConnectionString, obj.Key, NewName, InType, OutType, _
                        strSaveDir, obj.Text, obj.Project.TablePrefix)
                End Select

                If retPath <> "" Then
                    Process.Start(retPath)
                    'Shell("notepad.exe " & Quote(retPath, """"), AppWinStyle.NormalFocus)
                Else
                    MsgBox("There was a problem modeling " & obj.Text, MsgBoxStyle.Information, MsgTitle)
                    Log("Modeling of " & obj.Text & " was unsuccessful")
                End If
            Else
                Dim success As Integer = 0
                Dim Total As Integer = 0

                envobj = CType(tvExplorer.SelectedNode.Nodes(1).Tag, clsStructure).Environment

                envobj.LoadMe()
                'FolderBrowserDialog1.SelectedPath = envobj.LocalDBDDir
                'If FolderBrowserDialog1.ShowDialog(Me) <> Windows.Forms.DialogResult.OK Then
                '    Exit Sub
                'Else
                '    strSaveDir = FolderBrowserDialog1.SelectedPath
                'End If
                Select Case OutType
                    Case "DTD"
                        strSaveDir = envobj.LocalDTDDir
                    Case "DDL"
                        strSaveDir = envobj.LocalDDLDir
                    Case "H"
                        strSaveDir = envobj.LocalCDir
                    Case "LOD"
                        strSaveDir = envobj.LocalDMLDir
                    Case "SQL"
                        strSaveDir = envobj.LocalDMLDir
                    Case "MSSQL"
                        strSaveDir = envobj.LocalDMLDir
                    Case Else
                        strSaveDir = OutType
                End Select

                For Each StrNode As TreeNode In tvExplorer.SelectedNode.Nodes
                    StrObj = CType(StrNode.Tag, clsStructure)
                    NewName = "m" & StrObj.Text

                    'retPath = ModelStructure(StrObj, OutType, NewName, strSaveDir)
                    retPath = ModelScript(StrObj.Project.MetaConnectionString, StrObj.Key, NewName, InType, OutType, _
                    strSaveDir, StrObj.Text, StrObj.Project.TablePrefix)

                    If retPath = "" Then
                        MsgBox("There was a problem modeling " & StrObj.Text, MsgBoxStyle.Information, MsgTitle)
                        Log("Modeling of " & StrObj.StructureName & " was unsuccessful")
                    Else
                        success += 1
                    End If
                    Total += 1
                Next
                MsgBox("You Successfully created " & success & " models, out of " & Total & " attempted." & Chr(13) & "New Models were written to directory:" & Chr(13) & strSaveDir & Chr(13) & "See the log for a list of Failures", MsgBoxStyle.Information, "Structure Modelling")
            End If

        Catch ex As Exception
            LogError(ex, "frmMain DoModelAction")
        End Try

    End Sub

    Function GetTopLevelNode(ByVal nd As TreeNode) As TreeNode

        Do While Not nd Is Nothing
            If nd.Parent Is Nothing Then Exit Do
            nd = nd.Parent
        Loop

        GetTopLevelNode = nd

    End Function

    Function EnableTreeActionButton(ByVal bState As Boolean) As Boolean

        Dim i As Integer

        For i = 0 To ToolBar1.Buttons.Count - 1
            Select Case ToolBar1.Buttons(i).Text
                Case "Delete", "Properties", "Cut", "Copy", "Paste", "Query", "Map", "TreeToggle"
                    ToolBar1.Buttons(i).Enabled = bState
                Case "Save"
                    ToolBar1.Buttons(i).Enabled = IsModified
            End Select
        Next

    End Function

    Function SavePreviousScreen(ByVal cNode As TreeNode) As Boolean

        Dim IsFolder As Boolean
        Dim tb As ToolBarButton

        '//Before we switch to different view lets make sure that previous form is save
        Dim ret As MsgBoxResult
        ret = DoSave()

        If ret = MsgBoxResult.Cancel Then
            If Not prevNode Is Nothing Then
                IsEventFromCode = True
                tvExplorer.SelectedNode = prevNode
                IsEventFromCode = False
            End If
            Exit Function
        End If
        IsEventFromCode = False

        prevNode = cNode

        For Each tb In ToolBar1.Buttons
            tb.Enabled = False
        Next

        IsFolder = CType(cNode.Tag, INode).IsFolderNode

        ToolBar1.Buttons(enumToolBarButtons.TB_PROJECT).Enabled = True
        ToolBar1.Buttons(enumToolBarButtons.TB_OPEN).Enabled = True
        ToolBar1.Buttons(enumToolBarButtons.TB_HELP).Enabled = True
        ToolBar1.Buttons(enumToolBarButtons.TB_LOG).Enabled = True
        ToolBar1.Buttons(enumToolBarButtons.TB_MAP).Enabled = True
        ToolBar1.Buttons(enumToolBarButtons.TB_TOGGLE).Enabled = True

        EnableCopy(Not IsFolder)
        EnablePaste(IsClipboardAvailableforThisNodeType(CType(cNode.Tag, INode).Type))
        EnableDelete(Not IsFolder)

        HideAllUC()

        Return True

    End Function

    Sub ShowUsercontrol(ByVal nd As TreeNode, Optional ByVal ClearMsg As Boolean = False)

        '/// Modified Numerous times 12/06 thru 4/07 by Tom Karasch
        Dim NodeType As String
        Dim DSSelObj As clsDSSelection

        Try
            If CurLoadedProject Is Nothing Then
                CurLoadedProject = CType(nd.Tag, INode).Project
                cnn = New Odbc.OdbcConnection(CurLoadedProject.MetaConnectionString)
                cnn.Open()
            Else
                If CType(nd.Tag, INode).Project IsNot CurLoadedProject Then
                    CurLoadedProject = CType(nd.Tag, INode).Project
                    cnn = New Odbc.OdbcConnection(CurLoadedProject.MetaConnectionString)
                    cnn.Open()
                End If
            End If

            NodeType = CType(nd.Tag, INode).Type
            objTopMost = Nothing

            'If ActiveQueryDSlist.Count <= 0 Then
            '    lblStatusMsg.Text = ""
            'End If
            If ClearMsg = True Then
                lblStatusMsg.Text = ""
            End If

            Select Case NodeType
                Case NODE_PROJECT
                    ToolBar1.Buttons(enumToolBarButtons.TB_SCRIPT).Enabled = False
                    EnableCopy(False)
                    objTopMost = ctPrj.EditObj(nd.Tag)
                Case NODE_ENVIRONMENT
                    ToolBar1.Buttons(enumToolBarButtons.TB_ENV).Enabled = True
                    If NodeType = NODE_ENVIRONMENT Then objTopMost = ctEnv.EditObj(nd.Tag)
                Case NODE_STRUCT
                    ToolBar1.Buttons(enumToolBarButtons.TB_STRUCT).Enabled = True
                    If NodeType = NODE_STRUCT Then
                        objTopMost = ctStr.EditObj(nd.Tag)
                        nd.Expand()
                    End If
                Case NODE_STRUCT_SEL
                    ToolBar1.Buttons(enumToolBarButtons.TB_STRUCTSEL).Enabled = True
                    If NodeType = NODE_STRUCT_SEL Then
                        objTopMost = ctStrSel.EditObj(nd.Tag)
                        nd.Expand()
                    End If
                Case NODE_SYSTEM
                    ToolBar1.Buttons(enumToolBarButtons.TB_SYS).Enabled = True
                    If NodeType = NODE_SYSTEM Then objTopMost = ctSys.EditObj(nd.Tag)
                Case NODE_CONNECTION
                    ToolBar1.Buttons(enumToolBarButtons.TB_CONN).Enabled = True
                    If NodeType = NODE_CONNECTION Then objTopMost = ctConn.EditObj(nd.Tag)
                Case NODE_ENGINE
                    ToolBar1.Buttons(enumToolBarButtons.TB_ENGINE).Enabled = True
                    ToolBar1.Buttons(enumToolBarButtons.TB_SCRIPT).Enabled = True
                    If NodeType = NODE_ENGINE Then objTopMost = ctEng.EditObj(nd.Tag)
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    ToolBar1.Buttons(enumToolBarButtons.TB_DATASTORE).Enabled = True
                    If NodeType = NODE_SOURCEDATASTORE Or NodeType = NODE_TARGETDATASTORE Then
                        ToolBar1.Buttons(enumToolBarButtons.TB_QUERY).Enabled = False
                        nd.ExpandAll()
                        If CType(nd.Tag, clsDatastore).DatastoreType = enumDatastore.DS_INCLUDE Then
                            ctInc.Refresh()
                            objTopMost = ctInc.EditObj(nd.Tag)
                        Else
                            ctDs.Refresh()
                            objTopMost = ctDs.EditObj(nd.Tag, nd)
                        End If
                    End If
                Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                    ToolBar1.Buttons(enumToolBarButtons.TB_DATASTORE).Enabled = True
                    ToolBar1.Buttons(enumToolBarButtons.TB_QUERY).Enabled = True
                    nd.ExpandAll()
                    DSSelObj = nd.Tag
                    objTopMost = ctDs.EditObj(DSSelObj.ObjDatastore, DSSelObj.ObjDatastore.ObjTreeNode, nd)
                Case NODE_VARIABLE
                    ToolBar1.Buttons(enumToolBarButtons.TB_VAR).Enabled = True
                    If NodeType = NODE_VARIABLE Then objTopMost = ctVar.EditObj(nd.Tag)
                Case NODE_MAIN
                    ToolBar1.Buttons(enumToolBarButtons.TB_MAIN).Enabled = True
                    If NodeType = NODE_MAIN Then objTopMost = ctMain.EditObj(nd.Parent.Parent, nd.Tag)
                Case NODE_PROC
                    ToolBar1.Buttons(enumToolBarButtons.TB_PROC).Enabled = True
                    If CType(nd.Tag, clsTask).TaskType = enumTaskType.TASK_IncProc Then
                        ctInc.Refresh()
                        objTopMost = ctInc.EditObj(nd.Tag)
                    Else
                        ctTask.Refresh()
                        objTopMost = ctTask.EditObj(nd.Parent.Parent, nd.Tag)
                    End If
                    'If NodeType = NODE_PROC Then objTopMost = ctTask.EditObj(nd.Parent.Parent, nd.Tag)
                Case NODE_GEN
                    ToolBar1.Buttons(enumToolBarButtons.TB_MAIN).Enabled = True
                    If NodeType = NODE_GEN Then objTopMost = ctMain.EditObj(nd.Parent.Parent, nd.Tag)
                Case NODE_LOOKUP
                    ToolBar1.Buttons(enumToolBarButtons.TB_LOOKUP).Enabled = True
                    If NodeType = NODE_LOOKUP Then objTopMost = ctTask.EditObj(nd.Parent.Parent, nd.Tag)
                Case Else
                    If CType(nd.Tag, INode).IsFolderNode = True Then
                        objTopMost = ctFolder.EditObj(nd, nd.Tag)
                        nd.Expand()
                    End If
                    '//TODO
            End Select

        Catch ex As Exception
            LogError(ex, "frmMain ShowUserControl")
        End Try

    End Sub

    '//Make sure that opened for
    '//Returns : MsgBoxResult.Abort is no need to save 
    Function DoSave(Optional ByVal Prompt As Boolean = True) As MsgBoxResult

        Dim ret As MsgBoxResult = MsgBoxResult.Abort

        If objTopMost Is Nothing Then DoSave = MsgBoxResult.Abort : Exit Function

        If objTopMost.IsModified = True Then

            If Prompt = True Then
                ret = MsgBox("Do you want to save change(s) made to the opened form?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel, "Save Changes")
            Else
                ret = MsgBoxResult.Yes
            End If

            Select Case ret
                Case MsgBoxResult.Yes
                    Select Case objTopMost.Type
                        Case NODE_PROJECT
                            DoSave = IIf(ctPrj.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_ENVIRONMENT
                            DoSave = IIf(ctEnv.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_STRUCT
                            DoSave = IIf(ctStr.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_STRUCT_SEL
                            DoSave = IIf(ctStrSel.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_SYSTEM
                            DoSave = IIf(ctSys.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_CONNECTION
                            DoSave = IIf(ctConn.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_ENGINE
                            DoSave = IIf(ctEng.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_SOURCEDATASTORE
                            DoSave = IIf(ctDs.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_TARGETDATASTORE
                            DoSave = IIf(ctDs.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_VARIABLE
                            DoSave = IIf(ctVar.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_PROC
                            DoSave = IIf(ctTask.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_GEN
                            DoSave = IIf(ctTask.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_LOOKUP
                            DoSave = IIf(ctTask.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                        Case NODE_MAIN
                            DoSave = IIf(ctMain.Save(), MsgBoxResult.Yes, MsgBoxResult.No)
                            'Case NODE_INCLUDE

                    End Select
                    tbtnSave.Enabled = False
                Case MsgBoxResult.No
                    DoSave = MsgBoxResult.No
                    If objTopMost.Type = NODE_VARIABLE Or objTopMost.Type = NODE_PROC Or _
                       objTopMost.Type = NODE_GEN Or objTopMost.Type = NODE_LOOKUP Or _
                       objTopMost.Type = NODE_MAIN Then
                        CType(objTopMost, clsTask).IsDisturbed = True
                    End If
                Case MsgBoxResult.Cancel
                    DoSave = MsgBoxResult.Cancel
            End Select
        End If

    End Function

    '//////// Sub Added by TK and KS 11/6/2006 ////////
    Private Sub MapAs(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMapAsBinary.Click, mnuMapAsText.Click, mnuMapAsDelimited.Click, mnuMapAsXML.Click, mnuMapAsRelational.Click, mnuMapAsVSAM.Click, mnuMapAsIMS.Click, mnuMapAsDB2LOAD.Click, mnuMapAsHSSUNLOAD.Click, mnuMapAsIBMEvent.Click, MenuItem36.Click, mnuMapAsIMSCDC.Click, mnuMapAsDB2CDC.Click, mnuMapAsVSAMCDC.Click, mnuMapAsXMLCDC.Click, mnuMapAsTriggerCDC.Click, mnuMapAsOraCDC.Click, mnuMapAsSQDCDC.Click, mnuMapAsIMSLE.Click, mnuMapAsIMSLEBat.Click

        Dim obj As INode
        Dim cNode As TreeNode
        Dim destObj As clsDatastore = Nothing
        Dim dsObj As clsDatastore
        Dim selNew As clsDSSelection
        Dim selOld As clsDSSelection
        Dim index As Integer
        Dim i As Integer
        Dim j As Integer
        Dim dsType As Integer
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim newObjParent As INode = Nothing '//this is parent of new cloned object
        Dim taskMain As New clsTask
        Dim pNode As TreeNode
        Dim mainFunc As New clsSQFunction
        Dim tempNode As TreeNode            '// added 5/07 by Tom Karasch to hold current tree node
        Dim cancelMapAs As Boolean = False  '// added 5/07 by Tom Karasch

        Me.Cursor = Cursors.WaitCursor

        Try
            cNode = tvExplorer.SelectedNode
            tempNode = tvExplorer.SelectedNode

            Select Case GetClipboardObjectType()
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    '//Make parent node to one step up if node is not folder
                    If CType(cNode.Tag, INode).IsFolderNode = False Then
                        cNode = cNode.Parent
                        newObjParent = cNode.Parent.Tag '//parent for new cloned obj 
                    Else
                        newObjParent = cNode.Parent.Tag '//parent for new cloned obj 
                    End If
                Case Else
                    Exit Try
            End Select

            Try
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                Select Case (CType(sender, MenuItem).Text)
                    Case "Binary"
                        dsType = modDeclares.enumDatastore.DS_BINARY
                    Case "Text"
                        dsType = modDeclares.enumDatastore.DS_TEXT
                    Case "Delimited"
                        dsType = modDeclares.enumDatastore.DS_DELIMITED
                    Case "XML"
                        dsType = modDeclares.enumDatastore.DS_XML
                    Case "Relational"
                        dsType = modDeclares.enumDatastore.DS_RELATIONAL
                    Case "VSAM"
                        dsType = modDeclares.enumDatastore.DS_VSAM
                    Case "IMS DB"
                        dsType = modDeclares.enumDatastore.DS_IMSDB
                    Case "DB2LOAD"
                        dsType = modDeclares.enumDatastore.DS_DB2LOAD
                    Case "HSSUNLOAD"
                        dsType = modDeclares.enumDatastore.DS_HSSUNLOAD
                    Case "IMS CDC"
                        dsType = modDeclares.enumDatastore.DS_IMSCDC
                    Case "IMS CDC LE"
                        dsType = modDeclares.enumDatastore.DS_IMSCDCLE
                    Case "DB2 CDC"
                        dsType = modDeclares.enumDatastore.DS_DB2CDC
                    Case "VSAM CDC"
                        dsType = modDeclares.enumDatastore.DS_VSAMCDC
                        'Case "XML CDC"
                        '    dsType = modDeclares.enumDatastore.DS_XMLCDC
                        'Case "Trigger Based CDC"
                        '    dsType = modDeclares.enumDatastore.DS_TRBCDC
                    Case "Oracle CDC"
                        dsType = modDeclares.enumDatastore.DS_ORACLECDC
                    Case "Generic CDC"
                        dsType = modDeclares.enumDatastore.DS_GENERICCDC
                    Case "IBM Event"
                        dsType = modDeclares.enumDatastore.DS_IBMEVENT
                        'Case "IMS LE Batch"
                        '    dsType = modDeclares.enumDatastore.DS_IMSLEBATCH
                End Select

                obj = GetClipboardObject(newObjParent, index)
                dsObj = CType(obj, clsDatastore)

                For i = 0 To dsObj.ObjSelections.Count - 1

                    '******************** put m in here

                    destObj = dsObj.Clone(dsObj.Parent)
                    destObj.ObjSelections.Clear()
                    destObj.ObjSelections.Add(dsObj.ObjSelections(i))

                    selOld = dsObj.ObjSelections(i)
                    selNew = destObj.ObjSelections(0)
                    '// added by TK 1/18/2007 to clear child selections and make
                    '// foreign key availability correct
                    selNew.ObjDSSelections.Clear()
                    selNew.ObjDatastore = destObj

                    destObj.Text = "O_" & selNew.Text
                    destObj.DsPhysicalSource = "PHYS." & destObj.Text
                    destObj.DatastoreType = dsType
                    destObj.OperationType = DS_OPERATION_INSERT
                    destObj.DsDirection = modDeclares.DS_DIRECTION_TARGET

                    '/// added 6/07 by Tom Karasch so duplicate names are not allowed
ReName1:            If destObj.Engine.FindDupNames(destObj) = True Then
                        destObj.Text = InputBox("Please change name of New Datastore" & Chr(13) & _
                        "The Default Datastore Name already exists", "Change New Datastore Name", destObj.Text)
                        If destObj.Text = "" Then
                            If MsgBox("Are you sure you want to cancel this 'Map As' ?", MsgBoxStyle.YesNo, _
                            "Cancel Map As") = MsgBoxResult.Yes Then
                                tran.Rollback()
                                cancelMapAs = True
                                Exit Try
                            Else
                                GoTo ReName1
                            End If
                        Else
                            GoTo ReName1
                        End If
                    End If

                    If destObj.AddNew(cmd) = True Then '//8/15/05
                        Dim task As New clsTask

                        FillDataStoreFromClipboard(cNode, destObj, cmd)


                        task.SeqNo = i  '// added 11/13/2006 by TK and KS
                        task.TaskName = "P_" & selNew.Text
                        task.TaskType = modDeclares.enumTaskType.TASK_PROC
                        task.Engine = destObj.Engine

                        '/// added 6/07 by Tom Karasch so no duplicete names for tasks occur
renameTask:             If task.Engine.FindDupNames(task) = True Then
                            task.Text = InputBox("Please change name of New Procedure" & Chr(13) & _
                            "The Default Procedure Name already exists", "Change New Procedure Name", task.Text)
                            If task.Text = "" Then
                                If MsgBox("Are you sure you want to cancel this 'Map As' ?", MsgBoxStyle.YesNo, _
                                "Cancel Map As") = MsgBoxResult.Yes Then
                                    tran.Rollback()
                                    cancelMapAs = True
                                    Exit Try
                                Else
                                    GoTo renameTask
                                End If
                            Else
                                GoTo renameTask
                            End If
                        End If

                        task.ObjSources.Add(dsObj)
                        task.ObjTargets.Add(destObj)

                        For j = 0 To selNew.DSSelectionFields.Count - 1
                            '// If..Then added by TK 5/25/07 to ensure proper groupitem mapping
                            If CType(selNew.DSSelectionFields(j), clsField).GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) <> "GROUPITEM" Then
                                task.AutoMap(modDeclares.enumTaskType.TASK_PROC, j, dsObj.Text, destObj.Text, selOld.DSSelectionFields(j), selNew.DSSelectionFields(j), selNew.Text)
                            Else
                                If selNew.ObjDatastore.Engine.MapGroupItems = True Then
                                    task.AutoMap(modDeclares.enumTaskType.TASK_PROC, j, dsObj.Text, destObj.Text, selOld.DSSelectionFields(j), selNew.DSSelectionFields(j), selNew.Text)
                                End If
                            End If
                        Next

                        task.AddNew(cmd)

                        taskMain.ObjTargets.Add(destObj)

                        pNode = cNode.Parent.Nodes(3)
                        pNode = AddNode(pNode, task.Type, task)
                        task.ObjTreeNode = pNode
                        task.SeqNo = pNode.Index '//store position
                        task.SaveSeqNo(cmd)
                        pNode.Expand()
                    End If
                Next

                taskMain.TaskName = dsObj.Text & "_MAIN"
                taskMain.TaskType = modDeclares.enumTaskType.TASK_MAIN
                taskMain.Engine = destObj.Engine

                '/// added 6/07 by Tom Karasch so that duplicate Task_Main names don't occur
renameMain:     If taskMain.Engine.FindDupNames(taskMain) = True Then
                    taskMain.Text = InputBox("Please change name of New Main Procedure" & Chr(13) & _
                    "The Default Main Procedure Name already exists", "Change New Main Procedure Name", taskMain.Text)
                    If taskMain.Text = "" Then
                        If MsgBox("Are you sure you want to cancel this 'Map As' ?", MsgBoxStyle.YesNo, _
                        "Cancel Map As") = MsgBoxResult.Yes Then
                            tran.Rollback()
                            cancelMapAs = True
                            Exit Try
                        Else
                            GoTo renameMain
                        End If
                    Else
                        GoTo renameMain
                    End If
                End If

                taskMain.ObjSources.Add(dsObj)

                taskMain.AutoMap(modDeclares.enumTaskType.TASK_MAIN, , , , mainFunc)

                taskMain.AddNew(cmd)

                pNode = cNode.Parent.Nodes(4)
                pNode = AddNode(pNode, taskMain.Type, taskMain)
                taskMain.ObjTreeNode = pNode
                taskMain.SeqNo = pNode.Index
                taskMain.SaveSeqNo(cmd)
                pNode.Expand()

                tran.Commit()
            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "SQL Transaction in mapAs in Main Form")
            End Try

            '/// added 6/07 by Tom Karasch in case mapAs is cancelled
            If cancelMapAs = True Then
                FillProject(CurLoadedProject, False, True)
                tvExplorer.SelectedNode = SelectFirstMatchingNode(tvExplorer, tempNode.Text)
            End If

        Catch ex As Exception
            LogError(ex, "mapAs in Main Form")
        Finally
            Me.Cursor = Cursors.Default
            ClearClipboard()
        End Try

    End Sub

    '//////// added by TK to start Query tool ////////
    Function StartQueryTool() As Boolean

        'Dim frm As New frmQuery
        Dim QueryDS As clsDatastore
        Dim QuerySelection As clsDSSelection

        QuerySelection = tvExplorer.SelectedNode.Tag
        QueryDS = QuerySelection.ObjDatastore
        ActiveQueryDSlist.Add(QueryDS)
        Call ShowUsercontrol(tvExplorer.SelectedNode)
        lblStatusMsg.Text = "*** Some DATASTORES are presently being queried, and cannot be modified now ***"

        If Not QueryDS Is Nothing Then
            'frm = frm.Query(QueryDS, QuerySelection)
        Else
            MsgBox("this is not a valid Datastore Selection", MsgBoxStyle.Information, MsgTitle)
        End If

    End Function

    '//////// This Opens the MapList Form ////////
    Function OpenMapList() As Boolean

        Try
            If CurLoadedProject IsNot Nothing Then
                Dim frm As New frmMapList
                OpenMapList = frm.NewOrOpen(CurLoadedProject)
            End If

        Catch ex As Exception
            LogError(ex, "frmMain OpenMapList")
            Return False
        End Try

    End Function

    '//////// added by tk to open www.sqdata.com ////////
    Private Sub imgLogo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
#If CONFIG = "ETI" Then
            Shell("C:\Program Files\Internet Explorer\iexplore.exe http://www.eti.com/products/cdc.html", AppWinStyle.NormalFocus)
#Else
            Shell("C:\Program Files\Internet Explorer\iexplore.exe http://www.sqdata.com/portal/login.asp", AppWinStyle.NormalFocus)
#End If

        Catch ex As Exception
            LogError(ex)
        End Try

    End Sub

    '//////// added by TK for toggling split screen ////////
    Private Sub SCmain_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SCmain.DoubleClick

        Try
            If SCmain.Panel1Collapsed = True Then
                SCmain.Panel1Collapsed = False
            Else
                SCmain.Panel1Collapsed = True
            End If

        Catch ex As Exception
            LogError(ex, "frmMain SplittcontainerDoubleClick")
        End Try

    End Sub

    Private Sub mnuDSAlpha_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim SelObj As INode
        SelObj = tvExplorer.SelectedNode.Tag

        Dim SelNode As TreeNode = tvExplorer.SelectedNode

        Select Case SelObj.Type
            Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE

            Case NODE_FO_SOURCEDATASTORE, NODE_FO_TARGETDATASTORE

        End Select

    End Sub

#End Region

#Region "Generate Reports"

    Private Sub mnuGenRepStr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuGenRepStr.Click
        Dim frm As New frmReport
        Dim InVal As clsStructure

        Try
            InVal = tvExplorer.SelectedNode.Tag
            If InVal IsNot Nothing Then
                frm = frm.Report(InVal)
            End If

        Catch ex As Exception
            LogError(ex, "frmMain mnuGenRepStr_Click")
        End Try

    End Sub

#End Region

End Class