Public Module modDeclares

#Region "API Declare"
    
    Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As IntPtr, _
        ByVal wMsg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As Integer) As Integer

    Public Const EM_LINEINDEX As Integer = &HBB
    Public Const EM_POSFROMCHAR As Integer = &HD6
    Public Const EM_CHARFROMPOS As Integer = &HD7

    '///

    Public Declare Function CreateCaret Lib "user32" ( _
        ByVal hWnd As Integer, _
        ByVal hBitmap As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer) As Integer

    Public Declare Function ShowCaret Lib "user32" ( _
        ByVal hWnd As Integer) As Integer

    Public Declare Function SetCaretBlinkTime Lib "user32" ( _
        ByVal wMSeconds As Integer) As Integer
    Public Declare Function GetCaretBlinkTime Lib "user32" () As Integer


#End Region

#Region "Public Variables"

    Public EnableLogging As Boolean = True
    Public TraceFile As String
    Public errorTrace As String
    Public ODBCTrace As String
    ' These two for rename
    Public NameOfNodeBefore As String
    'Public PrevObjTreeNode As TreeNode

    Public IsEventFromCode As Boolean

    '//Points to the INode object of the topmost usercontrol
    '//Used to determine IsModified flag of currently opened form before we switch view
    '//This object is set when user select treenode for edit or new
    Public objTopMost As INode
    Public cnn As System.Data.Odbc.OdbcConnection
    '/// list of datastores that a currently being queried
    '/// this will be used to disable usercontrol for all datastores currently in query mode
    Public ActiveQueryDSlist As New ArrayList

    '//Some global controls used in frmMain, instead of dropping new controls on different we will use this global controls
    Public imgListSmall As ImageList
    Public imgListLarge As ImageList
    Public dlgSave As SaveFileDialog
    Public dlgOpen As OpenFileDialog
    Public dlgBrowseFolder As FolderBrowserDialog

    Public objLog As clsLogging
    Public mainwindow As frmMain

    Public SubstList As New Collection

    Public Procs As New Collection

#End Region

#Region "Private Variables"

    '//Project Save flag : if true then save project
    Private m_IsModified As Boolean

#End Region

#Region "Public Constants"

    '#If CONFIG = "ETI" Then
    '    Public Const MsgTitle As String = "ETI CDC Studio"
    '#Else
    Public Const MsgTitle As String = "Design Studio"
    '#End If

    '/// Used for inserting into database when sys, eng, or DS is null
    Public Const DBNULL As String = "<-null->"

    Public Const KEY_SAP As String = "." '// Added by npatel on 8/12/05

    Public Const SYSALL_NAME As String = "[SysAll]"

    Public Const MOUSE_LEAVE_COLOR As Integer = &HCED3D6
    Public Const MOUSE_ENTER_COLOR As Integer = &H840000

    '//Project specific constants used as key prefix of treenode
    Public Const NODE_PROJECT As String = "PRO"
    Public Const NODE_ENVIRONMENT As String = "ENV"
    Public Const NODE_STRUCT As String = "STR"
    Public Const NODE_STRUCT_FLD As String = "STF"
    Public Const NODE_STRUCT_SEL As String = "STS"
    Public Const NODE_SYSTEM As String = "SYS"
    Public Const NODE_ENGINE As String = "ENG"
    Public Const NODE_CONNECTION As String = "CNT"
    Public Const NODE_SOURCEDATASTORE As String = "SDS"
    Public Const NODE_SOURCEDSSEL As String = "SDR"
    Public Const NODE_TARGETDATASTORE As String = "TDS"
    Public Const NODE_TARGETDSSEL As String = "TDR"
    Public Const NODE_VARIABLE As String = "VAR"
    Public Const NODE_LOOKUP As String = "LOK"
    Public Const NODE_GEN As String = "GEN"
    Public Const NODE_PROC As String = "PRC"
    Public Const NODE_MAIN As String = "MAN"
    Public Const NODE_MAPPING As String = "MAP"
    Public Const NODE_FUN As String = "FUN" 'SQFunction
    Public Const NODE_CONST As String = "CON"
    Public Const NODE_TEMPLATE As String = "TMP"
    'Public Const NODE_INCLUDE As String = "INC"

    '//Variables for Folder node, used as key prefix of treenode
    Public Const NODE_FO_ENVIRONMENT As String = "FO_ENV"
    Public Const NODE_FO_STRUCT As String = "FO_STR"
    Public Const NODE_FO_STRUCT_SEL As String = "FO_STS"
    Public Const NODE_FO_SYSTEM As String = "FO_SYS"
    Public Const NODE_FO_ENGINE As String = "FO_ENG"
    Public Const NODE_FO_CONNECTION As String = "FO_CNT"
    Public Const NODE_FO_DATASTORE As String = "FO_DS"
    Public Const NODE_FO_SOURCEDATASTORE As String = "FO_SDS"
    Public Const NODE_FO_TARGETDATASTORE As String = "FO_TDS"
    Public Const DS_UNKNOWN As String = "FO_DS0"
    Public Const DS_BINARY As String = "FO_DS1"
    Public Const DS_TEXT As String = "FO_DS2"
    Public Const DS_DELIMITED As String = "FO_DS3"
    Public Const DS_XML As String = "FO_DS4"
    Public Const DS_RELATIONAL As String = "FO_DS5"
    Public Const DS_DB2LOAD As String = "FO_DS6"
    Public Const DS_HSSUNLOAD As String = "FO_DS7"
    Public Const DS_IMSDB As String = "FO_DS8"
    Public Const DS_VSAM As String = "FO_DS9"
    Public Const DS_IMSCDCLE As String = "FO_DS10"
    Public Const DS_DB2CDC As String = "FO_DS11"
    Public Const DS_VSAMCDC As String = "FO_DS12"
    'Public Const DS_XMLCDC As String = "FO_DS13"
    Public Const DS_IBMEVENT As String = "FO_DS14"
    'Trigger based cdc
    'Public Const DS_TRBCDC As String = "FO_DS15"
    'V3 additions
    Public Const DS_ORACLECDC As String = "FO_DS16"
    Public Const DS_GENERICCDC As String = "FO_DS17"
    Public Const DS_IMSCDC As String = "FO_DS18"
    'Public Const DS_IMSLEBATCH As String = "FO_DS19"
    Public Const DS_SUBVAR As String = "FO_DS20"
    Public Const DS_INCLUDE As String = "FO_DS21"
    Public Const NODE_FO_VARIABLE As String = "FO_VAR"
    Public Const NODE_FO_PROC As String = "FO_PRC"
    Public Const NODE_FO_JOIN As String = "FO_JON"
    Public Const NODE_FO_MAIN As String = "FO_MAN"
    Public Const NODE_FO_LOOKUP As String = "FO_LOK"
    Public Const NODE_FO_FUNCTION As String = "FO_FUN"
    Public Const NODE_FO_FUNCTION_RECENT As String = "FO_FUNREC"
    Public Const NODE_FO_TEMPLATE As String = "FO_TMP"
    Public Const NODE_FO_MAPPING As String = "FO_MAP"

    'Public Const NODE_FOLDERCLOSE As String = "CLOSE"
    Public Const NODE_FOLDEROPEN As String = "OPEN"

    '//Datastore related constants
    Public Const DS_DIRECTION_SOURCE As String = "S"
    Public Const DS_DIRECTION_TARGET As String = "T"

    Public Const DS_ACCESSMETHOD_FILE As String = "F"
    Public Const DS_ACCESSMETHOD_IP As String = "I"
    Public Const DS_ACCESSMETHOD_MQSERIES As String = "M"
    Public Const DS_ACCESSMETHOD_VSAM As String = "V"
    Public Const DS_ACCESSMETHOD_SQDCDC As String = "S"

    Public Const DS_CHARACTERCODE_EBCDIC As String = "E"
    Public Const DS_CHARACTERCODE_ASCII As String = "A"
    Public Const DS_CHARACTERCODE_SYSTEM As String = "S"

    Public Const DS_DELIMITER_CRLF As String = vbCrLf
    Public Const DS_DELIMITER_CR As String = vbCr
    Public Const DS_DELIMITER_LF As String = vbLf
    Public Const DS_DELIMITER_TAB As String = vbTab
    Public Const DS_DELIMITER_VERTICALBAR As String = "|"
    Public Const DS_DELIMITER_COMMA As String = ","
    Public Const DS_DELIMITER_SEMICOLON As String = ";"
    Public Const DS_DELIMITER_SQUOTE As String = "'" 'single quote
    Public Const DS_DELIMITER_DQUOTE As String = """" 'double quote
    Public Const DS_DELIMITER_TILDE As String = "~"
    Public Const DS_DELIMITER_NONE As String = ""

    Public Const DS_OPERATION_INSERT As String = "I"
    Public Const DS_OPERATION_REPLACE As String = "R"
    Public Const DS_OPERATION_UPDATE As String = "U"
    Public Const DS_OPERATION_MODIFY As String = "M"
    Public Const DS_OPERATION_CHANGE As String = "C"
    Public Const DS_OPERATION_DELETE As String = "D"

    '//

    Public Const MAPPING_TYPE_FIELD As String = "FLD"
    Public Const MAPPING_TYPE_WORKVAR As String = "WRK"
    Public Const MAPPING_TYPE_VAR As String = NODE_VARIABLE
    Public Const MAPPING_TYPE_PROC As String = NODE_PROC
    Public Const MAPPING_TYPE_JOIN As String = NODE_GEN
    Public Const MAPPING_TYPE_LOOKUP As String = NODE_LOOKUP
    Public Const MAPPING_TYPE_FUN As String = NODE_FUN

    '
    'Field attribute names
    '
    Public Const TXT_ISKEY As String = "Key field?"
    '// added 11/3/06
    Public Const TXT_FKEY As String = "Foreign Key (COBOL w/DBD)"
    Public Const TXT_INITVAL As String = "Initialization value"
    Public Const TXT_RETYPE As String = "Re-Type Data As"
    Public Const TXT_INVALID As String = "Action for invalid value"
    Public Const TXT_EXTTYPE As String = "External data type"
    Public Const TXT_DATEFORMAT As String = "Date Format"
    Public Const TXT_LABEL As String = "Label"
    Public Const TXT_IDENTVAL As String = "Record identifier value"
    Public Const TXT_NCHILDREN As String = "Number of Elements"
    Public Const TXT_LEVEL As String = "Level"
    Public Const TXT_TIMES As String = "Number of Occurs"
    Public Const TXT_OCCURS As String = "Occurance number"
    Public Const TXT_DATATYPE As String = "Data type"
    Public Const TXT_OFFSET As String = "Field offset"
    Public Const TXT_LENGTH As String = "Internal field length"
    Public Const TXT_SCALE As String = "Scale"
    Public Const TXT_CANNULL As String = "Null allowed?"

    Public Const TAB As String = "    "
#End Region

#Region "Public enumerations"

    '//This will tell loading function how to parse the attributes and which object to store in treeview node
    '//e.g if XMLFOR_FIELD then attributes will be parsed and stored in clsField object
    '//e.g if XMLFOR_SQFUNCTION then attributes will be parsed and stored in clsFunction object
    Public Enum enumXMLType
        XMLFOR_FIELD = 0
        XMLFOR_SEGMENTS = 1
        XMLFOR_SQFUNCTIONS = 2
    End Enum

    Public Enum enumVariableType
        VTYPE_LOCAL = 0
        VTYPE_GLOBAL = 1
        VTYPE_CONST = 2
    End Enum

    Public Enum enumMappingType
        MAPPING_TYPE_NONE = 0
        MAPPING_TYPE_FIELD = 1
        MAPPING_TYPE_VAR = 2
        MAPPING_TYPE_PROC = 3
        MAPPING_TYPE_JOIN = 4
        MAPPING_TYPE_LOOKUP = 5
        MAPPING_TYPE_FUN = 6
        MAPPING_TYPE_WORKVAR = 7 '//only for target
        MAPPING_TYPE_CONSTANT = 8 '//only for source
    End Enum

    Public Enum enumDirection
        DI_SOURCE = 0
        DI_TARGET = 1
        DI_SOURCE_TARGET = 2
    End Enum

    Public Enum enumAction
        ACTION_DELETE = 2
        ACTION_NEW = 3
        ACTION_OPEN = 4
        ACTION_CHANGE = 5
        ACTION_MERGE = 6
    End Enum

    '// changed 11/3/06 by TK and KS
    Public Enum enumFieldAttributes
        'Editable in datastore
        ATTR_ISKEY = 0      'Key?
        ATTR_FKEY = 1       'Foreign Key?
        ATTR_INITVAL = 2    'Initial Value
        ATTR_RETYPE = 3     'Change Datatype
        ATTR_EXTTYPE = 4    'External Datatype
        ATTR_INVALID = 5    'Action for invalid
        'ATTR_IFNULL = 6     'Action for Null value
        ATTR_DATEFORMAT = 6 'Date Format
        ATTR_LABEL = 7     'Label
        ATTR_IDENTVAL = 8   'Record ID

        'Read Only constants from structure
        ATTR_NCHILDREN = 9 'Number of Children
        ATTR_LEVEL = 10     'Level
        ATTR_TIMES = 11     'Number of cells
        ATTR_OCCURS = 12    'Cell number
        ATTR_DATATYPE = 13  'Datatype
        ATTR_OFFSET = 14  'Field Offset
        ATTR_LENGTH = 15    'Initial field length
        ATTR_SCALE = 16    'Scale
        ATTR_CANNULL = 17   'Null Allowed?

    End Enum

    Public Enum enumStructure
        STRUCT_UNKNOWN = 0
        STRUCT_COBOL = 1
        STRUCT_IMS = 2
        STRUCT_COBOL_IMS = 3
        STRUCT_XMLDTD = 4
        STRUCT_C = 5
        STRUCT_REL_DDL = 6
        STRUCT_REL_DML = 7     '/// Tablename in fPath1
        STRUCT_REL_DML_FILE = 8  '/// path in fPath1
    End Enum

    Public Enum enumCalledFrom
        BY_ENVIRONMENT = 1
        BY_STRUCTURE = 2
        BY_SCRIPTGEN = 3
    End Enum

    Public Enum enumDatastore
        DS_UNKNOWN = 0
        DS_BINARY = 1
        DS_TEXT = 2
        DS_DELIMITED = 3
        DS_XML = 4
        DS_RELATIONAL = 5
        DS_DB2LOAD = 6
        DS_HSSUNLOAD = 7
        DS_IMSDB = 8
        DS_VSAM = 9
        DS_IMSCDCLE = 18
        DS_IMSCDC = 10
        DS_DB2CDC = 11
        DS_VSAMCDC = 12
        'DS_XMLCDC = 13
        DS_IBMEVENT = 14
        'Trigger based cdc
        'DS_TRBCDC = 15
        'V3 additions
        DS_ORACLECDC = 16
        DS_UTSCDC = 17

        'DS_IMSLEBATCH = 19
        DS_SUBVAR = 20
        DS_INCLUDE = 21

    End Enum

    Public Enum enumTaskType
        TASK_MAIN = 0
        TASK_PROC = 1
        TASK_GEN = 2
        TASK_LOOKUP = 3
        TASK_IncProc = 4
    End Enum

    Public Enum enumControlTypes
        ControlUnknown = 0
        ControlTextBox = 1
        ControlListBox = 2
        ControlComboBox = 3
        ControlListView = 4
        ControlTreeView = 5
    End Enum

    Public Enum enumODBCtype
        ACCESS = 1
        SQL_SERVER = 2
        ORACLE = 3
        DB2 = 4
        OTHER = 20
    End Enum

    Public Enum enumParserReturnCode
        OK = 0
        PathWarning = 1
        Warning = 4
        Failed = 8
        NoCode = 20
    End Enum

    Public Enum enumErrorLocation
        ModGenHead = 0
        ModGenStruct = 5
        ModGenDS = 10
        ModGenJoin = 12
        ModGenLU = 13
        ModGenProc = 15
        ModGenVar = 17
        ModGenMap = 20
        ModGenMain = 25
        ModGenWindows = 30
        ModGenSQDParse = 40
        ModGenFileCreation = 50

        NoErrors = 100
    End Enum

    Public Enum enumMappingLevel
        ShowAll = 1
        ShowDesc = 2
        ShowFld = 3
    End Enum

    Public Enum enumOutMsg
        NoOutMsg = 0
        PrintOutMsg = 1
    End Enum

    Public Enum enumIncludeType
        SrcDS = 0
        TgtDS = 1
        Proc = 3
    End Enum

    Public Enum enumGenType
        DS = 0
        Proc = 1
        Eng = 2
    End Enum

    Public Enum enumMetaVersion
        V2 = 0
        V3 = 1
    End Enum

#End Region

#Region "Public Properties"

    Public Property IsModified() As Boolean
        Get
            Return m_IsModified
        End Get
        Set(ByVal Value As Boolean)
            m_IsModified = Value
        End Set
    End Property

#End Region

#Region "Public Functions"

    Public Function Log(ByVal sMsg As String, Optional ByVal AddNewLine As Boolean = True) As Boolean
        clsLogging.LogEvent(sMsg, AddNewLine)
    End Function

    Public Function ErrorLog(ByVal sMsg As String, Optional ByVal AddNewLine As Boolean = True) As Boolean
        clsLogging.ErrorEvent(sMsg, AddNewLine)
    End Function

    Public Function ODBCErrorLog(ByVal sMsg As String, Optional ByVal AddNewLine As Boolean = True) As Boolean
        clsLogging.ODBCEvent(sMsg, AddNewLine)
    End Function

    Public Function LogError(ByVal ex As Exception, Optional ByVal p1 As String = "", Optional ByVal p2 As String = "", Optional ByVal ThrowError As Boolean = False, Optional ByVal displayMSG As Boolean = False) As Boolean
        clsLogging.LogError(ex, p1, p2, ThrowError, displayMSG)
    End Function

    Public Function LogODBCError(ByVal ex As Exception, Optional ByVal p1 As String = "", Optional ByVal p2 As String = "", Optional ByVal ThrowError As Boolean = False, Optional ByVal displayMSG As Boolean = False) As Boolean
        clsLogging.LogODBCerror(ex, p1, p2, ThrowError, displayMSG)
    End Function

    Public Function LoadGlobalValues(Optional ByVal ClearLogOnStartUp As Boolean = True) As Boolean

        objLog = New clsLogging
        modDeclares.objLog = objLog

        Try
            If TraceFile Is Nothing Then
                TraceFile = "Trace.log"
            End If
            If errorTrace Is Nothing Then
                errorTrace = "ErrorTrc.log"
            End If
            If ODBCTrace Is Nothing Then
                ODBCTrace = "ODBCErrLog.log"
            End If
            If ClearLogOnStartUp = True Then
                System.IO.File.Delete(GetAppLog() & "\" & TraceFile)  '
                System.IO.File.Delete(GetAppLog() & "\" & errorTrace)   '& "\"
                System.IO.File.Delete(GetAppLog() & "\" & ODBCTrace)
                'System.IO.File.Delete(GetAppPath() & "*.log")
            End If
            EnableLogging = True

            Log("Trace Enabled")
            Return True

        Catch ex As Exception
            LogError(ex, "modDeclares LoadGlobalValues")
            Return False
        End Try

    End Function

#End Region

End Module