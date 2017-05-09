Public Class clsImpRetCode

    '/// Object class for a return code object to be used with Importer
    '/// Created July 2012 by TK
    Private mName As String                     'Return Code gets name of object that creates (e.g. Engine if script, TableName if Desc.)
    Private mHasError As Boolean                'Error Has Occured
    Private mErrorCount As Integer              'Number of Errors
    Private mHasFatal As Boolean                'Fatal Error Has Occured
    Private mPath As String                     'file path of file that caused problems (e.g. struct file, Script file)
    Private mErrorPath As String                'file path of the file where the error occured (Output File)
    Private mReturnCode As enumImpRetCode       'return code (if errors or if warnings or no errors)
    Private mLocalErrorMsg As String            'Error Message at that line of the script
    '/// Object Properties
    Private mImportType As enumImportType       'Type of import file is being read in
    Private mObjInode As INode                  'Object where the error occured
    Private mObjName As String                  'The Object's Name
    Private mObjType As String                  'The Inode Type in readable format (to use in error messages and log)
    Private mErrline As Integer                 'Line number in INL file where error occured

    '/// Collections and properties for Description Import
    Public d_DescName As String = ""              'Alias of the Description
    Public d_DescPath As String = ""              'Path of the Description File (may be multi-table file)
    Public d_DescType As enumStructure = enumStructure.STRUCT_UNKNOWN
    Public d_Tables As Collection = Nothing       'Collection of clsStructure Inodes
    Public d_CurrentStructure As clsStructure = Nothing    'Current Structure in use out of the collection

    '/// Collections and properties for Scripts
    Public Eng As clsEngine

    Enum enumImpRetCode
        None = -1
        Good = 0
        Warning = 4
        Fatal = 8
    End Enum

    Enum enumImportType
        None = 0
        ScriptRPT = 1
        ScriptSQD = 2
        ScriptINL = 3
        DescSQL = 5
        DescDDL = 6
        DescC = 7
        DescXML = 8
        DescDML = 9
        DescDMLFile = 10
    End Enum

#Region "Properties"

    Public Property Name() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value
        End Set
    End Property

    Public Property HasError() As Boolean
        Get
            Return mHasError
        End Get
        Set(ByVal value As Boolean)
            mHasError = value
        End Set
    End Property

    Public Property ErrorCount() As Integer
        Get
            Return mErrorCount
        End Get
        Set(ByVal value As Integer)
            mErrorCount = value
        End Set
    End Property

    Public Property HasFatal() As Boolean
        Get
            Return mHasFatal
        End Get
        Set(ByVal value As Boolean)
            mHasFatal = value
        End Set
    End Property

    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            mPath = value
        End Set
    End Property

    Public Property ErrorPath() As String
        Get
            Return mErrorPath
        End Get
        Set(ByVal value As String)
            mErrorPath = value
        End Set
    End Property

    Public Property ReturnCode() As enumImpRetCode
        Get
            Return mReturnCode
        End Get
        Set(ByVal value As enumImpRetCode)
            mReturnCode = value
        End Set
    End Property

    Public Property LocalErrorMsg() As String
        Get
            Return mLocalErrorMsg
        End Get
        Set(ByVal value As String)
            mLocalErrorMsg = value
        End Set
    End Property

    Public Property ImportType() As enumImportType
        Get
            Return mImportType
        End Get
        Set(ByVal value As enumImportType)
            mImportType = value
        End Set
    End Property

    Public Property ObjInode() As INode
        Get
            Return mObjInode
        End Get
        Set(ByVal value As INode)
            mObjInode = value
        End Set
    End Property

    Public ReadOnly Property ObjName() As String
        Get
            Return mObjInode.Text
        End Get
    End Property

    Public ReadOnly Property Objtype() As String
        Get
            Select Case mObjInode.Type
                Case NODE_PROJECT
                    Return "Project"
                Case NODE_ENVIRONMENT
                    Return "Environment"
                Case NODE_STRUCT
                    Return "Structure"
                Case NODE_STRUCT_FLD
                    Return "Field"
                Case NODE_STRUCT_SEL
                    Return "Structure Selection"
                Case NODE_SYSTEM
                    Return "System"
                Case NODE_ENGINE
                    Return "Engine"
                Case NODE_CONNECTION
                    Return "Connection"
                Case NODE_SOURCEDATASTORE
                    Return "Source Datastore"
                Case NODE_SOURCEDSSEL
                    Return "Source Datastore Selection"
                Case NODE_TARGETDATASTORE
                    Return "Target Datastore"
                Case NODE_TARGETDSSEL
                    Return "Target Datastore Selection"
                Case NODE_VARIABLE
                    Return "Variable"
                Case NODE_LOOKUP
                    Return "LookUp"
                Case NODE_GEN
                    Return "Join"
                Case NODE_PROC
                    Return "Procedure"
                Case NODE_MAIN
                    Return "Main Procedure"
                Case NODE_MAPPING
                    Return "Mapping"
                Case Else
                    Return ""
            End Select
        End Get
    End Property

    Public Property Errline() As Integer
        Get
            Return mErrline
        End Get
        Set(ByVal value As Integer)
            mErrline = value
        End Set
    End Property

#End Region

#Region "Methods"

    Function GetGUIErrorMsg() As String

        Dim MessageObj As String = ""
        Dim ParseFlag As Boolean = False

        Select Case Me.Objtype

            Case enumErrorLocation.NoErrors
                '/// No Errors --- Already Handled

            Case enumErrorLocation.ModGenWindows
                MessageObj = "A Windows error occured with this message:" & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenSQDParse
                '/// Already Handled

        End Select

        Log("Script Generation Finished with the following code:")
        Log(MessageObj)
        Log("Error occured at:")
        Log("Line No: " & Me.Errline)

        Return MessageObj

    End Function

#End Region

End Class
