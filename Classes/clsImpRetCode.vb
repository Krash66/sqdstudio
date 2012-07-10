Public Class clsImpRetCode

    '///Object class for a return code object

    Private mName As String                     'Return Code gets name of object that creates (e.g. Engine when gen. script)
    Private mHasError As Boolean                'Error Has Occured
    Private mErrorCount As Integer
    Private mHasFatal As Boolean                'Fatal Error Has Occured
    Private mPath As String                     'file path of file that caused problems (e.g. struct file)
    Private mErrorPath As String                'file path of the file where the error occured (Output File)
    Private mReturnCode As String               'return code (if errors or if warnings or no errors)
    Private mErrorLocation As enumErrorLocation 'part of the generator module that caused the error
    Private mParserCode As enumParserReturnCode 'the Parser's return code
    Private mParserRC As Integer
    Private mLocalErrorMsg As String
    Private mParserPath As String               'Path the Parser sent Back
    Private mObjInode As INode                  'Object where the error occured
    Private mObjName As String                  'The Object's Name
    Private mObjType As String                  'The Inode Type in readable format (to use in error messages and log)
    Private mINLline As Integer                 'Line number in INL file where error occured
    Private mSQDline As Integer                 'Line number in SQD file where error occured
    Private mTMPline As Integer                 'Line number in TMP file where error occured
    Private mSQDPath As String
    Private mTMPPath As String
    Private mINLPath As String
    Private mRPTPath As String
    Private mPRCPath As String
    Private mGenType As enumGenType

    Private mModPath As String                  'Path to the newly built Model
    Private mModErrPath As String               'Path to ModelError Log
    Private mModReturnCode As String                'Modeler Return Code

    'Dim Rkeys As ArrayList
    'Dim RDS As ArrayList
    'Dim RProcs As ArrayList

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

    Public Property ReturnCode() As String
        Get
            Return mReturnCode
        End Get
        Set(ByVal value As String)
            mReturnCode = value
        End Set
    End Property

    Public Property ErrorLocation() As enumErrorLocation
        Get
            Return mErrorLocation
        End Get
        Set(ByVal value As enumErrorLocation)
            mErrorLocation = value
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

    Public Property ParseCode() As enumParserReturnCode
        Get
            Return mParserCode
        End Get
        Set(ByVal value As enumParserReturnCode)
            mParserCode = value
        End Set
    End Property

    Public Property ParserRC() As Integer
        Get
            Return mParserRC
        End Get
        Set(ByVal value As Integer)
            mParserRC = value
        End Set
    End Property

    Public Property ParserPath() As String
        Get
            Return mParserPath
        End Get
        Set(ByVal value As String)
            mParserPath = value
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

    Public Property INLline() As Integer
        Get
            Return mINLline
        End Get
        Set(ByVal value As Integer)
            mINLline = value
        End Set
    End Property

    Public Property SQDline() As Integer
        Get
            Return mSQDline
        End Get
        Set(ByVal value As Integer)
            mSQDline = value
        End Set
    End Property

    Public Property TMPline() As Integer
        Get
            Return mTMPline
        End Get
        Set(ByVal value As Integer)
            mTMPline = value
        End Set
    End Property

    Public Property SQDPath() As String
        Get
            Return mSQDPath
        End Get
        Set(ByVal value As String)
            mSQDPath = value
        End Set
    End Property

    Public Property TMPPath() As String
        Get
            Return mTMPPath
        End Get
        Set(ByVal value As String)
            mTMPPath = value
        End Set
    End Property

    Public Property INLPath() As String
        Get
            Return mINLPath
        End Get
        Set(ByVal value As String)
            mINLPath = value
        End Set
    End Property

    Public Property RPTPath() As String
        Get
            Return mRPTPath
        End Get
        Set(ByVal value As String)
            mRPTPath = value
        End Set
    End Property

    Public Property PRCPath() As String
        Get
            Return mPRCPath
        End Get
        Set(ByVal value As String)
            mPRCPath = value
        End Set
    End Property

    Public Property GenType() As enumGenType
        Get
            Return mGenType
        End Get
        Set(ByVal value As enumGenType)
            mGenType = value
        End Set
    End Property

    Public Property ModPath() As String
        Get
            Return mModPath
        End Get
        Set(ByVal value As String)
            mModPath = value
        End Set
    End Property

    Public Property ModErrPath() As String
        Get
            Return mModErrPath
        End Get
        Set(ByVal value As String)
            mModErrPath = value
        End Set
    End Property

    Public Property ModReturnCode() As String
        Get
            Return mModReturnCode
        End Get
        Set(ByVal value As String)
            mModReturnCode = value
        End Set
    End Property

#End Region

#Region "Methods"

    Function GetGUIErrorMsg() As String

        Dim MessageObj As String = ""
        Dim ParseFlag As Boolean = False

        Select Case Me.ErrorLocation
            Case enumErrorLocation.NoErrors
                '/// No Errors --- Already Handled
            Case enumErrorLocation.ModGenFileCreation
                MessageObj = "While creating file " & Me.ErrorPath & Chr(13) & _
                "an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenHead
                MessageObj = "While generating the Script File Headers," & Chr(13) & _
                "an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenDS
                MessageObj = "While generating the Datastores," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenStruct
                MessageObj = "While generating the Descriptions," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenVar
                MessageObj = "While generating the Variables," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenJoin
                MessageObj = "While generating the Joins," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenLU
                MessageObj = "While generating the Lookups," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenProc
                MessageObj = "While generating the Procedures," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenMain
                MessageObj = "While generating the Routing Procedures," & Chr(13) & _
                "object: " & Me.ObjName & Chr(13) & _
                "of Type: " & Me.Objtype & Chr(13) & _
                "caused an error occured in Windows as follows: " & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenWindows
                MessageObj = "A Windows error occured with this message:" & Chr(13) & _
                Me.ReturnCode

            Case enumErrorLocation.ModGenSQDParse
                '/// Already Handled

        End Select

        Log("Script Generation Finished with the following code:")
        Log(MessageObj)
        Log("Error occured at:")
        Log("SQD Line No: " & Me.SQDline)
        Log("INL Line No: " & Me.INLline)
        Log("TMP Line No: " & Me.TMPline)

        Return MessageObj

    End Function

#End Region

End Class
