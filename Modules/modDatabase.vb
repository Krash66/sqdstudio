Public Module modDatabase

    Function GetNewId() As String
        Return System.Guid.NewGuid().ToString("N")
    End Function

    Function GetVal(ByVal obj As Object) As Object
        If (obj Is System.DBNull.Value) Then
            GetVal = Nothing
        Else
            GetVal = obj
        End If
    End Function

    Function GetStr(ByVal obj As Object) As String
        If obj Is Nothing Then
            GetStr = ""
        Else
            GetStr = obj.ToString
        End If
    End Function

    '/// UnUsed as of March 2007
#Region "DSN Methods - Added to keep DSN up to date after auto updates occur"

    'Constant Declaration
    Private Const ODBC_ADD_DSN As Integer = 1               ' Add data source
    Private Const ODBC_CONFIG_DSN As Integer = 2            ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN As Integer = 3            ' Remove data source
    Private Const vbAPINull As Long = 0                     ' NULL Pointer

    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
         (ByVal hwndParent As Integer, ByVal fRequest As Integer, _
         ByVal lpszDriver As String, ByVal lpszAttributes As String) As Integer

    ' JDM - 3/28/2006 - Update DSN (after program update)
    Function UpdateDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseFilePath As String = "", _
            Optional ByVal Description As String = "" _
            ) As Boolean

        If Description.Trim = "" Then
            Description = "This is default datasource for SQData Studio 1.0"
        End If

        If Not System.IO.File.Exists(DatabaseFilePath) Then ' Copy from parent
            If System.IO.File.Exists(GetParentPath() & "PROJECTS\SQDMeta.mdb") Then
                System.IO.File.Move(GetParentPath() & "PROJECTS\SQDMeta.mdb", DatabaseFilePath)
            Else
                System.IO.File.Move(GetParentPath() & "Temp\SQDMeta.mdb", DatabaseFilePath)
            End If
        End If

        RemoveDSN(DsnName)
        Return CreateDSN(DsnName, DatabaseFilePath, Description)

    End Function

    Function CreateDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseFilePath As String = "", _
            Optional ByVal Description As String = "" _
            ) As Boolean

        Dim intRet As Integer
        Dim strDriver As String
        Dim strAttributes As String

        If DatabaseFilePath.Trim = "" Then
            DatabaseFilePath = System.AppDomain.CurrentDomain.BaseDirectory() & "SQDMeta.mdb"
        End If
        If Description.Trim = "" Then
            Description = "This is default datasource for SQData Studio 1.0"
        End If

        strDriver = "Microsoft Access Driver (*.mdb)" & Chr(0)
        strAttributes = "DSN=" & DsnName & Chr(0) & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0) & "DBQ=" & DatabaseFilePath & Chr(0) & "Description=" & Description & Chr(0)
        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)

        If intRet Then
            Return True
        Else
            Return False
        End If

    End Function

    Function RemoveDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseName As String = "" _
            ) As Boolean

        Dim intRet As Integer
        Dim strDriver As String
        Dim strAttributes As String

        strDriver = "Microsoft Access Driver (*.mdb)"
        strAttributes = "DSN=" & DsnName & Chr(0)
        intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)

        If intRet Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

End Module