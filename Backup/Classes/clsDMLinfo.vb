Public Class clsDMLinfo

    '//// Created 6/2007 by Tom Karasch for Relational DML Structure Modeling
    Private m_DSNname As String             '/// ODBC DSN name of source for DML
    Private m_DSNtype As enumODBCtype       '/// Holds the type of ODBC source (i.e. DB2, Oracle)
    Private m_DSNDesc As String             '/// description of ODBC driver
    Private m_Schema As String              '/// User Chosen Schema for ODBC source
    Private m_TableName As String           '/// Name of chosen System catalog table
    Private m_PartialTableName As String    '/// Partial tablename and wildcard for searching sys catalog
    Private m_PartialSchema As String       '/// Partial SchemaName for wildcard search of SYS catalog
    Private m_TryAgain As Boolean           '/// Does the user want to start again?
    Private m_DoLogin As Boolean            '/// Does the ODBC source require a login
    Private m_Connection As clsConnection   '/// The connection Object for this DML Structure

    Public ColumnArray As New ArrayList     '/// Array of the columns chosen for the DML


    Property DSNname() As String
        Get
            Return m_DSNname
        End Get
        Set(ByVal value As String)
            m_DSNname = value
        End Set
    End Property

    Property DSNtype() As enumODBCtype
        Get
            Return m_DSNtype
        End Get
        Set(ByVal value As enumODBCtype)
            m_DSNtype = value
        End Set
    End Property

    Property DSNDesc() As String
        Get
            Return m_DSNDesc
        End Get
        Set(ByVal value As String)
            m_DSNDesc = value
        End Set
    End Property

    Property Schema() As String
        Get
            Return m_Schema
        End Get
        Set(ByVal value As String)
            m_Schema = value
        End Set
    End Property

    Property PartSchemaName() As String
        Get
            Return m_PartialSchema
        End Get
        Set(ByVal value As String)
            m_PartialSchema = value
        End Set
    End Property

    Property TableName() As String
        Get
            Return m_TableName
        End Get
        Set(ByVal value As String)
            m_TableName = value
        End Set
    End Property

    Property PartTabName() As String
        Get
            Return m_PartialTableName
        End Get
        Set(ByVal value As String)
            m_PartialTableName = value
        End Set
    End Property

    Property TryAgain() As Boolean
        Get
            Return m_TryAgain
        End Get
        Set(ByVal value As Boolean)
            m_TryAgain = value
        End Set
    End Property

    Property DoLogin() As Boolean
        Get
            Return m_DoLogin
        End Get
        Set(ByVal value As Boolean)
            m_DoLogin = value
        End Set
    End Property

    Public ReadOnly Property MetaConnectionString() As String
        Get
            Return "DSN=" & Connection.Database & "; UID=" & Connection.UserId & "; PWD=" & Connection.Password
        End Get
    End Property

    Public Property Connection() As clsConnection
        Get
            Return m_Connection
        End Get
        Set(ByVal value As clsConnection)
            m_Connection = value
        End Set
    End Property

End Class
