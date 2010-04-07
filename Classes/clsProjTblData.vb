Public Class clsProjTblData

    Private m_ProjectName As String
    Private m_ProjDesc As String
    Private m_ProjVer As String
    Private m_ProjcDate As String
    Private m_ProjMetaVersion As enumMetaVersion
    Private m_MetaConnStr As String
    Private m_ProjTblName As String
    Private m_DSNname As String
    Private m_MetaVersion As enumMetaVersion

    Property ProjectName() As String
        Get
            Return m_ProjectName
        End Get
        Set(ByVal value As String)
            m_ProjectName = value
        End Set
    End Property

    Property ProjectDesc() As String
        Get
            Return m_ProjDesc
        End Get
        Set(ByVal value As String)
            m_ProjDesc = value
        End Set
    End Property

    Property ProjectVer() As String
        Get
            Return m_ProjVer
        End Get
        Set(ByVal value As String)
            m_ProjVer = value
        End Set
    End Property

    Property ProjMetaVersion() As enumMetaVersion
        Get
            Return m_ProjMetaVersion
        End Get
        Set(ByVal value As enumMetaVersion)
            m_ProjMetaVersion = value
        End Set
    End Property

    Property ProjectCDate() As String
        Get
            Return m_ProjcDate
        End Get
        Set(ByVal value As String)
            m_ProjcDate = value
        End Set
    End Property

    Property MetaConnString() As String
        Get
            Return m_MetaConnStr
        End Get
        Set(ByVal value As String)
            m_MetaConnStr = value
        End Set
    End Property

    Property ProjectTableName() As String
        Get
            Return m_ProjTblName
        End Get
        Set(ByVal value As String)
            m_ProjTblName = value
        End Set
    End Property

    Property ProjectDSN() As String
        Get
            Return m_DSNname
        End Get
        Set(ByVal value As String)
            m_DSNname = value
        End Set
    End Property

    Property MetaVersion() As enumMetaVersion
        Get
            Return m_MetaVersion
        End Get
        Set(ByVal value As enumMetaVersion)
            m_MetaVersion = value
        End Set
    End Property

End Class
