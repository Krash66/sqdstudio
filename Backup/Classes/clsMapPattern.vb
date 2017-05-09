Public Class clsMapPattern

    Dim m_source As String = ""
    Dim m_target As String = ""
    Dim m_Project As clsProject
    Dim m_sourceIndex As Integer
    Dim m_targetIndex As Integer

    Property Source() As String
        Get
            Return m_source
        End Get
        Set(ByVal value As String)
            m_source = value
        End Set
    End Property

    Property Target() As String
        Get
            Return m_target
        End Get
        Set(ByVal value As String)
            m_target = value
        End Set
    End Property

    Public Property SrcIdx() As Integer
        Get
            Return m_sourceIndex
        End Get
        Set(ByVal value As Integer)
            m_sourceIndex = value
        End Set
    End Property

    Public Property TgtIdx() As Integer
        Get
            Return m_targetIndex
        End Get
        Set(ByVal value As Integer)
            m_targetIndex = value
        End Set
    End Property

    ReadOnly Property HasSource() As Boolean
        Get
            If m_source.Trim <> "" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ReadOnly Property HasTarget() As Boolean
        Get
            If m_target.Trim <> "" Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    ReadOnly Property ObjString() As String
        Get
            Return Source & "," & Target
        End Get
    End Property

    Property Project() As clsProject
        Get
            Return m_Project
        End Get
        Set(ByVal value As clsProject)
            m_Project = value
        End Set
    End Property

End Class
