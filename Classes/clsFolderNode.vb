<Serializable()> _
Public Class clsFolderNode
    Implements INode

    Private m_FolderNodeId As String
    Private m_FolderNodeType As String
    Private m_FolderNodeName As String = ""
    Private m_FolderNodeDescription As String = ""
    Private m_IsModified As Boolean
    Private m_ObjParent As INode
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False

#Region "INode Implementation"

    Public Function GetQuotedText(Optional ByVal bFix As Boolean = True) As String Implements INode.GetQuotedText
        If bFix = True Then
            Return "'" & FixStr(Me.Text.Trim) & "'"
        Else
            Return "'" & Me.Text.Trim & "'"
        End If
    End Function

    Public Property SeqNo() As Integer Implements INode.SeqNo
        Get
            Return m_SeqNo
        End Get
        Set(ByVal Value As Integer)
            m_SeqNo = Value
        End Set
    End Property

    Public ReadOnly Property GUID() As String Implements INode.GUID
        Get
            Return m_GUID
        End Get
    End Property

    Public Function IsFolderNode() As Boolean Implements INode.IsFolderNode
        If Strings.Left(Me.Type, 3) = "FO_" Then
            IsFolderNode = True
        Else
            IsFolderNode = False
        End If
    End Function

    Public Property ObjTreeNode() As System.Windows.Forms.TreeNode Implements INode.ObjTreeNode
        Get
            Return m_ObjTreeNode
        End Get
        Set(ByVal Value As System.Windows.Forms.TreeNode)
            m_ObjTreeNode = Value
        End Set
    End Property

    Public Property Parent() As INode Implements INode.Parent
        Get
            Return Me.m_ObjParent
        End Get
        Set(ByVal Value As INode)
            m_ObjParent = Value
        End Set
    End Property

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Return Nothing

    End Function

    Public Property IsModified() As Boolean Implements INode.IsModified
        Get
            Return m_IsModified
        End Get
        Set(ByVal Value As Boolean)
            m_IsModified = Value
        End Set
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.Parent.Project
            Return p
        End Get
    End Property

    Public ReadOnly Property Key() As String Implements INode.Key
        Get
            Key = FolderNodeId
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = FolderNodeName
        End Get
        Set(ByVal Value As String)
            FolderNodeName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return m_FolderNodeType
        End Get
    End Property

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        Return True

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete
        '//Always return true so wont stop subtree delete because of folder node Detele=False
        Delete = True

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    '// Function added 11/3/06 by KS and TK
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Public Property IsRenamed() As Boolean Implements INode.IsRenamed
        Get
            Return m_IsRenamed
        End Get
        Set(ByVal Value As Boolean)
            m_IsRenamed = Value
        End Set
    End Property

    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        Return True

    End Function

#End Region

#Region "Properties"

    Public Property FolderNodeId() As String
        Get
            Return m_FolderNodeId
        End Get
        Set(ByVal Value As String)
            m_FolderNodeId = Value
        End Set
    End Property

    Public Property FolderNodeName() As String
        Get
            Return m_FolderNodeName
        End Get
        Set(ByVal Value As String)
            m_FolderNodeName = Value
        End Set
    End Property

    Public Property FolderNodeDescription() As String
        Get
            Return m_FolderNodeDescription
        End Get
        Set(ByVal Value As String)
            m_FolderNodeDescription = Value
        End Set
    End Property

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub
    '//Constructure with parameters which sets 3 properties 
    Public Sub New(ByVal Text As String, ByVal NodeType As String)
        m_FolderNodeType = NodeType
        m_FolderNodeName = Text
        m_GUID = GetNewId()
        m_FolderNodeId = m_GUID
    End Sub

End Class
