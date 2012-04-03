<Serializable()> _
Public Class clsSQFunction
    Implements INode

    Private m_SQFunctionId As String
    Private m_SQFunctionType As String
    Private m_ParaCount As Integer = 0
    Private m_SQFunctionName As String = ""
    Private m_SQFunctionWithInnerText As String = ""
    Private m_SQFunctionSyntax As String = ""
    Private m_SQFunctionDescription As String = ""
    Private m_IsTemplate As Boolean
    Private m_IsModified As Boolean
    Private m_ObjParent As INode
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_IsLoaded As Boolean = True

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
        Try
            Dim obj As New clsSQFunction

            'obj.ObjTreeNode = Me.ObjTreeNode
            obj.IsTemplate = Me.IsTemplate
            obj.ParaCount = Me.ParaCount
            obj.SQFunctionDescription = Me.SQFunctionDescription
            obj.SQFunctionId = Me.SQFunctionId
            obj.SQFunctionName = Me.SQFunctionName
            obj.SQFunctionSyntax = Me.SQFunctionSyntax
            obj.SQFunctionWithInnerText = Me.SQFunctionWithInnerText
            obj.SeqNo = Me.SeqNo
            obj.Parent = NewParent 'Me.Parent

            Return obj

        Catch ex As Exception
            LogError(ex, "clsSQFunction Clone")
            Return Nothing
        End Try

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
            Key = SQFunctionId
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            'If SQFunctionSyntax.Contains("CDCIN") Then
            '    Text = SQFunctionSyntax
            'Else
            Text = SQFunctionName
            'End If
        End Get
        Set(ByVal Value As String)
            SQFunctionName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            If IsTemplate = True Then
                Return NODE_TEMPLATE
            Else
                Return NODE_FUN
            End If
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

    '// Function added 11/3/06 by TK and KS
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Return True

    End Function

    Property IsLoaded() As Boolean Implements INode.IsLoaded
        Get
            Return m_IsLoaded
        End Get
        Set(ByVal value As Boolean)
            m_IsLoaded = value
        End Set
    End Property

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

    Public Property SQFunctionId() As String
        Get
            Return m_SQFunctionId
        End Get
        Set(ByVal Value As String)
            m_SQFunctionId = Value
        End Set
    End Property

    Public Property IsTemplate() As Boolean
        Get
            Return m_IsTemplate
        End Get
        Set(ByVal Value As Boolean)
            m_IsTemplate = Value
        End Set
    End Property

    Public Property SQFunctionName() As String
        Get
            Return m_SQFunctionName
        End Get
        Set(ByVal Value As String)
            '//initializing first time
            If SQFunctionWithInnerText = "" Then
                If SQFunctionSyntax.Contains("CDCIN") Then
                    SQFunctionWithInnerText = SQFunctionSyntax
                Else
                    Dim tmpStr As String = Value
                    tmpStr = tmpStr & "("
                    For i As Integer = 1 To ParaCount
                        If i < ParaCount Then
                            tmpStr = tmpStr & " , "
                        End If
                    Next
                    tmpStr = tmpStr & ")"
                    If Value = "LOOP" Or Value = "IF" Then
                        tmpStr = tmpStr & vbCrLf & "DO" & vbCrLf & vbCrLf & "END"
                    End If
                    SQFunctionWithInnerText = tmpStr
                End If
            End If

            m_SQFunctionName = Value
        End Set
    End Property

    Public Property SQFunctionWithInnerText() As String
        Get
            Return m_SQFunctionWithInnerText
        End Get
        Set(ByVal Value As String)
            m_SQFunctionWithInnerText = Value
        End Set
    End Property

    Public Property SQFunctionSyntax() As String
        Get
            Return m_SQFunctionSyntax
        End Get
        Set(ByVal Value As String)
            m_SQFunctionSyntax = Value
        End Set
    End Property

    Public Property ParaCount() As Integer
        Get
            Return m_ParaCount
        End Get
        Set(ByVal Value As Integer)
            m_ParaCount = Value
        End Set
    End Property

    Public Property SQFunctionDescription() As String
        Get
            Return m_SQFunctionDescription
        End Get
        Set(ByVal Value As String)
            m_SQFunctionDescription = Value
        End Set
    End Property

#End Region

    '//Constructure with parameters which sets 3 properties 
    Public Sub New()
        m_GUID = GetNewId()
    End Sub

    Public Sub New(ByVal Text As String)
        m_GUID = GetNewId()
        SQFunctionName = Text
        '   SQFunctionId = GetNodeId(Me)  '// commented out by TK and KS 11/3/06
    End Sub

End Class
