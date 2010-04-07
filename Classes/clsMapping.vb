Public Class clsMapping
    Implements INode

    Private m_IsMapped As String = "0"
    Private m_MappingTargetValue As String
    Private m_MappingSource As Object '//This can be string or field object (if sourcetype is field)
    Private m_MappingTarget As Object '//This can be string or field object (if targettype is field)
    Private m_SourceType As enumMappingType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    Private m_TargetType As enumMappingType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    Private m_SourceDataStore As String '//new: 8/8/05
    Private m_TargetDataStore As String '//new: 8/8/05
    Private m_ObjTask As clsTask
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_MappingDesc As String = ""
    Private m_IsRenamed As Boolean = False
    Private m_OutMsg As enumOutMsg = enumOutMsg.NoOutMsg

    Public HasMissingDependency As Boolean = False
    Public Commment As String = ""

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
            Return Me.m_ObjTask
        End Get
        Set(ByVal Value As INode)
            m_ObjTask = Value
        End Set
    End Property

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsMapping
        Dim retTask As clsTask
        Dim retFld As clsField
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            obj.SeqNo = Me.SeqNo
            obj.OutMsg = Me.OutMsg
            obj.SourceType = Me.SourceType
            obj.TargetType = Me.TargetType
            obj.SourceDataStore = Me.SourceDataStore
            obj.TargetDataStore = Me.TargetDataStore
            obj.Parent = NewParent 'Me.Parent
            obj.IsMapped = Me.IsMapped

            If Not (Me.MappingSource Is Nothing) Then
                If Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
                    Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Or _
                    Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                    obj.MappingSource = CType(Me.MappingSource, INode).Clone(obj)

                ElseIf Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                    '//locate field under new parent engine
                    If obj.Task.Engine IsNot Nothing Then
                        retFld = SearchField(obj.Task.Engine, Me.MappingSource, enumDirection.DI_SOURCE, cmd)
                    Else
                        retFld = SearchField(obj.Task.Environment, Me.MappingSource)
                    End If

                    If retFld Is Nothing Then
                        MsgBox("This Engine could be missing a required Datastore" & Chr(13) & "Please add sources and targets required before adding Procedure", MsgBoxStyle.OkOnly, "Missing required Datastore(s)")
                        Throw (New Exception("Task [" & obj.Text & "] -> MappingSource Field[" & Me.MappingSource.Text & "] is not found in the target environment"))

                    Else
                        obj.MappingSource = retFld
                    End If

                ElseIf Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_VAR Then
                    '//locate variable under new parent env
                    obj.MappingSource = SearchVariable(obj.Task.Engine, Me.MappingSource)

                ElseIf Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then

                    '//This Mapping Source is Empty
                    obj.MappingSource = Nothing

                Else '/// then it's a task
                    retTask = SearchTask(obj.Task.Engine, Me.MappingSource)
                    If retTask Is Nothing Then
                        MsgBox("This Engine could be missing a required Datastore" & Chr(13) & "Please add sources and targets required before adding Procedure", MsgBoxStyle.OkOnly, "Missing required Datastore(s)")
                        Throw (New Exception("Task [" & obj.Text & "] -> MappingSource Task[" & Me.MappingSource.Text & "] is not found in the target environment"))
                    Else
                        obj.MappingSource = retTask
                    End If
                End If
            End If

            If Not (Me.MappingTarget Is Nothing) Then
                If Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
                    Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Or _
                    Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Or _
                    Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                    obj.MappingTarget = CType(Me.MappingTarget, INode).Clone(obj)

                ElseIf Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                    '//locate field under new parent engine
                    If obj.Task.Engine IsNot Nothing Then
                        retFld = SearchField(obj.Task.Engine, Me.MappingTarget, enumDirection.DI_TARGET, cmd)
                    Else
                        retFld = SearchField(obj.Task.Environment, Me.MappingTarget)
                    End If

                    If retFld Is Nothing Then
                        MsgBox("This Engine could be missing a required Datastore" & Chr(13) & "Please add sources and targets required before adding Procedure", MsgBoxStyle.OkOnly, "Missing required Datastore(s)")
                        Throw (New Exception("Task [" & obj.Text & "] -> MappingTarget Field[" & Me.MappingTarget.Text & "] is not found in the target environment"))
                    Else
                        obj.MappingTarget = retFld
                    End If

                ElseIf Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_VAR Then
                    '//locate variable under new parent env
                    If obj.Task.Engine IsNot Nothing Then
                        obj.MappingTarget = SearchVariable(obj.Task.Engine, Me.MappingTarget)
                    Else
                        obj.MappingTarget = SearchVariable(obj.Task.Environment, Me.MappingTarget)
                    End If

                ElseIf Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then

                    '// SourceType is nothing
                    obj.MappingSource = Nothing

                Else '//TASKS
                    '//locate task under new parent env
                    If obj.Task.Engine IsNot Nothing Then
                        retTask = SearchTask(obj.Task.Engine, Me.MappingTarget)
                    Else
                        retTask = SearchTask(obj.Task.Environment, Me.MappingTarget)
                    End If

                    If retTask Is Nothing Then
                        Throw (New Exception("Task [" & obj.Text & "] -> MappingTarget Task[" & Me.MappingTarget.Text & "] is not found in the target environment"))
                    Else
                        obj.MappingTarget = retTask
                    End If
                End If
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, ex.Message, "clsMapping Clone")
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

    Public ReadOnly Property Key() As String Implements INode.Key
        Get
            Key = 0
            '// all of this commented out by TK and KS 11/3/06

            'Key = MappingId
            'Key = Me.Task.Engine.ObjSystem.Environment.Project.Text  & KEY_SAP &  Me.Task.Engine.ObjSystem.Environment.Text  & KEY_SAP &  Me.Task.Engine.ObjSystem.Text  & KEY_SAP &  Me.Task.Engine.Text  & KEY_SAP &  Me.Task.Text  & KEY_SAP &  Me.Task.TaskType  & KEY_SAP &  Me.Text

        End Get
    End Property

    Public Property Text() As String Implements INode.Text
        Get
            Text = ""
        End Get
        Set(ByVal Value As String)

        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_MAPPING
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.Parent.Project
            Return p
        End Get
    End Property

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save
        '// commented out by TK and KS 11/3/06
        Return True

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew
        '// commented out by TK and KS 11/3/06
        Return True

    End Function

    '// Function added 11/3/06 by TK and KS
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

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

    Public Property MappingTargetValue() As String
        Get
            Return m_MappingTargetValue
        End Get
        Set(ByVal Value As String)
            m_MappingTargetValue = Value
        End Set
    End Property

    Public Property SourceDataStore() As String
        Get
            Return m_SourceDataStore
        End Get
        Set(ByVal Value As String)
            m_SourceDataStore = Value
        End Set
    End Property

    Public Property TargetDataStore() As String
        Get
            Return m_TargetDataStore
        End Get
        Set(ByVal Value As String)
            m_TargetDataStore = Value
        End Set
    End Property

    Public Property SourceParent() As String
        Get
            If SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                Return CType(Me.MappingSource, clsField).ParentStructureName
            Else
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            If SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                CType(Me.MappingSource, clsField).ParentStructureName = Value
            End If
        End Set
    End Property

    Public Property TargetParent() As String
        Get
            If TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                Return CType(Me.MappingTarget, clsField).ParentStructureName
            Else
                Return ""
            End If
        End Get
        Set(ByVal Value As String)
            If TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                CType(Me.MappingTarget, clsField).ParentStructureName = Value
            End If
        End Set
    End Property

    Public Property IsMapped() As String
        Get
            Return m_IsMapped
        End Get
        Set(ByVal Value As String)
            m_IsMapped = Value
        End Set
    End Property

    Public Property MappingSource() As Object
        Get
            Return m_MappingSource
        End Get
        Set(ByVal Value As Object)
            m_MappingSource = Value
        End Set
    End Property

    Public Property MappingTarget() As Object
        Get
            Return m_MappingTarget
        End Get
        Set(ByVal Value As Object)
            m_MappingTarget = Value
        End Set
    End Property

    Public Property SourceType() As enumMappingType
        Get
            Return m_SourceType
        End Get
        Set(ByVal Value As enumMappingType)
            m_SourceType = Value
        End Set
    End Property

    Public Property TargetType() As enumMappingType
        Get
            Return m_TargetType
        End Get
        Set(ByVal Value As enumMappingType)
            m_TargetType = Value
        End Set
    End Property

    Public Property Task() As clsTask
        Get
            Return m_ObjTask
        End Get
        Set(ByVal Value As clsTask)
            m_ObjTask = Value
        End Set
    End Property

    Public Property MappingDesc() As String
        Get
            Return m_MappingDesc
        End Get
        Set(ByVal Value As String)
            m_MappingDesc = Value
        End Set
    End Property

    Public Property OutMsg() As enumOutMsg
        Get
            Return m_OutMsg
        End Get
        Set(ByVal value As enumOutMsg)
            m_OutMsg = value
        End Set
    End Property

#End Region

#Region "Methods"

    '//============ ================== ================== =============== ================ ==========
    '//This function will return FieldID if Source is set to field object otherwise will return string
    '//============ ================== ================== =============== ================ ==========
    Public Function GetMappingSourceVal() As String

        Try
            If Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                '///Use corrected text insted of actual field name  : 9/7/05 issue #48
                If CType(Me.MappingSource, clsField).CorrectedFieldName <> "" Then
                    Return CType(Me.MappingSource, clsField).CorrectedFieldName
                Else
                    Return CType(Me.MappingSource, clsField).FieldName
                End If

            ElseIf Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                Return " "
            Else
                '//Other than Field all are string type
                Return CType(Me.MappingSource, INode).Text
            End If

        Catch ex As Exception
            LogError(ex, "clsMapping GetMappingSourceVal")
            Return ""
        End Try

    End Function

    '//modified 8/15/05
    '//This function will return Id 
    Public Function GetMappingSourceId() As Object

        Try
            If Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Or _
               Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Or _
               Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
               Me.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Then
                Return "''"
            Else
                If CType(Me.MappingSource, INode).GetQuotedText Is Nothing Then
                    Return "''"
                Else
                    Return CType(Me.MappingSource, INode).GetQuotedText
                End If
            End If

        Catch ex As Exception
            LogError(ex, "clsMapping GetMappingSourceId")
            Return Nothing
        End Try

    End Function

    '//modified 8/15/05
    Public Function GetMappingTargetId() As Object

        Try
            If Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT Then
                Return "''"
            Else
                If CType(Me.MappingTarget, INode).GetQuotedText Is Nothing Then
                    Return "''"
                Else
                    Return CType(Me.MappingTarget, INode).GetQuotedText
                End If
            End If

        Catch ex As Exception
            LogError(ex, "clsMapping GetMappingTargetId")
            Return Nothing
        End Try

    End Function

    '//This function will return FieldID if Souce is set to field object otherwise will return string
    Public Function GetMappingTargetVal() As String

        Try
            If Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                '///Use corrected text insted of actual field name  : 9/7/05 issue #48
                If CType(Me.MappingTarget, clsField).CorrectedFieldName <> "" Then
                    Return CType(Me.MappingTarget, clsField).CorrectedFieldName
                Else
                    Return CType(Me.MappingTarget, clsField).FieldName
                End If

            ElseIf Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                If Me.MappingTarget Is Nothing Then
                    Dim tempObj As New clsField
                    Me.MappingTarget = tempObj
                End If

                '// added 11/3/06 by KS and TK
                CType(Me.MappingTarget, clsField).CorrectedFieldName = "__COL" & Me.SeqNo

                Return CType(Me.MappingTarget, clsField).CorrectedFieldName
            ElseIf Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                If CType(Me.MappingTarget, clsVariable).CorrectedVariableName <> "" Then
                    Return CType(Me.MappingTarget, clsVariable).CorrectedVariableName
                Else
                    Return CType(Me.MappingTarget, clsVariable).Text
                End If
            Else
                '//Other than Field all are string type
                Return CType(Me.MappingTarget, INode).Text
            End If

        Catch ex As Exception
            LogError(ex, "clsMapping GetMappingTargetVal")
            Return ""
        End Try

    End Function

    Public Function GetOldMappingTargetVal() As String

        Try
            If Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                Return Me.MappingTargetValue
            Else
                Return CType(Me.MappingTarget, INode).Text
            End If

        Catch ex As Exception
            LogError(ex, "clsMapping GetOldMappingTargetVal")
            Return ""
        End Try

    End Function

    Public Function UpdateMappingTargetValue() As Boolean

        Try
            If Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Or _
               Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                If CType(Me.MappingTarget, clsField).CorrectedFieldName <> "" Then
                    Me.MappingTargetValue = CType(Me.MappingTarget, clsField).CorrectedFieldName
                End If
            ElseIf Me.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
                If CType(Me.MappingTarget, clsVariable).CorrectedVariableName <> "" Then
                    Me.MappingTargetValue = CType(Me.MappingTarget, clsVariable).CorrectedVariableName
                    CType(Me.MappingTarget, clsVariable).Text = CType(Me.MappingTarget, clsVariable).CorrectedVariableName
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "clsMapping UpdateMappingTargetValue")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
