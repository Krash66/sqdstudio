<Serializable()> _
Public Class clsTask
    Implements INode

    Private m_TaskType As enumTaskType
    Private m_TaskName As String = ""
    Private m_TaskDescription As String = ""
    Private m_ObjEngine As clsEngine = Nothing
    Private m_ObjEnv As clsEnvironment = Nothing
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_ObjParent As INode
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_IsLoaded As Boolean = False

    '//8/15/05 : we dont allow datastore update (add/delete task datastore) from usercontrol so when save 
    '//from usercontrol set this flag so we skip datastore operation
    Public CallFromUsercontrol As Boolean
    Public IsDisturbed As Boolean = False
    '// Sources and Targets can be one or more than one depending on type of task
    Public ObjSources As New ArrayList
    Public ObjTargets As New ArrayList
    Public ObjMappings As New ArrayList '//Array of selected Mappings

    '// Old mappings, this is helpful when doing Save operation, we can compare old mappings and new mappings 
    '// and perform insert/delete/update based on difference
    '// This array should be updated before we change ObjMappings so we can have old values for comparision
    Public OldObjMappings As New ArrayList

    '/// Last Source and Last Target Mapped
    Private g_LastSrcFld As String = ""
    Private g_LastTgtFld As String = ""

    '///Writing Last Mapped Fields Src and Tgt
    Dim fsLastFlds As System.IO.FileStream
    Dim objWriteLastFlds As System.IO.StreamWriter
    Dim objReadLastFlds As System.IO.StreamReader

    '/// AddFlow additions
    Private m_AFnode As Node
    Private m_Hloc As Integer
    Private m_Vloc As Integer
    'Public OutLinks As Collection


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

    '//This requires special handling. 
    Public Property Parent() As INode Implements INode.Parent
        Get
            If m_ObjParent Is Nothing Then
                If m_ObjEngine Is Nothing Then
                    Return Me.Environment
                Else
                    Return Me.Engine
                End If
            Else
                Return m_ObjParent
            End If
        End Get
        Set(ByVal Value As INode)
            If Value.Type = NODE_ENGINE Then
                m_ObjEngine = Value
                Me.Engine = Value
            End If
            If Value.Type = NODE_ENVIRONMENT Then
                m_ObjEnv = Value
                Me.Environment = Value
            End If
            m_ObjParent = Value
        End Set
    End Property

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsTask
        Dim retDS As clsDatastore
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            obj.TaskName = Me.TaskName
            obj.TaskDescription = Me.TaskDescription
            obj.TaskType = Me.TaskType
            obj.SeqNo = Me.SeqNo
            obj.IsModified = Me.IsModified
            obj.Engine = NewParent
            obj.LastSrcFld = Me.LastSrcFld
            obj.LastTgtFld = Me.LastTgtFld
            obj.Hloc = Me.Hloc
            obj.Vloc = Me.Vloc

            'obj.Parent = NewParent 'Me.Parent

            '// Make sure All items are loaded first
            Me.LoadDatastores(cmd)
            Me.LoadMappings(True, cmd)
            'Me.LoadMe()
            'Me.LoadItems()


            For Each ds As clsDatastore In Me.ObjSources
                'If Me.Engine IsNot Nothing Then
                retDS = SearchDatastore(obj.Engine, ds)
                'Else
                'retDS = SearchDatastore(Me.Environment, ds)
                'End If

                If retDS Is Nothing Then
                    Throw (New Exception("Source datastore " & ds.Text & " is not found in the target environment"))
                Else
                    'Dim NewDS As clsDatastore
                    'NewDS = ds.Clone(obj, True, cmd)
                    obj.ObjSources.Add(retDS)
                End If
            Next

            For Each ds As clsDatastore In Me.ObjTargets
                'If Me.Engine IsNot Nothing Then
                retDS = SearchDatastore(obj.Engine, ds)
                'Else
                'retDS = SearchDatastore(Me.Environment, ds)
                'End If

                If retDS Is Nothing Then
                    Throw (New Exception("Target datastore " & ds.Text & " is not found in the target environment"))
                Else
                    'Dim NewDS As clsDatastore
                    'NewDS = ds.Clone(obj, True, cmd)
                    obj.ObjTargets.Add(retDS)
                End If
            Next

            If Cascade = True Then
                '//clone all mappings for new object
                For Each map As clsMapping In Me.ObjMappings
                    Dim NewMap As New clsMapping
                    NewMap = map.Clone(obj, True, cmd)
                    NewMap.Task = obj
                    obj.ObjMappings.Add(NewMap)
                Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsTask Clone")
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
            If Me.Engine IsNot Nothing Then
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Engine.ObjSystem.Text & KEY_SAP & Me.Engine.Text & KEY_SAP & Me.Text '& KEY_SAP & Me.TaskType.ToString
            Else
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text '& KEY_SAP & Me.TaskType.ToString
            End If
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = TaskName
        End Get
        Set(ByVal Value As String)
            TaskName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Select Case Me.TaskType
                Case modDeclares.enumTaskType.TASK_MAIN
                    Return NODE_MAIN
                Case modDeclares.enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc
                    Return NODE_PROC
                Case modDeclares.enumTaskType.TASK_GEN
                    Return NODE_GEN
                Case modDeclares.enumTaskType.TASK_LOOKUP
                    Return NODE_LOOKUP
                Case Else
                    Return ""
            End Select
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

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            '//We need to put in transaction because we will add Datastore,
            '//DSSELECTIONS and fields in 
            '//three steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '//8/15/05 :skip this step if called from usercontrol
            'If Me.CallFromUsercontrol = False Then
            If UpdateTaskDataStores(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If
            'End If

            If DeleteMappings(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            If AddNewMappings(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            tran.Commit()

            Me.IsModified = False

            Save = True

        Catch OE As Odbc.OdbcException
            tran.Rollback()
            LogODBCError(OE, "clsTask Save")
            Save = False
        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "clsTask Save")
            Save = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '//first delete child records
            'If Me.Engine IsNot Nothing Then
            sql = "Delete From " & Me.Project.tblTaskDS & _
                        " where TaskType=" & Me.TaskType & _
                        " AND TaskName=" & Me.GetQuotedText & _
                        " AND EngineName=" & Me.Engine.GetQuotedText & _
                        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                        " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            'sql = "Delete From " & Me.Project.tblTaskDS & _
            '            " where TaskType=" & Me.TaskType & _
            '            " AND TaskName=" & Me.GetQuotedText & _
            '            " AND EngineName=" & Quote(DBNULL) & _
            '            " AND SystemName=" & Quote(DBNULL) & _
            '            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '            " AND ProjectName=" & Me.Project.GetQuotedText
            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            'If Me.Engine IsNot Nothing Then
            sql = "Delete From " & Me.Project.tblTaskMap & _
                        " where  TaskType=" & Me.TaskType & _
                        " AND TaskName=" & Me.GetQuotedText & _
                        " AND EngineName=" & Me.Engine.GetQuotedText & _
                        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                        " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            'sql = "Delete From " & Me.Project.tblTaskMap & _
            '            " where  TaskType=" & Me.TaskType & _
            '            " AND TaskName=" & Me.GetQuotedText & _
            '            " AND EngineName=" & Quote(DBNULL) & _
            '            " AND SystemName=" & Quote(DBNULL) & _
            '            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '            " AND ProjectName=" & Me.Project.GetQuotedText
            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            'If Me.Engine IsNot Nothing Then
            sql = "Delete From " & Me.Project.tblTasks & _
                       " where  TaskType=" & Me.TaskType & _
                       " AND TaskName=" & Me.GetQuotedText & _
                       " AND EngineName=" & Me.Engine.GetQuotedText & _
                       " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                       " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                       " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            'sql = "Delete From " & Me.Project.tblTasks & _
            '           " where  TaskType=" & Me.TaskType & _
            '           " AND TaskName=" & Me.GetQuotedText & _
            '           " AND EngineName=" & Quote(DBNULL) & _
            '           " AND SystemName=" & Quote(DBNULL) & _
            '           " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '           " AND ProjectName=" & Me.Project.GetQuotedText
            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                If Me.Engine IsNot Nothing Then
                    Select Case Me.Type
                        Case NODE_MAIN
                            RemoveFromCollection(Me.Engine.Mains, Me.GUID)
                            'Case NODE_GEN
                            '    RemoveFromCollection(Me.Engine.Gens, Me.GUID)
                        Case NODE_PROC, NODE_GEN
                            RemoveFromCollection(Me.Engine.Procs, Me.GUID)
                        Case NODE_LOOKUP
                            RemoveFromCollection(Me.Engine.Lookups, Me.GUID)
                    End Select
                Else
                    RemoveFromCollection(Me.Environment.Procedures, Me.GUID)
                End If

                '/// AddFlow Additions
                Me.AFnode.InLinks.Clear()
                Me.AFnode.OutLinks.Clear()
                Me.AFnode.Remove()

            End If

            Delete = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask Delete", sql)
            Delete = False
        Catch ex As Exception
            LogError(ex, "clsTask Delete", sql)
            Delete = False
        End Try

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            '//We need to put in transaction because we will add Datastore,
            '//DSSELECTIONS and fields in 
            '//three steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '/////////////////////////////////////////////////////////////////////////
            '//1) Insert in to TASKS table
            '/////////////////////////////////////////////////////////////////////////
            '//When we add new record we need to find Unique Database Record ID.
            If AddNewTask(cmd) = False Then
                tran.Rollback()
                Exit Try
            End If

            '/////////////////////////////////////////////////////////////////////////
            '//2) Insert in to TASKDS
            '/////////////////////////////////////////////////////////////////////////
            If UpdateTaskDataStores(cmd) = False Then
                tran.Rollback()
                Exit Try
            End If

            '/////////////////////////////////////////////////////////////////////////
            '//3) Insert in to TASKMAP table
            '/////////////////////////////////////////////////////////////////////////
            If Cascade = True Then
                If ObjMappings.Count > 0 Then
                    If AddNewMappings(cmd) = False Then
                        tran.Rollback()
                        Exit Try
                    End If
                End If
            End If


            tran.Commit()


            ' If Me.TaskType = modDeclares.enumTaskType.TASK_GEN Then
            '//Add Join into Engines's Proc collection
            'AddToCollection(Me.Engine.Procs, Me, Me.GUID)
            'Else
            If Me.TaskType = enumTaskType.TASK_LOOKUP Then
                '//Add Lookup into Engines's Lookup collection
                AddToCollection(Me.Engine.Lookups, Me, Me.GUID)
            ElseIf Me.TaskType = enumTaskType.TASK_MAIN Then
                '//Add Main into Engines's Main collection
                AddToCollection(Me.Engine.Mains, Me, Me.GUID)
            ElseIf Me.TaskType = enumTaskType.TASK_PROC Or _
            Me.TaskType = enumTaskType.TASK_IncProc Or _
            Me.TaskType = enumTaskType.TASK_GEN Then
                '//Add Proc into Engines's Proc collection
                'If Me.Engine IsNot Nothing Then
                AddToCollection(Me.Engine.Procs, Me, Me.GUID)
                If Me.TaskType = enumTaskType.TASK_GEN Then
                    AddToCollection(Me.Engine.Gens, Me, Me.GUID)
                End If
                'Else
                '    AddToCollection(Me.Environment.Procedures, Me, Me.GUID)
                'End If
            End If

            AddNew = True
            Me.IsModified = False

        Catch OE As Odbc.OdbcException
            tran.Rollback()
            LogODBCError(OE, "clsTask AddNew")
            AddNew = False
        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "clsTask AddNew")
            AddNew = False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// added by TK and KS 11/6/2006
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim success As Boolean = False
        Try
            Me.Text = Me.Text.Trim
            '/////////////////////////////////////////////////////////////////////////
            '//1) Insert in to TASKS table
            '/////////////////////////////////////////////////////////////////////////
            success = AddNewTask(cmd)
            '/////////////////////////////////////////////////////////////////////////
            '//2) Insert in to TASKDS
            '/////////////////////////////////////////////////////////////////////////
            If success = True Then
                success = UpdateTaskDataStores(cmd)
            End If
            '/////////////////////////////////////////////////////////////////////////
            '//3) Insert in to TASKMAP table
            '/////////////////////////////////////////////////////////////////////////
            If success = True Then
                If ObjMappings.Count > 0 Then
                    success = AddNewMappings(cmd)
                End If
            End If
           

            ' If Me.TaskType = modDeclares.enumTaskType.TASK_GEN Then
            'AddToCollection(Me.Engine.Gens, Me, Me.GUID)
            'Else
            If Me.TaskType = modDeclares.enumTaskType.TASK_LOOKUP Then
                AddToCollection(Me.Engine.Lookups, Me, Me.GUID)
            ElseIf Me.TaskType = modDeclares.enumTaskType.TASK_MAIN Then
                AddToCollection(Me.Engine.Mains, Me, Me.GUID)
            ElseIf Me.TaskType = modDeclares.enumTaskType.TASK_PROC Or _
            Me.TaskType = modDeclares.enumTaskType.TASK_IncProc Or _
            Me.TaskType = modDeclares.enumTaskType.TASK_GEN Then
                If Me.Engine IsNot Nothing Then
                    AddToCollection(Me.Engine.Procs, Me, Me.GUID)
                    If Me.TaskType = enumTaskType.TASK_GEN Then
                        AddToCollection(Me.Engine.Gens, Me, Me.GUID)
                    End If
                End If
                AddToCollection(Me.Environment.Procedures, Me, Me.GUID)
            End If

            AddNew = success
            Me.IsModified = False

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask AddNew")
            AddNew = False
        Catch ex As Exception
            LogError(ex, "clsTask AddNew")
            AddNew = False
        End Try

    End Function

    Public Property IsRenamed() As Boolean Implements INode.IsRenamed
        Get
            Return m_IsRenamed
        End Get
        Set(ByVal Value As Boolean)
            m_IsRenamed = Value
        End Set
    End Property

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Try
            'If Me.IsLoaded = True Then Exit Function

            If TreeLode = False Then
                Dim cmd As New System.Data.Odbc.OdbcCommand

                If Incmd IsNot Nothing Then
                    cmd = Incmd
                Else
                    cmd = New Odbc.OdbcCommand
                    cmd.Connection = cnn
                End If

                'cmd.Connection.Open()


                LoadDatastores(cmd)
                LoadMappings(Reload, cmd)
                Me.IsLoaded = Reload
            End If

            LoadItems = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask LoadItems")
            LoadItems = False
        Catch ex As Exception
            LogError(ex, "clsTask LoadItems")
            LoadItems = False
        End Try


    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then
                LoadMe = True
                Exit Try
            End If


            Dim cmd As New System.Data.Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            LoadDatastores(cmd)
            LoadMappings(True, cmd)

            Me.IsLoaded = True

            LoadMe = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask LoadMe")
            LoadMe = False
        Catch ex As Exception
            LogError(ex, "clsTask LoadMe")
            LoadMe = False
        End Try

    End Function

    Property IsLoaded() As Boolean Implements INode.IsLoaded
        Get
            Return m_IsLoaded
        End Get
        Set(ByVal value As Boolean)
            m_IsLoaded = value
        End Set
    End Property

    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""
        Dim readName As String
        Dim testName As String = NewName
        Dim NameValid As Boolean = True

        Try
            If testName = "" Then
                testName = Me.Text
            End If

            'cnn.Open()
            cmd = cnn.CreateCommand

            'If Me.Engine IsNot Nothing Then
            sql = "Select TASKNAME from " & Me.Project.tblTasks & _
                           " where PROJECTNAME=" & Me.Project.GetQuotedText & _
                           " AND ENVIRONMENTNAME=" & Me.Environment.GetQuotedText & _
                           " AND SYSTEMNAME=" & Me.Engine.ObjSystem.GetQuotedText & _
                           " AND ENGINENAME=" & Me.Engine.GetQuotedText
            'Else
            'sql = "Select TASKNAME from " & Me.Project.tblTasks & _
            '            " where PROJECTNAME= '" & Me.Project.ProjectName & _
            '            "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & "'"
            'End If


            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = GetVal(dr.Item("TASKNAME"))
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new Task name you have chosen already exists," & (Chr(13)) & "in this Engine, please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            ValidateNewObject = NameValid

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask ValidateNewObject", sql)
            ValidateNewObject = False
        Catch ex As Exception
            LogError(ex, "clsTask ValidateNewObject", sql)
            ValidateNewObject = False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property TaskName() As String
        Get
            Return m_TaskName
        End Get
        Set(ByVal Value As String)
            m_TaskName = Value
        End Set
    End Property

    Public Property TaskType() As enumTaskType
        Get
            Return m_TaskType
        End Get
        Set(ByVal Value As enumTaskType)
            m_TaskType = Value
        End Set
    End Property

    Public Property TaskDescription() As String
        Get
            Return m_TaskDescription
        End Get
        Set(ByVal Value As String)
            m_TaskDescription = Value
        End Set
    End Property

    Public Property Engine() As clsEngine
        Get
            Return m_ObjEngine
        End Get
        Set(ByVal Value As clsEngine)
            m_ObjEngine = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            If m_ObjEngine Is Nothing Then
                Return m_ObjEnv
            Else
                Return Me.Engine.ObjSystem.Environment
            End If
        End Get
        Set(ByVal value As clsEnvironment)
            m_ObjEnv = value
        End Set
    End Property

    Public Property LastSrcFld() As String
        Get
            Return g_LastSrcFld
        End Get
        Set(ByVal value As String)
            g_LastSrcFld = value
        End Set
    End Property

    Public Property LastTgtFld() As String
        Get
            Return g_LastTgtFld
        End Get
        Set(ByVal value As String)
            g_LastTgtFld = value
        End Set
    End Property

    Public Property AFnode() As Node
        Get
            Return m_AFnode
        End Get
        Set(ByVal value As Node)
            m_AFnode = value
        End Set
    End Property

    Public Property Hloc() As Integer
        Get
            Return m_Hloc
        End Get
        Set(ByVal value As Integer)
            m_Hloc = value
        End Set
    End Property

    Public Property Vloc() As Integer
        Get
            Return m_Vloc
        End Get
        Set(ByVal value As Integer)
            m_Vloc = value
        End Set
    End Property

    Public ReadOnly Property HasLocation() As Boolean
        Get
            If Hloc > 0 And Vloc > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

#End Region

#Region "Methods"

    Public Function SaveSeqNo(Optional ByRef P_cmd As Odbc.OdbcCommand = Nothing) As Boolean

        Dim sql As String = ""
        Try
            '/// modified by Tom Karasch 6/07
            Dim cmd As New System.Data.Odbc.OdbcCommand

            'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing

            If P_cmd Is Nothing Then
                cmd = New Odbc.OdbcCommand
            Else
                cmd = P_cmd
            End If

            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()

            cmd.Connection = cnn

            'If Me.Engine IsNot Nothing Then
            '    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '        sql = "Update " & Me.Project.tblTasks & _
            '        " set SeqNo=" & Me.SeqNo & _
            '        " where TaskType=" & Me.TaskType & _
            '        " AND TaskName=" & Me.GetQuotedText & _
            '        " AND EngineName=" & Me.Engine.GetQuotedText & _
            '        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            '        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '        " AND ProjectName=" & Me.Project.GetQuotedText
            '    Else
            sql = "Update " & Me.Project.tblTasks & _
            " set TASKSEQNO=" & Me.SeqNo & _
            " where TaskType=" & Me.TaskType & _
            " AND TaskName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText
            '    End If
            'Else
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set SeqNo=" & Me.SeqNo & _
            '    " where TaskType=" & Me.TaskType & _
            '    " AND TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set TASKSEQNO=" & Me.SeqNo & _
            '    " where TaskType=" & Me.TaskType & _
            '    " AND TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'End If
            'End If

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            Me.IsModified = False
            SaveSeqNo = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask SaveSeqNo", sql)
            SaveSeqNo = False
        Catch ex As Exception
            LogError(ex, "clsTask SaveSeqNo", sql)
            SaveSeqNo = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function SaveTaskType(Optional ByRef P_cmd As Odbc.OdbcCommand = Nothing) As Boolean

        '/// modified by Tom Karasch 6/07
        Dim sql As String = ""
        Try
            Dim cmd As New System.Data.Odbc.OdbcCommand

            'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing

            If P_cmd Is Nothing Then
                cmd = New Odbc.OdbcCommand
            Else
                cmd = P_cmd
            End If

            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()

            cmd.Connection = cnn

            'If Me.Engine IsNot Nothing Then
            '    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '        sql = "Update " & Me.Project.tblTasks & _
            '        " set TaskType=" & Me.TaskType & _
            '        " where TaskName=" & Me.GetQuotedText & _
            '        " AND EngineName=" & Me.Engine.GetQuotedText & _
            '        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            '        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '        " AND ProjectName=" & Me.Project.GetQuotedText
            '    Else
            Sql = "Update " & Me.Project.tblTasks & _
            " set TaskType=" & Me.TaskType & _
            " where TaskName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText
            '    End If
            'Else
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set TaskType=" & Me.TaskType & _
            '    " where TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set TaskType=" & Me.TaskType & _
            '    " where TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'End If
            'End If

            Log(Sql)
            cmd.CommandText = Sql
            cmd.ExecuteNonQuery()

            Me.IsModified = False
            SaveTaskType = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask SaveTaskType", sql)
            SaveTaskType = False
        Catch ex As Exception
            LogError(ex, "clsTask SaveTaskType", sql)
            SaveTaskType = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function LoadDatastores(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Integer

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        'Dim cmd As System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""

        Try
            '//check if already loaded ?
            If Me.ObjSources.Count > 0 And Me.ObjTargets.Count > 0 Then Exit Function
            'If  Then Exit Function

            Dim cmd As System.Data.Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            'If Me.Engine IsNot Nothing Then
            sql = "Select ds.DSTYPE,tds.PROJECTNAME,tds.ENVIRONMENTNAME,tds.SYSTEMNAME,tds.ENGINENAME,tds.TASKNAME,tds.TASKTYPE,tds.DATASTORENAME,tds.DSDIRECTION from " & Me.Project.tblTaskDS & " tds" & _
                        " inner join " & Me.Project.tblDatastores & " ds on tds.ProjectName=ds.ProjectName" & _
                        " AND tds.EnvironmentName=ds.EnvironmentName" & _
                        " AND tds.SystemName=ds.SystemName" & _
                        " AND tds.EngineName=ds.EngineName" & _
                        " AND tds.DataStoreName=ds.DataStoreName" & _
                        " where tds.ProjectName=" & Me.Project.GetQuotedText & _
                        " AND tds.EnvironmentName=" & Me.Environment.GetQuotedText & _
                        " AND tds.SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                        " AND tds.EngineName=" & Me.Engine.GetQuotedText & _
                        " AND tds.TaskName=" & Me.GetQuotedText & _
                        " AND tds.TaskType=" & Me.TaskType
            'ds.*,tds.DSDIRECTION
            'Else
            'sql = "Select ds.DSTYPE,tds.PROJECTNAME,tds.ENVIRONMENTNAME,tds.SYSTEMNAME,tds.ENGINENAME,tds.TASKNAME,tds.TASKTYPE,tds.DATASTORENAME,tds.DSDIRECTION from " & Me.Project.tblTaskDS & " tds" & _
            '            " inner join " & Me.Project.tblDatastores & " ds on tds.ProjectName=ds.ProjectName" & _
            '            " AND tds.EnvironmentName=ds.EnvironmentName" & _
            '            " AND tds.SystemName=ds.SystemName" & _
            '            " AND tds.EngineName=ds.EngineName" & _
            '            " AND tds.DataStoreName=ds.DataStoreName" & _
            '            " where tds.ProjectName=" & Me.Project.GetQuotedText & _
            '            " AND tds.EnvironmentName=" & Me.Environment.GetQuotedText & _
            '            " AND tds.SystemName=" & Quote(DBNULL) & _
            '            " AND tds.EngineName=" & Quote(DBNULL) & _
            '            " AND tds.TaskName=" & Me.GetQuotedText & _
            '            " AND tds.TaskType=" & Me.TaskType
            'End If


            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            Me.ObjTargets.Clear()
            Me.ObjSources.Clear()

            While dr.Read
                Dim objDs As clsDatastore

                'If Me.Engine IsNot Nothing Then
                objDs = SearchDatastoreByName(Me.Engine, GetVal(dr("DatastoreName")), GetVal(dr("EngineName")), _
                GetVal(dr("SystemName")), GetVal(dr("EnvironmentName")), GetVal(dr("ProjectName")), GetVal(dr("DsDirection")))

                'Else
                'objDs = SearchDatastoreByName(Me.Environment, GetVal(dr("DatastoreName")), GetVal(dr("EnvironmentName")), _
                'GetVal(dr("ProjectName")))
                'objDs.DsDirection = GetVal(dr("DsDirection"))
                'End If

                If objDs Is Nothing Then
                    clsLogging.LogEvent("Task Datastore [" & dr("DatastoreName") & "] is not found under this engine")
                ElseIf GetVal(dr("TASKNAME")) = Me.TaskName Then
                    If objDs.DsDirection = DS_DIRECTION_SOURCE Then
                        Me.ObjSources.Add(objDs)
                    Else 'If objDs.DsDirection = DS_DIRECTION_TARGET Then
                        Me.ObjTargets.Add(objDs)
                        'Else
                        '    ObjSources.Add(objDs)
                        '    ObjTargets.Add(objDs)
                        'MsgBox("Unknown datastore direction for DS[" & dr("DatastoreName") & "]", MsgBoxStyle.Exclamation, MsgTitle)
                    End If
                End If
            End While
            dr.Close()

            '/// don't know why this function is "as integer"
            '/// doesn't return any value
        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask LoadDatastores")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
        Catch ex As Exception
            LogError(ex, "clsTask LoadDatastores", sql)
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function LoadMappings(Optional ByVal Reload As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Integer

        Dim sql As String = ""

        Try
            If Reload = False Then
                '//check if already loaded ?
                If Me.IsDisturbed = False Then
                    If Me.ObjMappings.Count > 0 Then Exit Function
                Else
                    Me.IsDisturbed = False
                End If
            End If

            'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
                'Dim cmd As System.Data.Odbc.OdbcCommand
            Dim dr As System.Data.Odbc.OdbcDataReader
                'Dim objHash As New Hashtable

            Dim cmd As System.Data.Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

                'cmd.Connection.Open()

                'cmd = cmd.Connection.CreateCommand

                'If Me.Engine IsNot Nothing Then

            sql = "Select MAPPINGID,MAPPINGDESC,MAPPINGTARGET,SOURCETYPE,TARGETTYPE,ISMAPPED,MAPPINGSOURCEID,MAPPINGTARGETID,SOURCEPARENT,TARGETPARENT,SEQNO,SOURCEDATASTORE,TARGETDATASTORE,MAPPINGSOURCE from " & Me.Project.tblTaskMap & _
                           " where TaskType=" & Me.TaskType & _
                           " AND TaskName=" & Me.GetQuotedText & _
                           " AND EngineName=" & Me.Engine.GetQuotedText & _
                           " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                           " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                           " AND ProjectName=" & Me.Project.GetQuotedText & " order by MAPPINGID"
                'Else
                'sql = "Select MAPPINGID,MAPPINGDESC,MAPPINGTARGET,SOURCETYPE,TARGETTYPE,ISMAPPED,MAPPINGSOURCEID,MAPPINGTARGETID,SOURCEPARENT,TARGETPARENT,SEQNO,SOURCEDATASTORE,TARGETDATASTORE,MAPPINGSOURCE from " & Me.Project.tblTaskMap & _
                '           " where TaskType=" & Me.TaskType & _
                '           " AND TaskName=" & Me.GetQuotedText & _
                '           " AND EngineName=" & Quote(DBNULL) & _
                '           " AND SystemName=" & Quote(DBNULL) & _
                '           " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                '           " AND ProjectName=" & Me.Project.GetQuotedText & " order by MAPPINGID"
                'End If


            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            ObjMappings.Clear()

            If Me.TaskType <> enumTaskType.TASK_MAIN And Me.TaskType <> enumTaskType.TASK_GEN Then
                While dr.Read
                    Dim objMap As New clsMapping

                    objMap.IsMapped = GetVal(dr.Item("IsMapped"))
                    If GetVal(dr.Item("SeqNo")) = 1 Then
                        objMap.OutMsg = enumOutMsg.PrintOutMsg
                    Else
                        objMap.OutMsg = enumOutMsg.NoOutMsg
                    End If
                    objMap.SeqNo = GetVal(dr.Item("MappingID"))
                    objMap.MappingDesc = GetVal(dr.Item("MappingDesc"))
                    objMap.Parent = Me

                    objMap.SourceType = GetVal(dr.Item("SourceType"))
                    objMap.TargetType = GetVal(dr.Item("TargetType"))
                    objMap.SourceDataStore = GetVal(dr.Item("SourceDataStore"))
                    objMap.TargetDataStore = GetVal(dr.Item("TargetDataStore"))

                    objMap.MappingSource = BuildMappingSourceTarget(modDeclares.enumDirection.DI_SOURCE, dr)
                    objMap.MappingTarget = BuildMappingSourceTarget(modDeclares.enumDirection.DI_TARGET, dr)
                    objMap.MappingTargetValue = GetVal(dr.Item("MappingTarget"))

                    '//new: added on  7/23/05 to avoid blank mappings if no dependet object found
                    If objMap.MappingSource Is Nothing Or objMap.MappingTarget Is Nothing Then
                        'objMap.HasMissingDependency = True
                        If objMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And objMap.MappingSource Is Nothing Then
                            If objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                                objMap.Commment = "Parent object [" & GetVal(dr.Item("SourceParent")) & "] for MappingSource [" & _
                                GetVal(dr.Item("MappingSource")) & "] is missing" & vbCrLf
                            Else
                                objMap.Commment = "Object for MappingSource [" & GetVal(dr.Item("MappingSource")) & "] is missing" & _
                                vbCrLf
                            End If
                        End If

                        If objMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And objMap.MappingTarget Is Nothing Then
                            If objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
                                objMap.Commment = objMap.Commment & "Parent object [" & GetVal(dr.Item("TargetParent")) & _
                                "] for MappingTarget [" & GetVal(dr.Item("MappingTarget")) & "] is missing" & vbCrLf
                            Else
                                objMap.Commment = objMap.Commment & "Object for MappingTarget [" & GetVal(dr.Item("MappingTarget")) & _
                                "] is missing" & vbCrLf
                            End If
                        End If

                        objMap.MappingSource = BuildMappingSourceTarget(modDeclares.enumDirection.DI_SOURCE, dr, True)
                        objMap.MappingTarget = BuildMappingSourceTarget(modDeclares.enumDirection.DI_TARGET, dr, True)
                        objMap.MappingTargetValue = GetVal(dr.Item("MappingTarget"))
                    End If

                    ObjMappings.Add(objMap)
                End While
            Else
                While dr.Read
                    Dim objMap As New clsMapping

                    objMap.IsMapped = GetVal(dr.Item("IsMapped"))
                    If GetVal(dr.Item("SeqNo")) = 1 Then
                        objMap.OutMsg = enumOutMsg.PrintOutMsg
                    Else
                        objMap.OutMsg = enumOutMsg.NoOutMsg
                    End If
                    objMap.SeqNo = GetVal(dr.Item("MappingID"))
                    objMap.MappingDesc = GetVal(dr.Item("MappingDesc"))
                    objMap.Parent = Me

                    objMap.SourceType = enumMappingType.MAPPING_TYPE_FUN 'GetVal(dr.Item("SourceType"))
                    objMap.TargetType = enumMappingType.MAPPING_TYPE_NONE 'GetVal(dr.Item("TargetType"))
                    objMap.SourceDataStore = GetVal(dr.Item("SourceDataStore"))
                    objMap.TargetDataStore = Nothing 'GetVal(dr.Item("TargetDataStore"))

                    objMap.MappingSource = BuildMappingSourceTarget(modDeclares.enumDirection.DI_SOURCE, dr)
                    objMap.MappingTarget = Nothing 'BuildMappingSourceTarget(modDeclares.enumDirection.DI_TARGET, dr)
                    objMap.MappingTargetValue = Nothing 'GetVal(dr.Item("MappingTarget"))

                    ObjMappings.Add(objMap)
                End While
            End If

            dr.Close()

                '/// "As Integer" not used
                '///doesn't return anything
        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask LoadMappings", sql)
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
        Catch ex As Exception
            LogError(ex, "clsTask LoadMappings", sql)
        Finally
            'cnn.Close()
        End Try

    End Function

    '// added by TK and KS 11/6/2006
    Public Function AutoMap(ByVal taskType As modDeclares.enumTaskType, Optional ByVal seqNo As Integer = 0, Optional ByVal sDs As String = "", Optional ByVal tDs As String = "", Optional ByVal sFld As INode = Nothing, Optional ByVal tFld As INode = Nothing, Optional ByVal parentName As String = "") As Integer

        'Dim objHash As New Hashtable
        Dim objMap As New clsMapping
        Dim mainFunc As clsSQFunction

        Try
            If (taskType = modDeclares.enumTaskType.TASK_PROC) Then
                objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
                objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
                objMap.IsMapped = "1"
                objMap.MappingSource = CType(sFld, clsField)
                objMap.MappingTarget = CType(tFld, clsField)
                objMap.MappingTargetValue = CType(tFld, clsField).FieldName
            Else
                objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN
                objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
                mainFunc = CType(sFld, clsSQFunction)
                mainFunc.SQFunctionWithInnerText = GetMainText()
                mainFunc.SQFunctionSyntax = mainFunc.SQFunctionWithInnerText
                'mainFunc.SQFunctionName = mainFunc.SQFunctionSyntax

                objMap.MappingSource = mainFunc
                objMap.MappingTarget = Nothing
                objMap.MappingTargetValue = ""
            End If

            objMap.SourceDataStore = sDs
            objMap.TargetDataStore = tDs

            objMap.SeqNo = seqNo
            objMap.Parent = Me

            ObjMappings.Add(objMap)

        Catch ex As Exception
            LogError(ex, "clsTask AutoMap")
        End Try

    End Function

    '// no longer private function  by KS and TK 11/6/2006
    Function BuildMappingSourceTarget(ByVal Direction As enumDirection, ByVal dr As IDataReader, Optional ByVal ReturnDummyObj As Boolean = False) As INode

        Dim sType As enumMappingType
        Dim sInput As String = ""
        Dim retObj As INode
        Dim ID As String
        Dim ParentName As String
        Dim Text As String

        retObj = Nothing
        Try

            If Direction = modDeclares.enumDirection.DI_SOURCE Then
                sType = GetVal(dr.Item("SourceType"))
                ID = GetVal(dr.Item("MappingSourceId"))
                Text = GetVal(dr.Item("MappingSource"))
                ParentName = GetVal(dr.Item("SourceParent"))
            Else
                sType = GetVal(dr.Item("TargetType"))
                ID = GetVal(dr.Item("MappingTargetId"))
                Text = GetVal(dr.Item("MappingTarget"))
                ParentName = GetVal(dr.Item("TargetParent"))
            End If

            Select Case sType
                Case modDeclares.enumMappingType.MAPPING_TYPE_NONE
                    If ReturnDummyObj = False Then
                        retObj = New clsField
                        CType(retObj, clsField).FieldName = Text
                    End If
                Case modDeclares.enumMappingType.MAPPING_TYPE_FIELD
                    If ReturnDummyObj = False Then
                        If ID <> Text Then
                            'If Me.Engine IsNot Nothing Then
                            retObj = SearchFieldByName(Me.Engine, ID, ParentName, Me.Environment.Text, Me.Project.Text, Direction)
                            'Else
                            '    retObj = SearchFieldByName(Me.Environment, ID, ParentName, Me.Project.Text)
                            'End If

                            If retObj Is Nothing Then
                                Return Nothing
                                Exit Function
                            End If
                            CType(retObj, clsField).CorrectedFieldName = Text
                        Else
                            'If Me.Engine IsNot Nothing Then
                            retObj = SearchFieldByName(Me.Engine, Text, ParentName, Me.Environment.Text, Me.Project.Text, Direction)
                            '    Else
                            '    retObj = SearchFieldByName(Me.Environment, ID, ParentName, Me.Project.Text)
                            'End If

                            If retObj Is Nothing Then
                                Return Nothing
                                Exit Function
                            End If
                        End If

                    If ID <> Text Then
                        CType(retObj, clsField).CorrectedFieldName = Text
                    End If
                    Else
                    retObj = New clsField '//on 7/23/05

                    CType(retObj, clsField).FieldName = ID '//field name '//by npatel on 9/7/05
                    CType(retObj, clsField).ParentStructureName = ParentName '//structure name
                    CType(retObj, clsField).Parent = Me
                    If ID <> Text Then
                        CType(retObj, clsField).CorrectedFieldName = Text
                    End If
                    End If
                Case modDeclares.enumMappingType.MAPPING_TYPE_FUN
                    retObj = New clsSQFunction
                    '//some fake id so we wont get an weird error bcoz of no id
                    CType(retObj, clsSQFunction).SQFunctionId = Text
                    CType(retObj, clsSQFunction).SQFunctionName = Text 'ID
                    CType(retObj, clsSQFunction).SQFunctionWithInnerText = Text
                    CType(retObj, clsSQFunction).Parent = Me
                Case modDeclares.enumMappingType.MAPPING_TYPE_JOIN
                    If ReturnDummyObj = False Then
                        If Me.Engine IsNot Nothing Then
                            retObj = SearchTaskByName(Text, Me.Engine.Text, Me.Engine.ObjSystem.Text, _
                            Me.Environment.Text, Me.Project.Text, sType, Me.Engine)
                        Else
                            retObj = SearchTaskByName(Text, "", "", Me.Environment.Text, Me.Project.Text, sType, , Me.Environment)
                        End If

                    Else
                        retObj = New clsTask '//on 7/23/05
                        CType(retObj, clsTask).TaskName = Text
                        CType(retObj, clsTask).TaskType = modDeclares.enumTaskType.TASK_GEN
                        CType(retObj, clsTask).Parent = Me.Parent
                    End If
                Case modDeclares.enumMappingType.MAPPING_TYPE_LOOKUP
                    If ReturnDummyObj = False Then
                        If Me.Engine IsNot Nothing Then
                            retObj = SearchTaskByName(Text, Me.Engine.Text, Me.Engine.ObjSystem.Text, _
                            Me.Environment.Text, Me.Project.Text, sType, Me.Engine)
                        Else
                            retObj = SearchTaskByName(Text, "", "", Me.Environment.Text, Me.Project.Text, sType, , Me.Environment)
                        End If
                    Else
                        retObj = New clsTask
                        CType(retObj, clsTask).TaskName = Text
                        CType(retObj, clsTask).TaskType = modDeclares.enumTaskType.TASK_LOOKUP
                        CType(retObj, clsTask).Parent = Me.Parent
                    End If

                Case modDeclares.enumMappingType.MAPPING_TYPE_PROC
                    If ReturnDummyObj = False Then
                        If Me.Engine IsNot Nothing Then
                            retObj = SearchTaskByName(Text, Me.Engine.Text, Me.Engine.ObjSystem.Text, _
                             Me.Environment.Text, Me.Project.Text, sType, Me.Engine)
                        Else
                            retObj = SearchTaskByName(Text, "", "", Me.Environment.Text, Me.Project.Text, sType, , Me.Environment)
                        End If
                    Else
                        retObj = New clsTask
                        CType(retObj, clsTask).TaskName = Text
                        CType(retObj, clsTask).TaskType = modDeclares.enumTaskType.TASK_PROC
                        CType(retObj, clsTask).Parent = Me.Parent
                    End If
                Case modDeclares.enumMappingType.MAPPING_TYPE_VAR
                    If ReturnDummyObj = False Then
                        If Me.Engine IsNot Nothing Then
                            retObj = SearchVariableByName(Me.Engine, Text, Me.Engine.Text, Me.Engine.ObjSystem.Text, _
                            Me.Environment.Text, Me.Project.Text)
                        Else
                            retObj = SearchVariableByName(Me.Environment, Text, Me.Environment.Text, Me.Project.Text)
                        End If
                    Else
                        retObj = New clsVariable
                        CType(retObj, clsVariable).VariableName = Text
                        CType(retObj, clsVariable).Parent = Me.Parent
                    End If
                Case modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR
                    retObj = New clsVariable

                    CType(retObj, clsVariable).VariableName = Text
                    CType(retObj, clsVariable).Parent = Me.Parent
                Case modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT
                    retObj = New clsVariable(sInput, modDeclares.enumVariableType.VTYPE_CONST)

                    CType(retObj, clsVariable).VariableName = Text
                    CType(retObj, clsVariable).Parent = Me.Parent
                Case Else
                    retObj = Nothing
            End Select

            Return retObj

        Catch ex As Exception
            LogError(ex, "clsTask BuildMappingSourceTarget")
            Return Nothing
        End Try

    End Function

    Private Function UpdateTaskDataStores(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        '////////////////////////////////////////////////////////////////////////////
        '//Insert in to TaskDatastore table
        '////////////////////////////////////////////////////////////////////////////
        Dim sql As String = ""
        Dim i As Integer

        Try
            '//any datastores to add?
            If Me.ObjSources.Count > 0 Or Me.ObjTargets.Count > 0 Then
                '//Delete all previous entries
                'If Me.Engine IsNot Nothing Then
                sql = "DELETE FROM " & Me.Project.tblTaskDS & _
                                " WHERE TaskName=" & Me.GetQuotedText & _
                                " AND EngineName=" & Me.Engine.GetQuotedText & _
                                " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                                " AND ProjectName=" & Me.Project.GetQuotedText
                'TaskType=" & Me.TaskType & _
                '" AND
                'Else
                '    sql = "DELETE FROM " & Me.Project.tblTaskDS & _
                '                    " WHERE TaskName=" & Me.GetQuotedText & _
                '                    " AND EngineName=" & Quote(DBNULL) & _
                '                    " AND SystemName=" & Quote(DBNULL) & _
                '                    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                '                    " AND ProjectName=" & Me.Project.GetQuotedText
                '    ' TaskType=" & Me.TaskType & _
                '    ' " AND
                'End If

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()

                '//Now Loop through all sources and add
                For i = 0 To Me.ObjSources.Count - 1
                    Me.ObjSources(i).DsDirection = "S"
                    'If Me.Engine IsNot Nothing Then
                    sql = " INSERT INTO " & Me.Project.tblTaskDS & _
                    "(TaskType, TaskName, EngineName,SystemName,EnvironmentName,ProjectName, DatastoreName,DsDirection) Values( " & _
                                        Me.TaskType & "," & _
                                        Me.GetQuotedText & "," & _
                                        Me.Engine.GetQuotedText & "," & _
                                        Me.Engine.ObjSystem.GetQuotedText & "," & _
                                        Me.Environment.GetQuotedText & "," & _
                                        Me.Project.GetQuotedText & "," & _
                                        Me.ObjSources(i).GetQuotedText & "," & _
                                        Quote(Me.ObjSources(i).DsDirection) & ");"
                    'Else
                    'sql = " INSERT INTO " & Me.Project.tblTaskDS & _
                    '"(TaskType, TaskName, EngineName,SystemName,EnvironmentName,ProjectName, DatastoreName,DsDirection) Values( " & _
                    '                    Me.TaskType & "," & _
                    '                    Me.GetQuotedText & "," & _
                    '                    Quote(DBNULL) & "," & _
                    '                    Quote(DBNULL) & "," & _
                    '                    Me.Environment.GetQuotedText & "," & _
                    '                    Me.Project.GetQuotedText & "," & _
                    '                    Me.ObjSources(i).GetQuotedText & "," & _
                    '                    Quote(Me.ObjSources(i).DsDirection) & ");"
                    'End If


                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next

                '//Now Loop through all targets and add
                '//Fix: 8/12/05 by npatel
                For i = 0 To Me.ObjTargets.Count - 1
                    Me.ObjTargets(i).DsDirection() = "T"
                    'If Me.Engine IsNot Nothing Then
                    sql = " INSERT INTO " & Me.Project.tblTaskDS & _
                    "(TaskType, TaskName, EngineName,SystemName,EnvironmentName,ProjectName, DatastoreName,DsDirection ) Values( " & _
                                        Me.TaskType & "," & _
                                        Me.GetQuotedText & "," & _
                                        Me.Engine.GetQuotedText & "," & _
                                        Me.Engine.ObjSystem.GetQuotedText & "," & _
                                        Me.Environment.GetQuotedText & "," & _
                                        Me.Project.GetQuotedText & "," & _
                                        Me.ObjTargets(i).GetQuotedText & "," & _
                                        Quote(Me.ObjTargets(i).DsDirection) & ");"
                    'Else
                    'sql = " INSERT INTO " & Me.Project.tblTaskDS & _
                    '"(TaskType, TaskName, EngineName,SystemName,EnvironmentName,ProjectName, DatastoreName,DsDirection ) Values( " & _
                    '                    Me.TaskType & "," & _
                    '                    Me.GetQuotedText & "," & _
                    '                    Quote(DBNULL) & "," & _
                    '                    Quote(DBNULL) & "," & _
                    '                    Me.Environment.GetQuotedText & "," & _
                    '                    Me.Project.GetQuotedText & "," & _
                    '                    Me.ObjTargets(i).GetQuotedText & "," & _
                    '                    Quote(Me.ObjTargets(i).DsDirection) & ");"
                    'End If

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next
            End If

            UpdateTaskDataStores = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask UpdateTaskDataStores", sql)
            UpdateTaskDataStores = False
        Catch ex As Exception
            LogError(ex, "clsTask UpdateTaskDatastores", sql)
            UpdateTaskDataStores = False
        End Try

    End Function

    Function UpdateTaskDesc() As Boolean

        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim sql As String = ""

        Try
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            '//We need to put in transaction because we will add Datastore,
            '//DSSELECTIONS and fields in 
            '//three steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            'If Me.Engine IsNot Nothing Then
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set DESCRIPTION='" & Me.TaskDescription & _
            '    "' where TaskType=" & Me.TaskType & _
            '    " AND TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Me.Engine.GetQuotedText & _
            '    " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            sql = "Update " & Me.Project.tblTasks & _
            " set TASKDESCRIPTION='" & Me.TaskDescription & _
            "' where TaskType=" & Me.TaskType & _
            " AND TaskName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText
            'End If
            'Else
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set DESCRIPTION='" & Me.TaskDescription & _
            '    "' where TaskType=" & Me.TaskType & _
            '    " AND TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            '    sql = "Update " & Me.Project.tblTasks & _
            '    " set TASKDESCRIPTION='" & Me.TaskDescription & _
            '    "' where TaskType=" & Me.TaskType & _
            '    " AND TaskName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText
            'End If
            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()
            tran.Commit()

            UpdateTaskDesc = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask UpdateTaskDesc", sql)
            tran.Rollback()
            UpdateTaskDesc = False
        Catch ex As Exception
            LogError(ex, "clsTask UpdateTaskDesc")
            tran.Rollback()
            UpdateTaskDesc = False
        End Try

    End Function

    '// no longer private function  by TK and KS 11/6/2006
    Function AddNewTask(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim strSql As String = ""

        Try
            'If Me.Engine IsNot Nothing Then
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    strSql = "INSERT INTO " & Me.Project.tblTasks & "(" & _
            '                "ProjectName" & _
            '                ",EnvironmentName" & _
            '                ",SystemName" & _
            '                ",EngineName" & _
            '                ",TaskName" & _
            '                ",TaskType" & _
            '                ",Description" & _
            '                ",SeqNo" & _
            '                ") " & _
            '             " Values(" & _
            '                Me.Project.GetQuotedText & _
            '                "," & Me.Environment.GetQuotedText & _
            '                "," & Me.Engine.ObjSystem.GetQuotedText & _
            '                "," & Me.Engine.GetQuotedText & _
            '                "," & Me.GetQuotedText & _
            '                "," & Me.TaskType & _
            '                ",'" & FixStr(Me.TaskDescription) & "'" & _
            '                "," & Me.SeqNo & _
            '                ") "
            'Else
            strSql = "INSERT INTO " & Me.Project.tblTasks & "(" & _
                        "ProjectName" & _
                        ",EnvironmentName" & _
                        ",SystemName" & _
                        ",EngineName" & _
                        ",TaskName" & _
                        ",TaskType" & _
                        ",TASKDescription" & _
                        ",TASKSeqNo" & _
                        ") " & _
                     " Values(" & _
                        Me.Project.GetQuotedText & _
                        "," & Me.Environment.GetQuotedText & _
                        "," & Me.Engine.ObjSystem.GetQuotedText & _
                        "," & Me.Engine.GetQuotedText & _
                        "," & Me.GetQuotedText & _
                        "," & Me.TaskType & _
                        ",'" & FixStr(Me.TaskDescription) & "'" & _
                        "," & Me.SeqNo & _
                        ") "
            'End If
            'Else
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    strSql = "INSERT INTO " & Me.Project.tblTasks & "(" & _
            '                "ProjectName" & _
            '                ",EnvironmentName" & _
            '                ",SystemName" & _
            '                ",EngineName" & _
            '                ",TaskName" & _
            '                ",TaskType" & _
            '                ",Description" & _
            '                ",SeqNo" & _
            '                ") " & _
            '             " Values(" & _
            '                Me.Project.GetQuotedText & _
            '                "," & Me.Environment.GetQuotedText & _
            '                "," & Quote(DBNULL) & _
            '                "," & Quote(DBNULL) & _
            '                "," & Me.GetQuotedText & _
            '                "," & Me.TaskType & _
            '                ",'" & FixStr(Me.TaskDescription) & "'" & _
            '                "," & Me.SeqNo & _
            '                ") "
            'Else
            '    strSql = "INSERT INTO " & Me.Project.tblTasks & "(" & _
            '                "ProjectName" & _
            '                ",EnvironmentName" & _
            '                ",SystemName" & _
            '                ",EngineName" & _
            '                ",TaskName" & _
            '                ",TaskType" & _
            '                ",TASKDescription" & _
            '                ",TASKSeqNo" & _
            '                ") " & _
            '             " Values(" & _
            '                Me.Project.GetQuotedText & _
            '                "," & Me.Environment.GetQuotedText & _
            '                "," & Quote(DBNULL) & _
            '                "," & Quote(DBNULL) & _
            '                "," & Me.GetQuotedText & _
            '                "," & Me.TaskType & _
            '                ",'" & FixStr(Me.TaskDescription) & "'" & _
            '                "," & Me.SeqNo & _
            '                ") "
            'End If
            'End If



            Log(strSql)
            cmd.CommandText = strSql
            cmd.ExecuteNonQuery()

            AddNewTask = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask AddNewTask", strSql)
            AddNewTask = False
        Catch ex As Exception
            LogError(ex, "clsTask AddNewTask", strSql)
            AddNewTask = False
        End Try

    End Function

    Function AddNewMappings(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        '//Delete unselected selections
        Dim strSql As String = ""
        Dim i As Integer
        Dim IsAdded As Boolean
        Dim selNew As clsMapping
        'Dim dummy As New clsMapping

        Try
            'dummy.Parent = Me

            '//compare each new selection with old one
            For i = 0 To Me.ObjMappings.Count - 1
                selNew = ObjMappings(i)
                '//check new selection with old selection if not found means its added so we should add to database
                IsAdded = True '//first make flag true and if found in new then make it false

                '//Add mapping to taskmapping table
                'If Me.Engine IsNot Nothing Then
                strSql = "INSERT INTO " & Me.Project.tblTaskMap & " (" & _
                "ProjectName," & _
                "EnvironmentName," & _
                "SystemName," & _
                "EngineName," & _
                "TaskName," & _
                "MappingID," & _
                "TaskType," & _
                "SourceType, " & _
                "TargetType, " & _
                "MappingSource, " & _
                "MappingTarget, " & _
                "MappingSourceId, " & _
                "MappingTargetId, " & _
                "SourceParent, " & _
                "TargetParent, " & _
                "SourceDataStore, " & _
                "TargetDataStore, " & _
                "SeqNo, " & _
                "IsMapped, " & _
                "MappingDesc)" & _
                " Values( " & _
                Me.Project.GetQuotedText & "," & _
                Me.Environment.GetQuotedText & "," & _
                Me.Engine.ObjSystem.GetQuotedText & "," & _
                Me.Engine.GetQuotedText & "," & _
                Me.GetQuotedText & "," & _
                i & "," & _
                Me.TaskType & "," & _
                selNew.SourceType & "," & _
                selNew.TargetType & "," & _
                Quote(FixStr(selNew.GetMappingSourceVal), "'") & "," & _
                Quote(FixStr(selNew.GetMappingTargetVal), "'") & "," & _
                selNew.GetMappingSourceId & "," & _
                selNew.GetMappingTargetId & "," & _
                Quote(FixStr(selNew.SourceParent), "'") & "," & _
                Quote(FixStr(selNew.TargetParent), "'") & "," & _
                Quote(FixStr(selNew.SourceDataStore), "'") & "," & _
                Quote(FixStr(selNew.TargetDataStore), "'") & "," & _
                selNew.OutMsg & "," & _
                Quote(selNew.IsMapped, "'") & "," & _
                Quote(selNew.MappingDesc, "'") & _
                ");"
                'Else
                'strSql = "INSERT INTO " & Me.Project.tblTaskMap & " (" & _
                '"ProjectName," & _
                '"EnvironmentName," & _
                '"SystemName," & _
                '"EngineName," & _
                '"TaskName," & _
                '"MappingID," & _
                '"TaskType," & _
                '"SourceType, " & _
                '"TargetType, " & _
                '"MappingSource, " & _
                '"MappingTarget, " & _
                '"MappingSourceId, " & _
                '"MappingTargetId, " & _
                '"SourceParent, " & _
                '"TargetParent, " & _
                '"SourceDataStore, " & _
                '"TargetDataStore, " & _
                '"SeqNo, " & _
                '"IsMapped, " & _
                '"MappingDesc)" & _
                '" Values( " & _
                'Me.Project.GetQuotedText & "," & _
                'Me.Environment.GetQuotedText & "," & _
                'Quote(DBNULL) & "," & _
                'Quote(DBNULL) & "," & _
                'Me.GetQuotedText & "," & _
                'i & "," & _
                'Me.TaskType & "," & _
                'selNew.SourceType & "," & _
                'selNew.TargetType & "," & _
                'Quote(FixStr(selNew.GetMappingSourceVal), "'") & "," & _
                'Quote(FixStr(selNew.GetMappingTargetVal), "'") & "," & _
                'selNew.GetMappingSourceId & "," & _
                'selNew.GetMappingTargetId & "," & _
                'Quote(FixStr(selNew.SourceParent), "'") & "," & _
                'Quote(FixStr(selNew.TargetParent), "'") & "," & _
                'Quote(FixStr(selNew.SourceDataStore), "'") & "," & _
                'Quote(FixStr(selNew.TargetDataStore), "'") & "," & _
                'selNew.OutMsg & "," & _
                'Quote(selNew.IsMapped, "'") & "," & _
                'Quote(selNew.MappingDesc, "'") & _
                '");"
                'End If


                cmd.CommandText = strSql
                Log(strSql)
                cmd.ExecuteNonQuery()

                '//now save New Mapping in object 
                selNew.IsModified = False
                selNew.UpdateMappingTargetValue()
            Next

            AddNewMappings = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask AddNewMappings", strSql)
            AddNewMappings = False
        Catch ex As Exception
            LogError(ex, "clsTask AddNewMappings", strSql)
            AddNewMappings = False
        End Try

    End Function

    Function DeleteMappings(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim strSql As String = ""

        Try
            'If Me.Engine IsNot Nothing Then
            strSql = "delete from " & Me.Project.tblTaskMap & " " & _
                            " WHERE TaskName=" & Me.GetQuotedText & _
                            " AND EngineName=" & Me.Engine.GetQuotedText & _
                            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                            " AND ProjectName=" & Me.Project.GetQuotedText
            'TaskType=" & Me.TaskType & _   " AND 
            'Else
            'strSql = "delete from " & Me.Project.tblTaskMap & " " & _
            '                " WHERE TaskName=" & Me.GetQuotedText & _
            '                " AND EngineName=" & Quote(DBNULL) & _
            '                " AND SystemName=" & Quote(DBNULL) & _
            '                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '                " AND ProjectName=" & Me.Project.GetQuotedText
            ''TaskType=" & Me.TaskType & _ " AND 
            'End If

            cmd.CommandText = strSql
            Log(strSql)
            cmd.ExecuteNonQuery()

            DeleteMappings = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask DeleteMappings", strSql)
            DeleteMappings = False
        Catch ex As Exception
            LogError(ex, "clsTask DeleteMappings", strSql)
            DeleteMappings = False
        End Try

    End Function

    '// New Function added by KS and TK 11/6/2006 FOR AUTOMAP ONLY
    Private Function GetMainText() As String

        Dim sb As New System.Text.StringBuilder
        Dim sel As clsDSSelection
        'Dim i As Integer
        Dim taskName As String
        Dim sDs As clsDatastore

        Try
            sDs = Me.ObjSources(0)
            'sb.Append("CASE" & vbCrLf)
            'sb.Append("(" & vbCrLf)

            'For i = 0 To sDs.ObjSelections.Count - 1
            '    sel = sDs.ObjSelections(i)
            '    taskName = "P_" & sel.Text
            '    sb.Append(vbTab & IIf(i = 0, "", ",") & "EQ(RECNAME(" & sDs.Text & "),'" & sel.Text & "')" & vbCrLf)
            '    sb.Append(vbTab & ",CALLPROC(" & taskName & ")" & vbCrLf)
            '    sb.Append(vbCrLf)
            'Next

            'sb.Append(")" & vbCrLf)
            'sb.Append("{" & vbCrLf)
            'sb.Append("CASE" & vbCrLf)
            'For Each sel In sDs.ObjSelections
            '    taskName = "P_" & sel.Text
            '    sb.Append(TAB & "WHEN(RECNAME(" & sDs.Text & ") = '" & sel.Text & "')" & vbCrLf) '& IIf(nd.Index = 0, "", ",")
            '    sb.Append(TAB & "DO" & vbCrLf)
            '    sb.Append(TAB & TAB & "CALLPROC(" & taskName & ")" & vbCrLf)
            '    sb.Append(TAB & "END" & vbCrLf)
            '    sb.Append(vbCrLf)
            'Next
            ''sb.Append(")" & vbCrLf)
            'sb.AppendLine("}")   '& inDsNode.Text




            sb.AppendLine("CASE RECNAME(" & sDs.Text & ")")
            For Each sel In sDs.ObjSelections
                'count += 1
                'If count <= objThis.Procs.Count Then
                '    taskName = CType(objThis.Procs(count), clsTask).TaskName
                'Else
                taskName = ""
                'End If
                For Each proc As clsTask In Me.Engine.Procs
                    If proc.TaskName.Contains(sel.Text) = True Then
                        taskName = proc.TaskName
                        Exit For
                    End If
                Next

                sb.AppendLine(TAB & "WHEN '" & sel.Text & "'") '& IIf(nd.Index = 0, "", ",")
                sb.AppendLine(TAB & "DO")
                sb.AppendLine(TAB & TAB & "CALLPROC(" & taskName & ")")
                sb.AppendLine(TAB & "END") '& vbCrLf)
                'sb.Append(vbCrLf)
            Next
            GetMainText = sb.ToString

        Catch ex As Exception
            LogError(ex, "clsTask GetMainText")
            GetMainText = ""
        End Try

    End Function

    '//Write Global Datastore Properties to File
    'Function SaveLastFlds() As Boolean

    '    Try
    '        If System.IO.File.Exists(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.TaskName & ".txt") = True Then
    '            System.IO.File.Delete(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.TaskName & ".txt")
    '        End If
    '        fsLastFlds = System.IO.File.Create(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.TaskName & ".txt")
    '        objWriteLastFlds = New System.IO.StreamWriter(fsLastFlds)

    '        'Write Last Fields to file
    '        objWriteLastFlds.WriteLine(Me.LastSrcFld)
    '        objWriteLastFlds.WriteLine(Me.LastTgtFld)

    '        objWriteLastFlds.Close()
    '        fsLastFlds.Close()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDatastore SaveGlobals")
    '        Return False
    '    End Try

    'End Function

    'Function LoadLastFlds() As Boolean

    '    Try
    '        If System.IO.File.Exists(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.TaskName & ".txt") = False Then
    '            LoadLastFlds = True
    '            Exit Function
    '        End If
    '        fsLastFlds = System.IO.File.Open(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.TaskName & ".txt", IO.FileMode.Open)

    '        objReadLastFlds = New System.IO.StreamReader(fsLastFlds)

    '        'Write Globals to File
    '        Me.LastSrcFld = objReadLastFlds.ReadLine()
    '        Me.LastTgtFld = objReadLastFlds.ReadLine()

    '        objReadLastFlds.Close()
    '        fsLastFlds.Close()


    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDatastore LoadGlobals")
    '        Return False
    '    End Try

    'End Function

    '/// Added for Addflow
    Function UpdateTaskLocation() As Boolean

        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim sql As String = ""

        Try
            cmd.Connection = cnn
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            sql = "Update " & Me.Project.tblTasks & _
            " set CREATED_USER_ID='" & Me.Hloc.ToString & "', UPDATED_USER_ID='" & Me.Vloc.ToString & "'" & _
            " where TaskType=" & Me.TaskType & _
            " AND TaskName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()
            tran.Commit()

            UpdateTaskLocation = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsTask UpdateTaskLocation", sql)
            tran.Rollback()
            UpdateTaskLocation = False
        Catch ex As Exception
            LogError(ex, "clsTask UpdateTaskLocation")
            tran.Rollback()
            UpdateTaskLocation = False
        End Try

    End Function

    Function IsMapped() As Boolean

        Try
            Dim fldcnt As Integer = 0
            Dim mapcnt As Integer = 0

            '/// see if this proc has ANY targets
            If ObjTargets.Count = 0 Then
                IsMapped = False
                Exit Try
            End If

            '/// get a count of all the target fields
            For Each tgt As clsDatastore In ObjTargets
                For Each sel As clsDSSelection In tgt.ObjSelections
                    fldcnt = fldcnt + sel.DSSelectionFields.Count
                Next
            Next
            '/// get a count of the mappings that have a target AND a source field
            For Each map As clsMapping In Me.ObjMappings
                If map.IsMapped = True Then
                    mapcnt += 1
                End If
            Next

            If fldcnt = mapcnt Then
                IsMapped = True
            Else
                IsMapped = False
            End If

        Catch ex As Exception
            LogError(ex, "clsTask IsMapped")
            IsMapped = False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
