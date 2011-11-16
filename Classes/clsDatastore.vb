Option Compare Text
Public Class clsDatastore
    Implements INode

    '/// Object/Inode properties
    Private m_ObjEngine As clsEngine
    Private m_Environment As clsEnvironment
    Private m_ObjParent As INode
    Public OldObjSelections As New ArrayList '//Old DS selections, this is helpful when doing Save operation, we can compare old selection and new selection and perform insert/delete/update based on difference
    Public ObjSelections As New ArrayList '//Array of DS selections within Datastore
    Public ObjMissingMapList As New ArrayList '// array of DSSel that are mapped and no longer part of ME
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_IsLoaded As Boolean = False

    '/// Datastore Properties
    Private m_DatastoreName As String = ""
    Private m_DatastoreType As enumDatastore
    Private m_DsPhysicalSource As String = ""
    Private m_ExceptionDatastore As String = ""
    Private m_DsQueMgr As String = ""
    Private m_DsPort As String = ""
    Private m_DsDirection As String = "" '//Defined by constants DS_DIRECTION_XXXXXXXX
    Private m_DsAccessmethod As String = "" '//Defined by constants DS_ACCESSMETHOD_XXXXXXXX
    Private m_DatastoreDescription As String = ""

    '// Source Only
    Private m_DsCharacterCode As String = "" '//Defined by constants DS_CHARACTERCODE_XXXXXXXX
    'Global Field Properties
    Private g_ExtTypeChar As String = ""
    Private g_ExtTypeNum As String = ""
    Private g_IfNullChar As String = ""
    Private g_IfNullNum As String = ""
    Private g_InValidChar As String = ""
    Private g_InValidNum As String = ""

    '//Extended Properties
    Private m_IsIMSPathData As String = "0"
    Private m_IsSkipChangeCheck As String = "0"
    Private m_DsUOW As String = ""
    Private m_Restart As String = "" '// added by TK 11/9/2006
    Private m_Poll As String = ""  '// added by TK and KS 11/3/06
    'Private m_IsOrdered As String = "0"
    'Private m_IsCmmtKey As String = "0" '// added by TK and KS 11/3/06
    'Private m_IsBeforeImage As String = "0" '//new : 8/10/05 by npatel

    '// Target Only
    Private m_OperationType As String = ""
    Private m_IsKeyChng As String = "0"

    '/// Delimited Only
    Private m_TextQualifier As String = ""
    Private m_ColumnDelimiter As String = ""
    Private m_RowDelimiter As String = ""

    ' Lookup
    Private l_LookUp As Boolean = False

    '/// File writing for Global Datastore Field Values
    Dim fsDSglobal As System.IO.FileStream
    Dim objWriteGlobal As System.IO.StreamWriter
    Dim objReadGlobal As System.IO.StreamReader

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
            If m_ObjParent Is Nothing Then
                If m_ObjEngine Is Nothing Then
                    Return Nothing
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
                m_Environment = CType(Value, clsEngine).ObjSystem.Environment
            End If
            'If Value.Type = NODE_ENVIRONMENT Then
            '    m_Environment = Value
            'End If
            If Value.Type = NODE_PROC Or Value.Type = NODE_MAIN Or Value.Type = NODE_GEN Or Value.Type = NODE_LOOKUP Then
                m_ObjEngine = CType(Value, clsTask).Engine
                m_Environment = CType(Value, clsTask).Environment
            End If
            m_ObjParent = Value
        End Set
    End Property

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone
        '/// Fixed by TKarasch  May 07
        Try
            Dim obj As New clsDatastore
            Dim count As Integer = 0
            Dim cmd As Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadItems(, False, cmd)

            'obj.ObjTreeNode = Me.ObjTreeNode
            obj.DatastoreName = Me.DatastoreName
            obj.DatastoreDescription = Me.DatastoreDescription
            obj.DatastoreType = Me.DatastoreType
            obj.OperationType = Me.OperationType
            obj.ColumnDelimiter = Me.ColumnDelimiter
            obj.SeqNo = Me.SeqNo
            obj.RowDelimiter = Me.RowDelimiter
            obj.TextQualifier = Me.TextQualifier
            obj.DsCharacterCode = Me.DsCharacterCode
            obj.DsAccessMethod = Me.DsAccessMethod
            obj.DsDirection = Me.DsDirection
            obj.DsPhysicalSource = Me.DsPhysicalSource
            obj.Poll = Me.Poll '// added by TK and KS 11/3/06
            obj.DsPort = Me.DsPort
            obj.DsQueMgr = Me.DsQueMgr
            obj.ExceptionDatastore = Me.ExceptionDatastore
            'obj.IsCmmtKey = Me.IsCmmtKey '// added by TK and KS 11/3/06
            obj.IsIMSPathData = Me.IsIMSPathData
            'obj.IsOrdered = Me.IsOrdered
            obj.IsSkipChangeCheck = Me.IsSkipChangeCheck
            'obj.IsBeforeImage = Me.IsBeforeImage '// 8/10/05 by npatel
            obj.DsUOW = Me.DsUOW
            obj.IsLookUp = Me.IsLookUp
            'Field Props
            obj.ExtTypeChar = Me.ExtTypeChar
            obj.ExtTypeNum = Me.ExtTypeNum
            obj.IfNullChar = Me.IfNullChar
            obj.IfNullNum = Me.IfNullNum
            obj.InValidChar = Me.InValidChar
            obj.InValidNum = Me.InValidNum


            obj.IsModified = Me.IsModified
            obj.Parent = NewParent 'Me.Parent

            'obj.Engine = NewParent

            'Me.LoadItems()

            '/// completely redone to clone all child selections properly
            For Each DSSelObj As clsDSSelection In Me.ObjSelections
                DSSelObj.LoadMe(cmd)
                Dim dsselClone As New clsDSSelection
                'dsselClone = New clsDSSelection
                dsselClone = DSSelObj.Clone(obj, True, cmd)
                dsselClone.ObjDatastore = obj
                dsselClone.ObjSelection = FindStrSelforDSSel(dsselClone)
                dsselClone.ObjStructure = dsselClone.ObjSelection.ObjStructure
                obj.ObjSelections.Add(dsselClone)
            Next

            'May Not be necessary here, but won't Hurt anything
            obj.SetDSselParents()

            Return obj

        Catch ex As Exception
            LogError(ex, , , True)
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
                Key = Me.Engine.ObjSystem.Environment.Project.Text & KEY_SAP & Me.Engine.ObjSystem.Environment.Text & KEY_SAP & Me.Engine.ObjSystem.Text & KEY_SAP & Me.Engine.Text & KEY_SAP & Me.Text
            Else
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
            End If

        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = DatastoreName
        End Get
        Set(ByVal Value As String)
            DatastoreName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            If Me.m_DsDirection = DS_DIRECTION_TARGET Then
                Return NODE_TARGETDATASTORE
            Else
                Return NODE_SOURCEDATASTORE
            End If
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.Parent.Project
            Return p
        End Get
    End Property

    '///Done
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

            If UpdateDatastore(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            If DeleteUnselectedSelections(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            If AddNewSelections(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            If UpdateSelectionFields(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            tran.Commit()

            Me.IsModified = False

            Save = True

        Catch OE As Odbc.OdbcException
            tran.Rollback()
            LogODBCError(OE, " clsDatastore Inode Save() ")
            Save = False
        Catch ex As Exception
            tran.Rollback()
            LogError(ex, " clsDatastore Inode Save() ")
            Save = False
        Finally
            'cnn.Close()
        End Try

    End Function

    '///Done
    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd = New Odbc.OdbcCommand
            cmd.Connection = cnn

            '//We need to put in transaction because we will add Datastore,
            '//DSSELECTIONS and fields in 
            '//three steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '/////////////////////////////////////////////////////////////////////////
            '//1) Insert in to DATASTORES table
            '/////////////////////////////////////////////////////////////////////////

            If AddNewDatastore(cmd) = False Then
                tran.Rollback()
                AddNew = False
                Exit Try
            End If

            '/////////////////////////////////////////////////////////////////////////
            '//2) 3) Insert in to DSSELECTIONS and DSSELFIELDS tables
            '/////////////////////////////////////////////////////////////////////////

            If ObjSelections.Count > 0 Then
                If AddNewSelections(cmd) = False Then
                    tran.Rollback()
                    AddNew = False
                    Exit Try
                End If
            End If

            tran.Commit()


            'If Me.Engine IsNot Nothing Then
            If Me.DsDirection = DS_DIRECTION_TARGET Then
                '//Add  datastore in engine's targets collection
                AddToCollection(Me.Engine.Targets, Me, Me.GUID)
            ElseIf Me.DsDirection = DS_DIRECTION_SOURCE Then
                '//Add  datastore in engine's source collection
                AddToCollection(Me.Engine.Sources, Me, Me.GUID)
            End If
            'Else
            ''//Add  datastore in environment's datastore collection
            'AddToCollection(Me.Environment.Datastores, Me, Me.GUID)
            'End If

            Me.IsModified = False

            AddNew = True

        Catch OE As Odbc.OdbcException
            tran.Rollback()
            LogODBCError(OE, "clsDatastore AddNew")
            AddNew = False
        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "clsDatastore AddNew", , True)
            AddNew = False
        Finally
            'cnn.Close()
        End Try

    End Function

    '///Done
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Try
            Me.Text = Me.Text.Trim

            If AddNewDatastore(cmd) = False Then
                AddNew = False
                Exit Try
            End If

            '/////////////////////////////////////////////////////////////////////////
            '//2) 3) Insert in to DSSELECTIONS DSSELFIELDS table
            '/////////////////////////////////////////////////////////////////////////

            If ObjSelections.Count > 0 Then
                If AddNewSelections(cmd) = False Then
                    AddNew = False
                    Exit Try
                End If
            End If

            If Me.DsDirection = DS_DIRECTION_TARGET Then
                '//Add  datastore in engine's targets collection
                AddToCollection(Me.Engine.Targets, Me, Me.GUID)
            ElseIf Me.DsDirection = DS_DIRECTION_SOURCE Then
                '//Add  datastore in engine's source collection
                AddToCollection(Me.Engine.Sources, Me, Me.GUID)
                'Else
                '    '//Add  datastore in environment's datastore collection
                '    AddToCollection(Me.Environment.Datastores, Me, Me.GUID)
            End If

            AddNew = True
            Me.IsModified = False

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore AddNew")
            AddNew = False
        Catch ex As Exception
            LogError(ex, "clsDatastore AddNew", , True)
            AddNew = False
        End Try

    End Function

    '///Done
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            'If Me.Engine IsNot Nothing Then
            '/// delete from Datastore Selection Fields Table
            sql = "Delete From " & Me.Project.tblDSselFields & _
            " where DatastoreName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '/// delete from Datastore Selections Table
            sql = "Delete From " & Me.Project.tblDSselections & _
            " where  DatastoreName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '/// Delete from Task Datastores Table
            sql = "Delete From " & Me.Project.tblTaskDS & _
            " where  DatastoreName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '/// Delete From Datastores Table
            sql = "Delete From " & Me.Project.tblDatastores & _
            " where  DatastoreName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            'Else '/// datastore is under ENV
            '    '/// delete from Datastore Selection Fields Table
            '    sql = "Delete From " & Me.Project.tblDSselFields & _
            '    " where DatastoreName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText

            '    Log(sql)
            '    cmd.CommandText = sql
            '    cmd.ExecuteNonQuery()

            '    '/// delete from Datastore Selections Table
            '    sql = "Delete From " & Me.Project.tblDSselections & _
            '    " where  DatastoreName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText

            '    Log(sql)
            '    cmd.CommandText = sql
            '    cmd.ExecuteNonQuery()

            '    '/// Delete from Task Datastores Table
            '    sql = "Delete From " & Me.Project.tblTaskDS & _
            '    " where  DatastoreName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText

            '    Log(sql)
            '    cmd.CommandText = sql
            '    cmd.ExecuteNonQuery()

            '    '/// Delete From Datastores Table
            '    sql = "Delete From " & Me.Project.tblDatastores & _
            '    " where  DatastoreName=" & Me.GetQuotedText & _
            '    " AND EngineName=" & Quote(DBNULL) & _
            '    " AND SystemName=" & Quote(DBNULL) & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '    " AND ProjectName=" & Me.Project.GetQuotedText

            '    Log(sql)
            '    cmd.CommandText = sql
            '    cmd.ExecuteNonQuery()
            'End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '/// Delete from DatastoreATTR Table
            If Me.DeleteATTR(cmd) = False Then
                Delete = False
            End If
            'End If

            '/// the Datastores are removed from the Database, 
            '/// now we remove the child objects from the engines
            '/// IF we are removing it at the top (under environment)
            '/// and remove them from the tree
            'If Me.Engine Is Nothing Then
            '    For Each sys As clsSystem In Me.Environment.Systems
            '        For Each eng As clsEngine In sys.Engines
            '            For Each src As clsDatastore In eng.Sources
            '                If src.DatastoreName = Me.DatastoreName Then
            '                    src.ObjTreeNode.Remove()
            '                    RemoveFromCollection(eng.Sources, src.GUID)
            '                End If
            '            Next
            '            For Each tgt As clsDatastore In eng.Targets
            '                If tgt.DatastoreName = Me.DatastoreName Then
            '                    tgt.ObjTreeNode.Remove()
            '                    RemoveFromCollection(eng.Targets, tgt.GUID)
            '                End If
            '            Next
            '        Next
            '    Next
            'End If



            '/// Last ... remove from the engine object
            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                'If Me.Engine Is Nothing Then
                '    RemoveFromCollection(Me.Environment.Datastores, Me.GUID)
                'Else
                'RemoveFromCollection(Me.Environment.Datastores, Me.GUID)
                If Me.DsDirection = DS_DIRECTION_SOURCE Then
                    RemoveFromCollection(Me.Engine.Sources, Me.GUID)
                ElseIf Me.DsDirection = DS_DIRECTION_TARGET Then
                    RemoveFromCollection(Me.Engine.Targets, Me.GUID)
                End If
            End If
            'End If

            Delete = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore Delete")
            Delete = False
        Catch ex As Exception
            LogError(ex, sql, "clsDatastore Delete", False)
            Delete = False
        End Try

    End Function

    '///Done
    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Try
            '//check if fields already loaded in memory?
            If Me.ObjSelections.Count > 0 And Reload = False Then
                Exit Function
            End If
            'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
            Dim cmd As System.Data.Odbc.OdbcCommand
            Dim dr As System.Data.DataRow
            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As System.Data.DataTable
            Dim sql As String = ""
            Dim i As Integer
            Dim DSsel As clsDSSelection

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            If Me.ObjSelections.Count > 0 Then
                DSsel = Me.ObjSelections(0)
                If DSsel.DSSelectionFields.Count < 1 Then
                    If LoadAllDSSelFlds(cmd) = True Then
                        Exit Function
                    End If

                End If
            End If

            sql = "Select ds.PROJECTNAME,ds.ENVIRONMENTNAME,ds.SELECTIONNAME,ds.DESCRIPTIONNAME ,ss.selectdescription,ss.ISSYSTEMSEL from " & Me.Project.tblDSselections & _
                                " ds inner join " & Me.Project.tblDescriptionSelect & _
                                " ss on ss.selectionname=ds.selectionname and ss.descriptionname=ds.descriptionname and ss.environmentname=ds.environmentname and ss.projectname=ds.projectname " & _
                                " where ds.DatastoreName= " & Me.GetQuotedText & _
                                " And ds.EngineName= " & Me.Engine.GetQuotedText & _
                                " And ds.SystemName= " & Me.Engine.ObjSystem.GetQuotedText & _
                                " And ds.EnvironmentName= " & Me.Environment.GetQuotedText & _
                                " And ds.ProjectName= " & Me.Project.GetQuotedText & _
                                " order by ds.selectionname"

            cmd.CommandText = sql
            Log(sql)
            If Incmd IsNot Nothing Then
                cmd.Transaction = Incmd.Transaction
            End If
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            dt = New System.Data.DataTable("temp2")
            da.Fill(dt)
            da.Dispose()

            Me.ObjSelections.Clear()

            For i = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)
                '/// read in all DSselections
                Dim objSel As New clsDSSelection
                Dim ObjStrSel As clsStructureSelection

                ObjStrSel = SearchStructSelByName(Me.Environment, dr("SelectionName"), dr("DescriptionName"), dr("EnvironmentName"), dr("ProjectName"))

                If ObjStrSel Is Nothing Then
                    clsLogging.LogEvent("Datastore selection [" & dr("SelectionName") & "] not found under this environment")
                Else
                    objSel = CloneSSeltoDSSel(ObjStrSel, , , cmd, TreeLode)
                    objSel.LoadMe(cmd)
                    LoadDSSelectionFields(objSel, cmd, TreeLode)
                    Me.ObjSelections.Add(objSel)
                End If
            Next

            '/// Now set DSSelection Parents based on FKey
            LoadItems = SetDSselParents()


        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore LoadItems")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            LoadItems = False
        Catch ex As Exception
            LogError(ex, "clsDatastore LoadItems")
            LoadItems = False
        End Try

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then Exit Try

            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As New DataTable("temp")
            Dim dr As DataRow
            Dim sql As String = ""
            Dim Attrib As String = ""
            Dim Value As String = ""
            Dim tran As Odbc.OdbcTransaction = Nothing

            If Incmd Is Nothing Then
                Incmd = New Odbc.OdbcCommand
                Incmd.Connection = cnn
            End If

            sql = "SELECT DATASTOREATTRB,DATASTOREATTRBVALUE FROM " & Me.Project.tblDatastoresATTR & _
                        " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                        " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                        " AND SYSTEMNAME=" & Quote(Me.Engine.ObjSystem.SystemName) & _
                        " AND ENGINENAME=" & Quote(Me.Engine.EngineName) & _
                        " AND DATASTORENAME=" & Quote(Me.DatastoreName)

            Incmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetVal(dr("DATASTOREATTRB"))
                Select Case Attrib
                    Case "COLUMNDELIMITER"
                        Me.ColumnDelimiter = GetVal(dr("DATASTOREATTRBVALUE"))
                    Case "DATASTORETYPE"
                        Me.DatastoreType = GetVal(dr("DATASTOREATTRBVALUE"))
                    Case "DSACCESSMETHOD"
                        Me.DsAccessMethod = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "DSCHARACTERCODE"
                        Me.DsCharacterCode = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "DSDIRECTION"
                        Me.DsDirection = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "DSPHYSICALSOURCE"
                        Me.DsPhysicalSource = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "EXCEPTDATASTORE"
                        Me.ExceptionDatastore = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "RESTART"
                        Me.Restart = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "ISIMSPATHDATA"
                        Me.IsIMSPathData = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "POLL"
                        Me.Poll = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "ISSKIPCHGCHECK"
                        Me.IsSkipChangeCheck = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                        'Case "ONCMMTKEY"
                        '    Me.IsCmmtKey = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "OPERATIONTYPE"
                        Me.OperationType = GetStr(GetStr(GetVal(dr("DATASTOREATTRBVALUE"))))
                    Case "PORT"
                        Me.DsPort = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "QUEMGR"
                        Me.DsQueMgr = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "ROWDELIMITER"
                        Me.RowDelimiter = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                        'Case "SELECTEVERY"
                        '    Me.Poll = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "TEXTQUALIFIER"
                        Me.TextQualifier = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "UOW"
                        Me.DsUOW = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "EXTTYPECHAR"
                        Me.ExtTypeChar = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "EXTTYPENUM"
                        Me.ExtTypeNum = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "IFNULLCHAR"
                        Me.IfNullChar = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "IFNULLNUM"
                        Me.IfNullNum = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "INVALIDCHAR"
                        Me.InValidChar = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "INVALIDNUM"
                        Me.InValidNum = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "LOOKUP"
                        Me.IsLookUp = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                    Case "ISKEYCHNG"
                        Me.IsKeyChng = GetStr(GetVal(dr("DATASTOREATTRBVALUE")))
                End Select
            Next

            Me.IsLoaded = True

            LoadMe = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore LoadMe")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            LoadMe = False
        Catch ex As Exception
            LogError(ex, "clsDatastore LoadMe")
            LoadMe = False
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

    Property IsLoaded() As Boolean Implements INode.IsLoaded
        Get
            Return m_IsLoaded
        End Get
        Set(ByVal value As Boolean)
            m_IsLoaded = value
        End Set
    End Property

    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        Dim cmd As System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""
        Dim readName As String
        Dim testName As String = NewName
        Dim NameValid As Boolean = True

        Try
            If testName = "" Then
                testName = Me.Text
            End If

            cmd = cnn.CreateCommand

            'If Me.Engine IsNot Nothing Then
            sql = "Select DATASTORENAME from " & Me.Project.tblDatastores & _
            " where PROJECTNAME=" & Me.Project.GetQuotedText & _
            " AND ENVIRONMENTNAME=" & Me.Environment.GetQuotedText & _
            " AND SYSTEMNAME=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND ENGINENAME=" & Me.Engine.GetQuotedText
            'Else
            'sql = "Select DATASTORENAME from " & Me.Project.tblDatastores & _
            '" where PROJECTNAME=" & Me.Project.GetQuotedText & _
            '" AND ENVIRONMENTNAME=" & Me.Environment.GetQuotedText & _
            '" AND SYSTEMNAME=" & Quote(DBNULL) & _
            '" AND ENGINENAME=" & Quote(DBNULL)
            'End If

            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = GetVal(dr.Item("DATASTORENAME"))
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The New Datastore name you have chosen already exists" & (Chr(13)) & _
                    "in this Environment, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            ValidateNewObject = NameValid

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore ValidateNewObject")
            ValidateNewObject = False
        Catch ex As Exception
            LogError(ex, "clsDatastore ValidateNewObject")
            ValidateNewObject = False
        End Try

    End Function

#End Region

#Region "Properties"

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
                Return m_Environment
            Else
                Return Me.Engine.ObjSystem.Environment
            End If
        End Get
        Set(ByVal value As clsEnvironment)
            m_Environment = value
        End Set
    End Property

    Public Property DatastoreName() As String
        Get
            Return m_DatastoreName
        End Get
        Set(ByVal Value As String)
            m_DatastoreName = Value
        End Set
    End Property

    Public Property DsUOW() As String
        Get
            Return m_DsUOW
        End Get
        Set(ByVal Value As String)
            m_DsUOW = Value
        End Set
    End Property

    Public Property DatastoreType() As enumDatastore
        Get
            Return m_DatastoreType
        End Get
        Set(ByVal Value As enumDatastore)
            m_DatastoreType = Value
        End Set
    End Property

    Public Property DsPhysicalSource() As String
        Get
            Return m_DsPhysicalSource
        End Get
        Set(ByVal Value As String)
            m_DsPhysicalSource = Value
        End Set
    End Property

    '// added by TK 11/9/2006
    Public Property Restart() As String
        Get
            Return m_Restart
        End Get
        Set(ByVal Value As String)
            m_Restart = Value
        End Set
    End Property

    '// added by KS and TK 11/3/06
    '//Changed to Poll ... 12/22/09
    Public Property Poll() As String
        Get
            Return m_Poll
        End Get
        Set(ByVal Value As String)
            m_Poll = Value
        End Set
    End Property

    Public Property DsQueMgr() As String
        Get
            Return m_DsQueMgr
        End Get
        Set(ByVal Value As String)
            m_DsQueMgr = Value
        End Set
    End Property

    Public Property DsPort() As String
        Get
            Return m_DsPort
        End Get
        Set(ByVal Value As String)
            m_DsPort = Value
        End Set
    End Property

    Public Property DsDirection() As String
        Get
            Return m_DsDirection
        End Get
        Set(ByVal Value As String)
            m_DsDirection = Value
        End Set
    End Property

    Public Property DsAccessMethod() As String
        Get
            Return m_DsAccessmethod
        End Get
        Set(ByVal Value As String)
            m_DsAccessmethod = Value
        End Set
    End Property

    Public Property DsCharacterCode() As String
        Get
            Return m_DsCharacterCode
        End Get
        Set(ByVal Value As String)
            m_DsCharacterCode = Value
        End Set
    End Property

    'Public Property IsOrdered() As String
    '    Get
    '        Return m_IsOrdered
    '    End Get
    '    Set(ByVal Value As String)
    '        m_IsOrdered = Value
    '    End Set
    'End Property

    '// added by TK and KS 11/3/06
    'Public Property IsCmmtKey() As String
    '    Get
    '        Return m_IsCmmtKey
    '    End Get
    '    Set(ByVal Value As String)
    '        m_IsCmmtKey = Value
    '    End Set
    'End Property

    Public Property IsIMSPathData() As String
        Get
            Return m_IsIMSPathData
        End Get
        Set(ByVal Value As String)
            m_IsIMSPathData = Value
        End Set
    End Property

    Public Property IsSkipChangeCheck() As String
        Get
            Return m_IsSkipChangeCheck
        End Get
        Set(ByVal Value As String)
            m_IsSkipChangeCheck = Value
        End Set
    End Property

    '//new 8/10/05 by npatel
    'Public Property IsBeforeImage() As String
    '    Get
    '        Return m_IsBeforeImage
    '    End Get
    '    Set(ByVal Value As String)
    '        m_IsBeforeImage = Value
    '    End Set
    'End Property

    Public Property ExceptionDatastore() As String
        Get
            Return m_ExceptionDatastore
        End Get
        Set(ByVal Value As String)
            m_ExceptionDatastore = Value
        End Set
    End Property

    Public Property TextQualifier() As String
        Get
            Return m_TextQualifier
        End Get
        Set(ByVal Value As String)
            m_TextQualifier = Value
        End Set
    End Property

    Public Property OperationType() As String
        Get
            Return m_OperationType
        End Get
        Set(ByVal Value As String)
            m_OperationType = Value
        End Set
    End Property

    Public Property ColumnDelimiter() As String
        Get
            Return m_ColumnDelimiter
        End Get
        Set(ByVal Value As String)
            m_ColumnDelimiter = Value
        End Set
    End Property

    Public Property RowDelimiter() As String
        Get
            Return m_RowDelimiter
        End Get
        Set(ByVal Value As String)
            m_RowDelimiter = Value
        End Set
    End Property

    Public Property DatastoreDescription() As String
        Get
            Return m_DatastoreDescription
        End Get
        Set(ByVal Value As String)
            m_DatastoreDescription = Value
        End Set
    End Property

    Public Property ExtTypeChar() As String
        Get
            Return g_ExtTypeChar
        End Get
        Set(ByVal value As String)
            g_ExtTypeChar = value
        End Set
    End Property

    Public Property ExtTypeNum() As String
        Get
            Return g_ExtTypeNum
        End Get
        Set(ByVal value As String)
            g_ExtTypeNum = value
        End Set
    End Property

    Public Property IfNullChar() As String
        Get
            Return g_IfNullChar
        End Get
        Set(ByVal value As String)
            g_IfNullChar = value
        End Set
    End Property

    Public Property IfNullNum() As String
        Get
            Return g_IfNullNum
        End Get
        Set(ByVal value As String)
            g_IfNullNum = value
        End Set
    End Property

    Public Property InValidChar() As String
        Get
            Return g_InValidChar
        End Get
        Set(ByVal value As String)
            g_InValidChar = value
        End Set
    End Property

    Public Property InValidNum() As String
        Get
            Return g_InValidNum
        End Get
        Set(ByVal value As String)
            g_InValidNum = value
        End Set
    End Property

    Public Property IsLookUp() As Boolean
        Get
            Return l_LookUp
        End Get
        Set(ByVal value As Boolean)
            l_LookUp = value
        End Set
    End Property

    Public Property IsKeyChng() As String
        Get
            Return m_IsKeyChng
        End Get
        Set(ByVal value As String)
            m_IsKeyChng = value
        End Set
    End Property

#End Region

#Region "Methods"

    '///DONE
    Private Function UpdateSelectionFields(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim i, j As Integer
        Dim strSql As String = ""
        Dim objSel As clsDSSelection
        Dim objFld As clsField
        Dim isKey As String

        Try
            For i = 0 To Me.ObjSelections.Count - 1
                objSel = Me.ObjSelections(i)

                '//For each fields in selection
                For j = 0 To objSel.DSSelectionFields.Count - 1
                    objFld = objSel.DSSelectionFields(j)
                    'objFld.IsModified = True   "// this is just FOR TESTING metadata ONLY

                    '//save field only if modified by user
                    If objFld.IsModified = True Then
                        isKey = FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
                        If (isKey = "YES") Or (isKey = "1") Or (isKey = "Yes") Then
                            isKey = "Yes"
                        Else
                            isKey = "No"
                        End If

                        'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        '    If Me.Engine IsNot Nothing Then
                        '        strSql = "Update " & Me.Project.tblDSselFields & " Set " & _
                        '        "IsKey=" & Quote(isKey) & _
                        '        ",DateFormat=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATEFORMAT))) & _
                        '        ",Label=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LABEL))) & _
                        '        ",InitVal=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INITVAL))) & _
                        '        ",Retype=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_RETYPE))) & _
                        '        ",Invalid=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INVALID))) & _
                        '        ",ExtType=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_EXTTYPE))) & _
                        '        ",IdentVal=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_IDENTVAL))) & _
                        '        ",ForeignKey=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY))) & _
                        '        ",Parent=" & Quote(objSel.Parent.Text) & _
                        '        ",SEQNO=" & objFld.SeqNo & _
                        '        " WHERE  DatastoreName=" & Me.GetQuotedText & _
                        '        " AND EngineName=" & Me.Engine.GetQuotedText & _
                        '        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                        '        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                        '        " AND ProjectName=" & Me.Project.GetQuotedText & _
                        '        " AND StructureName=" & objSel.ObjStructure.GetQuotedText & _
                        '        " AND SelectionName=" & objSel.GetQuotedText & _
                        '        " AND FieldName=" & objFld.GetQuotedText '// modified to add FKey 11/8/2006 by TK
                        '    Else
                        '        strSql = "Update " & Me.Project.tblDSselFields & " Set " & _
                        '        "IsKey=" & Quote(isKey) & _
                        '        ",DateFormat=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATEFORMAT))) & _
                        '        ",Label=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LABEL))) & _
                        '        ",InitVal=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INITVAL))) & _
                        '        ",Retype=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_RETYPE))) & _
                        '        ",Invalid=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INVALID))) & _
                        '        ",ExtType=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_EXTTYPE))) & _
                        '        ",IdentVal=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_IDENTVAL))) & _
                        '        ",ForeignKey=" & Quote(FixStr(objFld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY))) & _
                        '        ",Parent=" & Quote(objSel.Parent.Text) & _
                        '        ",SEQNO=" & objFld.SeqNo & _
                        '        " WHERE  DatastoreName=" & Me.GetQuotedText & _
                        '        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                        '        " AND ProjectName=" & Me.Project.GetQuotedText & _
                        '        " AND StructureName=" & objSel.ObjStructure.GetQuotedText & _
                        '        " AND SelectionName=" & objSel.GetQuotedText & _
                        '        " AND FieldName=" & objFld.GetQuotedText '// modified to add FKey 11/8/2006 by TK
                        '    End If

                        '    cmd.CommandText = strSql
                        '    Log(strSql)
                        '    cmd.ExecuteNonQuery()

                        'Else '/// V3 Metadata
                        'If Me.Engine IsNot Nothing Then
                        strSql = "Update " & Me.Project.tblDSselFields & " set " & _
                        "ParentName= " & Quote(objSel.Parent.Text) & _
                        ", SEQNO= " & objFld.SeqNo & _
                        ", DescFieldDescription= " & Quote(objFld.FieldDesc) & _
                        ", NChildren= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & _
                        ", NLevel= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & _
                        ", Ntimes= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & _
                        ", NOccNo= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & _
                        ", Datatype= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & _
                        ", NOffSet= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & _
                        ", NLength= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & _
                        ", NScale= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & _
                        ", CanNull= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & _
                        ", ISKEY= " & Quote(isKey) & _
                        ", OrgName= " & Quote(objFld.OrgName) & _
                        ", DateFormat= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & _
                        ", Label= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & _
                        ", InitVal= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & _
                        ", ReType= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & _
                        ", InValid= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & _
                        ", ExtType= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & _
                        ", IdentVal= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & _
                        ", ForeignKey= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                        " WHERE ProjectName=" & Me.Project.GetQuotedText & _
                        " AND EnvironmentName=" & Me.Engine.ObjSystem.Environment.GetQuotedText & _
                        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                        " AND EngineName=" & Me.Engine.GetQuotedText & _
                        " AND DatastoreName = " & Me.GetQuotedText & _
                        " AND DescriptionName=" & objSel.ObjStructure.GetQuotedText & _
                        " AND SelectionName=" & objSel.GetQuotedText & _
                        " AND FieldName=" & objFld.GetQuotedText '// modified to add FKey 11/8/2006 by TK

                        cmd.CommandText = strSql
                        Log(strSql)
                        cmd.ExecuteNonQuery()

                        'Else
                        '    strSql = "Update " & Me.Project.tblDSselFields & " set " & _
                        '    "ParentName= " & Quote(objSel.Parent.Text) & _
                        '    ", SEQNO= " & objFld.SeqNo & _
                        '    ", DescFieldDescription= " & Quote(objFld.FieldDesc) & _
                        '    ", NChildren= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & _
                        '    ", NLevel= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & _
                        '    ", Ntimes= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & _
                        '    ", NOccNo= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & _
                        '    ", Datatype= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & _
                        '    ", NOffSet= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & _
                        '    ", NLength= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & _
                        '    ", NScale= " & objFld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & _
                        '    ", CanNull= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & _
                        '    ", ISKEY= " & Quote(isKey) & _
                        '    ", OrgName= " & Quote(objFld.OrgName) & _
                        '    ", DateFormat= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & _
                        '    ", Label= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & _
                        '    ", InitVal= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & _
                        '    ", ReType= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & _
                        '    ", InValid= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & _
                        '    ", ExtType= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & _
                        '    ", IdentVal= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & _
                        '    ", ForeignKey= " & Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                        '    " WHERE ProjectName=" & Me.Project.GetQuotedText & _
                        '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                        '    " AND DatastoreName = " & Me.GetQuotedText & _
                        '    " AND DescriptionName=" & objSel.ObjStructure.GetQuotedText & _
                        '    " AND SelectionName=" & objSel.GetQuotedText & _
                        '    " AND FieldName=" & objFld.GetQuotedText '// modified to add FKey 11/8/2006 by TK

                        '    cmd.CommandText = strSql
                        '    Log(strSql)
                        '    cmd.ExecuteNonQuery()

                        'End If
                        'End If

                        '//After we save reset the IsModified flag for field
                        objFld.IsModified = False
                    End If
                Next

                '//After we save entire selection reset the IsModified flag for selection
                objSel.IsModified = False
            Next

            UpdateSelectionFields = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore UpdateSelectionFields")
            UpdateSelectionFields = False
        Catch ex As Exception
            LogError(ex, strSql, "clsDatastore UpdateSelectionFields", False)
            UpdateSelectionFields = False
        End Try

    End Function

    '///DONE
    Private Function AddNewDatastore(ByVal cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim strSql As String = ""

        '//changed: IsBeforeImage by npatel on 8/10/05
        '// modified by TK 11/9/2006
        '// changed by TK and KS 11/3/06
        Try
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    If Me.Engine IsNot Nothing Then
            '        strSql = "INSERT INTO " & Me.Project.tblDatastores & _
            '        "(ProjectName" & _
            '        ",EnvironmentName" & _
            '        ",SystemName" & _
            '        ",EngineName" & _
            '        ",DatastoreName" & _
            '        ",DatastoreType" & _
            '        ",DsPhysicalSource" & _
            '        ",DsDirection" & _
            '        ",DsAccessMethod" & _
            '        ",DsCharacterCode" & _
            '        ",IsOrdered" & _
            '        ",IsIMSPathData" & _
            '        ",ISSKIPCHGCHECK" & _
            '        ",IsBeforeImage" & _
            '        ",EXCEPTDATASTORE" & _
            '        ",TextQualifier" & _
            '        ",ColumnDelimiter" & _
            '        ",RowDelimiter" & _
            '        ",Description" & _
            '        ",Operationtype" & _
            '        ",Port" & _
            '        ",QueMgr" & _
            '        ",UOW" & _
            '        ",SelectEvery" & _
            '        ",OnCmmtKey" & _
            '        ") " & _
            '        " Values(" & _
            '        Me.Project.GetQuotedText & _
            '        "," & Me.Environment.GetQuotedText & _
            '        "," & Me.Engine.ObjSystem.GetQuotedText & _
            '        "," & Me.Engine.GetQuotedText & _
            '        "," & Me.GetQuotedText & _
            '        "," & FixStr(Me.DatastoreType) & _
            '        ",'" & FixStr(Me.DsPhysicalSource) & "'" & _
            '        ",'" & FixStr(Me.DsDirection) & "'" & _
            '        ",'" & FixStr(Me.DsAccessMethod) & "'" & _
            '        ",'" & FixStr(Me.DsCharacterCode) & "'" & _
            '        ",'" & FixStr(Me.IsIMSPathData) & "'" & _
            '        ",'" & FixStr(Me.IsSkipChangeCheck) & "'" & _
            '        ",'" & FixStr("0") & "'" & _
            '        ",'" & FixStr(Me.ExceptionDatastore) & "'" & _
            '        ",'" & FixStr(Me.TextQualifier) & "'" & _
            '        ",'" & FixStr(Me.ColumnDelimiter) & "'" & _
            '        ",'" & FixStr(Me.RowDelimiter) & "'" & _
            '        ",'" & FixStr(Me.DatastoreDescription) & "'" & _
            '        ",'" & FixStr(Me.OperationType) & "'" & _
            '        ",'" & FixStr(Me.DsPort) & "'" & _
            '        ",'" & FixStr(Me.DsQueMgr) & "'" & _
            '        ",'" & FixStr(Me.DsUOW) & "'" & _
            '        ",'" & FixStr(Me.Poll) & "'" & _
            '        ",'" & FixStr("0") & "'" & _
            '        ")"
            '    Else
            '        strSql = "INSERT INTO " & Me.Project.tblDatastores & _
            '        "(ProjectName" & _
            '        ",EnvironmentName" & _
            '        ",SystemName" & _
            '        ",EngineName" & _
            '        ",DatastoreName" & _
            '        ",DatastoreType" & _
            '        ",DsPhysicalSource" & _
            '        ",DsDirection" & _
            '        ",DsAccessMethod" & _
            '        ",DsCharacterCode" & _
            '        ",IsOrdered" & _
            '        ",IsIMSPathData" & _
            '        ",ISSKIPCHGCHECK" & _
            '        ",IsBeforeImage" & _
            '        ",EXCEPTDATASTORE" & _
            '        ",TextQualifier" & _
            '        ",ColumnDelimiter" & _
            '        ",RowDelimiter" & _
            '        ",Description" & _
            '        ",Operationtype" & _
            '        ",Port" & _
            '        ",QueMgr" & _
            '        ",UOW" & _
            '        ",SelectEvery" & _
            '        ",OnCmmtKey" & _
            '        ") " & _
            '        " Values(" & _
            '        Me.Project.GetQuotedText & _
            '        "," & Me.Environment.GetQuotedText & _
            '        "," & Quote(DBNULL) & _
            '        "," & Quote(DBNULL) & _
            '        "," & Me.GetQuotedText & _
            '        "," & FixStr(Me.DatastoreType) & _
            '        ",'" & FixStr(Me.DsPhysicalSource) & "'" & _
            '        ",'" & FixStr(Me.DsDirection) & "'" & _
            '        ",'" & FixStr(Me.DsAccessMethod) & "'" & _
            '        ",'" & FixStr(Me.DsCharacterCode) & "'" & _
            '        ",'" & FixStr("0") & "'" & _
            '        ",'" & FixStr(Me.IsIMSPathData) & "'" & _
            '        ",'" & FixStr(Me.IsSkipChangeCheck) & "'" & _
            '        ",'" & FixStr("0") & "'" & _
            '        ",'" & FixStr(Me.ExceptionDatastore) & "'" & _
            '        ",'" & FixStr(Me.TextQualifier) & "'" & _
            '        ",'" & FixStr(Me.ColumnDelimiter) & "'" & _
            '        ",'" & FixStr(Me.RowDelimiter) & "'" & _
            '        ",'" & FixStr(Me.DatastoreDescription) & "'" & _
            '        ",'" & FixStr(Me.OperationType) & "'" & _
            '        ",'" & FixStr(Me.DsPort) & "'" & _
            '        ",'" & FixStr(Me.DsQueMgr) & "'" & _
            '        ",'" & FixStr(Me.DsUOW) & "'" & _
            '        ",'" & FixStr(Me.Poll) & "'" & _
            '        ",'" & FixStr("0") & "'" & _
            '        ")"
            '    End If

            'Else '/// V3 meta
            'If Me.Engine IsNot Nothing Then
            strSql = "INSERT INTO " & Me.Project.tblDatastores & "(" & _
                            "ProjectName" & _
                            ",EnvironmentName" & _
                            ",SystemName" & _
                            ",EngineName" & _
                            ",DatastoreName" & _
                            ",DSDIRECTION" & _
                            ",DSTYPE" & _
                            ",DATASTOREDESCRIPTION" & _
                            ") " & _
                            " Values(" & _
                            Me.Project.GetQuotedText & _
                            "," & Me.Environment.GetQuotedText & _
                            "," & Me.Engine.ObjSystem.GetQuotedText & _
                            "," & Me.Engine.GetQuotedText & _
                            "," & Me.GetQuotedText & _
                            "," & Quote(Me.DsDirection) & _
                            "," & Me.DatastoreType & _
                            "," & Quote(Me.DatastoreDescription) & _
                            ")"
            'Else
            'strSql = "INSERT INTO " & Me.Project.tblDatastores & "(" & _
            '                "ProjectName" & _
            '                ",EnvironmentName" & _
            '                ",SystemName" & _
            '                ",EngineName" & _
            '                ",DatastoreName" & _
            '                ",DSDIRECTION" & _
            '                ",DSTYPE" & _
            '                ",DATASTOREDESCRIPTION" & _
            '                ") " & _
            '                " Values(" & _
            '                Me.Project.GetQuotedText & _
            '                "," & Me.Environment.GetQuotedText & _
            '                "," & Quote(DBNULL) & _
            '                "," & Quote(DBNULL) & _
            '                "," & Me.GetQuotedText & _
            '                "," & Quote(Me.DsDirection) & _
            '                "," & Me.DatastoreType & _
            '                "," & Quote(Me.DatastoreDescription) & _
            '                ")"
            'End If

            'End If

            cmd.CommandText = strSql
            Log(strSql)
            cmd.ExecuteNonQuery()

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            'If Me.DeleteATTR(cmd) = False Then
            '    AddNewDatastore = False
            '    Exit Try
            'End If
            If Me.InsertATTR(cmd) = False Then
                AddNewDatastore = False
                Exit Try
            End If
            'End If

            AddNewDatastore = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore AddNewDatastore", strSql)
            AddNewDatastore = False
        Catch ex As Exception
            LogError(ex, strSql, "clsDatastore AddNewDatastore", False)
            AddNewDatastore = False
        End Try

    End Function

    '///DONE
    Private Function AddNewSelections(ByRef cmd As Data.Odbc.OdbcCommand, Optional ByVal idx As Integer = -1) As Boolean
        '// Changed by TK and KS 11/3/06
        '// Delete unselected selections
        Dim strSql As String = ""
        Dim i As Integer
        Dim j As Integer
        Dim IsAdded As Boolean
        Dim cnt As Integer
        Dim selOld As clsDSSelection
        Dim selNew As clsDSSelection
        Dim objFld As clsField
        Dim sidx As Integer
        Dim eidx As Integer

        Try
            '// added by TK and KS 11/3/06
            If (idx = -1) Then
                sidx = 0
                eidx = Me.ObjSelections.Count - 1
            Else
                sidx = idx
                eidx = idx
            End If


            '//compare each new selection with old one
            For i = sidx To eidx
                selNew = Me.ObjSelections(i)
                '//check new selection with old selection if not found means its added 
                'so we should add to database
                IsAdded = True '//first make flag true and if found in new then make it false

                For j = 0 To Me.OldObjSelections.Count - 1
                    selOld = Me.OldObjSelections(j)
                    If (selNew.SelectionName = selOld.SelectionName) And (selNew.ObjStructure.Text = selOld.ObjStructure.Text) And (selNew.ObjStructure.Environment.Text = selOld.ObjStructure.Environment.Text) And (selNew.Parent.Text = selOld.Parent.Text) And (selNew.IsModified = False) Then
                        '//if you come here means found Id in old collection means not new so we should not add it 
                        IsAdded = False
                        Exit For
                    End If
                Next

                If IsAdded = True Then
                    '//Add datastore selection and its fields related to datastore


                    '//delete any previous record to prevent any possible duplicate entries
                    strSql = "DELETE FROM " & Me.Project.tblDSselFields & _
                    " WHERE  DatastoreName= " & Me.GetQuotedText & _
                    " And EngineName= " & Me.Engine.GetQuotedText & _
                    " And SystemName= " & Me.Engine.ObjSystem.GetQuotedText & _
                    " And EnvironmentName= " & Me.Environment.GetQuotedText & _
                    " And ProjectName= " & Me.Project.GetQuotedText & _
                    " And DescriptionName= " & selNew.ObjStructure.GetQuotedText & _
                    " And SelectionName= " & selNew.GetQuotedText

                    cmd.CommandText = strSql
                    Log(strSql)
                    cmd.ExecuteNonQuery()

                    strSql = "DELETE FROM " & Me.Project.tblDSselections & _
                    " WHERE  DatastoreName= " & Me.GetQuotedText & _
                    " And EngineName= " & Me.Engine.GetQuotedText & _
                    " And SystemName= " & Me.Engine.ObjSystem.GetQuotedText & _
                    " And EnvironmentName= " & Me.Environment.GetQuotedText & _
                    " And ProjectName= " & Me.Project.GetQuotedText & _
                    " And DescriptionName= " & selNew.ObjStructure.GetQuotedText & _
                    " And SelectionName= " & selNew.GetQuotedText

                    cmd.CommandText = strSql
                    Log(strSql)
                    cmd.ExecuteNonQuery()

                    '//Add entry to DSSELECTIONS

                    strSql = "INSERT INTO " & Me.Project.tblDSselections & "(DatastoreName,EngineName,SystemName,EnvironmentName,ProjectName,DescriptionName,SelectionName,Parent) Values(" & Me.GetQuotedText & "," & Me.Engine.GetQuotedText & "," & Me.Engine.ObjSystem.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.Project.GetQuotedText & "," & selNew.ObjStructure.GetQuotedText & "," & selNew.GetQuotedText & "," & selNew.Parent.GetQuotedText & ")"

                    Log(strSql)
                    cmd.CommandText = strSql
                    cmd.ExecuteNonQuery()



                    '//Now add fields of current selection in the loop
                    'selNew.LoadItems()


                    '//Add entry to DSSELFIELDS
                    cnt = selNew.DSSelectionFields.Count
                    For j = 0 To cnt - 1
                        objFld = selNew.DSSelectionFields(j)

                        strSql = "INSERT INTO " & Me.Project.tblDSselFields & " (" & _
                        "ProjectName," & _
                        "EnvironmentName," & _
                        "SystemName," & _
                        "EngineName," & _
                        "DatastoreName," & _
                        "DescriptionName, " & _
                        "SelectionName, " & _
                        "FieldName, " & _
                        "ParentName, " & _
                        "SEQNO, " & _
                        "DescFieldDescription, " & _
                        "NChildren, " & _
                        "NLevel, " & _
                        "Ntimes, " & _
                        "NOccNo, " & _
                        "Datatype, " & _
                        "NOffSet, " & _
                        "NLength, " & _
                        "NScale, " & _
                        "CanNull, " & _
                        "ISKEY, " & _
                        "OrgName, " & _
                        "DateFormat, " & _
                        "Label, " & _
                        "InitVal, " & _
                        "ReType, " & _
                        "InValid, " & _
                        "ExtType, " & _
                        "IdentVal, " & _
                        "ForeignKey) " & _
                        "Values( " & _
                        Me.Project.GetQuotedText & "," & _
                        Me.Engine.ObjSystem.Environment.GetQuotedText & "," & _
                        Me.Engine.ObjSystem.GetQuotedText & "," & _
                        Me.Engine.GetQuotedText & "," & _
                        Me.GetQuotedText & "," & _
                        selNew.ObjStructure.GetQuotedText & "," & _
                        selNew.GetQuotedText & "," & _
                        objFld.GetQuotedText & "," & _
                        Quote(objFld.ParentName) & "," & _
                        objFld.SeqNo & "," & _
                        Quote(objFld.FieldDesc) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                        objFld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                        Quote(objFld.OrgName) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                        Quote(objFld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                        ");" '//TK 11/8/06 and 4/5/07 and 2/15/09

                        cmd.CommandText = strSql
                        Log(strSql)
                        cmd.ExecuteNonQuery()



                    Next '//end field loop
                End If
            Next '/// end selection loop

            AddNewSelections = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore AddNewSelections", strSql)
            AddNewSelections = False
        Catch ex As Exception
            LogError(ex, "clsDatastore AddNewSelections", strSql)
            AddNewSelections = False
        End Try

    End Function

    '///DONE
    Private Function UpdateDatastore(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim strSql As String = ""

        Try
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    If Me.Engine IsNot Nothing Then
            '        '//changed : IsbeforeImage added by npatel on 8/10/05
            '        strSql = "Update " & Me.Project.tblDatastores & " Set " & _
            '        "DatastoreName=" & "'" & FixStr(Me.DatastoreName) & "'" & _
            '        ",DatastoreType=" & FixStr(Me.DatastoreType) & _
            '        ",DsPhysicalSource=" & "'" & FixStr(Me.DsPhysicalSource) & "'" & _
            '        ",DsDirection=" & "'" & FixStr(Me.DsDirection) & "'" & _
            '        ",DsAccessMethod=" & "'" & FixStr(Me.DsAccessMethod) & "'" & _
            '        ",DsCharacterCode=" & "'" & FixStr(Me.DsCharacterCode) & "'" & _
            '        ",IsOrdered=" & "'" & FixStr("0") & "'" & _
            '        ",IsIMSPathData=" & "'" & FixStr(Me.IsIMSPathData) & "'" & _
            '        ",ISSKIPCHGCHECK=" & "'" & FixStr(Me.IsSkipChangeCheck) & "'" & _
            '        ",IsBeforeImage=" & "'" & FixStr("0") & "'" & _
            '        ",EXCEPTDATASTORE=" & "'" & FixStr(Me.ExceptionDatastore) & "'" & _
            '        ",TextQualifier=" & "'" & FixStr(Me.TextQualifier) & "'" & _
            '        ",ColumnDelimiter=" & "'" & FixStr(Me.ColumnDelimiter) & "'" & _
            '        ",RowDelimiter=" & "'" & FixStr(Me.RowDelimiter) & "'" & _
            '        ",Description=" & "'" & FixStr(Me.DatastoreDescription) & "'" & _
            '        ",Operationtype=" & "'" & FixStr(Me.OperationType) & "'" & _
            '        ",Port=" & "'" & FixStr(Me.DsPort) & "'" & _
            '        ",QueMgr=" & "'" & FixStr(Me.DsQueMgr) & "'" & _
            '        ",UOW=" & "'" & FixStr(Me.DsUOW) & "'" & _
            '        ",SelectEvery=" & "'" & FixStr(Me.Poll) & "'" & _
            '        ",OnCmmtKey=" & "'" & FixStr("0") & "'" & _
            '        " WHERE  DatastoreName=" & Me.GetQuotedText & _
            '        " AND EngineName=" & Me.Engine.GetQuotedText & _
            '        " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            '        " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '        " AND ProjectName=" & Me.Project.GetQuotedText

            '    Else
            '        '//changed : IsbeforeImage added by npatel on 8/10/05
            '        strSql = "Update " & Me.Project.tblDatastores & " Set " & _
            '         "DatastoreName=" & "'" & FixStr(Me.DatastoreName) & "'" & _
            '         ",DatastoreType=" & FixStr(Me.DatastoreType) & _
            '         ",DsPhysicalSource=" & "'" & FixStr(Me.DsPhysicalSource) & "'" & _
            '         ",DsDirection=" & "'" & FixStr(Me.DsDirection) & "'" & _
            '         ",DsAccessMethod=" & "'" & FixStr(Me.DsAccessMethod) & "'" & _
            '         ",DsCharacterCode=" & "'" & FixStr(Me.DsCharacterCode) & "'" & _
            '         ",IsOrdered=" & "'" & FixStr("0") & "'" & _
            '         ",IsIMSPathData=" & "'" & FixStr(Me.IsIMSPathData) & "'" & _
            '         ",ISSKIPCHGCHECK=" & "'" & FixStr(Me.IsSkipChangeCheck) & "'" & _
            '         ",IsBeforeImage=" & "'" & FixStr("0") & "'" & _
            '         ",EXCEPTDATASTORE=" & "'" & FixStr(Me.ExceptionDatastore) & "'" & _
            '         ",TextQualifier=" & "'" & FixStr(Me.TextQualifier) & "'" & _
            '         ",ColumnDelimiter=" & "'" & FixStr(Me.ColumnDelimiter) & "'" & _
            '         ",RowDelimiter=" & "'" & FixStr(Me.RowDelimiter) & "'" & _
            '         ",Description=" & "'" & FixStr(Me.DatastoreDescription) & "'" & _
            '         ",Operationtype=" & "'" & FixStr(Me.OperationType) & "'" & _
            '         ",Port=" & "'" & FixStr(Me.DsPort) & "'" & _
            '         ",QueMgr=" & "'" & FixStr(Me.DsQueMgr) & "'" & _
            '         ",UOW=" & "'" & FixStr(Me.DsUOW) & "'" & _
            '         ",SelectEvery=" & "'" & FixStr(Me.Poll) & "'" & _
            '         ",OnCmmtKey=" & "'" & FixStr("0") & "'" & _
            '         " WHERE  DatastoreName=" & Me.GetQuotedText & _
            '         " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '         " AND ProjectName=" & Me.Project.GetQuotedText
            '    End If

            '    cmd.CommandText = strSql
            '    Log(strSql)
            '    cmd.ExecuteNonQuery()

            'Else
            '    If Me.Engine IsNot Nothing Then
            strSql = "Update " & Me.Project.tblDatastores & " Set " & _
                            "DatastoreName=" & Quote(Me.DatastoreName) & _
                            ",DSDIRECTION=" & Quote(Me.DsDirection) & _
                            ",DSTYPE=" & Me.DatastoreType & _
                            ",DatastoreDescription=" & "'" & FixStr(Me.DatastoreDescription) & "'" & _
                            " WHERE  DatastoreName=" & Me.GetQuotedText & _
                            " AND EngineName=" & Me.Engine.GetQuotedText & _
                            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                            " AND ProjectName=" & Me.Project.GetQuotedText
            '    Else
            'strSql = "Update " & Me.Project.tblDatastores & " Set " & _
            '                "DatastoreName=" & Quote(Me.DatastoreName) & _
            '                ",DSDIRECTION=" & Quote(Me.DsDirection) & _
            '                ",DSTYPE=" & Me.DatastoreType & _
            '                ",DatastoreDescription=" & "'" & FixStr(Me.DatastoreDescription) & "'" & _
            '                " WHERE  DatastoreName=" & Me.GetQuotedText & _
            '                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '                " AND ProjectName=" & Me.Project.GetQuotedText
            '    End If

            cmd.CommandText = strSql
            Log(strSql)
            cmd.ExecuteNonQuery()

            If Me.DeleteATTR(cmd) = False Then
                UpdateDatastore = False
                Exit Try
            End If
            If Me.InsertATTR(cmd) = False Then
                UpdateDatastore = False
                Exit Try
            End If

            'End If

            UpdateDatastore = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore UpdateDatastore", strSql)
            UpdateDatastore = False
        Catch ex As Exception
            LogError(ex, "clsDatastore UpdateDatastore", strSql)
            UpdateDatastore = False
        End Try

    End Function

    '///DONE
    Private Function DeleteUnselectedSelections(ByRef cmd As Data.Odbc.OdbcCommand) As Boolean

        Dim strSql As String = ""
        Dim i As Integer
        Dim j As Integer
        Dim IsDeleted As Boolean
        Dim selOld As clsDSSelection
        Dim selNew As clsDSSelection

        Try
            For i = 0 To Me.OldObjSelections.Count - 1
                selOld = OldObjSelections(i)
                '//check old selection with current selection if not found means 
                'its unselected so we should delete from database
                IsDeleted = True '//first make flag true and if found in new then make it false
                For j = 0 To Me.ObjSelections.Count - 1
                    selNew = ObjSelections(j)
                    If (selNew.SelectionName = selOld.SelectionName) And (selNew.ObjStructure.Text = selOld.ObjStructure.Text) And (selNew.ObjStructure.Environment.Text = selOld.ObjStructure.Environment.Text) And (selNew.Parent.Text = selOld.Parent.Text) Then
                        '//if you come here means found oldid in new collection means 
                        'not unselected so we should not delete it 
                        IsDeleted = False
                        Exit For
                    End If
                Next

                If IsDeleted = True Then
                    '//Delete datastore selection and its fields related to datastore

                    '//First always children => delete entry from DSSELFIELDS
                    strSql = "DELETE FROM " & Me.Project.tblDSselFields & _
                    " Where  DatastoreName= " & Me.GetQuotedText & _
                    " And EngineName= " & Me.Engine.GetQuotedText & _
                    " And SystemName= " & Me.Engine.ObjSystem.GetQuotedText & _
                    " And EnvironmentName= " & Me.Environment.GetQuotedText & _
                    " And ProjectName= " & Me.Project.GetQuotedText & _
                    " And DescriptionName= " & selOld.ObjStructure.GetQuotedText & _
                    " And SelectionName= " & selOld.GetQuotedText                  '

                    cmd.CommandText = strSql
                    Log(strSql)
                    cmd.ExecuteNonQuery()

                    '//delete entry from DSSELECTIONS

                    strSql = "DELETE FROM " & Me.Project.tblDSselections & _
                    " Where DatastoreName= " & Me.GetQuotedText & _
                    " And EngineName= " & Me.Engine.GetQuotedText & _
                    " And SystemName= " & Me.Engine.ObjSystem.GetQuotedText & _
                    " And EnvironmentName= " & Me.Environment.GetQuotedText & _
                    " And ProjectName= " & Me.Project.GetQuotedText & _
                    " And DescriptionName= " & selOld.ObjStructure.GetQuotedText & _
                    " And SelectionName= " & selOld.GetQuotedText

                    cmd.CommandText = strSql
                    Log(strSql)
                    cmd.ExecuteNonQuery()


                    selOld.ObjSelection.ObjDSselections.Remove(selOld.Key)

                End If
            Next

            DeleteUnselectedSelections = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore DeleteUnselectedSelections", strSql)
            DeleteUnselectedSelections = False
        Catch ex As Exception
            LogError(ex, "clsDatastore DeleteUnselectedSelections", strSql)
            DeleteUnselectedSelections = False
        End Try

    End Function

    Function LoadAllDSSelFlds(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean

        Try
            For Each DSsel As clsDSSelection In Me.ObjSelections
                LoadDSSelectionFields(DSsel, Incmd)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsDatastore LoadDSSelFlds")
            Return False
        End Try

    End Function

    '///DONE
    Function LoadDSSelectionFields(ByRef objDSsel As clsDSSelection, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing, Optional ByVal Treelode As Boolean = False) As Boolean

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        Dim cmd As System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.DataRow
        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim sql As String = ""
        Dim i As Integer

        Try
            'If Treelode = True Then Exit Function

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If
            'cnn.Open()
            'cmd = cnn.CreateCommand

            '//we need to load extended attributes
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,FIELDNAME,PARENTNAME,SEQNO,DESCFIELDDESCRIPTION,NCHILDREN,NLEVEL,NTIMES,NOCCNO,DATATYPE,NOFFSET,NLENGTH,NSCALE,CANNULL,ISKEY,ORGNAME,DATEFORMAT,LABEL,INITVAL,RETYPE,INVALID,EXTTYPE,IDENTVAL,FOREIGNKEY from " & Me.Project.tblDSselFields & _
            " where selectionname=" & objDSsel.GetQuotedText & _
            " AND DescriptionName=" & objDSsel.ObjStructure.GetQuotedText & _
            " AND DatastoreName=" & Me.GetQuotedText & _
            " AND EngineName=" & Me.Engine.GetQuotedText & _
            " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText & _
            " ORDER BY SEQNO"


            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            dt = New System.Data.DataTable("temp2")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)
                Dim fld As clsField

                'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    fld = SearchDSFieldByName(objDSsel, GetVal(dr("FieldName")), GetVal(dr("StructureName")), GetVal(dr("EnvironmentName")), GetVal(dr("ProjectName")))
                'Else
                fld = SearchDSFieldByName(objDSsel, GetVal(dr("FieldName")), GetVal(dr("DescriptionName")), GetVal(dr("EnvironmentName")), GetVal(dr("ProjectName")))
                'End If

                If fld Is Nothing Then
                    Log("Field [" & dr("FieldName") & "] not found under this selection [" & objDSsel.SelectionName & "]")
                Else
                    'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    SetFieldExtendedAttributes(dr, fld)
                    fld.Parent = Me
                    'Else
                    For Each DSSfld As clsField In objDSsel.DSSelectionFields
                        If fld.FieldName = DSSfld.FieldName Then
                            objDSsel.DSSelectionFields.Remove(DSSfld)
                            GoTo nextfld
                        End If
                    Next
nextfld:            objDSsel.DSSelectionFields.Add(fld)
                End If
            Next

            LoadDSSelectionFields = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore LoadDSSelectionFields", sql)
            LoadDSSelectionFields = False
        Catch ex As Exception
            LogError(ex, "clsDatastore LoadDSSelectionFields", sql)
            LoadDSSelectionFields = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Sub SetFieldExtendedAttributes(ByVal dr As System.Data.DataRow, ByRef fld As clsField)

        Try

            fld.SeqNo = GetVal(dr("seqno"))
            fld.FieldDesc = GetStr(GetVal(dr("descfielddescription")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_NCHILDREN, GetVal(dr("NCHILDREN")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LEVEL, GetVal(dr("NLEVEL")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_TIMES, GetVal(dr("NTIMES")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OCCURS, GetVal(dr("NOCCNO")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, dr("DATATYPE"))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OFFSET, GetVal(dr("NOFFSET")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("NLENGTH")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("NSCALE")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetStr(GetVal(dr("CANNULL"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_ISKEY, GetStr(GetVal(dr("ISKEY"))))
            fld.OrgName = GetStr(GetVal(dr("ORGNAME")))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT, GetStr(GetVal(dr("DATEFORMAT"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LABEL, GetStr(GetVal(dr("LABEL"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INITVAL, GetStr(GetVal(dr("INITVAL"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_RETYPE, GetStr(GetVal(dr("RETYPE"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INVALID, GetStr(GetVal(dr("INVALID"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_EXTTYPE, GetStr(GetVal(dr("EXTTYPE"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_IDENTVAL, GetStr(GetVal(dr("IDENTVAL"))))
            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_FKEY, GetStr(GetVal(dr("FOREIGNKEY"))))


        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore SetFieldExtendedAttributes")
        Catch ex As Exception
            LogError(ex, "clsDatastore SetFieldExtendedAttributes")
        End Try

    End Sub

    '// added 12/29/2006 by TK to build new clsDSSelection object
    Public Function CloneSSeltoDSSel(ByVal ss As clsStructureSelection, Optional ByRef pDSSel As clsDSSelection = Nothing, Optional ByVal IsStrFileChg As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing, Optional ByVal TreeLode As Boolean = False) As clsDSSelection

        Dim NewDSS As New clsDSSelection
        Dim cmd As Odbc.OdbcCommand


        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
                cmd.Transaction = Incmd.Transaction
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If
            'Save Structure Selection properties, etc. to new DSSelection Object 
            'that belongs to the datastore not the structure
            'If ss.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    ss.ObjStructure.LoadDescAttr(cmd)
            'End If

            ss.ObjStructure.LoadItems(, , cmd)
            NewDSS.SelectionName = ss.SelectionName
            NewDSS.ObjStructure = ss.ObjStructure
            NewDSS.ObjSelection = ss
            NewDSS.ObjDatastore = Me
            NewDSS.SelectionDescription = ss.SelectionDescription
            NewDSS.IsModified = True
            NewDSS.Text = ss.Text
            NewDSS.ObjTreeNode = Me.ObjTreeNode
            NewDSS.IsRenamed = ss.IsRenamed
            NewDSS.SeqNo = ss.SeqNo

            'Parent can be Datastore or DSSelection
            If Not pDSSel Is Nothing Then
                NewDSS.Parent = pDSSel
            Else
                NewDSS.Parent = Me
            End If

            'Now add Fields to DSSelection based on whether or not Structure file is changed
            If IsStrFileChg = False And TreeLode = False Then
                ss.LoadMe(cmd)
                If ss.ObjSelectionFields.Count <> 0 Then
                    For Each fld As clsField In ss.ObjSelectionFields
                        Dim newfld As clsField     'create a new fld object
                        newfld = fld.Clone(NewDSS)     'clone old field into new field w/new parent
                        NewDSS.DSSelectionFields.Add(newfld)  'add new fld object to the DSSel object
                    Next
                Else
                    For Each fld As clsField In ss.ObjStructure.ObjFields
                        Dim newfld As clsField     'create a new fld object
                        newfld = fld.Clone(NewDSS)     'clone old field into new field w/new parent
                        NewDSS.DSSelectionFields.Add(newfld)  'add new fld object to the DSSel object
                    Next
                End If

                If ss.ObjDSselections.Contains(NewDSS.Key) Then
                    ss.ObjDSselections.Remove(NewDSS.Key)
                    ss.ObjDSselections.Add(NewDSS, NewDSS.Key)
                Else
                    ss.ObjDSselections.Add(NewDSS, NewDSS.Key)
                End If
            End If

            CloneSSeltoDSSel = NewDSS

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore CloneSSeltoDSSel")
            CloneSSeltoDSSel = Nothing
        Catch ex As Exception
            LogError(ex, "clsDatastore CloneSSeltoDSSel")
            CloneSSeltoDSSel = Nothing
        End Try

    End Function

    '/// New Function added by TKarasch 1/2007 for child DS Selection Handling based on Foreign Key
    Public Function SetDSselParents() As Boolean

        Dim i As Integer = 0  'Datastore Selections Loop
        Dim j As Integer = 0  'Fields Loop
        Dim k As Integer = 0  'Inner Datastore Selections Loop
        Dim selobj As clsDSSelection  'Selection object presently being Set
        Dim selobj2 As clsDSSelection  'Selection Object for comparison
        Dim fldobj As clsField  'Field to check for foreign key
        Dim fkeystring As String = ""  'Foreign key of present field in selobj
        Dim returnstring As String = ""  'Name of Foreign DSselection
        Dim NoFkeyFlag As Boolean = True  'Flag to set if Foreign DSsel found, else reset Fkey to nothing

        Try
            '// set all parents to datastore first
            For i = 0 To Me.ObjSelections.Count - 1
                selobj = Me.ObjSelections(i)
                selobj.Parent = Me
            Next

            '/// first set all the parents based on Foreign key
            For i = 0 To Me.ObjSelections.Count - 1
                selobj = Me.ObjSelections(i)
                For j = 0 To selobj.DSSelectionFields.Count - 1
                    fldobj = selobj.DSSelectionFields(j) '/// search each field of each selection for Fkeys
                    fkeystring = fldobj.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY).ToString
                    If Not fkeystring = "" Then '/// if there is a foriegn key then that is the new parent
                        returnstring = GetFKey(fkeystring) '// returnstring is the SelectionName of the FKey
                        For k = 0 To Me.ObjSelections.Count - 1
                            selobj2 = Me.ObjSelections(k)
                            If selobj2.SelectionName = returnstring Then
                                selobj.Parent = selobj2 '// now find the selection and set it as parent
                                selobj.IsModified = True '// set as modified so changes will save
                                selobj.SetAllFieldsModified() '// set all fields modified so changes will save
                                NoFkeyFlag = False  '//If foreign selection present, set and jump out
                                Exit For '// Parent Set so Jump out
                            End If
                        Next '//selection to see if it is parent
                        If NoFkeyFlag = True Then '// if fkey's parent is gone, then reset fkey to nothing
                            fldobj.SetSingleFieldAttr(enumFieldAttributes.ATTR_FKEY, Nothing)
                        End If
                        NoFkeyFlag = True
                        Exit For '// Jump out and do it again for the next selection
                    End If
                Next '// field to see if it has a foreign key
            Next '// selection in the datastore to check

            '/// now fill all the selections of selections based on parent
            For i = 0 To Me.ObjSelections.Count - 1
                selobj = Me.ObjSelections(i)
                selobj.ObjDSSelections.Clear()
                For j = 0 To Me.ObjSelections.Count - 1
                    selobj2 = Me.ObjSelections(j)
                    If selobj2.Parent Is selobj Then
                        selobj.ObjDSSelections.Add(selobj2)
                        selobj.IsModified = True
                    End If
                Next
            Next

            SetDSselParents = True

        Catch ex As Exception
            LogError(ex, "clsDatastore SetDSselParents")
            SetDSselParents = False
        End Try

    End Function

    Function SetIsMapped(Optional ByVal reload As Boolean = False, Optional ByVal TreeLoad As Boolean = False) As Boolean

        Try
            For Each DSsel As clsDSSelection In Me.ObjSelections
                '/// this section modified 9/07 to differentiate source and target
                For Each proc As clsTask In Me.Engine.Procs
                    proc.LoadItems(reload, TreeLoad)
                    'proc.LoadDatastores()
                    'proc.LoadMappings(reload)

                    If Me.DsDirection = DS_DIRECTION_SOURCE Then
                        For Each map As clsMapping In proc.ObjMappings
                            If map.SourceParent = DSsel.Text And map.SourceDataStore = Me.Text Then
                                DSsel.IsMapped = True
                                GoTo NextSel
                            Else
                                DSsel.IsMapped = False
                            End If
                        Next
                    Else
                        For Each map As clsMapping In proc.ObjMappings
                            If map.TargetParent = DSsel.Text And map.TargetDataStore = Me.Text Then
                                DSsel.IsMapped = True
                                GoTo NextSel
                            Else
                                DSsel.IsMapped = False
                            End If
                        Next
                    End If
                Next
                'For Each main As clsTask In Me.Engine.Mains
                '    main.LoadItems()
                '    main.LoadDatastores()
                '    main.LoadMappings()

                '    For Each map As clsMapping In main.ObjMappings
                '        If (map.TargetParent = DSsel.Text Or map.SourceParent = DSsel.Text) And (map.SourceDataStore = Me.Text Or map.TargetDataStore = Me.Text) Then
                '            DSsel.IsMapped = True
                '            GoTo NextSel
                '        Else
                '            DSsel.IsMapped = False
                '        End If
                '    Next
                'Next
                'For Each join As clsTask In Me.Engine.Joins
                '    join.LoadItems()
                '    join.LoadDatastores()
                '    join.LoadMappings()

                '    For Each map As clsMapping In join.ObjMappings
                '        If (map.TargetParent = DSsel.Text Or map.SourceParent = DSsel.Text) And (map.SourceDataStore = Me.Text Or map.TargetDataStore = Me.Text) Then
                '            DSsel.IsMapped = True
                '            GoTo NextSel
                '        Else
                '            DSsel.IsMapped = False
                '        End If
                '    Next
                'Next
                'For Each LU As clsTask In Me.Engine.Lookups
                '    LU.LoadItems()
                '    LU.LoadDatastores()
                '    LU.LoadMappings()

                '    For Each map As clsMapping In LU.ObjMappings
                '        If (map.TargetParent = DSsel.Text Or map.SourceParent = DSsel.Text) And (map.SourceDataStore = Me.Text Or map.TargetDataStore = Me.Text) Then
                '            DSsel.IsMapped = True
                '            GoTo NextSel
                '        Else
                '            DSsel.IsMapped = False
                '        End If
                '    Next
                'Next
NextSel:        '// next selection
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsDatastore SetIsMapped")
            Return False
        End Try

    End Function

    '//// USED ONLY FOR "MAP AS" --- WILL NOT PROCESS SOURCES -- ONLY TARGETS   
    'Public Function AddNewDs(ByVal dsType As Integer) As Boolean

    '    'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
    '    Dim cmd As New Odbc.OdbcCommand
    '    Dim tran As Odbc.OdbcTransaction = Nothing
    '    Dim selNew As clsDSSelection
    '    Dim i As Integer

    '    Try
    '        'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
    '        'cnn.Open()
    '        cmd.Connection = cnn

    '        '//We need to put in transaction because we will add Datastore,
    '        '//DSSELECTIONS and fields in 
    '        '//three steps so if one fails rollback all
    '        tran = cnn.BeginTransaction()
    '        cmd.Transaction = tran

    '        For i = 0 To Me.ObjSelections.Count - 1
    '            selNew = Me.ObjSelections(i)
    '            AddNewDatastore(cmd)
    '            AddNewSelections(cmd, i)
    '            AddToCollection(Me.Engine.Targets, Me, Me.GUID)
    '        Next

    '        tran.Commit()

    '        If Me.DsDirection = DS_DIRECTION_TARGET Then
    '            '//Add  datastore in engine's targets collection
    '            AddToCollection(Me.Engine.Targets, Me, Me.GUID)
    '        ElseIf Me.DsDirection = DS_DIRECTION_SOURCE Then
    '            '//Add  datastore in engine's source collection
    '            AddToCollection(Me.Engine.Sources, Me, Me.GUID)
    '        End If


    '        Me.IsModified = False
    '        AddNewDs = True

    '    Catch ex As Exception
    '        tran.Rollback()
    '        LogError(ex, ex.ToString, "clsDatastore AddNewDs")
    '        AddNewDs = False
    '    Finally
    '        'cnn.Close()
    '    End Try

    'End Function

    Public Function MapAsSave(ByRef cmd As Odbc.OdbcCommand) As Boolean

        Try
            '//We need to put in transaction because we will add Datastore,
            '//DSSELECTIONS and fields in 
            '//three steps so if one fails rollback all

            If UpdateDatastore(cmd) = False Then
                MapAsSave = False
                Exit Try
            End If
            If DeleteUnselectedSelections(cmd) = False Then
                MapAsSave = False
                Exit Try
            End If
            If AddNewSelections(cmd) = False Then
                MapAsSave = False
                Exit Try
            End If
            If UpdateSelectionFields(cmd) = False Then
                MapAsSave = False
                Exit Try
            End If

            Me.IsModified = False

            MapAsSave = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore MapAsSave")
            MapAsSave = False
        Catch ex As Exception
            LogError(ex, "clsDatastore MapAsSave")
            MapAsSave = False
        End Try

    End Function

    '//Write Global Datastore Properties to File
    'Function SaveGlobals() As Boolean

    '    Try
    '        If System.IO.File.Exists(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.Engine.EngineName & "." & _
    '        Me.DatastoreName & ".txt") = True Then
    '            System.IO.File.Delete(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.Engine.EngineName & "." & _
    '            Me.DatastoreName & ".txt")
    '        End If
    '        fsDSglobal = System.IO.File.Create(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.Engine.EngineName & "." _
    '        & Me.DatastoreName & ".txt")
    '        objWriteGlobal = New System.IO.StreamWriter(fsDSglobal)

    '        'Write Globals to File
    '        objWriteGlobal.WriteLine(Me.ExtTypeChar)
    '        objWriteGlobal.WriteLine(Me.ExtTypeNum)
    '        objWriteGlobal.WriteLine(Me.IfNullChar)
    '        objWriteGlobal.WriteLine(Me.IfNullNum)
    '        objWriteGlobal.WriteLine(Me.InValidChar)
    '        objWriteGlobal.WriteLine(Me.InValidNum)

    '        objWriteGlobal.Close()
    '        fsDSglobal.Close()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDatastore SaveGlobals")
    '        Return False
    '    End Try

    'End Function

    'Function LoadGlobals() As Boolean

    '    Try
    '        If System.IO.File.Exists(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.Engine.EngineName & "." _
    '        & Me.DatastoreName & ".txt") = False Then
    '            LoadGlobals = True
    '            Exit Function
    '        End If
    '        fsDSglobal = System.IO.File.Open(GetAppTemp() & "\" & Me.Project.ProjectName & "." & Me.Engine.EngineName & "." _
    '        & Me.DatastoreName & ".txt", IO.FileMode.Open)

    '        objReadGlobal = New System.IO.StreamReader(fsDSglobal)

    '        'Write Globals to File
    '        Me.ExtTypeChar = objReadGlobal.ReadLine()
    '        Me.ExtTypeNum = objReadGlobal.ReadLine()
    '        Me.IfNullChar = objReadGlobal.ReadLine()
    '        Me.IfNullNum = objReadGlobal.ReadLine()
    '        Me.InValidChar = objReadGlobal.ReadLine()
    '        Me.InValidNum = objReadGlobal.ReadLine()

    '        objReadGlobal.Close()
    '        fsDSglobal.Close()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDatastore LoadGlobals")
    '        Return False
    '    End Try

    'End Function

    Function InsertATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As Odbc.OdbcCommand
        Dim sql As String = ""
        Dim Attrib As String = ""
        Dim Value As String = ""

        Try
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            For i As Integer = 0 To 24
                Select Case i
                    Case 0
                        Attrib = "COLUMNDELIMITER"
                        Value = Me.ColumnDelimiter
                    Case 1
                        Attrib = "DATASTORETYPE"
                        Value = Me.DatastoreType
                    Case 2
                        Attrib = "DSACCESSMETHOD"
                        Value = Me.DsAccessMethod
                    Case 3
                        Attrib = "DSCHARACTERCODE"
                        Value = Me.DsCharacterCode
                    Case 4
                        Attrib = "DSDIRECTION"
                        Value = Me.DsDirection
                    Case 5
                        Attrib = "DSPHYSICALSOURCE"
                        Value = Me.DsPhysicalSource
                    Case 6
                        Attrib = "EXCEPTDATASTORE"
                        Value = Me.ExceptionDatastore
                    Case 7
                        Attrib = "RESTART"
                        Value = Me.Restart
                    Case 8
                        Attrib = "POLL"
                        Value = Me.Poll
                    Case 9
                        Attrib = "ISIMSPATHDATA"
                        Value = Me.IsIMSPathData
                    Case 10
                        Attrib = "ISSKIPCHGCHECK"
                        Value = Me.IsSkipChangeCheck
                    Case 11
                        Attrib = "OPERATIONTYPE"
                        Value = Me.OperationType
                    Case 12
                        Attrib = "PORT"
                        Value = Me.DsPort
                    Case 13
                        Attrib = "QUEMGR"
                        Value = Me.DsQueMgr
                    Case 14
                        Attrib = "ROWDELIMITER"
                        Value = Me.RowDelimiter
                    Case 15
                        Attrib = "TEXTQUALIFIER"
                        Value = Me.TextQualifier
                    Case 16
                        Attrib = "UOW"
                        Value = Me.DsUOW
                    Case 17
                        Attrib = "EXTTYPECHAR"
                        Value = Me.ExtTypeChar
                    Case 18
                        Attrib = "EXTTYPENUM"
                        Value = Me.ExtTypeNum
                    Case 19
                        Attrib = "IFNULLCHAR"
                        Value = Me.IfNullChar
                    Case 20
                        Attrib = "IFNULLNUM"
                        Value = Me.IfNullNum
                    Case 21
                        Attrib = "INVALIDCHAR"
                        Value = Me.InValidChar
                    Case 22
                        Attrib = "INVALIDNUM"
                        Value = Me.InValidNum
                    Case 23
                        Attrib = "LOOKUP"
                        Value = Me.IsLookUp.ToString
                    Case 24
                        Attrib = "ISKEYCHNG"
                        Value = Me.IsKeyChng
                        'Case 11
                        'Attrib = "ONCMMTKEY"
                        'Value = Me.IsCmmtKey
                        'Case 15
                        '    Attrib = "SELECTEVERY"
                        '    Value = Me.Poll
                End Select

                'If Me.Engine IsNot Nothing Then
                sql = "INSERT INTO " & Me.Project.tblDatastoresATTR & _
                                "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DATASTOREATTRB,DATASTOREATTRBVALUE) " & _
                                "Values('" & FixStr(Me.Project.ProjectName) & "','" & _
                                FixStr(Me.Environment.EnvironmentName) & "','" & _
                                FixStr(Me.Engine.ObjSystem.SystemName) & "','" & _
                                FixStr(Me.Engine.EngineName) & "','" & _
                                FixStr(Me.DatastoreName) & "','" & _
                                FixStr(Attrib) & "','" & _
                                FixStr(Value) & "')"
                'Else
                'sql = "INSERT INTO " & Me.Project.tblDatastoresATTR & _
                '                "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DATASTOREATTRB,DATASTOREATTRBVALUE) " & _
                '                "Values('" & FixStr(Me.Project.ProjectName) & "','" & _
                '                FixStr(Me.Environment.EnvironmentName) & "','" & _
                '                DBNULL & "','" & _
                '                DBNULL & "','" & _
                '                FixStr(Me.DatastoreName) & "','" & _
                '                FixStr(Attrib) & "','" & _
                '                FixStr(Value) & "')"
                'End If

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            InsertATTR = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore InsertATTR", sql)
            InsertATTR = False
        Catch ex As Exception
            LogError(ex, "clsDatastore InsertATTR", sql)
            InsertATTR = False
        End Try

    End Function

    Function DeleteATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As Odbc.OdbcCommand
        Dim sql As String = ""
        Dim Attrib As String = ""
        Dim Value As String = ""

        Try
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            'If Me.Engine IsNot Nothing Then
            sql = "DELETE FROM " & Me.Project.tblDatastoresATTR & _
                       " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                       " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                       " AND SYSTEMNAME=" & Quote(Me.Engine.ObjSystem.SystemName) & _
                       " AND ENGINENAME=" & Quote(Me.Engine.EngineName) & _
                       " AND DATASTORENAME=" & Quote(Me.DatastoreName)
            'Else
            'sql = "DELETE FROM " & Me.Project.tblDatastoresATTR & _
            '          " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            '          " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
            '          " AND SYSTEMNAME=" & Quote(DBNULL) & _
            '          " AND ENGINENAME=" & Quote(DBNULL) & _
            '          " AND DATASTORENAME=" & Quote(Me.DatastoreName)
            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            DeleteATTR = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsDatastore DeleteATTR", sql)
            DeleteATTR = False
        Catch ex As Exception
            LogError(ex, "clsDatastore DeleteATTR", sql)
            DeleteATTR = False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
