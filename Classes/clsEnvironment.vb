Public Class clsEnvironment
    Implements INode

    Private m_EnvironmentName As String = ""
    Private m_EnvironmentDescription As String = ""
    Private m_ObjProject As clsProject
    Private m_IsModified As Boolean
    Private m_LocalDTDDir As String
    Private m_LocalDDLDir As String
    Private m_LocalCobolDir As String
    Private m_LocalCDir As String
    Private m_LocalDMLDir As String
    Private m_LocalDBDDir As String
    Private m_LocalScriptDir As String
    Private m_DefaultStrDir As String = ""
    'Relative base project directory
    Private m_RelDir As String = ""
    'Have Paths Changed since last save
    Private m_PathChange As Boolean

    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_OldDMLobj As clsDMLinfo = Nothing
    Private m_IsLoaded As Boolean = False

    Public Connections As New Collection
    Public Structures As New Collection
    Public Systems As New Collection
    'Added 3/09 by TK
    'Public Datastores As New Collection
    Public Variables As New Collection
    'Public Joins As New Collection
    Public Procedures As New Collection

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
            Return Me.Project
        End Get
        Set(ByVal Value As INode)
            m_ObjProject = Value
        End Set
    End Property

    '/// OverHauled By TKarasch   April May  07
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsEnvironment
        Dim cmd As New Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)

            obj.EnvironmentName = Me.EnvironmentName
            obj.LocalDTDDir = Me.LocalDTDDir
            obj.LocalDDLDir = Me.LocalDDLDir
            obj.LocalCobolDir = Me.LocalCobolDir
            obj.LocalCDir = Me.LocalCDir
            obj.LocalDBDDir = Me.LocalDBDDir
            obj.LocalDMLDir = Me.LocalDMLDir
            obj.LocalScriptDir = Me.LocalScriptDir
            obj.SeqNo = Me.SeqNo
            obj.EnvironmentDescription = Me.EnvironmentDescription
            obj.IsModified = Me.IsModified
            obj.Parent = NewParent 'Me.Parent

            '//clone all environments under project
            If Cascade = True Then
                Dim NewConn As clsConnection
                Dim NewStr As clsStructure
                'Dim NewDS As clsDatastore
                Dim NewVar As clsVariable
                'Dim NewProc As clsTask
                Dim NewSys As clsSystem

                '//clone all connections under environment
                For Each conn As clsConnection In Me.Connections
                    'conn.LoadItems()
                    NewConn = conn.Clone(obj, True, cmd)
                    NewConn.Environment = obj
                    AddToCollection(obj.Connections, NewConn, NewConn.GUID)
                Next
                '//clone all structures under environment
                For Each str As clsStructure In Me.Structures
                    'If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                    '    str.LoadDescAttr()
                    'End If
                    'str.LoadItems()
                    NewStr = str.Clone(obj, True, cmd)
                    NewStr.Environment = obj
                    obj.Structures.Add(NewStr, NewStr.GUID)
                Next
                '// Clone all Datastores under env
                'For Each ds As clsDatastore In Me.Datastores
                '    'ds.LoadItems()

                '    NewDS = ds.Clone(obj, True, cmd)
                '    NewDS.Environment = obj
                '    AddToCollection(obj.Datastores, NewDS, NewDS.GUID)
                'Next
                '// Clone all Variables under env
                For Each var As clsVariable In Me.Variables
                    'var.LoadItems(True)

                    NewVar = var.Clone(obj, True, cmd)
                    NewVar.Environment = obj
                    AddToCollection(obj.Variables, NewVar, NewVar.GUID)
                Next
                '// Clone all Procedures under env
                'For Each proc As clsTask In Me.Procedures
                '    'proc.LoadDatastores(cmd)

                '    'proc.LoadMappings(True, cmd)

                '    NewProc = proc.Clone(obj, True, cmd)
                '    NewProc.Environment = obj
                '    AddToCollection(obj.Procedures, NewProc, NewProc.GUID)
                'Next
                '//clone all systems under environment
                For Each sys As clsSystem In Me.Systems
                    'sys.LoadItems()

                    NewSys = sys.Clone(obj, True, cmd)
                    NewSys.Environment = obj
                    obj.Systems.Add(NewSys, NewSys.GUID)
                Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsEnvironment Clone")
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
            Key = Me.Project.Text & KEY_SAP & Me.Text
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = EnvironmentName
        End Get
        Set(ByVal Value As String)
            EnvironmentName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_ENVIRONMENT
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Return m_ObjProject
        End Get
    End Property

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            
            cmd.Connection = cnn

            sql = "Update " & Me.Project.tblEnvironments & " set Environmentname='" & FixStr(Me.EnvironmentName) & _
            "',ENVIRONMENTDESCRIPTION='" & FixStr(Me.EnvironmentDescription) & "' where Environmentname='" & _
            Me.EnvironmentName & "' and Projectname='" & Me.Project.ProjectName & "'"

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.PathChange = True Then
                For Each Str As clsStructure In Me.Structures
                    If Str.ChangeFPath(cmd) = False Then
                        Save = False
                        Exit Try
                    End If
                Next
                Me.PathChange = False
            End If

            Save = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsEnv Save", sql)
            Return False
        End Try

    End Function

    '// modified by TK and KS 11/3/06
    '/// OverHauled By TKarasch   April May  07
    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '//When we add new record we need to find Unique Database Record ID.
                sql = "INSERT INTO " & Me.Project.tblEnvironments & "(" & _
                      "ProjectName,EnvironmentName,Description," & _
                      "LocalDTDDir,LocalDDLDir,LocalCobolDir,LocalCDir," & _
                      "LocalScriptDir,LocalModelDir) " & _
                      " Values(" & _
                      Me.Project.GetQuotedText & "," & _
                      Me.GetQuotedText & ",'" & _
                      FixStr(EnvironmentDescription) & "','" & _
                      FixStr(LocalDTDDir) & "','" & _
                      FixStr(LocalDDLDir) & "','" & _
                      FixStr(LocalCobolDir) & "','" & _
                      FixStr(LocalCDir) & "','" & _
                      FixStr(LocalScriptDir) & "','" & _
                      FixStr(LocalDBDDir) & "')"
            Else
                sql = "INSERT INTO " & Me.Project.tblEnvironments & "(ProjectName,Environmentname,ENVIRONMENTDESCRIPTION) Values('" & _
                FixStr(Me.Project.ProjectName) & "','" & Me.EnvironmentName & "','" & Me.EnvironmentDescription & "')"

                Me.InsertATTR(cmd)
            End If
           

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As clsConnection In Me.Connections
                    child.AddNew(True)
                Next
                For Each child As clsStructure In Me.Structures
                    If child.AddNew(True) = False Then
                        MsgBox("Description " & child.StructureName & " did not copy correctly")
                    End If
                Next
                'For Each child As clsDatastore In Me.Datastores
                '    If child.AddNew(True) = False Then
                '        MsgBox("Datastore " & child.DatastoreName & " did not copy correctly")
                '    End If
                'Next
                For Each child As clsVariable In Me.Variables
                    If child.AddNew(True) = False Then
                        MsgBox("Variable " & child.VariableName & " did not copy correctly")
                    End If
                Next
                For Each child As clsTask In Me.Procedures
                    If child.AddNew(True) = False Then
                        MsgBox("Procedure " & child.TaskName & " did not copy correctly")
                    End If
                Next
                For Each child As clsSystem In Me.Systems
                    If child.AddNew(True) = False Then
                        MsgBox("System " & child.SystemName & " did not copy correctly")
                    End If
                Next
            End If

            AddToCollection(Me.Project.Environments, Me, Me.GUID)

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsEnv AddNew", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// Function added 11/3/06 by KS and TK 
    '/// OverHauled By TKarasch   April May  07
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            '//When we add new record we need to find Unique Database Record ID.
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '//When we add new record we need to find Unique Database Record ID.
                sql = "INSERT INTO " & Me.Project.tblEnvironments & "(" & _
                      "ProjectName,EnvironmentName,Description," & _
                      "LocalDTDDir,LocalDDLDir,LocalCobolDir,LocalCDir," & _
                      "LocalScriptDir,LocalModelDir) " & _
                      " Values(" & _
                      Me.Project.GetQuotedText & "," & _
                      Me.GetQuotedText & ",'" & _
                      FixStr(EnvironmentDescription) & "','" & _
                      FixStr(LocalDTDDir) & "','" & _
                      FixStr(LocalDDLDir) & "','" & _
                      FixStr(LocalCobolDir) & "','" & _
                      FixStr(LocalCDir) & "','" & _
                      FixStr(LocalScriptDir) & "','" & _
                      FixStr(LocalDBDDir) & "')"

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Else
                sql = "INSERT INTO " & Me.Project.tblEnvironments & "(ProjectName,Environmentname,ENVIRONMENTDESCRIPTION) Values('" & _
                FixStr(Me.Project.ProjectName) & "','" & Me.EnvironmentName & "','" & Me.EnvironmentDescription & "')"

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()

                Me.InsertATTR(cmd)
            End If



            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                Dim child As INode
                For Each child In Connections
                    child.AddNew(cmd, True)
                Next
                For Each child In Structures
                    child.AddNew(cmd, True)
                Next
                'For Each child In Me.Datastores
                '    child.AddNew(cmd, True)
                'Next
                For Each child In Me.Variables
                    child.AddNew(cmd, True)
                Next
                For Each child In Me.Procedures
                    child.AddNew(cmd, True)
                Next
                For Each child In Systems
                    child.AddNew(cmd, True)
                Next
            End If

            AddToCollection(Me.Project.Environments, Me, Me.GUID)

            AddNew = True

            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsEnv AddNew", sql)
            Return False
        End Try

    End Function

    '/// OverHauled By TKarasch   April May  07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '//first delete all child objects
            If Cascade = True Then
                Dim o As INode
                For Each o In Me.Connections
                    o.Delete(cmd, cnn, Cascade)
                Next
                RemoveFromCollection(Me.Connections, "") '//clear collection
                For Each o In Me.Systems
                    o.Delete(cmd, cnn, Cascade, False)
                Next
                RemoveFromCollection(Me.Systems, "") '//clear collection
                For Each o In Me.Structures
                    o.Delete(cmd, cnn, Cascade, False)
                Next
                RemoveFromCollection(Me.Structures, "") '/// clear collection
                'For Each o In Me.Datastores
                '    o.Delete(cmd, cnn, Cascade, False)
                'Next
                'RemoveFromCollection(Me.Datastores, "")
                For Each o In Me.Variables
                    o.Delete(cmd, cnn, Cascade, False)
                Next
                RemoveFromCollection(Me.Variables, "")
                For Each o In Me.Procedures
                    o.Delete(cmd, cnn, Cascade, False)
                Next
                RemoveFromCollection(Me.Procedures, "")
            End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            Me.DeleteATTR(cmd)
            'End If

            sql = "Delete From " & Me.Project.tblEnvironments & " where EnvironmentName='" & FixStr(Me.EnvironmentName) & _
            "' AND ProjectName='" & Me.Project.ProjectName & "'"

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                RemoveFromCollection(Me.Project.Environments, Me.GUID)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsEnv Delete", sql)
            Delete = False
        End Try

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then Exit Try

            Dim cmd As New System.Data.Odbc.OdbcCommand
            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As New DataTable("temp")
            Dim dr As DataRow
            Dim sql As String = ""
            Dim Attrib As String = ""
            Dim Value As String = ""

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "SELECT ENVIRONMENTATTRB,ENVIRONMENTATTRBVALUE FROM " & Me.Project.tblEnvironmentsATTR & _
            " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.EnvironmentName)

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetStr(GetVal(dr("ENVIRONMENTATTRB")))
                Select Case Attrib
                    Case "DEFAULTSTRDIR"
                        Me.DefaultStrDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALCDIR"
                        Me.LocalCDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALCOBOLDIR"
                        Me.LocalCobolDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALDDLDIR"
                        Me.LocalDDLDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALDMLDIR"
                        Me.LocalDMLDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALDTDDIR"
                        Me.LocalDTDDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALDBDDIR"
                        Me.LocalDBDDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "LOCALSCRIPTDIR"
                        Me.LocalScriptDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                    Case "RELDIR"
                        Me.RelDir = GetStr(GetVal(dr("ENVIRONMENTATTRBVALUE")))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsEnvironment LoadMe")
            Return False
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

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
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

            'cnn.Open()
            cmd = cnn.CreateCommand

            sql = "Select ENVIRONMENTNAME from " & Me.Project.tblEnvironments & " where PROJECTNAME='" & Me.Project.ProjectName & "'"

            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("ENVIRONMENTNAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new Environment name you have chosen already exists" & (Chr(13)) & "in this Project, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsEnv ValidateNewObject", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property EnvironmentName() As String
        Get
            Return m_EnvironmentName
        End Get
        Set(ByVal Value As String)
            m_EnvironmentName = Value
        End Set
    End Property

    Public Property EnvironmentDescription() As String
        Get
            Return m_EnvironmentDescription
        End Get
        Set(ByVal Value As String)
            m_EnvironmentDescription = Value
        End Set
    End Property

    Public Property LocalCDir() As String
        Get
            Return m_LocalCDir
        End Get
        Set(ByVal Value As String)
            m_LocalCDir = Value
        End Set
    End Property

    Public Property LocalDBDDir() As String
        Get
            Return m_LocalDBDDir
        End Get
        Set(ByVal Value As String)
            m_LocalDBDDir = Value
        End Set
    End Property

    Public Property LocalDTDDir() As String
        Get
            Return m_LocalDTDDir
        End Get
        Set(ByVal Value As String)
            m_LocalDTDDir = Value
        End Set
    End Property

    Public Property LocalDDLDir() As String
        Get
            Return m_LocalDDLDir
        End Get
        Set(ByVal Value As String)
            m_LocalDDLDir = Value
        End Set
    End Property

    Public Property LocalDMLDir() As String
        Get
            Return m_LocalDMLDir
        End Get
        Set(ByVal Value As String)
            m_LocalDMLDir = Value
        End Set
    End Property

    Public Property LocalCobolDir() As String
        Get
            Return m_LocalCobolDir
        End Get
        Set(ByVal Value As String)
            m_LocalCobolDir = Value
        End Set
    End Property

    Public Property LocalScriptDir() As String
        Get
            Return m_LocalScriptDir
        End Get
        Set(ByVal Value As String)
            m_LocalScriptDir = Value
        End Set
    End Property

    Public Property DefaultStrDir() As String
        Get
            Return m_DefaultStrDir
        End Get
        Set(ByVal value As String)
            m_DefaultStrDir = value
        End Set
    End Property

    Public Property RelDir() As String
        Get
            Return m_RelDir
        End Get
        Set(ByVal value As String)
            m_RelDir = value
        End Set
    End Property

    Public Property PathChange() As Boolean
        Get
            Return m_PathChange
        End Get
        Set(ByVal value As Boolean)
            m_PathChange = value
        End Set
    End Property

    Public Property OldDMLobj() As clsDMLinfo
        Get
            Return m_OldDMLobj
        End Get
        Set(ByVal value As clsDMLinfo)
            m_OldDMLobj = value
        End Set
    End Property

#End Region

#Region "Methods"

    '/// added 6/5/07 by tk
    Function SetStructureDir() As Boolean

        Try
            Application.UserAppDataRegistry.SetValue("DefaultStrDir", Me.DefaultStrDir)
        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function GetStructureDir() As Boolean

        Try
            If Application.UserAppDataRegistry.GetValue("DefaultStrDir") IsNot Nothing Then
                Me.DefaultStrDir = Application.UserAppDataRegistry.GetValue("DefaultStrDir").ToString()
            End If

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

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

            'cmd.Connection = cnn

            For i As Integer = 0 To 8
                Select Case i
                    Case 0
                        Attrib = "DEFAULTSTRDIR"
                        Value = Me.DefaultStrDir
                    Case 1
                        Attrib = "LOCALCDIR"
                        Value = Me.LocalCDir
                    Case 2
                        Attrib = "LOCALCOBOLDIR"
                        Value = Me.LocalCobolDir
                    Case 3
                        Attrib = "LOCALDDLDIR"
                        Value = Me.LocalDDLDir
                    Case 4
                        Attrib = "LOCALDMLDIR"
                        Value = Me.LocalDMLDir
                    Case 5
                        Attrib = "LOCALDTDDIR"
                        Value = Me.LocalDTDDir
                    Case 6
                        Attrib = "LOCALDBDDIR"
                        Value = Me.LocalDBDDir
                    Case 7
                        Attrib = "LOCALSCRIPTDIR"
                        Value = Me.LocalScriptDir
                    Case 8
                        Attrib = "RELDIR"
                        Value = Me.RelDir
                End Select

                sql = "INSERT INTO " & Me.Project.tblEnvironmentsATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,ENVIRONMENTATTRB,ENVIRONMENTATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & "','" & FixStr(Me.EnvironmentName) & _
                "','" & FixStr(Attrib) & "','" & FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsEnvironment InsertATTR")
            Return False
        End Try

    End Function

    Function DeleteATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As New Odbc.OdbcCommand
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

            'cmd.Connection = cnn

            sql = "DELETE FROM " & Me.Project.tblEnvironmentsATTR & " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.EnvironmentName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsEnvironment DeleteATTR")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
