Public Class clsProject
    Implements INode

    Private m_ProjectName As String = ""
    Private m_ProjectVersion As String = ""
    Private m_SecurityAttr As String = ""

    Private m_ProjectMetaVersion As enumMetaVersion = enumMetaVersion.V3
    Private m_ProjectMetaDSN As String = ""
    Private m_ProjectMetaDSNUID As String
    Private m_ProjectMetaDSNPWD As String

    Private m_ProjectCreationDate As String = ""
    Private m_ProjectCustomerName As String = ""
    Private m_ProductVersion As String = ""
    Private m_ProjectDesc As String = ""
    Private m_ProjectLastUpdated As String = ""

    Private m_LoginReq As Boolean
    Private m_SchemaReq As Boolean
    Private m_IsModified As Boolean
    Private m_ObjParent As INode
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_MapListPath As String
    Private m_ChgVer As Boolean = False

    '/// For dynamic Table Names
    Private m_ODBCtype As enumODBCtype
    Private m_tablePrefix As String = ""

    Private m_connections As String = "CONNECTIONS"
    Private m_connectionsATTR As String = "CONNECTIONSATTR" 'newV3

    'Private m_structures As String = "STRUCTURES"  'old
    Private m_descriptions As String = "DESCRIPTIONS" 'newRV3
    Private m_descriptionsATTR As String = "DESCRIPTIONSATTR" 'newV3

    'Private m_structfields As String = "STRUCTFIELDS" 'old
    Private m_descriptionfields As String = "DESCRIPTIONFIELDS" 'newRV3
    Private m_descriptionfieldsATTR As String = "DESCRIPTIONFIELDSATTR" 'newV3 ***************

    'Private m_structsel As String = "STRUCTSEL" 'old
    Private m_descriptionselect As String = "DESCRIPTIONSELECT" 'newRV3

    'Private m_strselfields As String = "STRSELFIELDS" 'old
    Private m_descriptionselfields As String = "DESCRIPTSELFIELDS" 'newRV3

    Private m_datastores As String = "DATASTORES"
    Private m_datastoresATTR As String = "DATASTORESATTR" 'newV3

    Private m_dsselections As String = "DSSELECTIONS"
    Private m_dsselfields As String = "DSSELFIELDS"
    Private m_dsselfieldsATTR As String = "DSSELFIELDSATTR" 'newV3  ************************

    Private m_engines As String = "ENGINES"
    Private m_enginesATTR As String = "ENGINESATTR" 'newV3

    Private m_environments As String = "ENVIRONMENTS"
    Private m_environmentsATTR As String = "ENVIRONMENTSATTR" 'newV3

    Private m_projects As String = "PROJECTS"
    Private m_projectsATTR As String = "PROJECTATTR" 'newV3

    Private m_systems As String = "SYSTEMS"
    Private m_systemsATTR As String = "SYSTEMSATTR" 'newV3

    Private m_taskds As String = "TASKDS"
    Private m_taskmap As String = "TASKMAP"
    Private m_taskmapATTR As String = "TASKMAPATTR" 'newV3
    Private m_tasks As String = "TASKS"

    Private m_variables As String = "VARIABLES"
    Private m_variablesATTR As String = "VARIABLESATTR" 'newV3

    '//User Settings
    Public MainSeparatorX As Integer = 240
    Public Environments As New Collection
    Public Maplist As New ArrayList

    Private m_ENVexpanded As Boolean = True
    Private m_ENV_FOexpanded As Boolean = True
    Private m_CONNexpanded As Boolean = True
    Private m_STRexpanded As Boolean = True
    '/// individual Structure folder types
    Private m_COBOLexpanded As Boolean = True
    Private m_COBOLIMSexpanded As Boolean = True
    Private m_Cexpanded As Boolean = True
    Private m_DDLexpanded As Boolean = True
    Private m_DMLexpanded As Boolean = True
    Private m_XMLDTDexpanded As Boolean = True

    Private m_VARtopExpanded As Boolean = True
    Private m_SYS_FOexpanded As Boolean = True
    Private m_SYSexpanded As Boolean = True
    Private m_ENG_FOexpanded As Boolean = True
    Private m_ENGexpanded As Boolean = True
    Private m_SRCexpanded As Boolean = True
    Private m_SRCselExpanded As Boolean = False
    Private m_TGTexpanded As Boolean = True
    Private m_VARexpanded As Boolean = True
    Private m_PROCexpanded As Boolean = True
    Private m_MAINexpanded As Boolean = True

    Private m_IsLoaded As Boolean = False


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
            Return (Me)
        End Get
        Set(ByVal Value As INode)
            m_ObjParent = Value
        End Set
    End Property

    '// OverHauled By TKarasch   April May  07
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsProject
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            obj.ProjectName = Me.ProjectName
            obj.ProjectVersion = Me.ProjectVersion
            obj.ProjectCreationDate = Me.ProjectCreationDate
            obj.ProjectMetaDSN = Me.ProjectMetaDSN
            obj.ProductVersion = Me.ProductVersion
            obj.ProjectDesc = Me.ProjectDesc
            obj.IsModified = Me.IsModified
            obj.SeqNo = Me.SeqNo

            obj.Parent = NewParent 'Me.Parent

            '//clone all environments under project
            If Cascade = True Then
                Dim newenv As clsEnvironment
                For Each env As clsEnvironment In Me.Environments
                    newenv = env.Clone(obj, True, cmd)
                    AddToCollection(obj.Environments, newenv, newenv.GUID)
                Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsProject Clone")
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
            Key = Me.Text
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = ProjectName
        End Get
        Set(ByVal Value As String)
            ProjectName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_PROJECT
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Return Me
        End Get
    End Property

    '//This method will save Project to Database and to file as well
    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim sql As String = ""
        Dim cmd As New Odbc.OdbcCommand

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            
            sql = "Update " & Me.Project.tblProjects & " set PROJECTNAME='" & FixStr(Me.ProjectName) & _
            "',PROJECTDESCRIPTION='" & FixStr(Me.ProjectDesc) & "',SECURITYATTR='" & FixStr(Me.SecurityATTR) & _
            "' where ProjectName=" & Me.GetQuotedText

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Save = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsProject Save", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            If Me.ProjectMetaVersion = enumMetaVersion.V2 Then
                '//When we add new record we need to find Unique Database Record ID.
                sql = "INSERT INTO " & Me.Project.tblProjects & "(ProjectName,Description,Version,cDate) " & _
                " Values('" & FixStr(ProjectName) & "','" & FixStr(ProjectDesc) & "','" _
                & FixStr(ProjectVersion) & "','" & FixStr(Now().ToString) & "')"
            Else
                sql = "INSERT INTO " & Me.Project.tblProjects & "(ProjectName,ProjectDescription,SecurityATTR) " & _
                "Values('" & FixStr(Me.ProjectName) & "','" & FixStr(Me.ProjectDesc) & "','" & FixStr(Me.SecurityATTR) & "')"

                Me.DeleteATTR(cmd)
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
                For Each child As clsEnvironment In Environments
                    child.AddNew(True)
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsProject AddNew", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// Function added by KS and TK 11/3/06
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            '//When we add new record we need to find Unique Database Record ID.
            If Me.ProjectMetaVersion = enumMetaVersion.V2 Then
                '//When we add new record we need to find Unique Database Record ID.
                sql = "INSERT INTO " & Me.Project.tblProjects & "(ProjectName,Description,Version,cDate) " & _
                " Values('" & FixStr(ProjectName) & "','" & FixStr(ProjectDesc) & "','" _
                & FixStr(ProjectVersion) & "','" & FixStr(Now()) & "')"
            Else
                sql = "INSERT INTO " & Me.Project.tblProjects & "(ProjectName,ProjectDescription,SecurityATTR) " & _
                "Values('" & FixStr(Me.ProjectName) & "','" & FixStr(Me.ProjectDesc) & "','" & FixStr(Me.SecurityATTR) & "')"

                Me.DeleteATTR(cmd)
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
                Dim child As INode
                For Each child In Environments
                    child.AddNew(True)
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsProject AddNew", sql)
            Return False
        End Try

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '//first delete all child objects
            If Cascade = True Then
                Dim o As INode
                For Each o In Me.Environments
                    o.Delete(cmd, cnn, Cascade, False)
                Next
                RemoveFromCollection(Me.Environments, "") '//clear collection
            End If

            If Me.ProjectMetaVersion = enumMetaVersion.V2 Then
                '//When we add new record we need to find Unique Database Record ID.
                sql = "DELETE  FROM " & Me.Project.tblProjects & " WHERE PROJECTNAME='" & Me.ProjectName & "'"
            Else
                Me.DeleteATTR(cmd)
                sql = "DELETE  FROM " & Me.Project.tblProjects & " WHERE PROJECTNAME='" & Me.ProjectName & "'"
            End If
            
            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsProject Delete", sql)
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

            If cnn Is Nothing Then
                cnn = New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
                cnn.Open()
            End If

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "SELECT PROJECTNAME,PROJECTATTRB,PROJECTATTRBVALUE FROM " & Me.Project.tblProjectsATTR & _
            " WHERE PROJECTNAME='" & FixStr(Me.ProjectName) & "'"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetVal(dr("PROJECTATTRB"))
                Select Case Attrib
                    Case "CDATE"
                        Me.ProjectCreationDate = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "CUSTNAME"
                        Me.ProjectCustomerName = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "LASTUPDATED"
                        Me.ProjectLastUpdated = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "MAINSEP"
                        Me.MainSeparatorX = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "VERSION"
                        Me.ProjectVersion = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "ENV_FOexpanded"
                        Me.ENV_FOexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "ENVexpanded"
                        Me.ENVexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "CONNexpanded"
                        Me.CONNexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "STRexpanded"
                        Me.STRexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "COBOLexpanded"
                        Me.COBOLexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "COBOLIMSexpanded"
                        Me.COBOLIMSexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "Cexpanded"
                        Me.Cexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "XMLDTDexpanded"
                        Me.XMLDTDexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "DDLexpanded"
                        Me.DDLexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "DMLexpanded"
                        Me.DMLexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "SYS_FOexpanded"
                        Me.SYS_FOexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "SYSexpanded"
                        Me.SYSexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "ENG_FOexpanded"
                        Me.ENG_FOexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "ENGexpanded"
                        Me.ENGexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "SRCexpanded"
                        Me.SRCexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "SRCselExpanded"
                        Me.SRCselExpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "TGTexpanded"
                        Me.TGTexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "VARexpanded"
                        Me.VARexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "PROCexpanded"
                        Me.PROCexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                    Case "MAINexpanded"
                        Me.MAINexpanded = GetVal(dr("PROJECTATTRBVALUE"))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsProject LoadMe")
            Return False
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

    Public Property IsRenamed() As Boolean Implements INode.IsRenamed
        Get
            Return m_IsRenamed
        End Get
        Set(ByVal Value As Boolean)
            m_IsRenamed = Value
        End Set
    End Property

    '/// Added By TKarasch  Spring 07 for Rename and ODBC openProj
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
            sql = "Select * from " & Me.tblProjects
            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("PROJECTNAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    If InReg = False Then
                        MsgBox("The new Project name you have chosen already exists," & (Chr(13)) & "Please enter a different Project Name", MsgBoxStyle.Information, MsgTitle)
                    End If
                    Exit While
                End If
            End While
            dr.Close()

            If InReg = True Then
                Return Not NameValid
            Else
                Return NameValid
            End If

        Catch ex As Exception
            'MsgBox("The ODBC DSN of your last opened project has been removed, please select a datasource from the DSN list")
            LogError(ex, ex.Message, NewName, , False)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property ProjectDesc() As String
        Get
            Return m_ProjectDesc
        End Get
        Set(ByVal Value As String)
            m_ProjectDesc = Value
        End Set
    End Property

    Public Property ProjectName() As String
        Get
            Return m_ProjectName
        End Get
        Set(ByVal Value As String)
            m_ProjectName = Value
        End Set
    End Property

    Public Property SecurityATTR() As String
        Get
            Return m_SecurityAttr
        End Get
        Set(ByVal Value As String)
            m_SecurityAttr = Value
        End Set
    End Property

    Public Property ProjectVersion() As String
        Get
            Return m_ProjectVersion
        End Get
        Set(ByVal Value As String)
            m_ProjectVersion = Value
        End Set
    End Property

    Public Property ProductVersion() As String
        Get
            Return m_ProductVersion
        End Get
        Set(ByVal Value As String)
            m_ProductVersion = Value
        End Set
    End Property

    '///New for V2 V3 differentiation
    Public Property ProjectMetaVersion() As enumMetaVersion
        Get
            Return m_ProjectMetaVersion
        End Get
        Set(ByVal value As enumMetaVersion)
            m_ProjectMetaVersion = value
        End Set
    End Property

    Public Property ChangeVersion() As Boolean
        Get
            Return m_ChgVer
        End Get
        Set(ByVal value As Boolean)
            m_ChgVer = value
        End Set
    End Property

    Public Property ProjectCreationDate() As String
        Get
            Return m_ProjectCreationDate
        End Get
        Set(ByVal Value As String)
            m_ProjectCreationDate = Value
        End Set
    End Property

    Public Property ProjectCustomerName() As String
        Get
            Return m_ProjectCustomerName
        End Get
        Set(ByVal value As String)
            m_ProjectCustomerName = value
        End Set
    End Property

    Public Property ProjectLastUpdated() As String
        Get
            Return m_ProjectLastUpdated
        End Get
        Set(ByVal value As String)
            m_ProjectLastUpdated = value
        End Set
    End Property

    '/// ALL Properties below Created or Changed by TK May 07 for 
    '/// dynamic table naming for Oracle and DB2
    Public Property ProjectMetaDSNUID() As String
        Get
            Return m_ProjectMetaDSNUID
        End Get
        Set(ByVal Value As String)
            m_ProjectMetaDSNUID = Value
        End Set
    End Property

    Public Property ProjectMetaDSNPWD() As String
        Get
            Return m_ProjectMetaDSNPWD
        End Get
        Set(ByVal Value As String)
            m_ProjectMetaDSNPWD = Value
        End Set
    End Property

    Public Property ProjectMetaDSN() As String
        Get
            Return m_ProjectMetaDSN
        End Get
        Set(ByVal Value As String)
            m_ProjectMetaDSN = Value
        End Set
    End Property

    Public ReadOnly Property MetaConnectionString() As String
        Get
            Return "DSN=" & ProjectMetaDSN & "; UID=" & ProjectMetaDSNUID & "; PWD=" & ProjectMetaDSNPWD
        End Get
    End Property

    Public Property MapListPath() As String
        Get
            Return m_MapListPath
        End Get
        Set(ByVal value As String)
            value = Strings.Replace(value, "\\", "\", , , CompareMethod.Text)
            m_MapListPath = value
        End Set
    End Property

    Public Property ODBCtype() As Integer
        Get
            Return m_ODBCtype
        End Get
        Set(ByVal value As Integer)
            m_ODBCtype = value
        End Set
    End Property

    Public Property TablePrefix() As String
        Get
            Return m_tablePrefix
        End Get
        Set(ByVal value As String)
            m_tablePrefix = value
        End Set
    End Property

    Public Property LoginReq() As Boolean
        Get
            Return m_LoginReq
        End Get
        Set(ByVal value As Boolean)
            m_LoginReq = value
        End Set
    End Property

    Public Property SchemaReq() As Boolean
        Get
            Return m_SchemaReq
        End Get
        Set(ByVal value As Boolean)
            m_SchemaReq = value
        End Set
    End Property

    Public ReadOnly Property tblConnections() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_connections
            Else
                Return m_connections
            End If
        End Get
    End Property

    Public ReadOnly Property tblConnectionsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_connectionsATTR
            Else
                Return m_connectionsATTR
            End If
        End Get
    End Property

    'Public ReadOnly Property tblStructures() As String
    '    Get
    '        If TablePrefix <> "" Then
    '            Return TablePrefix & "." & m_structures
    '        Else
    '            Return m_structures
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property tblDescriptions() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptions
            Else
                Return m_descriptions
            End If
        End Get
    End Property

    Public ReadOnly Property tblDescriptionsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptionsATTR
            Else
                Return m_descriptionsATTR
            End If
        End Get
    End Property

    'Public ReadOnly Property tblStructFields() As String
    '    Get
    '        If TablePrefix <> "" Then
    '            Return TablePrefix & "." & m_structfields
    '        Else
    '            Return m_structfields
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property tblDescriptionFields() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptionfields
            Else
                Return m_descriptionfields
            End If
        End Get
    End Property

    Public ReadOnly Property tblDescriptionFieldsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptionfieldsATTR
            Else
                Return m_descriptionfieldsATTR
            End If
        End Get
    End Property

    'Public ReadOnly Property tblStructSel() As String
    '    Get
    '        If TablePrefix <> "" Then
    '            Return TablePrefix & "." & m_structsel
    '        Else
    '            Return m_structsel
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property tblDescriptionSelect() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptionselect
            Else
                Return m_descriptionselect
            End If
        End Get
    End Property

    'Public ReadOnly Property tblStrSelFields() As String
    '    Get
    '        If TablePrefix <> "" Then
    '            Return TablePrefix & "." & m_strselfields
    '        Else
    '            Return m_strselfields
    '        End If
    '    End Get
    'End Property

    Public ReadOnly Property tblDescriptionSelFields() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_descriptionselfields
            Else
                Return m_descriptionselfields
            End If
        End Get
    End Property

    Public ReadOnly Property tblDatastores() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_datastores
            Else
                Return m_datastores
            End If
        End Get
    End Property

    Public ReadOnly Property tblDatastoresATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_datastoresATTR
            Else
                Return m_datastoresATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblDSselections() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_dsselections
            Else
                Return m_dsselections
            End If
        End Get
    End Property

    Public ReadOnly Property tblDSselFields() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_dsselfields
            Else
                Return m_dsselfields
            End If
        End Get
    End Property

    Public ReadOnly Property tblDSselFieldsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_dsselfieldsATTR
            Else
                Return m_dsselfieldsATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblEngines() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_engines
            Else
                Return m_engines
            End If
        End Get
    End Property

    Public ReadOnly Property tblEnginesATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_enginesATTR
            Else
                Return m_enginesATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblEnvironments() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_environments
            Else
                Return m_environments
            End If
        End Get
    End Property

    Public ReadOnly Property tblEnvironmentsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_environmentsATTR
            Else
                Return m_environmentsATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblProjects() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_projects
            Else
                Return m_projects
            End If
        End Get
    End Property

    Public ReadOnly Property tblProjectsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_projectsATTR
            Else
                Return m_projectsATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblSystems() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_systems
            Else
                Return m_systems
            End If
        End Get
    End Property

    Public ReadOnly Property tblSystemsATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_systemsATTR
            Else
                Return m_systemsATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblTaskDS() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_taskds
            Else
                Return m_taskds
            End If
        End Get
    End Property

    Public ReadOnly Property tblTaskMap() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_taskmap
            Else
                Return m_taskmap
            End If
        End Get
    End Property

    Public ReadOnly Property tblTaskMapATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_taskmapATTR
            Else
                Return m_taskmapATTR
            End If
        End Get
    End Property

    Public ReadOnly Property tblTasks() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_tasks
            Else
                Return m_tasks
            End If
        End Get
    End Property

    Public ReadOnly Property tblVariables() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_variables
            Else
                Return m_variables
            End If
        End Get
    End Property

    Public ReadOnly Property tblVariablesATTR() As String
        Get
            If TablePrefix <> "" Then
                Return TablePrefix & "." & m_variablesATTR
            Else
                Return m_variablesATTR
            End If
        End Get
    End Property

    Public Property ENV_FOexpanded() As Boolean
        Get
            Return m_ENV_FOexpanded
        End Get
        Set(ByVal value As Boolean)
            m_ENV_FOexpanded = value
        End Set
    End Property

    Public Property ENVexpanded() As Boolean
        Get
            Return m_ENVexpanded
        End Get
        Set(ByVal value As Boolean)
            m_ENVexpanded = value
        End Set
    End Property

    Public Property CONNexpanded() As Boolean
        Get
            Return m_CONNexpanded
        End Get
        Set(ByVal value As Boolean)
            m_CONNexpanded = value
        End Set
    End Property

    Public Property STRexpanded() As Boolean
        Get
            Return m_STRexpanded
        End Get
        Set(ByVal value As Boolean)
            m_STRexpanded = value
        End Set
    End Property

    Public Property COBOLexpanded() As Boolean
        Get
            Return m_COBOLexpanded
        End Get
        Set(ByVal value As Boolean)
            m_COBOLexpanded = value
        End Set
    End Property

    Public Property COBOLIMSexpanded() As Boolean
        Get
            Return m_COBOLIMSexpanded
        End Get
        Set(ByVal value As Boolean)
            m_COBOLIMSexpanded = value
        End Set
    End Property

    Public Property Cexpanded() As Boolean
        Get
            Return m_Cexpanded
        End Get
        Set(ByVal value As Boolean)
            m_Cexpanded = value
        End Set
    End Property

    Public Property DDLexpanded() As Boolean
        Get
            Return m_DDLexpanded
        End Get
        Set(ByVal value As Boolean)
            m_DDLexpanded = value
        End Set
    End Property

    Public Property DMLexpanded() As Boolean
        Get
            Return m_DMLexpanded
        End Get
        Set(ByVal value As Boolean)
            m_DMLexpanded = value
        End Set
    End Property

    Public Property XMLDTDexpanded() As Boolean
        Get
            Return m_XMLDTDexpanded
        End Get
        Set(ByVal value As Boolean)
            m_XMLDTDexpanded = value
        End Set
    End Property

    Public Property VARtopExpanded() As Boolean
        Get
            Return m_VARtopExpanded
        End Get
        Set(ByVal value As Boolean)
            m_VARtopExpanded = value
        End Set
    End Property

    Public Property SYS_FOexpanded() As Boolean
        Get
            Return m_SYS_FOexpanded
        End Get
        Set(ByVal value As Boolean)
            m_SYS_FOexpanded = value
        End Set
    End Property

    Public Property SYSexpanded() As Boolean
        Get
            Return m_SYSexpanded
        End Get
        Set(ByVal value As Boolean)
            m_SYSexpanded = value
        End Set
    End Property

    Public Property ENG_FOexpanded() As Boolean
        Get
            Return m_ENG_FOexpanded
        End Get
        Set(ByVal value As Boolean)
            m_ENG_FOexpanded = value
        End Set
    End Property

    Public Property ENGexpanded() As Boolean
        Get
            Return m_ENGexpanded
        End Get
        Set(ByVal value As Boolean)
            m_ENGexpanded = value
        End Set
    End Property

    Public Property SRCexpanded() As Boolean
        Get
            Return m_SRCexpanded
        End Get
        Set(ByVal value As Boolean)
            m_SRCexpanded = value
        End Set
    End Property

    Public Property SRCselExpanded() As Boolean
        Get
            Return m_SRCselExpanded
        End Get
        Set(ByVal value As Boolean)
            m_SRCselExpanded = value
        End Set
    End Property

    Public Property TGTexpanded() As Boolean
        Get
            Return m_TGTexpanded
        End Get
        Set(ByVal value As Boolean)
            m_TGTexpanded = value
        End Set
    End Property

    Public Property VARexpanded() As Boolean
        Get
            Return m_VARexpanded
        End Get
        Set(ByVal value As Boolean)
            m_VARexpanded = value
        End Set
    End Property

    Public Property PROCexpanded() As Boolean
        Get
            Return m_PROCexpanded
        End Get
        Set(ByVal value As Boolean)
            m_PROCexpanded = value
        End Set
    End Property

    Public Property MAINexpanded() As Boolean
        Get
            Return m_MAINexpanded
        End Get
        Set(ByVal value As Boolean)
            m_MAINexpanded = value
        End Set
    End Property

#End Region

#Region "Methods"

    '//New 3/16/2007 by TK
    Public Function SaveToRegistry() As Boolean

        Dim ReqLogin As String
        Dim ReqSchema As String
        Try
            ' Save the Last Project and Main Separator to the registry
            Application.UserAppDataRegistry.SetValue("ProjName", Me.ProjectName)
            Application.UserAppDataRegistry.SetValue("LastDSN", Me.ProjectMetaDSN)
            Application.UserAppDataRegistry.SetValue("LastDesc", Me.ProjectDesc)
            Application.UserAppDataRegistry.SetValue("LastVer", Me.ProjectVersion)
            Application.UserAppDataRegistry.SetValue("MainSeparator", Me.MainSeparatorX)
            Application.UserAppDataRegistry.SetValue("ODBCType", Me.ODBCtype)
            Application.UserAppDataRegistry.SetValue("TablePrefix", Me.TablePrefix)
            Application.UserAppDataRegistry.SetValue("CreationDate", Me.ProjectCreationDate)
            Application.UserAppDataRegistry.SetValue(Me.ProjectName & "MapListPath", Me.MapListPath)
            Application.UserAppDataRegistry.SetValue("MetaVer", Me.ProjectMetaVersion.ToString)
            If Me.LoginReq = True Then
                ReqLogin = "true"
            Else
                ReqLogin = "false"
            End If
            Application.UserAppDataRegistry.SetValue("LoginReq", ReqLogin)
            If Me.SchemaReq = True Then
                ReqSchema = "true"
            Else
                ReqSchema = "false"
            End If
            Application.UserAppDataRegistry.SetValue("SchemaReq", ReqSchema)

            Return True

        Catch ex As Exception
            LogError(ex, ex.Message, "clsProject SaveToRegistry")
            Return False
        End Try

    End Function

    '//New 3/16/2007 by TK
    Public Function RetrieveFromRegistry() As Boolean

        Dim ReqLogin As String
        Dim ReqSchema As String
        Dim ver As String

        Try
            ' Get the Last Project and Main Separator from the registry.
            If Not (Application.UserAppDataRegistry.GetValue("ProjName") Is Nothing) Then
                Me.ProjectName = Application.UserAppDataRegistry.GetValue("ProjName").ToString()
                Me.ProjectMetaDSN = Application.UserAppDataRegistry.GetValue("LastDSN").ToString
                Me.ProjectDesc = Application.UserAppDataRegistry.GetValue("LastDesc").ToString
                Me.ProjectVersion = Application.UserAppDataRegistry.GetValue("LastVer").ToString
                Me.MainSeparatorX = CInt(Application.UserAppDataRegistry.GetValue("MainSeparator").ToString)
                Me.ODBCtype = CInt(Application.UserAppDataRegistry.GetValue("ODBCtype").ToString)
                Me.TablePrefix = Application.UserAppDataRegistry.GetValue("TablePrefix").ToString
                Me.ProjectCreationDate = Application.UserAppDataRegistry.GetValue("CreationDate").ToString
                If Application.UserAppDataRegistry.GetValue("MetaVer") IsNot Nothing Then
                    ver = Application.UserAppDataRegistry.GetValue("MetaVer").ToString()
                    If ver = "V2" Or ver = "0" Then
                        Me.ProjectMetaVersion = enumMetaVersion.V2
                    Else
                        Me.ProjectMetaVersion = enumMetaVersion.V3
                    End If
                Else
                    Me.ProjectMetaVersion = enumMetaVersion.V2
                End If
                If Application.UserAppDataRegistry.GetValue(Me.ProjectName & "MapListPath") IsNot Nothing Then
                    Me.MapListPath = Application.UserAppDataRegistry.GetValue(Me.ProjectName & "MapListPath").ToString
                    Me.MapListPath = Strings.Replace(Me.MapListPath, "\\", "\", , , CompareMethod.Text)
                Else
                    Me.MapListPath = ""
                End If
                If Application.UserAppDataRegistry.GetValue("LoginReq") IsNot Nothing Then
                    ReqLogin = Application.UserAppDataRegistry.GetValue("LoginReq").ToString
                Else
                    ReqLogin = ""
                End If
                If ReqLogin = "true" Then
                    Me.LoginReq = True
                Else
                    Me.LoginReq = False
                End If
                If Application.UserAppDataRegistry.GetValue("SchemaReq") IsNot Nothing Then
                    ReqSchema = Application.UserAppDataRegistry.GetValue("SchemaReq").ToString
                Else
                    ReqSchema = ""
                End If
                If ReqSchema = "true" Then
                    Me.SchemaReq = True
                Else
                    Me.SchemaReq = False
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, ex.Message, "clsProject RetrieveFromRegistry")
            Return False
        End Try

    End Function

    Function InsertATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

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

            cmd.Connection = cnn

            For i As Integer = 0 To 24
                Select Case i
                    Case 0
                        Attrib = "CDATE"
                        Value = Me.ProjectCreationDate
                    Case 1
                        Attrib = "CUSTNAME"
                        Value = Me.ProjectCustomerName
                    Case 2
                        Attrib = "LASTUPDATED"
                        Value = Now.ToString
                    Case 3
                        Attrib = "MAINSEP"
                        Value = Me.MainSeparatorX
                    Case 4
                        Attrib = "VERSION"
                        Value = Me.ProjectVersion
                    Case 5
                        Attrib = "ENV_FOexpanded"
                        Value = Me.ENV_FOexpanded
                    Case 6
                        Attrib = "ENVexpanded"
                        Value = Me.ENVexpanded
                    Case 7
                        Attrib = "CONNexpanded"
                        Value = Me.CONNexpanded
                    Case 8
                        Attrib = "STRexpanded"
                        Value = Me.STRexpanded
                    Case 9
                        Attrib = "COBOLexpanded"
                        Value = Me.COBOLexpanded
                    Case 10
                        Attrib = "COBOLIMSexpanded"
                        Value = Me.COBOLIMSexpanded
                    Case 11
                        Attrib = "Cexpanded"
                        Value = Me.Cexpanded
                    Case 12
                        Attrib = "XMLDTDexpanded"
                        Value = Me.XMLDTDexpanded
                    Case 13
                        Attrib = "DDLexpanded"
                        Value = Me.DDLexpanded
                    Case 14
                        Attrib = "DMLexpanded"
                        Value = Me.DMLexpanded
                    Case 15
                        Attrib = "SYS_FOexpanded"
                        Value = Me.SYS_FOexpanded
                    Case 16
                        Attrib = "SYSexpanded"
                        Value = Me.SYSexpanded
                    Case 17
                        Attrib = "ENG_FOexpanded"
                        Value = Me.ENG_FOexpanded
                    Case 18
                        Attrib = "ENGexpanded"
                        Value = Me.ENGexpanded
                    Case 19
                        Attrib = "SRCexpanded"
                        Value = Me.SRCexpanded
                    Case 20
                        Attrib = "SRCselExpanded"
                        Value = Me.SRCselExpanded
                    Case 21
                        Attrib = "TGTexpanded"
                        Value = Me.TGTexpanded
                    Case 22
                        Attrib = "VARexpanded"
                        Value = Me.VARexpanded
                    Case 23
                        Attrib = "PROCexpanded"
                        Value = Me.PROCexpanded
                    Case 24
                        Attrib = "MAINexpanded"
                        Value = Me.MAINexpanded
                    Case Else

                End Select

                sql = "INSERT INTO " & Me.Project.tblProjectsATTR & "(PROJECTNAME,PROJECTATTRB,PROJECTATTRBVALUE) " & _
                "Values('" & FixStr(Me.ProjectName) & "','" & FixStr(Attrib) & "','" & FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsProject InsertATTR")
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

            cmd.Connection = cnn

            sql = "DELETE FROM " & Me.Project.tblProjectsATTR & " WHERE PROJECTNAME=" & Quote(Me.ProjectName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsProject DeleteATTR")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class