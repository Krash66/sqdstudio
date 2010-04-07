Option Compare Text
Public Class clsDSSelection
    Implements INode
    '/// new class by Tom Karasch, December 2006 & Jan 2007 for Datastore Selection Object.
    '/// This is to separate Structures from the Datastore. Once a structure is
    '/// added to the datastore it is "cloned" into a DatastoreSelection.
    '/// The new DSselection can then be manipulated and changed specifically
    '/// for each Datastore without effecting the Original Structure.
    '/// Some of the features this allows are: Adding of changable field
    '/// properties, such as key and foreign key. It also allows parent child 
    '/// relationships between different Datastore Selections, based on foreign keys.
    '/// The Structure, Structure Selection, and Datastore Selection remain
    '/// Independent, yet related to each other referencially.

    Private m_ObjDatastore As clsDatastore '// the datastore
    Private m_ObjParent As INode '//clsDataStore or clsDSelection
    Private m_ObjStructure As clsStructure '//structure of the DS Selection
    Private m_ObjSelection As clsStructureSelection '//Structure Selection of the DSSelection
    Private m_SelectionName As String = ""
    Private m_SelectionDescription As String = ""
    Private m_IsRenamed As Boolean = False
    Private m_IsModified As Boolean = False
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String = ""
    Private m_SeqNo As Integer = 0
    Private m_IsMapped As Boolean = False
    Private m_ObjEngine As clsEngine
    Private m_Environment As clsEnvironment

    Public DSSelectionFields As New ArrayList '//Array of Fields selected with in structure
    Public OldDSSelectionFields As New ArrayList '//Array of Old fields
    Public ObjDSSelections As New ArrayList '// Array of child DSSelections
    Public OldObjDSSelections As New ArrayList '// Array of old child DSSelections

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
        '/// Fixed By TKarasch May 07
        Try
            Dim obj As New clsDSSelection
            Dim cmd As Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadItems(False, False, cmd)

            obj.Engine = Me.Engine
            obj.Environment = Me.Environment
            obj.SelectionName = Me.SelectionName
            obj.SelectionDescription = Me.SelectionDescription
            obj.IsModified = Me.IsModified
            obj.SeqNo = Me.SeqNo
            obj.Text = Me.Text
            obj.Parent = NewParent 'Me.Parent

            '///////MustUpdate to new Obj Pointers in datastore
            obj.ObjStructure = Me.ObjStructure
            obj.ObjSelection = Me.ObjSelection

            Dim fld As clsField
            Dim retfld As clsField

            '//Clone the Fields in the DS Selection
            For Each fld In Me.DSSelectionFields
                retfld = SearchDSFieldByName(Me, fld.FieldName, Me.ObjStructure.StructureName, Me.ObjStructure.Environment.EnvironmentName, Me.Project.ProjectName)
                If retfld Is Nothing Then
                    Throw (New Exception("Selection Field [" & fld.Text & "] is not found in parent structure [" & fld.Struct.Text & "]"))
                Else
                    Dim NewFld As clsField
                    NewFld = New clsField
                    NewFld = fld.Clone(obj, True, cmd)
                    NewFld.Parent = obj
                    obj.DSSelectionFields.Add(NewFld)
                End If
            Next

            '/// Now Clone all Child DS Selections in this Selection
            '// **** This will be recursive if dsSel has Child dsSel
            Dim newSel As clsDSSelection
            For Each DSSel As clsDSSelection In Me.ObjDSSelections
                newSel = DSSel.Clone(obj, True, cmd)
                obj.ObjDSSelections.Add(newSel)
            Next

            If Cascade = True Then
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsDSSelection Clone")
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
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.ObjDatastore.Engine.ObjSystem.Text & KEY_SAP & Me.Engine.Text & KEY_SAP & Me.ObjDatastore.Text & KEY_SAP & Me.Text
            Else
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
            End If

        End Get
    End Property

    Public Property Text() As String Implements INode.Text
        Get
            Text = SelectionName
        End Get
        Set(ByVal Value As String)
            SelectionName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            If (Parent.Type = NODE_SOURCEDATASTORE Or Parent.Type = NODE_SOURCEDSSEL) Then
                Type = NODE_SOURCEDSSEL
            Else
                Type = NODE_TARGETDSSEL
            End If
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.ObjDatastore.Project
            Return p
        End Get
    End Property

    '/// save is done in Datastore
    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        Return True

    End Function

    '//// this done in the Datastore
    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    '//// this done in the Datastore
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    '//// this is only used if Structure or structure selection is deleted
    '/// if DSselections are deleted from the datastore, delete is done in the 
    '/// datastore class
    '/// Fixed by TKarasch   May 07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '/// delete from Datastore Selection Fields Table
            sql = "Delete From " & Me.Project.tblDSselFields & " where SelectionName='" & Me.ObjSelection.SelectionName & "' AND EnvironmentName=" & Me.ObjDatastore.Engine.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '/// delete from Datastore Selections Table
            sql = "Delete From " & Me.Project.tblDSselections & " where SelectionName='" & Me.ObjSelection.SelectionName & "' AND EnvironmentName=" & Me.ObjDatastore.Engine.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '// Remove Selections from Task Mappings
            If Cascade = True And RemoveFromParentCollection = False Then
                If RemoveFromTaskMappings(cmd, cnn, Cascade, RemoveFromParentCollection) = False Then
                    Return False
                    Exit Function
                End If
            End If
            
            '/// Last ... remove from the Datastore Object
            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                Dim ObjDS As clsDatastore = Me.ObjDatastore
                ObjDS.ObjSelections.Remove(Me)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, sql, "clsDSSelection Delete", , False)
            Return False
        End Try

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        Dim cmd As System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.DataRow
        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim sql As String = ""
        Dim i As Integer

        Try
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If
            '//check if already loaded ?
            If TreeLode = True Then
                Exit Function
            End If
            If Reload = False Then
                If Me.DSSelectionFields.Count > 0 Then Exit Function
            End If

            'cnn.Open()
            'cmd = cmd.Connection.CreateCommand

            Me.ObjDatastore.LoadItems()
            Me.ObjStructure.LoadItems()
            Me.DSSelectionFields.Clear()

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                If Me.Engine IsNot Nothing Then
                    sql = "Select dssf.PROJECTNAME,dssf.ENVIRONMENTNAME,dssf.StructureName,dssf.FIELDNAME from " & Me.Project.tblDSselFields & " dssf " & _
                    "inner join " & Me.Project.tblDSselections & " dss " & _
                    "on dssf.projectname=dss.projectname " & _
                    "AND dssf.environmentname=dss.environmentname " & _
                    "AND dssf.SystemName=dss.SystemName " & _
                    "AND dssf.EngineName=dss.EngineName " & _
                    "AND dssf.DatastoreName=dss.DatastoreName " & _
                    "AND dssf.StructureName=dss.StructureName " & _
                    "AND dssf.SelectionName=dss.SelectionName " & _
                    "AND dssf.DSDirection=dss.DSDirection " & _
                    "AND dssf.Parent=dss.Parent " & _
                    "where  dssf.ProjectName=" & Me.Project.GetQuotedText & _
                    " AND dssf.EnvironmentName=" & Me.Environment.GetQuotedText & _
                    " AND dssf.SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                    " AND dssf.EngineName=" & Me.Engine.GetQuotedText & _
                    " AND dssf.DatastoreName = " & Me.ObjDatastore.GetQuotedText & _
                    " AND dssf.DSDirection=" & Quote(Me.ObjDatastore.DsDirection) & _
                    " AND dssf.SelectionName=" & Me.GetQuotedText & _
                    " AND dssf.StructureName=" & Me.ObjStructure.GetQuotedText & _
                    " AND dssf.Parent=" & Me.Parent.GetQuotedText & _
                    " ORDER BY dssf.SEQNO"
                Else
                    sql = "Select dssf.PROJECTNAME,dssf.ENVIRONMENTNAME,dssf.StructureName,dssf.FIELDNAME from " & Me.Project.tblDSselFields & " dssf " & _
                    "inner join " & Me.Project.tblDSselections & " dss " & _
                    "on dssf.projectname=dss.projectname " & _
                    "AND dssf.environmentname=dss.environmentname " & _
                    "AND dssf.DatastoreName=dss.DatastoreName " & _
                    "AND dssf.SelectionName=dss.SelectionName " & _
                    "AND dssf.StructureName=dss.StructureName " & _
                    "AND dssf.DSDirection=dss.DSDirection " & _
                    "AND dssf.Parent=dss.Parent " & _
                    "where dssf.ProjectName=" & Me.Project.GetQuotedText & _
                    " AND dssf.EnvironmentName=" & Me.Environment.GetQuotedText & _
                    " AND dssf.DatastoreName=" & Me.ObjDatastore.GetQuotedText & _
                    " AND dssf.SelectionName=" & Me.GetQuotedText & _
                    " AND dssf.StructureName=" & Me.ObjStructure.GetQuotedText & _
                    " AND dssf.DSDirection=" & Quote(Me.ObjDatastore.DsDirection) & _
                    " AND dssf.Parent=" & Me.Parent.GetQuotedText & _
                    " ORDER BY dssf.SEQNO"
                End If
            Else
                If Me.Engine IsNot Nothing Then
                    sql = "Select dssf.PROJECTNAME,dssf.ENVIRONMENTNAME,dssf.DESCRIPTIONNAME,dssf.FIELDNAME from " & Me.Project.tblDSselFields & " dssf " & _
                    "inner join " & Me.Project.tblDSselections & " dss " & _
                    "on dssf.projectname=dss.projectname " & _
                    "AND dssf.environmentname=dss.environmentname " & _
                    "AND dssf.SystemName=dss.SystemName " & _
                    "AND dssf.EngineName=dss.EngineName " & _
                    "AND dssf.DatastoreName=dss.DatastoreName " & _
                    "AND dssf.SelectionName=dss.SelectionName " & _
                    "AND dssf.DescriptionName=dss.DescriptionName " & _
                    "AND dssf.ParentName=dss.Parent " & _
                    "where dssf.ProjectName=" & Me.Project.GetQuotedText & _
                    " AND dssf.EnvironmentName=" & Me.Environment.GetQuotedText & _
                    " AND dssf.SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                    " AND dssf.EngineName=" & Me.Engine.GetQuotedText & _
                    " AND dssf.DatastoreName=" & Me.ObjDatastore.GetQuotedText & _
                    " AND dssf.SelectionName=" & Me.GetQuotedText & _
                    " AND dssf.DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                    " AND dssf.ParentName=" & Me.Parent.GetQuotedText & _
                    " ORDER BY dssf.SEQNO"
                Else
                    sql = "Select dssf.PROJECTNAME,dssf.ENVIRONMENTNAME,dssf.DESCRIPTIONNAME,dssf.FIELDNAME,dss.parent from " & Me.Project.tblDSselFields & " dssf " & _
                    "inner join " & Me.Project.tblDSselections & " dss " & _
                    " ON dssf.projectname=dss.projectname " & _
                    " AND dssf.environmentname=dss.environmentname " & _
                    " AND dssf.SystemName=dss.SystemName" & _
                    " AND dssf.EngineName=dss.EngineName" & _
                    " AND dssf.DatastoreName=dss.DatastoreName " & _
                    " AND dssf.SelectionName=dss.SelectionName " & _
                    " AND dssf.DescriptionName=dss.DescriptionName " & _
                    " AND dssf.Parentname=dss.parent " & _
                    " where dssf.ProjectName=" & Me.Project.GetQuotedText & _
                    " AND dssf.EnvironmentName=" & Me.Environment.GetQuotedText & _
                    " AND dssf.SystemName=" & Quote(DBNULL) & _
                    " AND dssf.EngineName=" & Quote(DBNULL) & _
                    " AND dssf.DatastoreName=" & Me.ObjDatastore.GetQuotedText & _
                    " AND dssf.SelectionName=" & Me.GetQuotedText & _
                    " AND dssf.DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                    " AND dssf.ParentName=" & Me.Parent.GetQuotedText & _
                    " ORDER BY dssf.SEQNO"
                End If
            End If


            ', dssf.fieldname
            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            dt = New System.Data.DataTable("temp2")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)
                Dim fld As New clsField
                If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    fld = SearchDSFieldByName(Me, dr("FieldName"), dr("StructureName"), dr("environmentname"), dr("projectname"))
                Else
                    fld = SearchDSFieldByName(Me, dr("FieldName"), dr("DescriptionName"), dr("environmentname"), dr("projectname"))
                End If

                If fld Is Nothing Then
                    Log("Field [" & dr("fieldname") & "] is missing in the parent structure - " & Me.ObjStructure.Text)
                Else
                    fld.Parent = Me
                    If Me.DSSelectionFields.Contains(fld) = False Then
                        DSSelectionFields.Add(fld)
                    End If
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsDSSelection LoadItems", sql)
            Return False
        Finally
            'cnn.Close()
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

    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        '/// since it is cloned from a structure selection, it has been checked already
        Return True

    End Function

#End Region

#Region "Properties"

    Public Property SelectionName() As String
        Get
            Return m_SelectionName
        End Get
        Set(ByVal Value As String)
            m_SelectionName = Value
        End Set
    End Property

    Public Property SelectionDescription() As String
        Get
            Return m_SelectionDescription
        End Get
        Set(ByVal Value As String)
            m_SelectionDescription = Value
        End Set
    End Property

    Public Property ObjDatastore() As clsDatastore
        Get
            Return m_ObjDatastore
        End Get
        Set(ByVal Value As clsDatastore)
            m_ObjDatastore = Value
        End Set
    End Property

    Public Property ObjStructure() As clsStructure
        Get
            Return m_ObjStructure
        End Get
        Set(ByVal Value As clsStructure)
            m_ObjStructure = Value
        End Set
    End Property

    Public Property ObjSelection() As clsStructureSelection
        Get
            Return m_ObjSelection
        End Get
        Set(ByVal Value As clsStructureSelection)
            m_ObjSelection = Value
        End Set
    End Property

    Public Property Engine() As clsEngine
        Get
            If m_ObjEngine Is Nothing Then
                m_ObjEngine = Me.ObjDatastore.Engine
            End If
            Return m_ObjEngine
        End Get
        Set(ByVal Value As clsEngine)
            m_ObjEngine = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            If m_Environment Is Nothing Then
                m_Environment = Me.ObjDatastore.Environment
            End If
            Return m_Environment
        End Get
        Set(ByVal Value As clsEnvironment)
            m_Environment = Value
        End Set
    End Property

    Public ReadOnly Property IsChildDSSelection() As Boolean
        Get
            If (Me.Parent.Type = NODE_SOURCEDATASTORE Or Me.Parent.Type = NODE_TARGETDATASTORE) Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public Property IsMapped() As Boolean
        Get
            Return m_IsMapped
        End Get
        Set(ByVal Value As Boolean)
            m_IsMapped = Value
        End Set
    End Property

#End Region

#Region "Methods"

    '/// sets all fields as modified so metadata updates parents correctly for DSSelection
    Public Function SetAllFieldsModified() As Boolean

        Dim i As Integer = 0
        Dim fld As clsField

        For i = 0 To Me.DSSelectionFields.Count - 1
            fld = Me.DSSelectionFields(i)
            fld.IsModified = True
        Next

    End Function

    '// Used when a structure or DSselection is deleted to maintain referential integrity in Metadata
    Function RemoveFromTaskMappings(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean

        Dim sql As String = ""

        Try
            '/// delete from Task Mappings Table (Source Parent is the Selection/Structure Name)
            sql = "Delete From " & Me.Project.tblTaskMap & " where SourceParent='" & Me.ObjSelection.SelectionName & "' AND EnvironmentName=" & Me.ObjDatastore.Engine.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '/// delete from Task Mappings Table (Target Parent is the Selection/Structure Name)
            sql = "Delete From " & Me.Project.tblTaskMap & " where TargetParent='" & Me.ObjSelection.SelectionName & "' AND EnvironmentName=" & Me.ObjDatastore.Engine.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            RemoveFromTaskMappings = True

        Catch ex As Exception
            LogError(ex, "clsDSSelection RemoveFromTaskMappings", sql, False)
            RemoveFromTaskMappings = False
        End Try

    End Function

    'Function InsertFldATTR(ByVal Fld As clsField, Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

    '    Dim cmd As Odbc.OdbcCommand
    '    Dim sql As String = ""
    '    Dim Attrib As String = ""
    '    Dim Value As String = ""

    '    Try
    '        If INcmd IsNot Nothing Then
    '            cmd = INcmd
    '        Else
    '            cmd = New Odbc.OdbcCommand
    '            cmd.Connection = cnn
    '        End If

    '        For i As Integer = 0 To 20

    '            Select Case i
    '                Case 0
    '                    Attrib = "CANNULL"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL).ToString
    '                Case 1
    '                    Attrib = "DATATYPE"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
    '                Case 2
    '                    Attrib = "ISKEY"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY).ToString
    '                Case 3
    '                    Attrib = "NCHILDREN"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN).ToString
    '                Case 4
    '                    Attrib = "NLENGTH"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString
    '                Case 5
    '                    Attrib = "NLEVEL"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL).ToString
    '                Case 6
    '                    Attrib = "NOCCNO"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS).ToString
    '                Case 7
    '                    Attrib = "NOFFSET"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET).ToString
    '                Case 8
    '                    Attrib = "NSCALE"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE).ToString
    '                Case 9
    '                    Attrib = "NTIMES"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES).ToString
    '                Case 10
    '                    Attrib = "FKEY"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY).ToString
    '                Case 11
    '                    Attrib = "INITVAL"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL).ToString
    '                Case 12
    '                    Attrib = "RETYPE"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE).ToString
    '                Case 13
    '                    Attrib = "EXTTYPE"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE).ToString
    '                Case 14
    '                    Attrib = "INVALID"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID).ToString
    '                Case 15
    '                    Attrib = "DATEFORMAT"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT).ToString
    '                Case 16
    '                    Attrib = "LABEL"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL).ToString
    '                Case 17
    '                    Attrib = "IDENTVAL"
    '                    Value = Fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL).ToString
    '                Case 18
    '                    Attrib = "ORGNAME"
    '                    Value = Fld.OrgName
    '                Case 19
    '                    Attrib = "PARENTNAME"
    '                    Value = Fld.ParentName
    '                Case 20
    '                    Attrib = "SEQNO"
    '                    Value = Fld.SeqNo.ToString
    '            End Select

    '            sql = "INSERT INTO " & Me.Project.tblDSselFieldsATTR & _
    '            "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DESCRIPTIONNAME,SELECTIONNAME,DATASTORENAME,FIELDNAME,DSSELFIELDATTRB,DSSELFIELDATTRBVALUE) " & _
    '            "Values(" & _
    '            Quote(Me.Project.ProjectName) & "," & _
    '            Quote(Me.ObjDatastore.Engine.ObjSystem.Environment.EnvironmentName) & "," & _
    '            Quote(Me.ObjDatastore.Engine.ObjSystem.SystemName) & "," & _
    '            Quote(Me.ObjDatastore.Engine.EngineName) & "," & _
    '            Quote(Me.ObjStructure.StructureName) & "," & _
    '            Quote(Me.SelectionName) & "," & _
    '            Quote(Me.ObjDatastore.DatastoreName) & "," & _
    '            Quote(Fld.FieldName) & "," & _
    '            Quote(Attrib) & "," & _
    '            Quote(Value) & ")"


    '            cmd.CommandText = sql
    '            Log(sql)
    '            cmd.ExecuteNonQuery()
    '        Next

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDSSelection InsertATTR")
    '        Return False
    '    End Try

    'End Function

    'Function DeleteFldATTR(ByVal Fld As clsField, Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

    '    Dim cmd As Odbc.OdbcCommand
    '    Dim sql As String = ""
    '    Dim Attrib As String = ""
    '    Dim Value As String = ""

    '    Try
    '        If INcmd IsNot Nothing Then
    '            cmd = INcmd
    '        Else
    '            cmd = New Odbc.OdbcCommand
    '            cmd.Connection = cnn
    '        End If

    '        sql = "DELETE FROM " & Me.Project.tblDSselFieldsATTR & _
    '        " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
    '        " AND ENVIRONMENTNAME=" & Quote(Me.ObjDatastore.Engine.ObjSystem.Environment.EnvironmentName) & _
    '        " AND SYSTEMNAME=" & Quote(Me.ObjDatastore.Engine.ObjSystem.SystemName) & _
    '        " AND ENGINENAME=" & Quote(Me.ObjDatastore.Engine.EngineName) & _
    '        " AND DESCRIPTIONNAME=" & Quote(Me.ObjStructure.StructureName) & _
    '        " AND SELECTIONNAME=" & Quote(Me.SelectionName) & _
    '        " AND DATASTORENAME=" & Quote(Me.ObjDatastore.DatastoreName) & _
    '        " AND FIELDNAME=" & Quote(Fld.FieldName)

    '        cmd.CommandText = sql
    '        Log(sql)
    '        cmd.ExecuteNonQuery()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDSSelection DeleteATTR")
    '        Return False
    '    End Try

    'End Function

    'Function LoadFieldAttr(ByRef Fld As clsField) As Boolean

    '    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
    '        Return True
    '        Exit Function
    '    End If

    '    Dim cmd As New System.Data.Odbc.OdbcCommand
    '    Dim da As System.Data.Odbc.OdbcDataAdapter
    '    Dim dt As New DataTable("temp")
    '    Dim dr As DataRow
    '    Dim sql As String = ""
    '    'Dim strAttrs As String
    '    Dim Attrib As String = ""
    '    Dim Value As String = ""

    '    Try
    '        sql = "SELECT * FROM " & Me.Project.tblDSselFieldsATTR & _
    '        " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
    '        " AND ENVIRONMENTNAME=" & Quote(Me.ObjDatastore.Engine.ObjSystem.Environment.EnvironmentName) & _
    '        " AND SYSTEMNAME=" & Quote(Me.ObjDatastore.Engine.ObjSystem.SystemName) & _
    '        " AND ENGINENAME=" & Quote(Me.ObjDatastore.Engine.EngineName) & _
    '        " AND DESCRIPTIONNAME=" & Quote(Me.ObjStructure.StructureName) & _
    '        " AND SELECTIONNAME=" & Quote(Me.SelectionName) & _
    '        " AND DATASTORENAME=" & Quote(Me.ObjDatastore.DatastoreName) & _
    '        " AND FIELDNAME=" & Quote(Fld.FieldName)

    '        cmd.CommandText = sql
    '        Log(sql)
    '        da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
    '        da.Fill(dt)
    '        da.Dispose()

    '        For i As Integer = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows(i)

    '            Attrib = GetVal(dr("DSSELFIELDATTRB"))
    '            Select Case Attrib
    '                Case "CANNULL"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "DATATYPE"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "ISKEY"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_ISKEY, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NCHILDREN"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_NCHILDREN, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NLENGTH"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NLEVEL"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LEVEL, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NOCCNO"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OCCURS, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NOFFSET"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OFFSET, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NSCALE"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "NTIMES"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_TIMES, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "FKEY"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_FKEY, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "INITVAL"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INITVAL, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "RETYPE"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_RETYPE, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "EXTTYPE"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_EXTTYPE, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "INVALID"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INVALID, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "DATEFORMAT"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "LABEL"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LABEL, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "IDENTVAL"
    '                    Fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_IDENTVAL, GetVal(dr("DSSELFIELDATTRBVALUE")))
    '                Case "ORGNAME"
    '                    Fld.OrgName = GetVal(dr("DSSELFIELDATTRBVALUE"))
    '                Case "PARENTNAME"
    '                    Fld.ParentName = GetVal(dr("DSSELFIELDATTRBVALUE"))
    '                Case "SEQNO"
    '                    Fld.SeqNo = GetVal(dr("DSSELFIELDATTRBVALUE"))
    '            End Select
    '        Next

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsDSSelection LoadFieldAttr")
    '        Return False
    '    End Try

    'End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
