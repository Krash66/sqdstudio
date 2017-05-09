Public Class clsStructure
    Implements INode

    Public fPath1 As String = ""
    Public fPath2 As String = ""
    Public IMSDBName As String = ""
    Public ObjFields As New ArrayList '//Array of Fields with in structure
    Public StructureSelections As New Collection

    Private m_StructureType As enumStructure
    Private m_ObjEnvironment As clsEnvironment
    Private m_Connection As clsConnection
    Private m_ObjParent As INode
    Private m_StructureName As String = ""
    Private m_SegmentName As String = ""
    Private m_StructureDescription As String = ""
    Private m_IsModified As Boolean
    Private m_IsRenamed As Boolean = False
    Private m_SysAllSelection As clsStructureSelection
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsLoaded As Boolean = False
    Private m_Tablespace As String = ""

    
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

    '//This requires special handling. In case of Datastore we need DSSELECTIONS, 
    '//Structs whose parent is Datastore not structure or environment
    Public Property Parent() As INode Implements INode.Parent
        Get
            If m_ObjEnvironment Is Nothing Then
                Return Me.m_ObjParent
            Else
                Return Me.m_ObjEnvironment
            End If
        End Get
        Set(ByVal Value As INode)
            If Value.Type = NODE_ENVIRONMENT Then
                m_ObjEnvironment = Value
            Else
                m_ObjParent = Value
            End If

        End Set
    End Property

    '/// OverHauled by TKarasch   April May 07
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsStructure
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)
            Me.LoadItems(, False, cmd)

            obj.StructureName = Me.StructureName
            obj.StructureDescription = Me.StructureDescription
            obj.SegmentName = Me.SegmentName
            obj.StructureType = Me.StructureType
            obj.fPath1 = Me.fPath1
            obj.fPath2 = Me.fPath2
            obj.IMSDBName = Me.IMSDBName
            obj.SeqNo = Me.SeqNo
            obj.Parent = NewParent 'Me.Parent '//set the parent environment
            obj.IsModified = Me.IsModified
            obj.Connection = Me.Connection
            obj.Tablespace = Me.Tablespace

            'If Me.ObjFields IsNot Nothing Then

            '//clone all fields
            For Each fld As clsField In Me.ObjFields
                Dim newfld As New clsField
                newfld = fld.Clone(obj, True, cmd)
                newfld.ParentStructureName = obj.StructureName
                newfld.Struct = obj
                obj.ObjFields.Add(newfld)
            Next
            ' End If

            '//clone all Selections under structure
            If Cascade = True Then

                '//clone all structures under Structure
                For Each sel As clsStructureSelection In Me.StructureSelections
                    Dim newsel As New clsStructureSelection
                    newsel = sel.Clone(obj, True, cmd)
                    newsel.ObjStructure = obj
                    AddToCollection(obj.StructureSelections, newsel, newsel.GUID)
                Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsStructure Clone", ex.ToString)
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
            Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = StructureName
        End Get
        Set(ByVal Value As String)
            StructureName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_STRUCT
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

        Dim sql As String = ""
        Dim ConnText As String = ""
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand

        Try
            Me.Text = Me.Text.Trim
            If Me.Connection Is Nothing Then
                ConnText = ""
            Else
                ConnText = Me.Connection.Text
            End If
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblStructures & " " & _
            '                    "set StructureName='" & FixStr(Me.StructureName.Trim) & "' " & _
            '                    ", Description='" & FixStr(Me.StructureDescription) & "' " & _
            '                    ", StructureType=" & Me.StructureType & " " & _
            '                    ", fPath1='" & FixStr(Me.fPath1) & "' " & _
            '                    ", fPath2='" & FixStr(Me.fPath2) & "' " & _
            '                    ", IMSDBName='" & FixStr(Me.IMSDBName) & "' " & _
            '                    ", ConnectionName='" & ConnText & "' " & _
            '                    "where StructureName=" & Me.GetQuotedText & _
            '                    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            '                    " AND ProjectName=" & Me.Environment.Project.GetQuotedText
            'Else
            sql = "Update " & Me.Project.tblDescriptions & _
            " set DescriptionName=" & Quote(FixStr(Me.StructureName.Trim)) & _
            ", DESCRIPTIONTYPE= " & Quote(FixStr(Me.StructureType)) & _
            ", DESCRIPTIONDescription=" & Quote(FixStr(Me.StructureDescription)) & _
            " where ProjectName=" & Me.Project.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND DescriptionName=" & Me.GetQuotedText

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)
            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.StructureName <> Me.SysAllSelection.SelectionName Then
                'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    sql = "Update " & Me.Project.tblStructSel & " " & _
                '    "set SelectionName='" & FixStr(Me.StructureName.Trim) & "' " & _
                '    "where StructureName=" & Me.GetQuotedText & _
                '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                '    " AND ProjectName=" & Me.Environment.Project.GetQuotedText

                'Else
                sql = "Update " & Me.Project.tblDescriptionSelect & _
                " set SelectionName=" & Quote(FixStr(Me.StructureName.Trim)) & _
                " where ProjectName=" & Me.Project.GetQuotedText & _
                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                " AND DescriptionName=" & Me.GetQuotedText

                'End If

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            End If

            Save = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsStructure Save", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim tran As System.Data.Odbc.OdbcTransaction = Nothing
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            '//////////////////////////////////////////////////////////////////////////
            '//Insert in to structures table
            '//////////////////////////////////////////////////////////////////////////
            '//We need to put in transaction because we will add structure 
            '//and fields in two steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            Dim ConnText As String = ""
            If Me.Connection IsNot Nothing Then
                ConnText = Me.Connection.Text
            End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "INSERT INTO " & Me.Project.tblStructures & _
            '    "(ProjectName,EnvironmentName,StructureName,Description,StructureType,fPath1,fPath2,SegmentName,IMSDBName,ConnectionName) " & _
            '    "Values(" & Me.Environment.Project.GetQuotedText & _
            '    "," & Me.Environment.GetQuotedText & _
            '    "," & Me.GetQuotedText & ",'" & FixStr(StructureDescription) & "'," & _
            '    StructureType & ",'" & FixStr(fPath1) & _
            '    "','" & FixStr(fPath2) & "','" & FixStr(SegmentName) & "','" & FixStr(IMSDBName) & "','" & FixStr(ConnText) & "')"

            '    Log(sql)
            '    cmd.CommandText = sql
            '    cmd.ExecuteNonQuery()
            'Else
            sql = "INSERT INTO " & Me.Project.tblDescriptions & _
            "(ProjectName,EnvironmentName,DescriptionName,DESCRIPTIONTYPE,DescriptionDescription)" & _
            " Values(" & Me.Project.GetQuotedText & "," & _
            Me.Environment.GetQuotedText & "," & _
            Me.GetQuotedText & "," & _
            Me.StructureType & "," & _
            Quote(FixStr(Me.StructureDescription)) & ")"

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)
            'End If

            '//Delete any previous fields : Safe
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "DELETE FROM " & Me.Project.tblStructFields & " WHERE StructureName=" & Me.GetQuotedText & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText
            'Else
            sql = "DELETE FROM " & Me.Project.tblDescriptionFields & " WHERE DESCRIPTIONName=" & Me.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            'Log(sql)
            'cmd.CommandText = sql
            'cmd.ExecuteNonQuery()
            ''remove attr too
            'sql = "DELETE FROM " & Me.Project.tblDescriptionFieldsATTR & " WHERE DESCRIPTIONName=" & Me.GetQuotedText & _
            '" AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText
            'End If


            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '//////////////////////////////////////////////////////////////////////////
            '//Insert new entries to STRUCTFIELDS table
            '//////////////////////////////////////////////////////////////////////////

            '//any fields to add?
            If ObjFields.Count > 0 Then
                Dim i As Integer

                '//Now Loop through all fields and build sql insert 
                '//script and submit to DB
                For i = 0 To Me.ObjFields.Count - 1
                    Dim fld As clsField = CType(Me.ObjFields(i), clsField)

                    'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "INSERT INTO " & Me.Project.tblStructFields & " (" & _
                    '    "ProjectName, " & _
                    '    "EnvironmentName, " & _
                    '    "StructureName, " & _
                    '    "ParentName, " & _
                    '    "FieldName, " & _
                    '    "nChildren, " & _
                    '    "nLevel, " & _
                    '    "nTimes, " & _
                    '    "nOccNo, " & _
                    '    "DataType, " & _
                    '    "nOffset, " & _
                    '    "nLength, " & _
                    '    "nScale, " & _
                    '    "CanNull, " & _
                    '    "IsKey, " & _
                    '    "OrgName, " & _
                    '    "SeqNo, " & _
                    '    "FieldDesc " & _
                    '    ")" & _
                    '    "Values(" & _
                    '    Me.Project.GetQuotedText & "," & _
                    '    Me.Environment.GetQuotedText & "," & _
                    '    Me.GetQuotedText & "," & _
                    '    Quote(fld.ParentName) & "," & _
                    '    Quote(fld.FieldName) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LEVEL) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_TIMES) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OCCURS) & "," & _
                    '    Quote(fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OFFSET) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LENGTH) & "," & _
                    '    fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_SCALE) & "," & _
                    '    Quote(fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_CANNULL)) & "," & _
                    '    Quote(fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY)) & "," & _
                    '    Quote(fld.OrgName) & "," & _
                    '    fld.SeqNo & "," & _
                    '    Quote(fld.FieldDesc) & _
                    '    ");"

                    '    Log(sql)
                    '    cmd.CommandText = sql
                    '    cmd.ExecuteNonQuery()
                    'Else
                    sql = "INSERT INTO " & Me.Project.tblDescriptionFields & " (" & _
                    "ProjectName, " & _
                    "EnvironmentName, " & _
                    "DescriptionName, " & _
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
                    "Values(" & _
                    Me.Project.GetQuotedText & "," & _
                    Me.Environment.GetQuotedText & "," & _
                    Me.GetQuotedText & "," & _
                    Quote(fld.FieldName) & "," & _
                    Quote(fld.ParentName) & "," & _
                    fld.SeqNo & "," & _
                    Quote(fld.FieldDesc) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                    Quote(fld.OrgName) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                    ");"

                    Log(sql)
                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()

                    'Me.InsertFldATTR(CType(ObjFields(i), clsField), cmd)
                    'End If
                Next  '/// field loop
            End If

            tran.Commit()

            Try
                '//When we add new structure then by default we add new Selection
                SysAllSelection.AddNew(True)
            Catch ex As Exception
                LogError(ex, "clsStructure AddNew .. Add addSysAllSelection")
                SysAllSelection.Delete(cmd, cnn)
            End Try

            '//Add structure in to environment's structure collection
            AddToCollection(Me.Environment.Structures, Me, Me.GUID)

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As clsStructureSelection In Me.StructureSelections
                    '//dont add sys selection again coz its already added in previous operation
                    If child.IsSystemSelection <> "1" Then
                        child.AddNew(True)
                    End If
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsStructure AddNew", sql)
            tran.Rollback()
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""
        Dim connect As String

        Try
            Me.Text = Me.Text.Trim
            '/////////////////////////////////////////////////////////////////////////
            '//Insert in to structures table
            '/////////////////////////////////////////////////////////////////////////
            '//We need to put in transaction because we will add structure and 
            '//fields in two steps so if one fails rollback all
            If Me.Connection Is Nothing Then
                connect = ""
            Else
                connect = FixStr(Me.Connection.Text)
            End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "INSERT INTO " & Me.Project.tblStructures & _
            '    "(ProjectName,EnvironmentName,StructureName,Description,StructureType,fPath1,fPath2,SegmentName,IMSDBName,ConnectionName) " & _
            '    "Values(" & Me.Environment.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & _
            '    "," & Me.GetQuotedText & ",'" & FixStr(StructureDescription) & "'," & StructureType & ",'" & FixStr(fPath1) & _
            '    "','" & FixStr(fPath2) & "','" & FixStr(SegmentName) & "','" & FixStr(IMSDBName) & "','" & FixStr(connect) & "')"
            'Else
            sql = "INSERT INTO " & Me.Project.tblDescriptions & _
            "(ProjectName,EnvironmentName,DescriptionName,DESCRIPTIONTYPE,DescriptionDescription)" & _
            " Values(" & Me.Project.GetQuotedText & "," & _
            Me.Environment.GetQuotedText & "," & _
            Me.GetQuotedText & "," & _
            Quote(Me.StructureType) & ",'" & _
            FixStr(Me.StructureDescription) & "')"

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)
            'End If

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '//Delete any previous fields : Safe
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "DELETE FROM " & Me.Project.tblStructFields & " WHERE StructureName=" & Me.GetQuotedText & _
            '    " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText
            'Else
            sql = "DELETE FROM " & Me.Project.tblDescriptionFields & " WHERE DESCRIPTIONName=" & Me.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            'Log(sql)
            'cmd.CommandText = sql
            'cmd.ExecuteNonQuery()
            ''remove attr too
            'sql = "DELETE FROM " & Me.Project.tblDescriptionFieldsATTR & " WHERE DESCRIPTIONName=" & Me.GetQuotedText & _
            '" AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText
            'End If

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            '/////////////////////////////////////////////////////////////////////////
            '//Insert new entries to STRUCTFIELDS table
            '/////////////////////////////////////////////////////////////////////////

            '//any fields to add?
            If ObjFields.Count > 0 Then
                Dim i As Integer


                '//Now Loop through all fields and build sql insert script 
                '//and submit to DB
                For i = 0 To ObjFields.Count - 1
                    Dim fld As clsField = CType(Me.ObjFields(i), clsField)

                    'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "INSERT INTO " & Me.Project.tblStructFields & " (" & _
                    '    "StructureName, " & _
                    '    "EnvironmentName, " & _
                    '    "ProjectName, " & _
                    '    "ParentName, " & _
                    '    "FieldName, " & _
                    '    "nChildren, " & _
                    '    "nLevel, " & _
                    '    "nTimes, " & _
                    '    "nOccNo, " & _
                    '    "DataType, " & _
                    '    "nOffset, " & _
                    '    "nLength, " & _
                    '    "nScale, " & _
                    '    "CanNull, " & _
                    '    "IsKey, " & _
                    '    "OrgName, " & _
                    '    "SeqNo, " & _
                    '    "FieldDesc " & _
                    '    ")" & _
                    '    "Values(" & _
                    '    Me.GetQuotedText & "," & _
                    '    Me.Environment.GetQuotedText & "," & _
                    '    Me.Environment.Project.GetQuotedText & "," & _
                    '    Quote(CType(ObjFields(i), clsField).ParentName) & "," & _
                    '    Quote(CType(ObjFields(i), clsField).FieldName) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LEVEL) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_TIMES) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OCCURS) & "," & _
                    '    Quote(CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OFFSET) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LENGTH) & "," & _
                    '    CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_SCALE) & "," & _
                    '    Quote(CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_CANNULL)) & "," & _
                    '    Quote(CType(ObjFields(i), clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY)) & "," & _
                    '    Quote(CType(ObjFields(i), clsField).OrgName) & "," & _
                    '    CType(ObjFields(i), clsField).SeqNo & "," & _
                    '    "'" & CType(ObjFields(i), clsField).FieldDesc & "'" & _
                    '    ");"
                    'Else
                    sql = "INSERT INTO " & Me.Project.tblDescriptionFields & " (" & _
                   "ProjectName, " & _
                   "EnvironmentName, " & _
                   "DescriptionName, " & _
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
                   "Values(" & _
                   Me.Project.GetQuotedText & "," & _
                   Me.Environment.GetQuotedText & "," & _
                   Me.GetQuotedText & "," & _
                   Quote(fld.FieldName) & "," & _
                   Quote(fld.ParentName) & "," & _
                   fld.SeqNo & "," & _
                   Quote(fld.FieldDesc) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                   fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                   Quote(fld.OrgName) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                   Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                   ");"

                    'End If

                    Log(sql)
                    cmd.CommandText = sql
                    cmd.ExecuteNonQuery()
                Next '// field loop
            End If

            Try
                '//When we add new structure then by default we add new Selection
                SysAllSelection.AddNew(cmd)
            Catch ex As Exception
                LogError(ex, "clsStructure AddNew .. AddNew SysAllSelection")
                SysAllSelection.Delete(cmd, cmd.Connection, False, True)
            End Try

            AddToCollection(Me.Environment.Structures, Me, Me.GUID)

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                Dim child As INode
                For Each child In StructureSelections
                    '//dont add sys selection again coz its already 
                    '//added in previous operation
                    If CType(child, clsStructureSelection).IsSystemSelection <> "1" Then
                        child.AddNew(cmd)
                    End If
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsStructure AddNew", sql)
            Return False
        End Try

    End Function

    '/// OverHauled by TKarasch April May 07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '//first delete child records
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Delete From " & Me.Project.tblStructFields & " where StructureName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            '    cmd.CommandText = sql
            '    Log(sql)
            '    cmd.ExecuteNonQuery()

            '    sql = "Delete From " & Me.Project.tblStructures & " where StructureName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            '    cmd.CommandText = sql
            '    Log(sql)
            '    cmd.ExecuteNonQuery()
            'Else
            sql = "Delete From " & Me.Project.tblDescriptionFields & " where DescriptionName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            sql = "Delete From " & Me.Project.tblDescriptions & " where DescriptionName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            DeleteATTR(cmd)
            'For Each fld As clsField In Me.ObjFields
            '    DeleteFldATTR(fld, cmd)
            'Next

            'End If



            If Cascade = True Then
                For Each SS As clsStructureSelection In Me.StructureSelections
                    SS.Delete(cmd, cnn, Cascade, RemoveFromParentCollection)
                    RemoveFromCollection(Me.StructureSelections, SS.GUID)
                Next
            Else '/// if we are not cascading down, sysAllSelection must be deleted from the STRUCTSEL table
                'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    sql = "Delete From " & Me.Project.tblStructSel & " where StructureName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText & " AND SelectionName= " & Me.SysAllSelection.GetQuotedText & " AND IsSystemSelect='1'"
                'Else
                sql = "Delete From " & Me.Project.tblDescriptionSelect & " where DescriptionName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText & " AND SelectionName= " & Me.SysAllSelection.GetQuotedText & " AND IsSystemSel=1"
                'End If

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            End If

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                RemoveFromCollection(Me.Environment.Structures, Me.GUID)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, "cls Structure Delete", sql)
            Return False
        End Try

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New DataTable("temp")
        Dim dr As DataRow
        Dim sql As String = ""
        Dim i As Integer

        Try
            If Reload = False Then
                '//check if fields already loaded in memory?
                If ObjFields.Count > 0 Then
                    Exit Function
                End If
            End If
            If TreeLode = True Then
                Exit Function
            End If

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            '/// Load Structure Attributes
            LoadMe(cmd)

            '/// Load Field Attributes
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,FIELDNAME,PARENTNAME,SEQNO,DESCFIELDDESCRIPTION,NCHILDREN,NLEVEL,NTIMES,NOCCNO,DATATYPE,NOFFSET,NLENGTH,NSCALE,CANNULL,ISKEY,ORGNAME,DATEFORMAT,LABEL,INITVAL,RETYPE,INVALID,EXTTYPE,IDENTVAL,FOREIGNKEY from " & Me.Project.tblDescriptionFields & _
            " where ProjectName=" & Me.Project.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND DescriptionName=" & Me.GetQuotedText & _
            " ORDER BY SEQNO"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            ObjFields.Clear()
            For i = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Dim fld As New clsField

                fld.Struct = Me
                fld.FieldName = dr("FieldName")
                fld.ParentName = GetStr(GetVal(dr("PARENTNAME")))
                fld.SeqNo = GetVal(dr("SEQNO"))
                fld.FieldDesc = GetVal(dr("DescFieldDescription"))
                fld.ParentStructureName = Me.Text
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_NCHILDREN, GetVal(dr("NCHILDREN")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LEVEL, GetVal(dr("NLEVEL")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_TIMES, GetVal(dr("NTIMES")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OCCURS, GetVal(dr("NOCCNO")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetVal(dr("DATATYPE")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OFFSET, GetVal(dr("NOFFSET")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("NLENGTH")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("NSCALE")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetVal(dr("CANNULL")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_ISKEY, GetVal(dr("ISKEY")))
                fld.OrgName = GetStr(GetVal(dr("ORGNAME")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT, GetVal(dr("DATEFORMAT")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LABEL, GetVal(dr("LABEL")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INITVAL, GetVal(dr("INITVAL")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_RETYPE, GetVal(dr("RETYPE")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_INVALID, GetVal(dr("INVALID")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_EXTTYPE, GetVal(dr("EXTTYPE")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_IDENTVAL, GetVal(dr("IDENTVAL")))
                fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_FKEY, GetVal(dr("FOREIGNKEY")))

                Me.ObjFields.Add(fld)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure LoadItems", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then Exit Try

            Dim cmd As System.Data.Odbc.OdbcCommand
            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As New DataTable("temp")
            Dim dr As DataRow
            Dim sql As String = ""
            'Dim strAttrs As String
            Dim Attrib As String = ""
            Dim Value As String = ""

            If Incmd IsNot Nothing Then
                cmd = Incmd
                'cmd.Transaction = INcmd.Transaction
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "SELECT DESCRIPTIONATTRB,DESCRIPTIONATTRBVALUE FROM " & Me.Project.tblDescriptionsATTR & _
            " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & _
            "' AND ENVIRONMENTNAME='" & FixStr(Me.Environment.EnvironmentName) & _
            "' AND DESCRIPTIONNAME='" & FixStr(Me.StructureName) & "'"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetVal(dr("DESCRIPTIONATTRB"))
                Select Case Attrib
                    Case "CONNECTIONNAME"
                        Me.Connection = SearchEnvForConn(Me, GetVal(dr("DESCRIPTIONATTRBVALUE")))
                    Case "FPATH1"
                        Me.fPath1 = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                    Case "FPATH2"
                        Me.fPath2 = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                    Case "IMDSBNAME"
                        Me.IMSDBName = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                    Case "SEGMENTNAME"
                        Me.SegmentName = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                    Case "STRUCTURETYPE"
                        Me.StructureType = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                    Case "TABLESPACE"
                        Me.Tablespace = GetVal(dr("DESCRIPTIONATTRBVALUE"))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure LoadMe")
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

    '/// Added by TKarasch for all validation
    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""
        Dim readName As String
        Dim readStrName As String
        Dim testName As String = NewName
        Dim NameValid As Boolean = True

        Try
            If testName = "" Then
                testName = Me.Text
            End If

            'cnn.Open()
            cmd = cnn.CreateCommand

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select SELECTIONNAME,STRUCTURENAME from " & Me.Project.tblStructSel & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & "'"
            'Else
            sql = "Select SELECTIONNAME,DESCRIPTIONNAME from " & Me.Project.tblDescriptionSelect & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & "'"
            'End If


            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("SELECTIONNAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new name you have chosen already exists as a Subset of a Structure" & (Chr(13)) & "in this Environment, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
                'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    readStrName = dr("STRUCTURENAME")
                'Else
                readStrName = dr("DESCRIPTIONNAME")
                'End If

                If testName.Equals(readStrName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new name you have chosen already exists as a Structure" & (Chr(13)) & "in this Environment, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsStructure ValidateNewObject", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property StructureType() As enumStructure
        Get
            Return m_StructureType
        End Get
        Set(ByVal Value As enumStructure)
            m_StructureType = Value
        End Set
    End Property

    Public Property StructureName() As String
        Get
            Return m_StructureName
        End Get
        Set(ByVal Value As String)
            m_StructureName = Value
            Me.SysAllSelection.SelectionName = Me.StructureName
        End Set
    End Property

    Public Property SegmentName() As String
        Get
            Return m_SegmentName
        End Get
        Set(ByVal Value As String)
            m_SegmentName = Value
        End Set
    End Property

    Public Property StructureDescription() As String
        Get
            Return m_StructureDescription
        End Get
        Set(ByVal Value As String)
            m_StructureDescription = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            Return m_ObjEnvironment
        End Get
        Set(ByVal Value As clsEnvironment)
            m_ObjEnvironment = Value
        End Set
    End Property

    Public Property SysAllSelection() As clsStructureSelection
        Get
            If m_SysAllSelection Is Nothing Then
                m_SysAllSelection = New clsStructureSelection
                m_SysAllSelection.ObjStructure = Me
                m_SysAllSelection.IsSystemSelection = "1"
                m_SysAllSelection.SelectionName = Me.Text

                Return m_SysAllSelection
            Else
                Return m_SysAllSelection
            End If
        End Get
        Set(ByVal Value As clsStructureSelection)
            If m_SysAllSelection Is Nothing Then
                m_SysAllSelection = New clsStructureSelection
            End If
            m_SysAllSelection = Value
        End Set
    End Property

    Public Property Connection() As clsConnection
        Get
            Return m_Connection
        End Get
        Set(ByVal value As clsConnection)
            m_Connection = value
        End Set
    End Property

    Public Property Tablespace() As String
        Get
            Return m_Tablespace
        End Get
        Set(ByVal value As String)
            m_Tablespace = value
        End Set
    End Property

#End Region

#Region "Methods"

    '/// Show Structure File in Notepad
    Function showStrFile() As Boolean

        Try
            If Me.StructureType <> enumStructure.STRUCT_REL_DML Then
                If System.IO.File.Exists(Me.fPath1) = True Then
                    Dim OpenProcess As New System.Diagnostics.Process
                    Process.Start(Me.fPath1)
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure showStrFile")
            Return False
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

            For i As Integer = 0 To 6
                Select Case i
                    Case 0
                        Attrib = "CONNECTIONNAME"
                        If Me.Connection IsNot Nothing Then
                            Value = Me.Connection.ConnectionName
                        Else
                            Value = ""
                        End If
                    Case 1
                        Attrib = "FPATH1"
                        Value = Me.fPath1
                    Case 2
                        Attrib = "FPATH2"
                        Value = Me.fPath2
                    Case 3
                        Attrib = "IMDSBNAME"
                        Value = Me.IMSDBName
                    Case 4
                        Attrib = "SEGMENTNAME"
                        Value = Me.SegmentName
                    Case 5
                        Attrib = "STRUCTURETYPE"
                        Value = Me.StructureType
                    Case 6
                        Attrib = "TABLESPACE"
                        Value = Me.Tablespace
                End Select

                sql = "INSERT INTO " & Me.Project.tblDescriptionsATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONATTRB,DESCRIPTIONATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & "','" & FixStr(Me.Environment.EnvironmentName) & _
                "','" & FixStr(Me.StructureName) & "','" & FixStr(Attrib) & "','" & FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure InsertATTR")
            Return False
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

            sql = "DELETE FROM " & Me.Project.tblDescriptionsATTR & " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & " AND DESCRIPTIONNAME=" & Quote(Me.StructureName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure DeleteATTR")
            Return False
        End Try

    End Function

    Function ChangeFPath(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As Odbc.OdbcCommand
        Dim sqlstr1 As String = ""
        Dim sqlstr2 As String = ""
        Dim sql As String = ""
        Dim Attrib1 As String = "FPATH1"
        Dim Attrib2 As String = "FPATH2"
        'Dim Value As String = ""
        Dim DoReplace As Boolean = False

        Try
            '/// If DML, no file path used
            If Me.StructureType = enumStructure.STRUCT_REL_DML Or Me.StructureType = enumStructure.STRUCT_UNKNOWN Then
                ChangeFPath = True
                Exit Try
            End If

            '/// Init ODBC Command
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            '/// Compare Description Paths with Environment Paths
            Select Case Me.StructureType
                Case enumStructure.STRUCT_C
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalCDir Then
                        sqlstr1 = Me.Environment.LocalCDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If

                Case enumStructure.STRUCT_COBOL
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalCobolDir Then
                        sqlstr1 = Me.Environment.LocalCobolDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If

                Case enumStructure.STRUCT_COBOL_IMS
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalCobolDir Then
                        sqlstr1 = Me.Environment.LocalCobolDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If
                    If GetDirFromPath(Me.fPath2) <> Me.Environment.LocalDBDDir Then
                        sqlstr2 = Me.Environment.LocalDBDDir & "\" & GetFileNameFromPath(Me.fPath2)
                        DoReplace = True
                    Else
                        sqlstr2 = Me.fPath2
                    End If

                Case enumStructure.STRUCT_REL_DDL
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalDDLDir Then
                        sqlstr1 = Me.Environment.LocalDDLDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If

                Case enumStructure.STRUCT_REL_DML_FILE
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalDMLDir Then
                        sqlstr1 = Me.Environment.LocalDMLDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If

                Case enumStructure.STRUCT_XMLDTD
                    If GetDirFromPath(Me.fPath1) <> Me.Environment.LocalDTDDir Then
                        sqlstr1 = Me.Environment.LocalDTDDir & "\" & GetFileNameFromPath(Me.fPath1)
                        DoReplace = True
                    Else
                        sqlstr1 = Me.fPath1
                    End If

                Case Else
                    ChangeFPath = True
                    Exit Try
            End Select

            If DoReplace = True Then
                '/// Delete old file path(s)
                sql = "DELETE FROM " & Me.Project.tblDescriptionsATTR & _
                " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                " AND DESCRIPTIONNAME=" & Quote(Me.StructureName) & _
                " AND DESCRIPTIONATTRB=" & Quote(Attrib1)

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()

                If Me.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                    sql = "DELETE FROM " & Me.Project.tblDescriptionsATTR & _
                    " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                    " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                    " AND DESCRIPTIONNAME=" & Quote(Me.StructureName) & _
                    " AND DESCRIPTIONATTRB=" & Quote(Attrib2)

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                End If

                '/// Insert new File Path(s)
                sql = "INSERT INTO " & Me.Project.tblDescriptionsATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONATTRB,DESCRIPTIONATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & _
                "','" & FixStr(Me.Environment.EnvironmentName) & _
                "','" & FixStr(Me.StructureName) & _
                "','" & FixStr(Attrib1) & _
                "','" & FixStr(sqlstr1) & "')"

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()

                If Me.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                    sql = "INSERT INTO " & Me.Project.tblDescriptionsATTR & _
                    "(PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONATTRB,DESCRIPTIONATTRBVALUE) " & _
                    "Values('" & FixStr(Me.Project.ProjectName) & _
                    "','" & FixStr(Me.Environment.EnvironmentName) & _
                    "','" & FixStr(Me.StructureName) & _
                    "','" & FixStr(Attrib2) & _
                    "','" & FixStr(sqlstr2) & "')"

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                End If
                Me.IsLoaded = False
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "clsStructure InsertATTR")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class