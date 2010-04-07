Public Class clsStructureSelection
    Implements INode

    Private m_ObjStructure As clsStructure
    Private m_ObjParent As INode 'clsStructure
    Private m_SelectionName As String = ""
    Private m_IsSystemSelection As String = "0"
    Private m_SelectionDescription As String = ""
    Private m_IsRenamed As Boolean = False
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0

    Public ObjDSselections As New Collection '// collection of dsselections using this selection design
    Public ObjSelectionFields As New ArrayList '//Array of Fields selected with in structure
    '//Old selectionfieldss, this is helpful when doing Save operation, we can compare old fields and 
    '//new fields and perform insert/delete/update based on difference
    '//This array should be updated before we change ObjSelectionFields so we can have old values for comparision
    Public OldObjSelectionFields As New ArrayList

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
    '//This requires special handling. In case of Datastore we need 
    '//DSSELECTIONS whose parent is Datastore not structure
    Public Property Parent() As INode Implements INode.Parent
        Get
            If m_ObjStructure Is Nothing Then
                Return Me.m_ObjParent
            Else
                Return Me.m_ObjStructure
            End If
        End Get
        Set(ByVal Value As INode)
            If Value.Type = NODE_STRUCT Then
                m_ObjStructure = Value
            Else
                m_ObjParent = Value
            End If

        End Set
    End Property

    '/// OverHauled by TKarasch   April May 07
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsStructureSelection
        Dim ret As clsField
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadItems(True, False, cmd)

            obj.SelectionName = Me.SelectionName
            obj.SelectionDescription = Me.SelectionDescription
            obj.IsSystemSelection = Me.IsSystemSelection
            obj.IsModified = Me.IsModified
            obj.SeqNo = Me.SeqNo
            obj.Parent = NewParent 'Me.Parent

            '//commented by npatel on 7/25/05
            For Each fld As clsField In Me.ObjSelectionFields
                ret = SearchField(obj.ObjStructure, fld)
                If ret Is Nothing Then
                    Throw (New Exception("Selection Field [" & fld.Text & "] is not found in parent structure [" & fld.Struct.Text & "]"))
                Else
                    obj.ObjSelectionFields.Add(ret.Clone(obj, True, cmd))
                End If
            Next

            If obj.IsSystemSelection = "1" Then
                obj.ObjStructure.SysAllSelection = obj
            End If

            If Cascade = True Then
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsSS Clone")
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
            Key = Me.ObjStructure.Environment.Project.Text & KEY_SAP & Me.ObjStructure.Environment.Text & KEY_SAP & Me.ObjStructure.Text & KEY_SAP & Me.Text
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
            Return NODE_STRUCT_SEL
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
        Dim sql As String = ""
        Dim i As Integer = 0

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            '//We need to put in transaction because we will add structure and 
            '//fields in two steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '//Here We will save selection info and selection fields
            '//Note: Use transaction
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Update " & Me.Project.tblStructSel & " set SelectionName='" & FixStr(Me.SelectionName) & "', Description='" & _
                FixStr(Me.SelectionDescription) & "' where SelectionName=" & Me.GetQuotedText & _
                " AND StructureName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                sql = "Update " & Me.Project.tblDescriptionSelect & " set SelectionName='" & FixStr(Me.SelectionName) & _
                "' , SelectDescription='" & FixStr(Me.SelectionDescription) & _
                "' where SelectionName=" & Me.GetQuotedText & _
                " AND DESCRIPTIONName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If
            

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Save Selected Fields
            '//First delete old field selection and then add new selection when we save fields
            '//Now Loop through all fields and build sql insert script and submit to DB

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "DELETE FROM " & Me.Project.tblStrSelFields & " Where SelectionName=" & Me.GetQuotedText & _
                " AND StructureName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                sql = "DELETE FROM " & Me.Project.tblDescriptionSelFields & " Where SelectionName=" & Me.GetQuotedText & _
                " AND DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.ObjSelectionFields.Count > 0 Then
                '//Now Loop through all fields and build sql insert script and submit to DB
                For i = 0 To ObjSelectionFields.Count - 1

                    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "INSERT INTO " & Me.Project.tblStrSelFields & _
                        " (SelectionName,StructureName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    Else
                        sql = "INSERT INTO " & Me.Project.tblDescriptionSelFields & _
                        " (SelectionName,DescriptionName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    End If
                    

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next
            End If

            tran.Commit()

            Save = True
            Me.IsModified = False

        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "clsSS Save", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim strsql As String = ""
        Dim i As Integer = 0

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn
            '//We need to put in transaction because we will add structure and 
            '//fields in two steps so if one fails rollback all
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '/////////////////////////////////////////////////////////////////////////
            '//Insert in to structures table
            '/////////////////////////////////////////////////////////////////////////

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                strsql = "INSERT INTO " & Me.Project.tblStructSel & "(ProjectName,EnvironmentName,StructureName,SelectionName,Description,ISSYSTEMSELECT) " & " Values(" & Me.Project.GetQuotedText & "," & Me.ObjStructure.Environment.GetQuotedText & "," & Me.ObjStructure.GetQuotedText & "," & Me.GetQuotedText & ",'" & FixStr(SelectionDescription) & "'" & ",'" & FixStr(IsSystemSelection) & "')"
            Else
                strsql = "INSERT INTO " & Me.Project.tblDescriptionSelect & _
                "(ProjectName,EnvironmentName,DescriptionName,SelectionName,SelectDescription,ISSYSTEMSEL) Values(" & _
                Me.Project.GetQuotedText & "," & _
                Me.ObjStructure.Environment.GetQuotedText & "," & _
                Me.ObjStructure.GetQuotedText & "," & _
                Me.GetQuotedText & ",'" & _
                FixStr(Me.SelectionDescription) & "','" & _
                FixStr(Me.IsSystemSelection) & "')"
            End If


            cmd.CommandText = strsql
            Log(strsql)
            cmd.ExecuteNonQuery()

            '///Moved this part after Inser in StructureSelections on 8/15/05 by npatel 
            '//delete any previous record to prevent any possible duplicate entries
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                strsql = "DELETE FROM " & Me.Project.tblStrSelFields & " WHERE  SelectionName=" & Me.GetQuotedText & " AND StructureName=" & Me.ObjStructure.GetQuotedText & " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                strsql = "DELETE FROM " & Me.Project.tblDescriptionSelFields & " WHERE SelectionName=" & Me.GetQuotedText & _
                " AND DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If

            cmd.CommandText = strsql
            Log(strsql)
            cmd.ExecuteNonQuery()

            '//////////////////////////////////////////////////////////////////////////
            '//Also update parent table' SysAllSelection if this is SysAll Selection
            '//////////////////////////////////////////////////////////////////////////
            If Me.IsSystemSelection = "1" Then
                '//if ok then update structure object
                Me.ObjStructure.SysAllSelection = Me
            End If

            '/////////////////////////////////////////////////////////////////////////
            '//Insert in to STRUCTFIELDS table
            '/////////////////////////////////////////////////////////////////////////

            '//any fields to add for non sys selection?
            If Me.ObjSelectionFields.Count > 0 And Me.IsSystemSelection = "0" Then
                '//Now Loop through all fields and build sql insert script and submit to DB
                For i = 0 To ObjSelectionFields.Count - 1

                    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        strsql = "INSERT INTO " & Me.Project.tblStrSelFields & _
                        " (SelectionName,StructureName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    Else
                        strsql = "INSERT INTO " & Me.Project.tblDescriptionSelFields & _
                        " (SelectionName,DescriptionName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    End If

                    cmd.CommandText = strsql
                    Log(strsql)
                    cmd.ExecuteNonQuery()
                Next
            End If

            tran.Commit()

            AddToCollection(Me.ObjStructure.StructureSelections, Me, Me.GUID)

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            tran.Rollback()
            LogError(ex, "clsSS AddNew", strsql)
            Return False
        Finally
            'cnn.Close()
        End Try
       
    End Function

    '// Added by TK and KS 11/6/2006
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim strsql As String = ""
        Dim i As Integer = 0

        Try
            Me.Text = Me.Text.Trim
            '//////////////////////////////////////////////////////////////////////////
            '//Insert in to structures table
            '//////////////////////////////////////////////////////////////////////////
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                strsql = "INSERT INTO " & Me.Project.tblStructSel & "(ProjectName,EnvironmentName,StructureName,SelectionName,Description,ISSYSTEMSELECT) " & " Values(" & Me.Project.GetQuotedText & "," & Me.ObjStructure.Environment.GetQuotedText & "," & Me.ObjStructure.GetQuotedText & "," & Me.GetQuotedText & ",'" & FixStr(SelectionDescription) & "'" & ",'" & FixStr(IsSystemSelection) & "')"
            Else
                strsql = "INSERT INTO " & Me.Project.tblDescriptionSelect & _
                "(ProjectName,EnvironmentName,DescriptionName,SelectionName,SelectDescription,ISSYSTEMSEL) Values(" & _
                Me.Project.GetQuotedText & "," & _
                Me.ObjStructure.Environment.GetQuotedText & "," & _
                Me.ObjStructure.GetQuotedText & "," & _
                Me.GetQuotedText & ",'" & _
                FixStr(SelectionDescription) & "','" & _
                FixStr(IsSystemSelection) & "')"
            End If

            cmd.CommandText = strsql
            Log(strsql)
            cmd.ExecuteNonQuery()

            '///Moved this part after Inser in StructureSelections on 8/15/05 by npatel 
            '//delete any previous record to prevent any possible duplicate entries
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                strsql = "DELETE FROM " & Me.Project.tblStrSelFields & " WHERE  SelectionName=" & Me.GetQuotedText & " AND StructureName=" & Me.ObjStructure.GetQuotedText & " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                strsql = "DELETE FROM " & Me.Project.tblDescriptionSelFields & " WHERE SelectionName=" & Me.GetQuotedText & _
                " AND DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If

            cmd.CommandText = strsql
            Log(strsql)
            cmd.ExecuteNonQuery()

            '//////////////////////////////////////////////////////////////////////////
            '//Also update parent table' SysAllSelection if this is SysAll Selection
            '//////////////////////////////////////////////////////////////////////////
            If Me.IsSystemSelection = "1" Then
                '//if ok then update structure object
                Me.ObjStructure.SysAllSelection = Me
            End If

            '//////////////////////////////////////////////////////////////////////////
            '//Insert in to STRUCTFIELDS table
            '//////////////////////////////////////////////////////////////////////////

            '//any fields to add for non sys selection?
            If Me.ObjSelectionFields.Count > 0 And Me.IsSystemSelection = "0" Then
                '//Now Loop through all fields and build sql insert script 
                '//and submit to DB
                For i = 0 To ObjSelectionFields.Count - 1
                    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        strsql = "INSERT INTO " & Me.Project.tblStrSelFields & _
                        " (SelectionName,StructureName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    Else
                        strsql = "INSERT INTO " & Me.Project.tblDescriptionSelFields & _
                        " (SelectionName,DescriptionName,EnvironmentName,ProjectName,FieldName) Values( " & _
                        Me.GetQuotedText & "," & _
                        Me.ObjStructure.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.GetQuotedText & "," & _
                        Me.ObjStructure.Environment.Project.GetQuotedText & "," & _
                        ObjSelectionFields(i).GetQuotedText & ");"
                    End If

                    cmd.CommandText = strsql
                    Log(strsql)
                    cmd.ExecuteNonQuery()
                Next
            End If

            '//Add structuresel in to structure's structuresel collection
            AddToCollection(Me.ObjStructure.StructureSelections, Me, Me.GUID)

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsSS AddNew", strsql)
            Return False
        End Try

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            '//delete from STRSELFIELDS
            '//first delete child records
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Delete From " & Me.Project.tblStrSelFields & " where  SelectionName=" & Me.GetQuotedText & " AND StructureName=" & Me.ObjStructure.GetQuotedText & " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & " AND ProjectName=" & Me.ObjStructure.Environment.Project.GetQuotedText
            Else
                sql = "Delete From " & Me.Project.tblDescriptionSelFields & _
                " where SelectionName=" & Me.GetQuotedText & _
                " AND DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.ObjStructure.Environment.Project.GetQuotedText
            End If
           
            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//delete from StructureSelections
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Delete From " & Me.Project.tblStructSel & " where SelectionName=" & Me.GetQuotedText & " AND StructureName=" & Me.ObjStructure.GetQuotedText & " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                sql = "Delete From " & Me.Project.tblDescriptionSelect & _
                " where SelectionName=" & Me.GetQuotedText & _
                " AND DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                " AND EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '/// Now delete all DSselections that correspond to this structure(selection)
            If Cascade = True Then
                Dim DSselList As New ArrayList
                DSselList.Clear()
                DSselList = SearchDSselForStructureRef(Me)
                For Each dsSel As clsDSSelection In DSselList
                    dsSel.Delete(cmd, cnn, Cascade, RemoveFromParentCollection)
                Next
            End If

            '//Remove from parent collection
            If RemoveFromParentCollection = True Then
                RemoveFromCollection(Me.ObjStructure.StructureSelections, Me.GUID)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, sql)
            'MsgBox("craps at StrSelection")
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
            '//check if already loaded ?
            If Reload = False Then
                If Me.ObjSelectionFields.Count > 0 Then Exit Function
            End If

            'cnn.Open()
            'cmd = cnn.CreateCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.ObjStructure.LoadItems(Reload, False, cmd)
            ObjSelectionFields.Clear()

            If Me.IsSystemSelection = "1" Then
                ObjSelectionFields = Me.ObjStructure.ObjFields.Clone
            Else
                If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    sql = "Select sf.PROJECTNAME,sf.ENVIRONMENTNAME,sf.STRUCTURENAME,sf.FIELDNAME from " & Me.Project.tblStrSelFields & " ssf " & _
                    "inner join " & Me.Project.tblStructFields & " sf " & _
                    "on ssf.fieldname=sf.fieldname " & _
                    "AND  ssf.structurename=sf.structurename " & _
                    "AND  ssf.environmentname=sf.environmentname " & _
                    "AND  ssf.projectname=sf.projectname " & _
                    "where  ssf.SelectionName=" & Me.GetQuotedText & _
                    " AND ssf.StructureName=" & Me.ObjStructure.GetQuotedText & _
                    " AND ssf.EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                    " AND ssf.ProjectName=" & Me.ObjStructure.Environment.Project.GetQuotedText & _
                    " ORDER BY sf.SEQNO"
                Else
                    sql = "Select sf.PROJECTNAME,sf.ENVIRONMENTNAME,sf.DESCRIPTIONNAME,sf.FIELDNAME from " & Me.Project.tblDescriptionSelFields & " ssf inner join " & _
                    Me.Project.tblDescriptionFields & " sf " & _
                    "on ssf.fieldname=sf.fieldname " & _
                    "AND  ssf.descriptionname=sf.descriptionname " & _
                    "AND ssf.environmentname=sf.environmentname " & _
                    "AND  ssf.projectname=sf.projectname " & _
                    "where  ssf.SelectionName=" & Me.GetQuotedText & _
                    " AND ssf.DescriptionName=" & Me.ObjStructure.GetQuotedText & _
                    " AND ssf.EnvironmentName=" & Me.ObjStructure.Environment.GetQuotedText & _
                    " AND ssf.ProjectName=" & Me.ObjStructure.Environment.Project.GetQuotedText & _
                    " ORDER BY sf.SEQNO"

                End If


                cmd.CommandText = sql
                Log(sql)
                da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
                dt = New System.Data.DataTable("temp2")
                da.Fill(dt)
                da.Dispose()

                For i = 0 To dt.Rows.Count - 1
                    dr = dt.Rows(i)

                    Dim fld As clsField

                    If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        fld = SearchFieldByName(Me.ObjStructure, dr("fieldname"), _
                        dr("structurename"), dr("environmentname"), dr("projectname"))
                    Else
                        fld = SearchFieldByName(Me.ObjStructure, dr("fieldname"), _
                        dr("descriptionname"), dr("environmentname"), dr("projectname"))
                    End If

                    If fld Is Nothing Then
                        clsLogging.LogEvent("Field [" & dr("fieldname") & "] is missing in the parent structure - " & Me.ObjStructure.Text)
                    Else
                        ObjSelectionFields.Add(fld)
                    End If
                Next
            End If

        Catch ex As Exception
            LogError(ex, sql)
        Finally
            'cnn.Close()
        End Try

    End Function

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
            If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select SELECTIONNAME,STRUCTURENAME from " & Me.Project.tblStructSel & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.ObjStructure.Environment.EnvironmentName & "'"
            Else
                sql = "Select SELECTIONNAME,DESCRIPTIONNAME from " & Me.Project.tblDescriptionSelect & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.ObjStructure.Environment.EnvironmentName & "'"
            End If

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
                If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    readStrName = dr("STRUCTURENAME")
                Else
                    readStrName = dr("DESCRIPTIONNAME")
                End If

                If testName.Equals(readStrName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new name you have chosen already exists as a Description(Structure)" & (Chr(13)) & "in this Environment, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex)
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

#End Region

#Region "Properties"

    Public Property IsSystemSelection() As String
        Get
            Return m_IsSystemSelection
        End Get
        Set(ByVal Value As String)
            m_IsSystemSelection = Value
        End Set
    End Property

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

    Public Property ObjStructure() As clsStructure
        Get
            Return m_ObjStructure
        End Get
        Set(ByVal Value As clsStructure)
            m_ObjStructure = Value
        End Set
    End Property

#End Region

#Region "Methods"

    '/// All Methods are Part of Inode
    
#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class