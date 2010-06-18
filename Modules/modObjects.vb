Module modObjects

#Region "Search By Name"

    Function SearchStructSelByName(ByRef objEnv As clsEnvironment, ByRef SelectionName As String, ByRef StructureName As String, ByRef EnvironmentName As String, ByRef ProjectName As String) As clsStructureSelection

        Try
            For Each objstr As clsStructure In objEnv.Structures
                For Each objsel As clsStructureSelection In objstr.StructureSelections
                    If objsel.SelectionName = SelectionName And _
                    objsel.ObjStructure.StructureName = StructureName And _
                    objsel.ObjStructure.Environment.EnvironmentName = EnvironmentName And _
                    objsel.Project.ProjectName = ProjectName Then
                        Return objsel
                        Exit Function
                    End If
                Next
            Next

            '// if it makes it out of the loop
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchStructSelByName")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    '//Description : This function will find matching field from struct field list. 
    '//Input : objEng - Engine which contains the variable list which needs to be search
    '//        VariableId - VariableId of a Variable which needs to be located in engine's variable list
    '//Return : Variable reference if found
    Function SearchVariableByName(ByRef objEng As clsEngine, ByRef VariableName As String, ByRef EngineName As String, ByRef SystemName As String, ByRef EnvironmentName As String, ByRef ProjectName As String) As clsVariable

        Dim var As clsVariable
        Dim i As Integer

        Try
            For i = 1 To objEng.Variables.Count
                var = objEng.Variables(i)
                If var.VariableName = VariableName _
                And var.Engine.EngineName = EngineName _
                And var.Engine.ObjSystem.SystemName = SystemName _
                And var.Environment.EnvironmentName = EnvironmentName _
                And var.Project.ProjectName = ProjectName Then
                    Return objEng.Variables(i)
                    Exit Function
                End If
            Next

            '// if it makes it out of the loop
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchVariableByName")
            Return Nothing
        End Try

    End Function

    '//Created : May 2009
    '//Description : This function will find matching field from struct field list. 
    '//Input : objEng - Engine which contains the variable list which needs to be search
    '//        VariableId - VariableId of a Variable which needs to be located in engine's variable list
    '//Return : Variable reference if found
    Function SearchVariableByName(ByRef objEnv As clsEnvironment, ByRef VariableName As String, _
    ByRef EnvironmentName As String, ByRef ProjectName As String) As clsVariable

        Try
            For Each var As clsVariable In objEnv.Variables
                If var.VariableName = VariableName _
                And var.Environment.EnvironmentName = EnvironmentName _
                And var.Project.ProjectName = ProjectName Then
                    Return var
                    Exit Function
                End If
            Next

            '// if it makes it out of the loop
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchVariableByName")
            Return Nothing
        End Try

    End Function

    '//Created : 7/21/05
    '//Description : This function will find matching field from struct field list. 
    '//Input : objStruct - Structure which contains the field which needs to be search
    '//        FieldId - FieldId of a field which needs to be located in structure field list
    '//Return : Field reference if found
    Function SearchTaskByName(ByRef TaskName As String, ByRef EngineName As String, ByRef SystemName As String, ByRef EnvironmentName As String, ByRef ProjectName As String, ByRef TaskType As enumTaskType, Optional ByRef objEng As clsEngine = Nothing, Optional ByRef ObjEnv As clsEnvironment = Nothing) As clsTask

        Dim tsk As clsTask
        Dim i As Integer

        Try
            Select Case TaskType
                'Case modDeclares.enumTaskType.TASK_GEN
                '    If objEng IsNot Nothing Then
                '        For i = 1 To objEng.Gens.Count
                '            tsk = objEng.Gens(i)
                '            If tsk.TaskName = TaskName _
                '            And tsk.Engine.EngineName = EngineName _
                '            And tsk.Engine.ObjSystem.SystemName = SystemName _
                '            And tsk.Environment.EnvironmentName = EnvironmentName _
                '            And tsk.Project.ProjectName = ProjectName Then
                '                Return objEng.Gens(i)
                '                Exit Function
                '            End If
                '        Next
                '    Else
                '        For Each proc As clsTask In ObjEnv.Procedures
                '            If proc.TaskName = TaskName _
                '            And proc.Environment.EnvironmentName = EnvironmentName _
                '            And proc.Project.ProjectName = ProjectName Then
                '                Return proc
                '                Exit Function
                '            End If
                '        Next
                '    End If

                Case modDeclares.enumTaskType.TASK_LOOKUP
                    If objEng IsNot Nothing Then
                        For i = 1 To objEng.Lookups.Count
                            tsk = objEng.Lookups(i)
                            If tsk.TaskName = TaskName And _
                            tsk.Engine.EngineName = EngineName _
                            And tsk.Engine.ObjSystem.SystemName = SystemName _
                            And tsk.Environment.EnvironmentName = EnvironmentName _
                            And tsk.Project.ProjectName = ProjectName Then
                                Return objEng.Lookups(i)
                                Exit Function
                            End If
                        Next
                    Else
                        For Each proc As clsTask In ObjEnv.Procedures
                            If proc.TaskName = TaskName _
                            And proc.Environment.EnvironmentName = EnvironmentName _
                            And proc.Project.ProjectName = ProjectName Then
                                Return proc
                                Exit Function
                            End If
                        Next
                    End If

                Case modDeclares.enumTaskType.TASK_MAIN
                    If objEng IsNot Nothing Then
                        For i = 1 To objEng.Mains.Count
                            tsk = objEng.Mains(i)
                            If tsk.TaskName = TaskName _
                            And tsk.Engine.EngineName = EngineName _
                            And tsk.Engine.ObjSystem.SystemName = SystemName _
                            And tsk.Environment.EnvironmentName = EnvironmentName _
                            And tsk.Project.ProjectName = ProjectName Then
                                Return objEng.Mains(i)
                                Exit Function
                            End If
                        Next
                    Else
                        For Each proc As clsTask In ObjEnv.Procedures
                            If proc.TaskName = TaskName _
                            And proc.Environment.EnvironmentName = EnvironmentName _
                            And proc.Project.ProjectName = ProjectName Then
                                Return proc
                                Exit Function
                            End If
                        Next
                    End If

                Case modDeclares.enumTaskType.TASK_PROC, modDeclares.enumTaskType.TASK_GEN
                    If objEng IsNot Nothing Then
                        For i = 1 To objEng.Procs.Count
                            tsk = objEng.Procs(i)
                            If tsk.TaskName = TaskName _
                            And tsk.Engine.EngineName = EngineName _
                            And tsk.Engine.ObjSystem.SystemName = SystemName _
                            And tsk.Environment.EnvironmentName = EnvironmentName _
                            And tsk.Project.ProjectName = ProjectName Then
                                Return objEng.Procs(i)
                                Exit Function
                            End If
                        Next
                    Else
                        For Each proc As clsTask In ObjEnv.Procedures
                            If proc.TaskName = TaskName _
                            And proc.Environment.EnvironmentName = EnvironmentName _
                            And proc.Project.ProjectName = ProjectName Then
                                Return proc
                                Exit Function
                            End If
                        Next
                    End If

            End Select

            '// if it makes it out of the loop
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchTaskByName")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    '//Description : This function will find matching field from selections if datastores defined under engine. 
    '//Input : objEng - Engine which contains the datastore->selections which needs to be search
    '//        FieldId - FieldId of a field which needs to be located in selection list
    '//Return : Field reference if found
    Function SearchFieldByName(ByRef objEng As clsEngine, ByRef FieldName As String, ByRef StructureName As String, ByRef EnvironmentName As String, ByRef ProjectName As String, ByRef Dir As enumDirection) As clsField

        Dim i, j, k As Integer
        Dim objDs As clsDatastore
        Dim objSel As clsDSSelection

        Try
            If Dir = modDeclares.enumDirection.DI_SOURCE Or Dir = modDeclares.enumDirection.DI_SOURCE_TARGET Then
                For i = 1 To objEng.Sources.Count
                    objDs = objEng.Sources(i)
                    objDs.LoadItems()
                    For j = 0 To objDs.ObjSelections.Count - 1
                        objSel = objDs.ObjSelections(j)
                        'objSel.LoadItems()
                        For k = 0 To objSel.DSSelectionFields.Count - 1
                            If CType(objSel.DSSelectionFields(k), clsField).FieldName = FieldName _
                               And CType(objSel.DSSelectionFields(k), clsField).Struct.StructureName = StructureName _
                               And CType(objSel.DSSelectionFields(k), clsField).Struct.Environment.EnvironmentName = EnvironmentName _
                               And CType(objSel.DSSelectionFields(k), clsField).Struct.Environment.Project.ProjectName = ProjectName Then
                                'objSel.IsMapped = True  '/// added 5/9/07 by TKarasch
                                Return objSel.DSSelectionFields(k)
                                Exit Function
                                'Else
                                '    objSel.IsMapped = False
                            End If
                        Next
                    Next
                Next
            End If

            If Dir = modDeclares.enumDirection.DI_TARGET Or Dir = modDeclares.enumDirection.DI_SOURCE_TARGET Then
                For i = 1 To objEng.Targets.Count
                    objDs = objEng.Targets(i)
                    objDs.LoadItems()
                    For j = 0 To objDs.ObjSelections.Count - 1
                        objSel = objDs.ObjSelections(j)
                        'objSel.LoadItems()
                        For k = 0 To objSel.DSSelectionFields.Count - 1
                            If CType(objSel.DSSelectionFields(k), clsField).FieldName = FieldName _
                                    And CType(objSel.DSSelectionFields(k), clsField).Struct.StructureName = StructureName _
                                    And CType(objSel.DSSelectionFields(k), clsField).Struct.Environment.EnvironmentName = EnvironmentName _
                                    And CType(objSel.DSSelectionFields(k), clsField).Struct.Environment.Project.ProjectName = ProjectName Then
                                'objSel.IsMapped = True   '/// added 5/9/07  by TKarasch
                                Return objSel.DSSelectionFields(k)
                                Exit Function
                                'Else
                                '    objSel.IsMapped = False
                            End If
                        Next
                    Next
                Next
            End If

            '// if it makes it out of the loops
            Return Nothing

        Catch ex As Exception
            LogError(ex, "SearchFieldByName")
            Return Nothing
        End Try

    End Function

    '//Created : May 2009
    '//Description : This function will find matching field from selections if datastores defined under engine. 
    '//Input : objEnv - Environment which contains the datastore->selections which needs to be search
    '//        FieldId - FieldId of a field which needs to be located in selection list
    '//Return : Field reference if found
    Function SearchFieldByName(ByRef objEnv As clsEnvironment, ByRef FieldName As String, _
    ByRef selName As String, ByRef ProjectName As String) As clsField

        Try
            For Each DSobj As clsDatastore In objEnv.Datastores
                DSobj.LoadItems()
                For Each SelObj As clsDSSelection In DSobj.ObjSelections
                    SelObj.LoadItems()
                    For Each FldObj As clsField In SelObj.DSSelectionFields
                        If FldObj.FieldName = FieldName _
                        And FldObj.Struct.StructureName = selName _
                        And FldObj.Struct.Environment.EnvironmentName = objEnv.EnvironmentName _
                        And FldObj.Project.ProjectName = ProjectName Then
                            Return FldObj
                            Exit Function
                        End If
                    Next
                Next
            Next

            '// if it makes it out of the loops
            Return Nothing

        Catch ex As Exception
            LogError(ex, "SearchFieldByName")
            Return Nothing
        End Try

    End Function

    '//Created : 7/18/05
    '//Description : This function will find matching field from struct field list. 
    '//Input : objStruct - Structure which contains the field which needs to be search
    '//        FieldName - FieldName of a field which needs to be located in structure field list
    '//        StructureName -  
    '//        EnvironmentName -
    '//        ProjectName -
    '//Return : Field reference if found
    Function SearchFieldByName(ByRef objStruct As clsStructure, ByRef FieldName As String, ByRef StructureName As String, ByVal EnvironmentName As String, ByVal ProjectName As String) As clsField

        Dim fld As clsField
        Dim i As Integer

        Try
            objStruct.LoadItems()
            For i = 0 To objStruct.ObjFields.Count - 1
                fld = objStruct.ObjFields(i)
                If (fld.FieldName = FieldName) And (fld.Struct.Text = StructureName) Then
                    Return objStruct.ObjFields(i)
                    Exit Function
                End If
            Next

            '// if it makes it out of the loop
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchFieldByName-OL2")
            Return Nothing
        End Try

    End Function

    '//Created : 7/19/05
    '//Description : This function will find matching field from struct field list. 
    '//Input : objSelection - Structure Selection which contains the field which needs to be search
    '//        FieldName - FieldName of a field which needs to be located in structure sel field list
    '//        StructureName -  
    '//        EnvironmentName -
    '//        ProjectName -
    '//Return : Field reference if found
    Function SearchFieldByName(ByRef objSel As clsStructureSelection, ByRef FieldName As String, ByRef StructureName As String, ByVal EnvironmentName As String, ByVal ProjectName As String) As clsField

        Dim fld As clsField
        Dim i As Integer

        Try
            objSel.LoadItems()
            For i = 0 To objSel.ObjSelectionFields.Count - 1
                fld = objSel.ObjSelectionFields(i)
                If (fld.FieldName = FieldName) And (fld.Struct.Text = StructureName) And (fld.Struct.Environment.Text = EnvironmentName) And (fld.Struct.Environment.Project.Text = ProjectName) Then
                    Return objSel.ObjSelectionFields(i)
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchFieldByName-OL3")
            Return Nothing
        End Try

    End Function

    '//Created : 7/21/05
    '//Description : This function will find matching datastore from the datastore list of a specified engine. 
    '//Input : objEng - Engine which contains the selection which needs to be search
    '//DatastoreName,.... - DatastoreName of a datastore which needs to be located in engine's datastore list 
    '//Return : datastore reference if found
    Function SearchDatastoreByName(ByRef objEng As clsEngine, _
            ByRef DatastoreName As String, _
            ByRef EngineName As String, _
            ByRef SystemName As String, _
            ByRef EnvironmentName As String, _
            ByRef ProjectName As String, _
            ByRef DsDirection As String) As clsDatastore

        Dim i As Integer

        Try
            If DsDirection = DS_DIRECTION_SOURCE Then
                For i = 1 To objEng.Sources.Count

                    If CType(objEng.Sources(i), clsDatastore).DatastoreName = DatastoreName _
                        And CType(objEng.Sources(i), clsDatastore).Engine.EngineName = EngineName _
                        And CType(objEng.Sources(i), clsDatastore).Engine.ObjSystem.SystemName = SystemName _
                        And CType(objEng.Sources(i), clsDatastore).Engine.ObjSystem.Environment.EnvironmentName = EnvironmentName _
                        And CType(objEng.Sources(i), clsDatastore).Engine.ObjSystem.Environment.Project.ProjectName = ProjectName Then

                        Return objEng.Sources(i)
                        Exit Function
                    End If
                Next
            End If

            If DsDirection = DS_DIRECTION_TARGET Then

                For i = 1 To objEng.Targets.Count
                    If CType(objEng.Targets(i), clsDatastore).DatastoreName = DatastoreName _
                        And CType(objEng.Targets(i), clsDatastore).Engine.EngineName = EngineName _
                        And CType(objEng.Targets(i), clsDatastore).Engine.ObjSystem.SystemName = SystemName _
                        And CType(objEng.Targets(i), clsDatastore).Engine.ObjSystem.Environment.EnvironmentName = EnvironmentName _
                        And CType(objEng.Targets(i), clsDatastore).Engine.ObjSystem.Environment.Project.ProjectName = ProjectName Then

                        Return objEng.Targets(i)
                        Exit Function
                    End If
                Next
            End If

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDatastoreByName")
            Return Nothing
        End Try

    End Function

    '/// Added May 2009 for New tree
    Function SearchDatastoreByName(ByRef objEnv As clsEnvironment, _
                ByRef DatastoreName As String, _
                ByRef EnvironmentName As String, _
                ByRef ProjectName As String) As clsDatastore

        Try
            For Each objDS As clsDatastore In objEnv.Datastores
                If objDS.DatastoreName = DatastoreName _
                And objDS.Environment.EnvironmentName = EnvironmentName _
                And objDS.Project.ProjectName = ProjectName Then
                    Return objDS
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDatastoreByName")
            Return Nothing
        End Try

    End Function

    'added by TK 12/29/2006
    Function SearchDSFieldByName(ByRef objSel As clsDSSelection, ByRef FieldName As String, ByRef StructureName As String, ByVal EnvironmentName As String, ByVal ProjectName As String) As clsField

        'Dim fld As clsField 'new???? try this
        'Dim i As Integer

        Try
            'objSel.ObjStructure.LoadItems()
            For Each fld As clsField In objSel.ObjStructure.ObjFields
                'fld = objSel.DSSelectionFields(i)
                If (fld.FieldName = FieldName) And (fld.Struct.StructureName = StructureName) And (fld.Struct.Environment.EnvironmentName = EnvironmentName) And (fld.Project.ProjectName = ProjectName) And (fld.ParentName = objSel.ObjStructure.StructureName) Then
                    Return fld
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDSfieldByName")
            Return Nothing
        End Try

    End Function

    '/// Function added by TK on 5/4/07 
    '/// This function is used when setting ObjSelection for the DSSelection
    '/// Input: clsDSselection
    '/// Output: It's corresponding clsStructureSelection in the same environment
    Function FindStrSelforDSSel(ByRef DSsel As clsDSSelection) As clsStructureSelection

        Try
            Dim dsSelName As String = DSsel.SelectionName
            Dim SrchEnv As clsEnvironment = DSsel.ObjDatastore.Environment

            'If DSsel.Parent.Type = NODE_ENVIRONMENT Then
            '    SrchEnv = DSsel.Parent
            'Else
            '    SrchEnv = DSsel.ObjDatastore.Environment
            'End If


            For Each str As clsStructure In SrchEnv.Structures
                For Each sel As clsStructureSelection In str.StructureSelections
                    If sel.SelectionName = dsSelName Then
                        Return sel
                        Exit Function
                    End If
                Next
            Next
            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects FindStrSelforDSsel")
            Return Nothing
        End Try

    End Function

#End Region

#Region "Search by Object"

    '//Created : 4/29/05
    '//Description : This function will find matching field from struct field list. 
    '//Input : objStruct - Structure which contains the field which needs to be search
    '//        FieldId - FieldId of a field which needs to be located in structure field list
    '//Return : Field reference if found
    Function SearchField(ByRef objStruct As clsStructure, ByRef objFld As clsField) As clsField

        Return SearchFieldByName(objStruct, objFld.FieldName, objFld.Struct.Text, objFld.Struct.Environment.Text, objFld.Struct.Environment.Project.Text)

    End Function

    Function SearchField(ByRef objSel As clsStructureSelection, ByRef objFld As clsField) As clsField

        Return SearchFieldByName(objSel, objFld.Text, objFld.Struct.Text, objFld.Struct.Environment.Text, objFld.Struct.Environment.Project.Text)

    End Function

    Function SearchField(ByRef objEnv As clsEnvironment, ByRef objFld As clsField) As clsField

        Dim str As clsStructure
        Dim i As Integer

        Try
            For i = 1 To objEnv.Structures.Count
                str = objEnv.Structures(i)
                If str.StructureName = objFld.ParentStructureName Then
                    Return SearchFieldByName(str, objFld.FieldName, objFld.Struct.Text, objFld.Struct.Environment.Text, objFld.Struct.Environment.Project.Text)
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchField-OL3")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    '//Description : This function will find matching selection from struct list of a specified env. 
    '//Input : objEnv - Environment which contains the selection which needs to be search
    '//        objSel - selection object which needs to be located in structure's selection 
    '///                list under a specified env
    '//Return : Selection reference if found
    Function SearchStructSel(ByRef objEnv As clsEnvironment, ByRef objSel As clsStructureSelection) As clsStructureSelection

        Dim objStr As clsStructure
        Dim i, j As Integer

        Try
            For i = 1 To objEnv.Structures.Count
                objStr = objEnv.Structures(i)
                If objStr.StructureName = objSel.ObjStructure.StructureName Then
                    For j = 1 To objStr.StructureSelections.Count
                        If CType(objStr.StructureSelections(j), clsStructureSelection).SelectionName = objSel.SelectionName Then
                            Return objStr.StructureSelections(j)
                            Exit Function
                        End If
                    Next
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchStructSel")
            Return Nothing
        End Try

    End Function

    Function SearchEnvForConn(ByRef objStr As clsStructure, ByVal ConnName As String) As clsConnection

        Dim i As Integer
        Dim objConn As clsConnection
        Dim objEnv As clsEnvironment

        Try
            objEnv = objStr.Environment
            For i = 1 To objEnv.Connections.Count
                objConn = objEnv.Connections(i)
                If objConn.ConnectionName = ConnName Then
                    Return objConn
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchStructSel")
            Return Nothing
        End Try

    End Function

    Function SearchForConn(ByRef objEng As clsEngine, ByVal ConnName As String) As clsConnection

        Dim i As Integer
        Dim objConn As clsConnection
        Dim objEnv As clsEnvironment

        Try
            objEnv = objEng.ObjSystem.Environment
            For i = 1 To objEnv.Connections.Count
                objConn = objEnv.Connections(i)
                If objConn.ConnectionName = ConnName Then
                    Return objConn
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchStructSel")
            Return Nothing
        End Try

    End Function

    '//created by: Tom Karasch 
    '// Use to find datastore selections matching structure selections
    '// Input: Structure selection
    '// Return: Corresponding Datastore selection(s) as arraylist
    '// Usage: Used when Structure or StructSelection is deleted
    '//        so that corresponding datastore selections will also be deleted.
    '//        or to find all structure selection references in it's environment
    Function SearchDSselForStructureRef(ByRef objsel As clsStructureSelection) As ArrayList

        Try
            Dim TempList As New ArrayList
            TempList.Clear()

            For Each sys As clsSystem In objsel.ObjStructure.Environment.Systems
                For Each eng As clsEngine In sys.Engines
                    For Each sds As clsDatastore In eng.Sources
                        For Each sdsSel As clsDSSelection In sds.ObjSelections
                            If sdsSel.ObjSelection Is objsel Then
                                TempList.Add(sdsSel)
                            End If
                        Next
                    Next
                    For Each tds As clsDatastore In eng.Targets
                        For Each tdsSel As clsDSSelection In tds.ObjSelections
                            If tdsSel.ObjSelection Is objsel Then
                                TempList.Add(tdsSel)
                            End If
                        Next
                    Next
                Next
            Next

            SearchDSselForStructureRef = TempList

        Catch ex As Exception
            LogError(ex, "modObjects SearchDSselForStructureRef")
            Return Nothing
        End Try

    End Function

    '//Created : 1/30/2007 By TKarasch
    '//Description : This function will find matching selection from struct list of a specified env. 
    '//Input : objEnv - Environment which contains the selection which needs to be search
    '//        objSel - DSselection object which needs to be located in structure's selection list 
    '//                 under a specified env
    '//Return : DSSelection reference if found
    Function SearchDSSel(ByRef objEnv As clsEnvironment, ByRef objSel As clsDSSelection) As clsDSSelection

        Dim objStr As clsStructure
        Dim i, j As Integer

        Try
            For i = 1 To objEnv.Structures.Count
                objStr = objEnv.Structures(i)
                If objStr.StructureName = objSel.ObjStructure.StructureName Then
                    For j = 1 To objStr.StructureSelections.Count
                        If CType(objStr.StructureSelections(j), clsStructureSelection).SelectionName = objSel.SelectionName Then
                            Return objSel
                            Exit Function
                        End If
                    Next
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDSsel")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    '//Description : This function will find matching datastore from the datastore list of a specified engine. 
    '//Input : objEng - Engine which contains the selection which needs to be search
    '//        objDs - datastore which needs to be located in engine's datastore list 
    '//Return : datastore reference if found
    Function SearchDatastore(ByRef objEng As clsEngine, ByRef objDs As clsDatastore) As clsDatastore

        Dim i As Integer

        Try
            If objDs.DsDirection = DS_DIRECTION_SOURCE Then
                For i = 1 To objEng.Sources.Count
                    If CType(objEng.Sources(i), clsDatastore).DatastoreName = objDs.DatastoreName Then
                        Return objEng.Sources(i)
                        Exit Function
                    End If
                Next
            End If

            If objDs.DsDirection = DS_DIRECTION_TARGET Then
                For i = 1 To objEng.Targets.Count
                    If CType(objEng.Targets(i), clsDatastore).DatastoreName = objDs.DatastoreName Then
                        Return objEng.Targets(i)
                        Exit Function
                    End If
                Next
            End If

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDatastore")
            Return Nothing
        End Try

    End Function

    '//Created : May 2009
    '//Description : This function will find matching datastore from the datastore list of a specified engine. 
    '//Input : objEng - Engine which contains the selection which needs to be search
    '//        objDs - datastore which needs to be located in engine's datastore list 
    '//Return : datastore reference if found
    Function SearchDatastore(ByRef objEnv As clsEnvironment, ByRef objDs As clsDatastore) As clsDatastore

        Try
            For Each DS As clsDatastore In objEnv.Datastores
                If DS.DatastoreName = objDs.DatastoreName Then
                    Return DS
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchDatastore")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    Function SearchTask(ByRef objEng As clsEngine, ByRef objTask As clsTask) As clsTask

        Dim tsk As clsTask
        Dim i As Integer
        Dim TaskType As enumTaskType

        Try
            TaskType = objTask.TaskType
            objTask.LoadMappings()

            Select Case TaskType
                'Case modDeclares.enumTaskType.TASK_GEN
                '    For i = 1 To objEng.Gens.Count
                '        tsk = objEng.Gens(i)
                '        tsk.LoadMappings()
                '        If tsk.TaskName = objTask.TaskName And tsk.ObjMappings.Count = objTask.ObjMappings.Count Then
                '            Return objEng.Gens(i)
                '            Exit Function
                '        End If
                '    Next
                Case modDeclares.enumTaskType.TASK_LOOKUP
                    For i = 1 To objEng.Lookups.Count
                        tsk = objEng.Lookups(i)
                        tsk.LoadMappings()
                        If tsk.TaskName = objTask.TaskName And tsk.ObjMappings.Count = objTask.ObjMappings.Count Then
                            Return objEng.Lookups(i)
                            Exit Function
                        End If
                    Next
                Case modDeclares.enumTaskType.TASK_MAIN
                    For i = 1 To objEng.Mains.Count
                        tsk = objEng.Mains(i)
                        tsk.LoadMappings()

                        If tsk.TaskName = objTask.TaskName And tsk.ObjMappings.Count = objTask.ObjMappings.Count Then
                            Return objEng.Mains(i)
                            Exit Function
                        End If
                    Next
                Case modDeclares.enumTaskType.TASK_PROC, modDeclares.enumTaskType.TASK_GEN
                    For i = 1 To objEng.Procs.Count
                        tsk = objEng.Procs(i)
                        tsk.LoadMappings()

                        If tsk.TaskName = objTask.TaskName And tsk.ObjMappings.Count = objTask.ObjMappings.Count Then
                            Return objEng.Procs(i)
                            Exit Function
                        End If
                    Next
            End Select

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchTask")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    Function SearchTask(ByRef objEnv As clsEnvironment, ByRef objTask As clsTask) As clsTask

        Try
            objTask.LoadMappings()

            For Each task As clsTask In objEnv.Procedures
                task.LoadMappings()

                If task.TaskName = objTask.TaskName And task.ObjMappings.Count = objTask.ObjMappings.Count Then
                    Return task
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchTask")
            Return Nothing
        End Try

    End Function

    Function SearchVariable(ByRef objEng As clsEngine, ByRef objVar As clsVariable) As clsVariable

        Dim var As clsVariable
        Dim i As Integer

        Try
            For i = 1 To objEng.Variables.Count
                var = objEng.Variables(i)
                If var.VariableName = var.VariableName And var.VariableType = objVar.VariableType Then
                    Return objEng.Variables(i)
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchVariable")
            Return Nothing
        End Try

    End Function

    Function SearchVariable(ByRef objEnv As clsEnvironment, ByRef objVar As clsVariable) As clsVariable

        Try
            For Each var As clsVariable In objEnv.Variables
                If var.VariableName = var.VariableName And var.VariableType = objVar.VariableType Then
                    Return var
                    Exit Function
                End If
            Next

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchVariable")
            Return Nothing
        End Try

    End Function

    '//Created : 4/29/05
    '//Description : This function will find matching field from selections if datastores defined under engine. 
    '//Input : objEng - Engine which contains the datastore->selections which needs to be search
    '//        FieldId - FieldId of a field which needs to be located in selection list
    '//Return : Field reference if found
    Function SearchField(ByRef objEng As clsEngine, ByRef objFld As clsField, ByRef Dir As enumDirection, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As clsField

        Dim fld As clsField
        Dim i, j, k As Integer
        Dim objDs As clsDatastore
        Dim objSel As clsDSSelection
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            If Dir = modDeclares.enumDirection.DI_SOURCE Or Dir = modDeclares.enumDirection.DI_SOURCE_TARGET Then
                For i = 1 To objEng.Sources.Count
                    objDs = objEng.Sources(i)
                    objDs.LoadItems(, , cmd)
                    For j = 0 To objDs.ObjSelections.Count - 1
                        objSel = objDs.ObjSelections(j)
                        objSel.LoadItems(, , cmd)
                        For k = 0 To objSel.DSSelectionFields.Count - 1
                            fld = objSel.DSSelectionFields(k)
                            If fld.FieldName = objFld.FieldName And fld.ParentStructureName = objFld.ParentStructureName Then
                                Return objSel.DSSelectionFields(k)
                                Exit Function
                            End If
                        Next
                    Next
                Next
            End If

            If Dir = modDeclares.enumDirection.DI_TARGET Or Dir = modDeclares.enumDirection.DI_SOURCE_TARGET Then
                For i = 1 To objEng.Targets.Count
                    objDs = objEng.Targets(i)
                    objDs.LoadItems(, , cmd)
                    For j = 0 To objDs.ObjSelections.Count - 1
                        objSel = objDs.ObjSelections(j)
                        objSel.LoadItems(, , cmd)
                        For k = 0 To objSel.DSSelectionFields.Count - 1
                            fld = objSel.DSSelectionFields(k)
                            If fld.FieldName = objFld.FieldName And fld.ParentStructureName = objFld.ParentStructureName Then
                                Return objSel.DSSelectionFields(k)
                                Exit Function
                            End If
                        Next
                    Next
                Next
            End If

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modObjects SearchField")
            Return Nothing
        End Try

    End Function

#End Region

    '// This gets the Structure/selection Name from the foreign key
    Public Function GetFKey(ByVal fkey As String) As String
        Dim element As String
        element = Split(fkey, ".")(0)
        Return element
    End Function

    '/// Used to add any Object to a Parent Object's Collection
    Function AddToCollection(ByVal col As Collection, ByVal obj As INode, Optional ByVal key As String = "") As Boolean
        ' Necesary in case a duplicate key value happens to be calculated
        On Error Resume Next
        If key <> "" Then
            col.Add(obj, key)
        Else
            col.Add(obj)
        End If
    End Function

    '//if key is not passed then all items will be removed
    Function RemoveFromCollection(ByRef col As Collection, ByVal key As String) As Boolean
        ' Necesary in no element found for this key 
        On Error Resume Next
        If key <> "" Then
            col.Remove(key)
        Else
            col = Nothing
            col = New Collection
        End If

    End Function

    Public Function GetObjCopy(ByVal NewParent As INode, ByRef srcObj As INode) As INode

        ''Makes Deep Copy of src to dest 
        ''Note both objects must be of same type and serialzable 
        Return CType(srcObj, INode).Clone(NewParent)

    End Function

    '/// Created by Tom Karasch June 2007 to create a new structure from relational database system tables
    Public Function MakeDMLStructure(ByVal ObjDML As clsDMLinfo, ByRef Env As clsEnvironment) As clsStructure

        Dim conn As New System.Data.Odbc.OdbcConnection(ObjDML.MetaConnectionString)
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""

        Try
            Dim Objthis As New clsStructure

            '/// Initialize new Object properties
            Objthis.StructureName = ObjDML.Schema.Trim & "_" & ObjDML.TableName.Trim
            Objthis.StructureType = enumStructure.STRUCT_REL_DML
            Objthis.StructureDescription = ""
            Objthis.Environment = Env
            Objthis.Parent = Env
            Objthis.Connection = ObjDML.Connection

            '/// open a connection to the Database where the DML is to come from
            conn.Open()
            cmd = conn.CreateCommand
            cmd.Connection = conn

            For Each col As String In ObjDML.ColumnArray

                '/// Get the properties for each field chosen for the DML
                Select Case ObjDML.DSNtype
                    Case enumODBCtype.DB2
                        sql = "SELECT TYPENAME, LENGTH, SCALE, NULLS, COLNO FROM SYSCAT.COLUMNS WHERE COLNAME= '" & col & "' AND TABSCHEMA = '" & ObjDML.Schema & "' AND TABNAME = '" & ObjDML.TableName & "' ORDER BY COLNO"
                    Case enumODBCtype.ORACLE
                        sql = "SELECT DATA_TYPE, DATA_LENGTH, DATA_SCALE, NULLABLE, COLUMN_ID FROM USER_TAB_COLUMNS WHERE COLUMN_NAME ='" & col & "' AND TABLE_NAME = '" & ObjDML.TableName & "' ORDER BY COLUMN_ID"
                    Case enumODBCtype.SQL_SERVER
                        sql = "select data_type, character_maximum_length, numeric_scale, is_nullable, ordinal_position from information_schema.columns where column_name= '" & col & "' and table_name = '" & ObjDML.TableName & "' order by ordinal_position"
                End Select

                cmd.CommandText = sql
                Log(sql)
                dr = cmd.ExecuteReader

                While dr.Read
                    Dim fld As New clsField
                    '/// Create a new field for each chosen column in the sys table
                    '/// Read in all of the properties for each field
                    If ObjDML.DSNtype = enumODBCtype.DB2 Then
                        fld.FieldName = col
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetVal(dr("TYPENAME")))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("LENGTH")))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("SCALE")))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetVal(dr("NULLS")))
                        '/// SYNC field sequence numbers with column numbers so fields stay in order
                        fld.SeqNo = GetVal(dr("COLNO"))

                    ElseIf ObjDML.DSNtype = enumODBCtype.ORACLE Then
                        fld.FieldName = col
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetVal(dr("DATA_TYPE")))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("DATA_LENGTH")))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("DATA_SCALE")))
                        '/// Oracle puts in Nulls instead of ZEROs for scale, so correct this here
                        If fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) = Nothing Then
                            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, 0)
                        End If
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetVal(dr("NULLABLE")))
                        '/// SYNC field sequence numbers with column numbers so fields stay in order
                        fld.SeqNo = GetVal(dr("COLUMN_ID"))

                    ElseIf ObjDML.DSNtype = enumODBCtype.SQL_SERVER Then
                        fld.FieldName = col
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetSQLsvrDataType(GetVal(dr("Data_Type"))))
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("character_maximum_length")))
                        '/// SQL server puts in Nulls instead of ZEROs for length, so correct this here
                        If fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) = Nothing Then
                            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, 0)
                        End If
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, GetVal(dr("numeric_scale")))
                        '/// SQL server puts in Nulls instead of ZEROs for scale, so correct this here
                        If fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) = Nothing Then
                            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_SCALE, 0)
                        End If
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_CANNULL, GetVal(dr("is_nullable")))
                        fld.SeqNo = GetVal(dr("ordinal_position"))

                    Else
                        MsgBox("There are no fields in you selected table", MsgBoxStyle.Information, MsgTitle)
                        Return Nothing
                        Exit Try
                    End If

                    '/// initialize these settings to defaults
                    '/// these do not apply to Relational Databases
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_NCHILDREN, 0)
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LEVEL, 0)
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_TIMES, 1)
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OCCURS, 0)
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_OFFSET, 0)
                    fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_ISKEY, "no")
                    '/// Set Field Parents
                    fld.Parent = Objthis
                    fld.ParentName = Objthis.Text
                    '/// Add the field to the structure
                    Objthis.ObjFields.Add(fld)
                End While
                dr.Close()
            Next

            '/// This select will insert Schema.TableName for fpath1 to know the origin of the DML
            Objthis.fPath1 = ObjDML.Schema.Trim & "." & ObjDML.TableName.Trim

            '/// return the new structure Object to the Main form
            MakeDMLStructure = Objthis

        Catch ex As Exception
            LogError(ex, "modObjects MakeDMLStructure")
            Return Nothing
        Finally
            conn.Close()
        End Try

    End Function

    Function GetSQLsvrDataType(ByVal obj As Object) As Object

        Try
            Select Case obj.ToString
                Case "INT", "int"
                    obj = "INTEGER"
                Case "TINYINT", "tinyint"
                    obj = "SMALLINT"
                Case "DATETIME", "datetime"
                    obj = "TIMESTAMP"
                Case "MONEY", "money"
                    obj = "DECIMAL"
                Case Else
                    obj = obj
            End Select

            Return obj

        Catch ex As Exception
            LogError(ex, "modObjects GetSQLsvrDataType")
            Return Nothing
        End Try

    End Function

#Region "MyListener Class"
    ' NativeWindow class to listen to operating system messages. 
    Public Class MyListener
        Inherits NativeWindow

        Public Event MyScroll(ByVal sender As Object, ByVal e As EventArgs)
        Const WM_MOUSEACTIVATE As Integer = &H21
        Const WM_MOUSEMOVE As Integer = &H200
        Private ctrl As Control

        Public Sub New(ByVal ctrl As Control)
            AssignHandle(ctrl.Handle)
        End Sub

        Protected Overrides Sub WndProc(ByRef m As Message)
            ' Listen for operating system messages 
            Const WM_HSCROLL As Integer = &H114
            Const WM_VSCROLL As Integer = &H115

            If m.Msg = WM_HSCROLL Or m.Msg = WM_VSCROLL Then
                RaiseEvent MyScroll(ctrl, New EventArgs)
            End If

            MyBase.WndProc(m)
        End Sub

        Protected Overrides Sub Finalize()
            ReleaseHandle()
            MyBase.Finalize()
        End Sub
    End Class

#End Region

#Region "ObjectLoad"

    Function LoadProject(ByVal obj As clsProject, Optional ByVal Renamed As Boolean = False, Optional ByVal Updated As Boolean = False) As Boolean
        '// Modified 2/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp")
        Dim i As Integer
        Dim sql As String
        Dim MsgClear As Boolean

        Try
            cnn = New System.Data.Odbc.OdbcConnection(obj.Project.MetaConnectionString)
            cnn.Open()
            Dim cNode As TreeNode = Nothing

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            obj.LoadItems()

            '// Optional Passed Vars added By Tom Karasch for change(i.e. delete, etc...) and rename
            '/// to reload project according to what occured in the object tree.
            If Renamed = False Then
                If Updated = False Then  '// opening new project
                    '//Add project node
                    'cNode = AddNode(tvExplorer.Nodes, NODE_PROJECT, obj)
                    obj.SeqNo = cNode.Index '//store position
                    IsEventFromCode = True
                    'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    IsEventFromCode = False
                    MsgClear = True
                Else   '// Reloading project and tree to reflect MetaData Changes
                    cNode = obj.ObjTreeNode
                    cNode.Nodes.Clear()
                    IsEventFromCode = True
                    'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    IsEventFromCode = False
                    MsgClear = False
                End If
            Else
                '// Obj is existing project with renamed objects so refill with name changes
                cNode = obj.ObjTreeNode
                cNode.Nodes.Clear()
                cNode.Text = obj.Project.ProjectName
                'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                CType(obj, clsProject).Save()
                CType(obj, clsProject).SaveToRegistry()
                MsgClear = False
            End If

            'ShowUsercontrol(cNode, MsgClear)
            'tvExplorer.SelectedNode = cNode

            '//Now add and process each environment 
            If obj.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblEnvironments & " where ProjectName=" & obj.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,ENVIRONMENTDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblEnvironments & " where ProjectName=" & obj.GetQuotedText
                Log(sql)
            End If

            'Log(obj.Project.MetaConnectionString)    //// commented so Password is not in the log
            'Log("cnn.connectionstring >> " & cnn.ConnectionString)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            '//Add Environment folder node
            Dim objF As clsFolderNode
            objF = New clsFolderNode("Environments", NODE_FO_ENVIRONMENT)
            objF.Parent = CType(cNode.Tag, INode)
            cNode = AddNode(cNode.Nodes, objF.Type, objF)
            objF.SeqNo = cNode.Index '//store position

            For i = 0 To dt.Rows.Count - 1
                '//Process this env
                LoadEnv(cNode, dt.Rows(i), cnn)
            Next


            cNode = cNode.Parent
            cNode.EnsureVisible()


        Catch ex As Exception
            LogError(ex, "fillProject")
        Finally
            ShowStatusMessage("")
        End Try

    End Function

    Function LoadEnv(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim i As Integer

        Dim obj As New clsEnvironment
        Dim sql As String = ""
        Dim Dir As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent '//Project->[Env Folder]->Env
            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
                obj.LocalDTDDir = GetVal(dr.Item("LocalDTDDir"))
                obj.LocalDDLDir = GetVal(dr.Item("LocalDDLDir"))
                obj.LocalCobolDir = GetVal(dr.Item("LocalCobolDir"))
                obj.LocalCDir = GetVal(dr.Item("LocalCDir"))
                obj.LocalScriptDir = GetVal(dr.Item("LocalScriptDir"))
                obj.LocalDBDDir = GetVal(dr.Item("LocalModelDir"))

                obj.EnvironmentDescription = GetVal(dr.Item("Description"))

            Else
                obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
                obj.EnvironmentDescription = GetVal(dr.Item("ENVIRONMENTDESCRIPTION"))
            End If


            AddToCollection(obj.Project.Environments, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            '/// Added 6/4/07 by TK
            obj.GetStructureDir()

        Catch ex As Exception
            LogError(ex, "fillEnv")
        End Try
        '///////////////////////////////////////////////
        '//Now add connections
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under Env
            Dim cNodeCnn As TreeNode
            Dim objCnn As INode

            objCnn = New clsFolderNode("Connections", NODE_FO_CONNECTION)
            objCnn.Parent = CType(cNode.Tag, INode)
            cNodeCnn = AddNode(cNode.Nodes, objCnn.Type, objCnn)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & objCnn.Project.tblConnections & " where EnvironmentName=" & obj.GetQuotedText & " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,CONNECTIONNAME,CONNECTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objCnn.Project.tblConnections & _
                            " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName =" & obj.GetQuotedText

                Log(sql)
            End If


            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillConn(cNodeCnn, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "fillConn")
        End Try
        '///////////////////////////////////////////////
        '//Now add structures
        '///////////////////////////////////////////////
        Try
            '//Add Structure folder node under env
            Dim cNodeStruct As TreeNode
            Dim objStruct As INode
            objStruct = New clsFolderNode("Descriptions", NODE_FO_STRUCT)
            objStruct.Parent = CType(cNode.Tag, INode)
            cNodeStruct = AddNode(cNode.Nodes, objStruct.Type, objStruct)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblStructures & _
                " where EnvironmentName=" & obj.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText & _
                " order by structuretype, structureName"
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONTYPE,DESCRIPTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblDescriptions & _
                " where ProjectName=" & obj.Project.GetQuotedText & _
                " AND EnvironmentName = " & obj.GetQuotedText & _
                " order by DESCRIPTIONTYPE, DescriptionName"
                Log(sql)
            End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New System.Data.DataTable("temp2")

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                FillStruct(cNodeStruct, dt.Rows(i), cnn)
            Next

            cNodeStruct.Expand()

        Catch ex As Exception
            LogError(ex, "fillEnv-fillStr")
        End Try

       
        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add variables folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode.Nodes, objVar.Type, objVar, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from variables where EngineName=" & obj.GetQuotedText & _
                " AND ENGINENAME=" & Quote(DBNULL) & _
                " AND SystemName=" & Quote(DBNULL) & _
                " AND EnvironmentName=" & obj.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables " & _
                            "where EnvironmentName=" & obj.GetQuotedText & _
                            " AND ProjectName=" & obj.Project.GetQuotedText & _
                            " AND SYSTEMNAME=" & Quote(DBNULL) & _
                            " AND ENGINENAME=" & Quote(DBNULL) & _
                            " ORDER BY VARIABLENAME"
                Log(sql)
            End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVar(cNodeVar, dt.Rows(i), cnn, True)
            Next
        Catch ex As Exception
            LogError(ex, "modObjects LoadEngine>VariableLoad")
        End Try



        '///////////////////////////////////////////////
        '//Now add systems
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Systems", NODE_FO_SYSTEM)
            objSys.Parent = CType(cNode.Tag, INode)
            cNodeSys = AddNode(cNode.Nodes, objSys.Type, objSys)

            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblSystems & _
                " where EnvironmentName=" & obj.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,SYSTEMDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblSystems & " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName=" & obj.GetQuotedText
                Log(sql)
            End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillSys(cNodeSys, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "FillEnv-FillSys")
        End Try

    End Function

    Function FillStruct(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsStructure
        Dim sql As String = ""
        '///////////////////////////////////////////////
        '//Construct structure object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent
            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.StructureName = GetVal(dr.Item("StructureName"))
                obj.StructureType = GetVal(dr.Item("StructureType"))
                obj.fPath1 = GetVal(dr.Item("fPath1"))
                obj.fPath2 = GetVal(dr.Item("fPath2"))
                obj.IMSDBName = GetVal(dr.Item("IMSDBName"))
                obj.SegmentName = GetVal(dr.Item("SegmentName"))
                obj.StructureDescription = GetVal(dr.Item("Description"))
                obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->StructFolder->Struct
                Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
                If ConnName <> "" Then
                    For Each conn As clsConnection In obj.Environment.Connections
                        If conn.Text = ConnName Then
                            obj.Connection = conn
                        End If
                    Next
                End If
            Else
                obj.StructureName = GetVal(dr.Item("DescriptionName"))
                obj.StructureDescription = GetStr(GetVal(dr.Item("DescriptionDescription")))
                obj.StructureType = GetVal(dr.Item("DescriptionTYPE"))
            End If


            AddToCollection(obj.Environment.Structures, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddStructNode(obj, cNode)
            cNode.Expand()

        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

        '///////////////////////////////////////////////
        '//Now add fieldselections
        '///////////////////////////////////////////////
        Try
            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblStructSel & " where StructureName=" & obj.GetQuotedText & _
                " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & _
                obj.Environment.Project.GetQuotedText & " order by selectionname"
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,SELECTIONNAME,ISSYSTEMSEL,SELECTDESCRIPTION from " & obj.Project.tblDescriptionSelect & _
                " where ProjectName=" & obj.Project.GetQuotedText & _
                " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
                " AND DescriptionName = " & obj.GetQuotedText & _
                " order by selectionname"
            End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode is root node under which we add other nodes)
                FillStructSel(cNode, dt.Rows(i), cnn)
            Next
            cNode.Collapse()


        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

    End Function

    Function FillStructSel(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsStructureSelection

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.ObjStructure = CType(cNode.Tag, INode)
            If obj.ObjStructure.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSELECT"))
                obj.SelectionName = GetVal(dr.Item("SelectionName"))
                obj.SelectionDescription = GetVal(dr.Item("Description"))
            Else
                obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSEL"))
                obj.SelectionName = GetVal(dr.Item("SelectionName"))
                obj.SelectionDescription = GetVal(dr.Item("SELECTDescription"))
            End If

            AddToCollection(obj.ObjStructure.StructureSelections, obj, obj.GUID)

            If obj.IsSystemSelection = "1" Then
                obj.ObjStructure.SysAllSelection = obj
                Exit Function '//dont add this node if system selection
            End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            cNode.Collapse()

        Catch ex As Exception
            LogError(ex, "fillstrSel")
        End Try

    End Function

    Function FillSys(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsSystem
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.SystemName = GetVal(dr.Item("SystemName"))
                obj.Port = GetVal(dr.Item("Port"))
                obj.Host = GetVal(dr.Item("Host"))
                obj.QueueManager = GetVal(dr.Item("QMgr"))
                obj.OSType = GetVal(dr.Item("OSType"))

                obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
                obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
                obj.DTDLib = GetVal(dr.Item("DTDLib"))
                obj.DDLLib = GetVal(dr.Item("DDLLib"))

                obj.SystemDescription = GetVal(dr.Item("Description"))
            Else
                obj.SystemName = GetVal(dr.Item("SystemName"))
                obj.SystemDescription = GetVal(dr.Item("SystemDescription"))
            End If

            AddToCollection(obj.Environment.Systems, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex, "fillSys")
        End Try

        '///////////////////////////////////////////////
        '//Now add engines
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Engines", NODE_FO_ENGINE)
            objSys.Parent = CType(cNode.Tag, INode)
            cNodeSys = AddNode(cNode.Nodes, objSys.Type, objSys)

            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & objSys.Project.tblEngines & " where SystemName=" & obj.GetQuotedText & " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & obj.Environment.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,ENGINEDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objSys.Project.tblEngines & _
                            " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
                            " AND SystemName=" & obj.GetQuotedText
                Log(sql)
            End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                FillEngine(cNodeSys, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "fillSys")
        End Try

    End Function

    Function FillEngine(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer
        Dim Dir As String = ""
        Dim obj As New clsEngine
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.EngineName = GetVal(dr.Item("EngineName"))
            obj.ObjSystem = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.ReportEvery = GetVal(dr.Item("ReportEvery"))
                obj.CommitEvery = GetVal(dr.Item("CommitEvery"))
                obj.ReportFile = GetVal(dr.Item("ReportFile"))
                obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
                obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
                obj.DTDLib = GetVal(dr.Item("DTDLib"))
                obj.DDLLib = GetVal(dr.Item("DDLLib"))
                obj.EngineDescription = GetVal(dr.Item("Description"))
                Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
                If ConnName <> "" Then
                    For Each conn As clsConnection In obj.ObjSystem.Environment.Connections
                        If conn.Text = ConnName Then
                            obj.Connection = conn
                        End If
                    Next
                End If
                obj.LoadEngProps()
            Else
                obj.EngineDescription = GetVal(dr.Item("EngineDescription"))
            End If

            AddToCollection(obj.ObjSystem.Engines, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now source datastore
        '///////////////////////////////////////////////
        Try
            Dim cNode1 As TreeNode
            Dim obj1 As INode

            obj1 = New clsFolderNode("Sources", NODE_FO_SOURCEDATASTORE)
            obj1.Parent = CType(cNode.Tag, INode)
            cNode1 = AddNode(cNode.Nodes, obj1.Type, obj1, False)


            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='S' AND EngineName=" & _
                obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
                obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
                " where projectname=" & Quote(obj.Project.ProjectName) & _
                " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
                " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
                " and enginename=" & Quote(obj.EngineName) & _
                " and DSDIRECTION='S'" & _
                " order by datastorename"
                Dir = "S"
            End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()



            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStore(cNode1, dt.Rows(i), cnn, Dir)
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add targets
        '///////////////////////////////////////////////
        Try
            Dim cNode2 As TreeNode
            Dim obj2 As INode


            obj2 = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
            obj2.Parent = CType(cNode.Tag, INode)
            cNode2 = AddNode(cNode.Nodes, obj2.Type, obj2, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='T' AND EngineName=" & _
                obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
                obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
                " where projectname=" & Quote(obj.Project.ProjectName) & _
                " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
                " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
                " and enginename=" & Quote(obj.EngineName) & _
                " and DSDIRECTION='T'" & _
                " order by datastorename"
                Dir = "T"
            End If
            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                FillDataStore(cNode2, dt.Rows(i), cnn, Dir)
            Next

        Catch ex As Exception
            LogError(ex, sql)
        End Try

        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode.Nodes, objVar.Type, objVar, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
                Log(sql)
            End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                FillVar(cNodeVar, dt.Rows(i), cnn)
            Next
        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>VariableLoad")
        End Try

        '///////////////////////////////////////////////
        '//Now add tasks (main, join, lookup, proc)
        '///////////////////////////////////////////////
        Try
            Dim cNodeMain As TreeNode
            'Dim cNodeLookup As TreeNode
            'Dim cNodeJoin As TreeNode
            Dim cNodeProc As TreeNode
            Dim objMain As INode
            'Dim objLook As INode
            'Dim objGen As INode
            Dim objProc As INode
            Dim ttype As enumTaskType

            '//Add Join folder
            'objJoin = New clsFolderNode("Join", NODE_FO_JOIN)
            'objJoin.Parent = CType(cNode.Tag, INode)
            'cNodeJoin = AddNode(cNode.Nodes, objJoin.Type, objJoin, False)

            '//Add Proc folder
            objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
            objProc.Parent = CType(cNode.Tag, INode)
            cNodeProc = AddNode(cNode.Nodes, objProc.Type, objProc, False)

            '//Add Lookup folder
            'objLook = New clsFolderNode("Lookup", NODE_FO_LOOKUP)
            'objLook.Parent = CType(cNode.Tag, INode)
            'cNodeLookup = AddNode(cNode.Nodes, objLook.Type, objLook, False)

            '//Add Main folder
            objMain = New clsFolderNode("Main Procedure(s)", NODE_FO_MAIN)
            objMain.Parent = CType(cNode.Tag, INode)
            cNodeMain = AddNode(cNode.Nodes, objMain.Type, objMain, False)
            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                If obj.ObjSystem IsNot Nothing Then
                    sql = "Select * from " & obj.Project.tblTasks & _
                    " where EngineName=" & obj.GetQuotedText & _
                    " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
                    " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
                    " AND ProjectName=" & obj.Project.GetQuotedText & _
                    " Order by SeqNo"
                Else
                    sql = "Select * from " & obj.Project.tblTasks & _
                    " where EnvironmentName = " & obj.ObjSystem.Environment.GetQuotedText & _
                    " AND ProjectName=" & obj.Project.GetQuotedText & _
                    " Order by SeqNo"
                End If
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,TASKNAME,TASKTYPE,TASKSEQNO,TASKDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblTasks & _
                " where EngineName=" & obj.GetQuotedText & _
                " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
                " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText & _
                " Order by TASKSEQNO"
            End If

            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode is root node under which we add other nodes)
                ttype = GetVal(dt.Rows(i).Item("TaskType"))
                Select Case ttype
                    Case enumTaskType.TASK_MAIN
                        FillTasks(cNodeMain, dt.Rows(i), cnn)
                    Case enumTaskType.TASK_GEN, enumTaskType.TASK_LOOKUP, enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc
                        FillTasks(cNodeProc, dt.Rows(i), cnn)
                        'Case enumTaskType.TASK_JOIN, enumTaskType.TASK_LOOKUP
                        '    FillTasks(cNodeJoin, dt.Rows(i), cnn)
                        'Case 
                        '    FillTasks(cNodeLookup, dt.Rows(i), cnn)
                End Select
            Next
            cNode.Expand()

            For Each ds As clsDatastore In obj.Sources
                ds.SetIsMapped(False, True)
            Next
            For Each ds As clsDatastore In obj.Targets
                ds.SetIsMapped(False, True)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>fillTasks", sql)
            Return False
        End Try

    End Function

    Function FillTasks(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsTask
        Dim Otype As String

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try

            Otype = CType(cNode.Parent.Tag, INode).Type
            If Otype = NODE_ENVIRONMENT Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->TaskFolder->Any Task
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.TaskName = GetVal(dr.Item("TaskName"))
            obj.TaskType = GetVal(dr.Item("TaskType"))
            If obj.Parent.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.TaskDescription = GetVal(dr.Item("Description"))
            Else
                obj.TaskDescription = GetStr(GetVal(dr("TASKDESCRIPTION")))
            End If

            Select Case obj.TaskType
                'Case modDeclares.enumTaskType.TASK_JOIN, modDeclares.enumTaskType.TASK_LOOKUP
                '    If obj.Engine IsNot Nothing Then
                '        AddToCollection(obj.Engine.Joins, obj, obj.GUID)
                '    Else
                '        AddToCollection(obj.Environment.Joins, obj, obj.GUID)
                '    End If
                Case modDeclares.enumTaskType.TASK_MAIN
                    If obj.Engine IsNot Nothing Then
                        AddToCollection(obj.Engine.Mains, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    End If
                Case modDeclares.enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc, _
                modDeclares.enumTaskType.TASK_GEN, modDeclares.enumTaskType.TASK_LOOKUP
                    If obj.Engine IsNot Nothing Then
                        AddToCollection(obj.Engine.Procs, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    End If
            End Select

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.LoadDatastores()
                obj.LoadMappings()
            End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store treeview node index

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillVar(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal EnvLoad As Boolean = False) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsVariable
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            If EnvLoad = True Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->VarFolder->Var
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.VariableName = GetVal(dr.Item("VariableName"))


            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.VarInitVal = GetVal(dr.Item("VarInitVal"))
                obj.VarSize = GetVal(dr.Item("VarSize"))
                obj.VariableDescription = GetVal(dr.Item("Description"))
            Else
                obj.VariableDescription = GetVal(dr.Item("VariableDescription"))
            End If

            If EnvLoad = True Then
                AddToCollection(obj.Environment.Variables, obj, obj.GUID)
            Else
                AddToCollection(obj.Engine.Variables, obj, obj.GUID)
            End If


            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex, "frmMain fillVar")
        End Try

    End Function

    Function FillConn(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsConnection
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            If CType(cNode.Tag, INode).Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
                obj.ConnectionType = GetVal(dr.Item("ConnectionType"))
                obj.ConnectionDescription = GetVal(dr.Item("Description"))
                obj.Database = GetVal(dr.Item("DBName"))
                obj.UserId = GetVal(dr.Item("UserId"))
                obj.Password = GetVal(dr.Item("Password"))
                obj.DateFormat = GetVal(dr.Item("Dateformat"))
            Else
                obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
                obj.ConnectionDescription = GetVal(dr.Item("ConnectionDescription"))
            End If

            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->ConnFolder->Conn
            AddToCollection(obj.Environment.Connections, obj, obj.GUID)

            cNode = AddNode(cNode.Nodes, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function FillDataStore(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal DSD As String = "") As Boolean

        '/// Modified by Tom Karasch 12/06
        Dim obj As New clsDatastore

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Parent.Tag, INode) 'Engine->Folder(source/target)->ds
            obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

            If DSD = "" Then
                obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
                obj.DsDirection = GetVal(dr.Item("DsDirection"))
                obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
                obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
                'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
                obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
                obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
                'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
                obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

                obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
                obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
                obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.OperationType = GetVal(dr.Item("OperationType"))
                obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
                obj.DsPort = GetVal(dr.Item("Port"))
                obj.DsUOW = GetVal(dr.Item("UOW"))
                obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
                'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
                '/// Load Globals From File
                If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    obj.LoadGlobals()
                End If

                '//If AnyTree Object is renamed, then reload all datastore items to 
                '// Make sure renaming is propagated to all nodes
                obj.LoadItems()
            Else
                obj.DatastoreDescription = GetVal(dr.Item("DATASTOREDESCRIPTION"))

                If obj.DatastoreDescription Is Nothing Then
                    obj.DatastoreDescription = ""
                End If
                If DSD = "S" Then
                    obj.DsDirection = DS_DIRECTION_SOURCE
                Else
                    obj.DsDirection = DS_DIRECTION_TARGET
                End If

                obj.LoadItems(, True)

                obj.LoadAttr()

            End If

            Select Case obj.DsDirection
                Case DS_DIRECTION_SOURCE
                    If obj.Engine Is Nothing Then
                        AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Engine.Sources, obj, obj.GUID)
                    End If
                Case DS_DIRECTION_TARGET
                    If obj.Engine Is Nothing Then
                        AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Engine.Targets, obj, obj.GUID)
                    End If
            End Select

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode.Nodes, obj.Type, obj, False, obj.Text)
            obj.SeqNo = cNode.Index '//store position

            cNode.Text = obj.DsPhysicalSource

            '// now add Datastore selections to tree
            AddDSstructuresToTree(cNode, obj)

            cNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillDatastore")
            Return False
        End Try

    End Function

    Function FillDSbyType(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal Ver As enumMetaVersion = enumMetaVersion.V2) As Boolean

        Try
            Dim obj As New clsDatastore

            obj.Parent = CType(cNode.Tag, INode).Parent
            obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

            If Ver = enumMetaVersion.V2 Then
                obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
                obj.DsDirection = GetVal(dr.Item("DsDirection"))
                obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
                obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
                'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
                obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
                obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
                'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
                obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

                obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
                obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
                obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.OperationType = GetVal(dr.Item("OperationType"))
                obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
                obj.DsPort = GetVal(dr.Item("Port"))
                obj.DsUOW = GetVal(dr.Item("UOW"))
                obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
                'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
                '/// Load Globals From File
                If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    obj.LoadGlobals()
                End If

                '//If AnyTree Object is renamed, then reload all datastore items to 
                '// Make sure renaming is propagated to all nodes
                obj.LoadItems()
            Else
                obj.DatastoreDescription = GetStr(GetVal(dr.Item("DATASTOREDESCRIPTION")))
                obj.DatastoreType = GetVal(dr.Item("DSTYPE"))

                obj.LoadItems(, True)

                obj.LoadAttr()
            End If


            AddToCollection(obj.Environment.Datastores, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddDSNode(obj, cNode)
            obj.SeqNo = cNode.Index '//store position

            cNode.Text = obj.DsPhysicalSource

            '// now add Datastore selections to tree
            AddDSstructuresToTree(cNode, obj)
            cNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillDSbyType")
            Return False
        End Try

    End Function

    '//This will add Datastore node under proper Category Folder
    '//cNode : is "Datastores" folder node
    '//obj : is object to represent struct node
    '// [Datastores]
    '//     |
    '//    [+]-[<DSType>]
    '//         |
    '//         [+]-DS1
    '//         [+]-DS2
    Function AddDSNode(ByVal obj As clsDatastore, ByVal cNode As TreeNode) As TreeNode
        '//Now look for Datastore type and insert it in proper folder
        Dim nd As TreeNode = Nothing
        Dim IsFound As Boolean

        Try
            For Each nd In cNode.Nodes
                If nd.Text = GetDSFolderText(obj.DatastoreType) Then
                    IsFound = True
                    Exit For
                End If
            Next

            '//if folder for structure category is not found then add new folder 
            '//and then add structure node under it
            If IsFound = True Then
                '//Add struct node under struct category
                cNode = AddNode(nd.Nodes, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            Else
                Dim objFol As INode
                Dim objFolType As String = GetFolType(obj.DatastoreType)

                objFol = New clsFolderNode(GetDSFolderText(obj.DatastoreType), objFolType)

                objFol.Parent = CType(cNode.Parent.Tag, INode)
                '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
                nd = AddNode(cNode.Nodes, objFol.Type, objFol, False)
                '//Add struct node under struct category
                cNode = AddNode(nd.Nodes, obj.Type, obj, False, obj.DatastoreName)
                obj.SeqNo = cNode.Index '//store position
            End If
            Return cNode

        Catch ex As Exception
            LogError(ex, "frmMain AddDSNode")
            Return Nothing
        End Try

    End Function

    Function GetDSFolderText(ByVal DSType As enumDatastore) As String

        Select Case DSType
            Case modDeclares.enumDatastore.DS_BINARY
                Return "BINARY"
            Case modDeclares.enumDatastore.DS_DB2CDC
                Return "DB2CDC"
            Case modDeclares.enumDatastore.DS_DB2LOAD
                Return "DB2LOAD"
            Case modDeclares.enumDatastore.DS_DELIMITED
                Return "DELIMITED"
            Case modDeclares.enumDatastore.DS_GENERICCDC
                Return "GENERICCDC"
            Case modDeclares.enumDatastore.DS_HSSUNLOAD
                Return "HSSUNLOAD"
            Case modDeclares.enumDatastore.DS_IBMEVENT
                Return "IBMEVENT"
                'Case modDeclares.enumDatastore.DS_IMSCDC
                '    Return "IMSCDC"
            Case modDeclares.enumDatastore.DS_IMSDB
                Return "IMSDB"
                'Case modDeclares.enumDatastore.DS_IMSLE
                '    Return "IMSLE"
                'Case modDeclares.enumDatastore.DS_IMSLEBATCH
                '    Return "IMSLEBATCH"
            Case modDeclares.enumDatastore.DS_INCLUDE
                Return "INCLUDE"
            Case modDeclares.enumDatastore.DS_ORACLECDC
                Return "ORACLECDC"
            Case modDeclares.enumDatastore.DS_RELATIONAL
                Return "RELATIONAL"
            Case modDeclares.enumDatastore.DS_SUBVAR
                Return "SUBVAR"
            Case modDeclares.enumDatastore.DS_TEXT
                Return "TEXT"
                'Case modDeclares.enumDatastore.DS_TRBCDC
                '    Return "TRBCDC"
            Case modDeclares.enumDatastore.DS_UNKNOWN
                Return "UNKNOWN"
            Case modDeclares.enumDatastore.DS_VSAM
                Return "VSAM"
            Case modDeclares.enumDatastore.DS_VSAMCDC
                Return "VSAMCDC"
            Case modDeclares.enumDatastore.DS_XML
                Return "XML"
                'Case modDeclares.enumDatastore.DS_XMLCDC
                '    Return "XMLCDC"
            Case Else
                Return "Unknown"
        End Select

    End Function

    Function GetFolType(ByVal DStype As enumDatastore) As String

        Select Case DStype
            Case enumDatastore.DS_BINARY
                Return DS_BINARY
            Case enumDatastore.DS_DB2CDC
                Return DS_DB2CDC
            Case enumDatastore.DS_DB2LOAD
                Return DS_DB2LOAD
            Case enumDatastore.DS_DELIMITED
                Return DS_DELIMITED
            Case enumDatastore.DS_GENERICCDC
                Return DS_GENERICCDC
            Case enumDatastore.DS_HSSUNLOAD
                Return DS_HSSUNLOAD
            Case enumDatastore.DS_IBMEVENT
                Return DS_IBMEVENT
            Case enumDatastore.DS_IMSCDCLE
                Return DS_IMSCDCLE
            Case enumDatastore.DS_IMSDB
                Return DS_IMSDB
                'Case enumDatastore.DS_IMSLE
                '    Return DS_IMSLE
                'Case enumDatastore.DS_IMSLEBATCH
                '    Return DS_IMSLEBATCH
            Case enumDatastore.DS_INCLUDE
                Return DS_INCLUDE
            Case enumDatastore.DS_ORACLECDC
                Return DS_ORACLECDC
            Case enumDatastore.DS_RELATIONAL
                Return DS_RELATIONAL
            Case enumDatastore.DS_SUBVAR
                Return DS_SUBVAR
            Case enumDatastore.DS_TEXT
                Return DS_TEXT
                'Case enumDatastore.DS_TRBCDC
                '    Return DS_TRBCDC
            Case enumDatastore.DS_UNKNOWN
                Return DS_UNKNOWN
            Case enumDatastore.DS_VSAM
                Return DS_VSAM
            Case enumDatastore.DS_VSAMCDC
                Return DS_VSAMCDC
            Case enumDatastore.DS_XML
                Return DS_XML
                'Case enumDatastore.DS_XMLCDC
                '    Return DS_XMLCDC
            Case Else
                Return NODE_FO_DATASTORE
        End Select

    End Function

    Function AddDSstructuresToTree(ByVal pNode As TreeNode, ByVal obj As clsDatastore, Optional ByVal MapAs As Boolean = False) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            pNode.Nodes.Clear()
            obj.LoadItems(, True)


            For i = 0 To obj.ObjSelections.Count - 1
                Dim DSselobj As clsDSSelection = obj.ObjSelections(i)

                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                DSselobj.LoadItems(, True)
                'End If

                If Not DSselobj.IsChildDSSelection = True Or MapAs = True Then
                    cnode = AddNode(pNode.Nodes, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    Call addDSSelChildrenToTree(cnode, DSselobj)
                End If
            Next

            pNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "main AddDSstructuresToTree")
            Return False
        End Try

    End Function

    Private Function addDSSelChildrenToTree(ByRef pNode As TreeNode, ByRef obj As clsDSSelection) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            For i = 0 To obj.ObjDSSelections.Count - 1
                Dim DSselobj As clsDSSelection
                DSselobj = obj.ObjDSSelections(i)
                DSselobj.LoadItems(, True)
                If DSselobj.Parent Is obj Then
                    cnode = AddNode(pNode.Nodes, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    Call addDSSelChildrenToTree(cnode, DSselobj) '// recurse for all children
                End If
            Next
            pNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "main addDSselChildrenToTree")
            Return False
        End Try

    End Function

    '//This will add Structure node under proper Category Folder
    '//cNode : is "Structures" folder node
    '//obj : is object to represent struct node
    '// [Structures]
    '//     |
    '//    [+]-[<category>]
    '//         |
    '//         [+]-Struct1
    '//         [+]-struct2
    Function AddStructNode(ByVal obj As clsStructure, ByVal cNode As TreeNode) As TreeNode
        '//Now look for Structure type and insert it in proper folder
        Dim nd As TreeNode = Nothing
        Dim IsFound As Boolean

        Try
            For Each nd In cNode.Nodes
                If nd.Text = GetStructureFolderText(obj.StructureType) Then
                    IsFound = True
                    Exit For
                End If
            Next

            '//if folder for structure category is not found then add new folder 
            '//and then add structure node under it
            If IsFound = True Then
                '//Add struct node under struct category
                cNode = AddNode(nd.Nodes, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            Else
                Dim objFol As INode
                objFol = New clsFolderNode(GetStructureFolderText(obj.StructureType), NODE_FO_STRUCT)
                objFol.Parent = CType(cNode.Parent.Tag, INode)
                '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
                nd = AddNode(cNode.Nodes, objFol.Type, objFol, False)
                '//Add struct node under struct category
                cNode = AddNode(nd.Nodes, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            End If
            Return cNode

        Catch ex As Exception
            Log("AddStructNode=>" & ex.Message)
            Return Nothing
        End Try

    End Function

    Function GetStructureFolderText(ByVal stType As enumStructure) As String

        Select Case stType
            Case modDeclares.enumStructure.STRUCT_C
                Return "CHeader"
            Case modDeclares.enumStructure.STRUCT_COBOL
                Return "COBOL"
            Case modDeclares.enumStructure.STRUCT_COBOL_IMS
                Return "COBOLIMS"
            Case modDeclares.enumStructure.STRUCT_IMS
                Return "IMS"
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                Return "DDL"
            Case modDeclares.enumStructure.STRUCT_REL_DML, enumStructure.STRUCT_REL_DML_FILE
                Return "DML"
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                Return "XMLDTD"
            Case Else
                Return "Unknown"
        End Select

    End Function

    '//////////////////////////////////////////////////////////////////
    '//// Now the tree is built from loaded data in the MetaDataDB ////
    '//////////////////////////////////////////////////////////////////

#End Region

End Module
