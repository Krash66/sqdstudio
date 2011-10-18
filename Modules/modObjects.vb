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
    'Function SearchFieldByName(ByRef objEnv As clsEnvironment, ByRef FieldName As String, _
    'ByRef selName As String, ByRef ProjectName As String) As clsField

    '    Try
    '        For Each DSobj As clsDatastore In objEnv.Datastores
    '            DSobj.LoadItems()
    '            For Each SelObj As clsDSSelection In DSobj.ObjSelections
    '                SelObj.LoadMe()
    '                For Each FldObj As clsField In SelObj.DSSelectionFields
    '                    If FldObj.FieldName = FieldName _
    '                    And FldObj.Struct.StructureName = selName _
    '                    And FldObj.Struct.Environment.EnvironmentName = objEnv.EnvironmentName _
    '                    And FldObj.Project.ProjectName = ProjectName Then
    '                        Return FldObj
    '                        Exit Function
    '                    End If
    '                Next
    '            Next
    '        Next

    '        '// if it makes it out of the loops
    '        Return Nothing

    '    Catch ex As Exception
    '        LogError(ex, "SearchFieldByName")
    '        Return Nothing
    '    End Try

    'End Function

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
            objSel.LoadMe()
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
    'Function SearchDatastoreByName(ByRef objEnv As clsEnvironment, _
    '            ByRef DatastoreName As String, _
    '            ByRef EnvironmentName As String, _
    '            ByRef ProjectName As String) As clsDatastore

    '    Try
    '        For Each objDS As clsDatastore In objEnv.Datastores
    '            If objDS.DatastoreName = DatastoreName _
    '            And objDS.Environment.EnvironmentName = EnvironmentName _
    '            And objDS.Project.ProjectName = ProjectName Then
    '                Return objDS
    '                Exit Function
    '            End If
    '        Next

    '        Return Nothing

    '    Catch ex As Exception
    '        LogError(ex, "modObjects SearchDatastoreByName")
    '        Return Nothing
    '    End Try

    'End Function

    'added by TK 12/29/2006
    Function SearchDSFieldByName(ByRef objSel As clsDSSelection, ByRef FieldName As String, ByRef StructureName As String, ByVal EnvironmentName As String, ByVal ProjectName As String) As clsField

        'Dim fld As clsField 'new???? try this
        'Dim i As Integer

        Try
            'objSel.ObjStructure.LoadItems()
            For Each fld As clsField In objSel.ObjStructure.ObjFields
                'fld = objSel.DSSelectionFields(i)
                If (fld.FieldName = FieldName) And _
                (fld.Struct.StructureName = StructureName) And _
                (fld.Struct.Environment.EnvironmentName = EnvironmentName) And _
                (fld.Project.ProjectName = ProjectName) Then  'And _(fld.ParentName = objSel.ObjStructure.StructureName)
                    Return fld.Clone(objSel)
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

    Function SearchSSForDSSelRef(ByRef objsel As clsStructureSelection) As Collection

        Try
            Dim TempList As New Collection
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

            SearchSSForDSSelRef = TempList

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
    'Function SearchDatastore(ByRef objEnv As clsEnvironment, ByRef objDs As clsDatastore) As clsDatastore

    '    Try
    '        For Each DS As clsDatastore In objEnv.Datastores
    '            If DS.DatastoreName = objDs.DatastoreName Then
    '                Return DS
    '                Exit Function
    '            End If
    '        Next

    '        Return Nothing

    '    Catch ex As Exception
    '        LogError(ex, "modObjects SearchDatastore")
    '        Return Nothing
    '    End Try

    'End Function

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
                If var.VariableName = objVar.VariableName And var.VariableType = objVar.VariableType Then
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
                If var.VariableName = objVar.VariableName And var.VariableType = objVar.VariableType Then
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
                        objSel.LoadMe(cmd)
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
                        objSel.LoadMe(cmd)
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
            Objthis.StructureName = ObjDML.TableName.Trim  'ObjDML.Schema.Trim & "_" & 
            '/// truncate Structure names to 128 characters so they will fit in Metadata
            If Objthis.StructureName.Length > 128 Then
                Objthis.StructureName = Strings.Left(Objthis.StructureName, 128)
            End If
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
                        sql = "select data_type, character_maximum_length, numeric_precision, numeric_scale, is_nullable, ordinal_position from INFORMATION_SCHEMA.COLUMNS where column_name= '" & col & "' and table_name = '" & ObjDML.TableName & "' order by ordinal_position"
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
                        fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_DATATYPE, GetVal(dr("Data_Type"))) 'GetSQLsvrDataType(
                        Dim dtype As String = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                        If dtype = "varchar" Or dtype = "char" Or dtype = "nchar" Or dtype = "nvarchar" _
                            Or dtype = "binary" Or dtype = "varbinary" Then
                            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("character_maximum_length")))
                        Else
                            fld.SetSingleFieldAttr(enumFieldAttributes.ATTR_LENGTH, GetVal(dr("numeric_precision")))
                        End If
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

End Module
