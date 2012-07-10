Imports System.IO
Module modModeler
    '////////////////////////////////////////////////////
    '/// Created Sept. 2010 by Tom Karasch            ///
    '/// This module generates Models of Descriptions ///
    '/// in various formats eg. DDL, COBOL, C, XML    ///
    '/// and replaces sqdumodl.exe                    ///
    '////////////////////////////////////////////////////

    'One of these three will come in ONLY
    Dim ObjSel As clsStructureSelection = Nothing
    Dim ObjStr As clsStructure = Nothing
    Dim ObjDSSel As clsDSSelection = Nothing

    '/// Field array
    Dim FldArray As New ArrayList

    Private semi As String = ";"

    Private success As Boolean = True
    Dim InObj As INode = Nothing
    Dim InInodeType As String
    Dim InStructureType As String
    Dim OutPath As String = ""
    Dim NameModl As String = ""
    Dim TableName As String = ""
    Dim SchemaName As String = ""
    Dim idx As Integer
    Dim Len As Integer
    Dim OutType As String = ""
    Dim FileExt As String = ""
    Dim FullPathModl As String = ""
    Dim ErrFile As String = GetAppLog() & "ModelErr.Log"


    'FileStreams and stream writers
    Dim fsOut As System.IO.FileStream
    Dim objWriteOut As System.IO.StreamWriter
    Dim fsERR As System.IO.FileStream
    Dim objWriteERR As System.IO.StreamWriter

    'Future Use
    Dim DBDalias As String
    Dim DBDfile As String

    '/// Enum for open and closed braces and brackets
    Enum OpenClose As Integer
        OPEN = 0
        CLOSE = 1
    End Enum

#Region "Main Processes"

    Public Function ModelStructure(ByVal Obj As INode, ByVal TypeToModl As String, ByVal ModName As String, ByVal ModPath As String) As String

        Try
            Log("********* Modeler Start *********")
            '/// set module-wide variables
            InInodeType = Obj.Type
            NameModl = ModName
            OutType = TypeToModl
            FileExt = getFileExt()
            OutPath = ModPath
            SchemaName = ""

            ObjSel = Nothing
            ObjStr = Nothing
            ObjDSSel = Nothing


            '/// create new, complete model path
            If Strings.Right(OutPath, 1) <> "\" Then
                FullPathModl = OutPath & "\" & NameModl & FileExt
            Else
                FullPathModl = OutPath & NameModl & FileExt
            End If

            '/// create new model file
            success = CreateFile()

            If success Then
                '/// determine input object
                Select Case InInodeType
                    Case NODE_STRUCT
                        ObjStr = CType(Obj, clsStructure)
                        TableName = ObjStr.StructureName
                        ObjStr.LoadMe()
                        ObjStr.LoadItems()
                        InStructureType = ObjStr.StructureType
                        If InStructureType = enumStructure.STRUCT_REL_DML Then
                            idx = ObjStr.fPath1.LastIndexOf(".")
                            Len = ObjStr.fPath1.Length
                            SchemaName = Strings.Left(ObjStr.fPath1, idx)
                            TableName = Strings.Right(ObjStr.fPath1, Len - (idx + 1))
                        End If
                    Case NODE_STRUCT_SEL
                        ObjSel = CType(Obj, clsStructureSelection)
                        TableName = ObjSel.ObjStructure.StructureName
                        ObjSel.LoadMe()
                        ObjSel.LoadItems()
                        InStructureType = ObjSel.ObjStructure.StructureType
                        If InStructureType = enumStructure.STRUCT_REL_DML Then
                            idx = ObjSel.ObjStructure.fPath1.LastIndexOf(".")
                            Len = ObjSel.ObjStructure.fPath1.Length
                            SchemaName = Strings.Left(ObjSel.ObjStructure.fPath1, idx)
                            TableName = Strings.Right(ObjSel.ObjStructure.fPath1, Len - (idx + 1))
                        End If
                    Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                        ObjDSSel = CType(Obj, clsDSSelection)
                        TableName = ObjDSSel.SelectionName
                        ObjDSSel.LoadMe()
                        ObjDSSel.LoadItems()
                        InStructureType = ObjDSSel.ObjStructure.StructureType
                    Case Else
                        success = False
                        GoTo ErrorGoTo
                End Select
            End If

            If success Then
                success = PopulateFldArrayFromObject()
            End If

            If success Then
                Select Case OutType
                    Case "DB2DDL"
                        success = objModelDB2DDL()
                    Case "ORADDL"
                        success = objModelORADDL()
                    Case "SQLDDL"
                        success = objModelSQLDDL()
                    Case "SQDDDL"
                        success = objModelSQDDDL()
                    Case "H"
                        success = objModelC()
                    Case "DTD"
                        success = objModelDTD()
                    Case "LOD"
                        success = objModelLOD()
                    Case "SQL"
                        success = objModelSQL()
                    Case "MSSQL"
                        success = objModelMSSQL()
                    Case "DB2"
                        success = objModelDB2Trig()
                    Case Else
                        success = False
                        GoTo ErrorGoTo
                End Select
            End If

            If success Then
                Return FullPathModl
                Log("********* Modeler Finished *********")
            Else
ErrorGoTo:      '/// on errors
                Return ""
                Log("********* Modeler Finished with Errors *********")
            End If

        Catch ex As Exception
            LogError(ex, "modModeler ModelStructure")
            Return ""
        Finally
            objWriteOut.Close()
            fsOut.Close()
        End Try

    End Function

    Function ModelStructures(ByVal Col As Collection, ByVal TypeToModl As String, ByVal ModName As String, ByVal ModPath As String, ByVal typeIn As String) As String

        Try
            Log("********* Modeler Start *********")
            '/// set module-wide variables
            InInodeType = typeIn
            NameModl = ModName
            OutType = TypeToModl
            FileExt = getFileExt()
            OutPath = ModPath
            SchemaName = ""

            '/// create new, complete model path
            If Strings.Right(OutPath, 1) <> "\" Then
                FullPathModl = OutPath & "\" & NameModl & FileExt
            Else
                FullPathModl = OutPath & NameModl & FileExt
            End If

            '/// create new model file
            success = CreateFile()

            For Each ObjStr In Col
                TableName = ObjStr.StructureName
                ObjStr.LoadMe()
                ObjStr.LoadItems()
                If ObjStr.StructureType = enumStructure.STRUCT_REL_DML Then
                    idx = ObjStr.fPath1.LastIndexOf(".")
                    Len = ObjStr.fPath1.Length
                    SchemaName = Strings.Left(ObjStr.fPath1, idx)
                    TableName = Strings.Right(ObjStr.fPath1, Len - (idx + 1))
                End If
                If success Then
                    success = PopulateFldArrayFromObject()
                End If
                If success Then
                    Select Case OutType
                        Case "DB2DDL"
                            success = objModelDB2DDL()
                        Case "ORADDL"
                            success = objModelORADDL()
                        Case "SQLDDL"
                            success = objModelSQLDDL()
                        Case "H"
                            success = objModelC()
                        Case "DTD"
                            success = objModelDTD()
                        Case "LOD"
                            success = objModelLOD()
                        Case "SQL"
                            success = objModelSQL()
                        Case "MSSQL"
                            success = objModelMSSQL()
                        Case "DB2"
                            success = objModelDB2Trig()
                        Case Else
                            success = False
                            GoTo NextStr
                    End Select
                End If
                wBlankLine()
                wBlankLine()
NextStr:
            Next


            If success Then
                Return FullPathModl
                Log("********* Modeler Finish *********")
            Else
ErrorGoTo2:     '/// on errors
                Return ""
                Log("********* Modeler Finish with Errors *********")
            End If

        Catch ex As Exception
            LogError(ex, "modModeler ModelStructures")
            Return ""
        Finally
            objWriteOut.Close()
            fsOut.Close()
        End Try

    End Function

#End Region

#Region "Create Output Files"

    Function CreateFile() As Boolean

        Try
            '//Create new file and filestream for Output file
            If System.IO.File.Exists(FullPathModl) Then
                System.IO.File.Delete(FullPathModl)
            End If
            fsOut = System.IO.File.Create(FullPathModl)
            objWriteOut = New System.IO.StreamWriter(fsOut)

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler CreateFile")
            Return False
        End Try

    End Function

#End Region

    '/// Take info from input objects and send to proper output path
#Region "Object Functions"

    Function PopulateFldArrayFromObject() As Boolean

        Try
            FldArray.Clear()

            If ObjDSSel IsNot Nothing Then
                For Each fld As clsField In ObjDSSel.DSSelectionFields
                    FldArray.Add(fld)
                Next
            ElseIf ObjSel IsNot Nothing Then
                For Each fld As clsField In ObjSel.ObjSelectionFields
                    FldArray.Add(fld)
                Next
            ElseIf ObjStr IsNot Nothing Then
                For Each fld As clsField In ObjStr.ObjFields
                    FldArray.Add(fld)
                Next
            Else
                Return False
                Exit Try
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler PopulateFromObject")
            Return False
        End Try

    End Function

#End Region

#Region "Output Logic by Model Type"

#Region "DDL"

    Function objModelDB2DDL() As Boolean

        Try
            Dim fldLen As Integer
            Dim fldScl As Integer
            Dim FORcreate As String = String.Format("{0}{1}{2}", "CREATE ", "TABLE ", RDash(TableName))
            Dim FORKey As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim first As Boolean = True

            objWriteOut.WriteLine(FORcreate)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "," & RDash(fld.FieldName)
                End If
                fldLen = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "DB2DDL", fldLen, fldScl)


                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Or _
                fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL) = "No" Then
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}{3,-12}", " ", NameFld, OutAttr, "NOT NULL")
                Else
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}", " ", NameFld, OutAttr)
                End If
                objWriteOut.WriteLine(FORKey)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelDB2DDL")
            Return False
        End Try

    End Function

    Function objModelORADDL() As Boolean

        Try
            Dim fldLen As Integer
            Dim fldScl As Integer
            Dim FORcreate As String = String.Format("{0}{1}{2}", "CREATE ", "TABLE ", RDash(TableName))
            Dim FORKey As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim first As Boolean = True

            objWriteOut.WriteLine(FORcreate)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "," & RDash(fld.FieldName)
                End If
                fldLen = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "ORADDL", fldLen, fldScl)


                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Or _
                fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL) = "No" Then
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}{3,-12}", " ", NameFld, OutAttr, "NOT NULL")
                Else
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}", " ", NameFld, OutAttr)
                End If
                objWriteOut.WriteLine(FORKey)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelORADDL")
            Return False
        End Try

    End Function

    Function objModelSQLDDL() As Boolean

        Try
            Dim fldLen As Integer
            Dim fldScl As Integer
            Dim FORcreate As String
            Dim FORKey As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim first As Boolean = True

            If SchemaName <> "" Then
                FORcreate = String.Format("{0}{1}{2}{3}", "CREATE TABLE ", SchemaName, ".", RDash(TableName))
                '"[" & SchemaName & "]" "[" & RDash(TableName) & "]"
            Else
                FORcreate = String.Format("{0}{1}{2}", "CREATE ", "TABLE ", RDash(TableName))   '"[" &  & "]"
            End If

            objWriteOut.Write(FORcreate)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & RDash(fld.FieldName & "  ")   '"[" &  & "] "
                    first = False
                Else
                    NameFld = "," & RDash(fld.FieldName & "  ")   '"[" & fld.FieldName & "] "
                End If
                fldLen = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "SQLDDL", fldLen, fldScl)

                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Or _
                fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL) = "No" Then
                    FORKey = String.Format("{0,4}{1}", " ", NameFld & OutAttr & " NOT NULL")
                Else
                    FORKey = String.Format("{0,4}{1}", " ", NameFld & OutAttr & " NULL")
                End If
                objWriteOut.WriteLine(FORKey)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            'wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelSQLDDL")
            Return False
        End Try

    End Function

    Function objModelSQDDDL() As Boolean

        Try
            Dim fldLen As Integer
            Dim fldScl As Integer
            Dim FORcreate As String = String.Format("{0}{1}{2}", "CREATE ", "TABLE ", RDash(TableName))
            Dim FORKey As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim first As Boolean = True

            objWriteOut.WriteLine(FORcreate)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "," & RDash(fld.FieldName)
                End If
                fldLen = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "SQDDDL", fldLen, fldScl)


                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Or _
                fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL) = "No" Then
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}{3,-12}", " ", NameFld, OutAttr, "NOT NULL")
                Else
                    FORKey = String.Format("{0,4}{1,-36}{2,-20}", " ", NameFld, OutAttr)
                End If
                objWriteOut.WriteLine(FORKey)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelDB2DDL")
            Return False
        End Try

    End Function

#End Region

#Region "C"

    Function objModelC() As Boolean

        Try
            Dim fldLen As String
            Dim fldScl As Integer
            Dim FORcreate As String = String.Format("{0}{1}", "struct ", TableName)
            Dim FORfld As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim OutLen As String
            Dim first As Boolean = True

            objWriteOut.WriteLine(FORcreate)
            wBrace(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto

                NameFld = fld.FieldName
                fldLen = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
                If fldLen < 1 Then
                    OutLen = ""
                Else
                    OutLen = "(" & fldLen & ")"
                End If
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "H", fldLen, fldScl)

                FORfld = String.Format("{0,4}{1,-20}{2,30}{3,6}", " ", OutAttr, NameFld, OutLen)

                objWriteOut.WriteLine(FORfld)
nextgoto:   Next

            wBrace(OpenClose.CLOSE, False)
            wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelC")
            Return False
        End Try

    End Function

#End Region

#Region "DTD"

    Function objModelDTD() As Boolean

        Try
            Dim fldLen As String
            Dim FORcreate As String = String.Format("{0}{1}", "<!ELEMENT ", TableName)
            Dim FORfld As String
            Dim NameFld As String
            Dim fldattr As String
            Dim first As Boolean = True

            objWriteOut.WriteLine(FORcreate)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & fld.FieldName
                    first = False
                Else
                    NameFld = "," & fld.FieldName
                End If
                fldLen = "(" & fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString & ");"
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString

                FORfld = String.Format("{0,4}{1,-30}", " ", NameFld)

                objWriteOut.WriteLine(FORfld)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            objWriteOut.WriteLine(">")
            wBlankLine()

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto1

                NameFld = fld.FieldName
                fldLen = "(" & fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString & ");"
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString

                FORfld = String.Format("{0}{1,-40}{2}", "<!ELEMENT ", NameFld, " (#PCDATA)>")

                objWriteOut.WriteLine(FORfld)
nextgoto1:  Next

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelDTD")
            Return False
        End Try

    End Function

#End Region

#Region "Oracle Load"

    Function objModelLOD() As Boolean

        Try
            Dim fldLen As String
            Dim fldScl As Integer
            Dim FORinfile As String = String.Format("{0}{1}{2}", "INFILE ", "'" & RDash(TableName) & ".dat' ", Chr(34) & "FIX 39" & Chr(34))
            Dim FORInto As String = String.Format("{0}{1}", "INTO TABLE ", RDash(TableName))
            Dim FORfld As String
            Dim NameFld As String
            Dim fldattr As String
            Dim OutAttr As String
            Dim first As Boolean = True
            Dim Pos As Integer = 1

            objWriteOut.WriteLine("LOAD DATA")
            objWriteOut.WriteLine(FORinfile)
            objWriteOut.WriteLine(FORInto)
            wBracket(OpenClose.OPEN, True)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = " " & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "," & RDash(fld.FieldName)
                End If
                fldLen = "(" & Pos.ToString & ":" & fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString & ")"
                fldScl = fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE)
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "ORADDL", 0, fldScl)

                FORfld = String.Format("{0,4}{1,-36}{2,-30}{3}", " ", NameFld, "POSITION" & fldLen, OutAttr)

                objWriteOut.WriteLine(FORfld)
                Pos = Pos + fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH)
nextgoto:   Next

            wBracket(OpenClose.CLOSE, False)
            wSemiLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelLOD")
            Return False
        End Try

    End Function

#End Region

#Region "Oracle Trigger"

    Function objModelSQL() As Boolean

        Try
            Dim FORCreateReplace As String = String.Format("{0}{1}", "CREATE OR REPLACE TRIGGER sqdaudit_", RDash(TableName))
            Dim FORAfter As String = String.Format("{0}{1}", "AFTER INSERT OR DELETE OR UPDATE ON ", RDash(TableName))
            Dim NameFld As String
            Dim first As Boolean = True
            Dim ArrayNoGroupItemCnt As Integer = 0

            '/// Get array count without Groupitems
            For Each fld As clsField In FldArray
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) <> "GROUPITEM" Then
                    ArrayNoGroupItemCnt += 1
                End If
            Next

            objWriteOut.WriteLine(FORCreateReplace)
            objWriteOut.WriteLine(FORAfter)
            objWriteOut.WriteLine("FOR EACH ROW")
            objWriteOut.WriteLine("DECLARE")
            objWriteOut.WriteLine("    op CHAR(1);")
            objWriteOut.WriteLine("BEGIN()")
            objWriteOut.WriteLine("if updating then op := 'R'; END IF;")
            objWriteOut.WriteLine("if inserting then op := 'I'; END IF;")
            objWriteOut.WriteLine("if deleting then op := 'D'; END IF;")
            objWriteOut.WriteLine("INSERT INTO SQDENGCD (")
            objWriteOut.WriteLine("       EYECATCHER()")
            objWriteOut.WriteLine("      ,CHANGEOP")
            objWriteOut.WriteLine("      ,UPDATE_TSTMP")
            objWriteOut.WriteLine("      ,UPDATE_RBA")
            objWriteOut.WriteLine("      ,USER_UPDATED")
            objWriteOut.WriteLine("      ,TABLE_ALIAS")
            objWriteOut.WriteLine("      ,TABLE_UPDATED")
            objWriteOut.WriteLine("      ,NUM_COLS")
            objWriteOut.WriteLine("      ,CDC_AFTER_DATA")
            objWriteOut.WriteLine("      ,CDC_BEFORE_DATA")
            objWriteOut.WriteLine(") VALUES (")
            objWriteOut.WriteLine("      'SQOR'")
            objWriteOut.WriteLine("      ,op")
            objWriteOut.WriteLine("      ,sysdate")
            objWriteOut.WriteLine("      ,DBMS_TRANSACTION.LOCAL_TRANSACTION_ID")
            objWriteOut.WriteLine("      ,user") '
            objWriteOut.WriteLine("      ,'" & RDash(TableName) & "'")
            objWriteOut.WriteLine("      ,'" & RDash(TableName) & "'")
            objWriteOut.WriteLine("      ," & ArrayNoGroupItemCnt.ToString)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = "      ,:NEW." & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "||','||:NEW." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto:   Next
            first = True
            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto1
                If first = True Then
                    NameFld = "      ,:OLD." & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "||','||:OLD." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto1:  Next

            wBracket(OpenClose.CLOSE, False)
            wSemiLine()

            objWriteOut.WriteLine("END sqdaudit_" & RDash(TableName) & ";")
            objWriteOut.WriteLine("/")
            objWriteOut.WriteLine("COMMIT;")

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelSQL(oracle Trigger)")
            Return False
        End Try

    End Function

#End Region

#Region "MSSQLServer Trigger"

    Function objModelMSSQL() As Boolean

        Try
            Dim FORCreateTrig As String = String.Format("{0}{1}", "CREATE TRIGGER sqdaudit_I_", RDash(TableName))
            Dim FORAfter As String = String.Format("{0}{1}", "ON ", RDash(TableName))
            Dim NameFld As String
            Dim firstFld As String = ""
            Dim first As Boolean = True
            Dim ArrayNoGroupItemCnt As Integer = 0

            '/// Get array count without Groupitems
            For Each fld As clsField In FldArray
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) <> "GROUPITEM" Then
                    ArrayNoGroupItemCnt += 1
                End If
            Next

            objWriteOut.WriteLine(FORCreateTrig)
            objWriteOut.WriteLine(FORAfter)
            objWriteOut.WriteLine("AFTER UPDATE")
            objWriteOut.WriteLine("AS")
            objWriteOut.WriteLine("BEGIN")
            objWriteOut.WriteLine("INSERT SQDENGCD")
            objWriteOut.WriteLine("SELECT")
            objWriteOut.WriteLine(" 'SQMS'    AS EYECATCHER")
            objWriteOut.WriteLine(",'I'       AS CHANGEOP")
            objWriteOut.WriteLine(",GETDATE() AS UPDATE_TSTMP")
            objWriteOut.WriteLine(",user      AS USER_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_ALIAS")
            objWriteOut.WriteLine("," & ArrayNoGroupItemCnt.ToString)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto
                If first = True Then
                    NameFld = "    ,inserted." & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "+','+inserted." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto:   Next
            first = True

            objWriteOut.WriteLine(" AS CDC_AFTER_DATA")
            objWriteOut.WriteLine(",'' AS CDC_BEFORE_DATA")
            'objWriteOut.WriteLine(");")
            objWriteOut.WriteLine("FROM inserted")
            'objWriteOut.WriteLine(";")
            objWriteOut.WriteLine("END")
            objWriteOut.WriteLine("go")
            objWriteOut.WriteLine("--------------------")
            objWriteOut.WriteLine("CREATE TRIGGER sqdaudit_R_" & RDash(TableName))
            objWriteOut.WriteLine("ON " & RDash(TableName))
            objWriteOut.WriteLine("AFTER UPDATE")
            objWriteOut.WriteLine("AS") '
            objWriteOut.WriteLine("BEGIN")
            objWriteOut.WriteLine("INSERT SQDENGCD")
            objWriteOut.WriteLine("SELECT")
            objWriteOut.WriteLine(" 'SQMS'    AS EYECATCHER")
            objWriteOut.WriteLine(",'R'       AS CHANGEOP")
            objWriteOut.WriteLine(",GETDATE() AS UPDATE_TSTMP")
            objWriteOut.WriteLine(",user      AS USER_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_ALIAS")
            objWriteOut.WriteLine("," & ArrayNoGroupItemCnt.ToString)

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto1
                If first = True Then
                    NameFld = "    ,inserted." & RDash(fld.FieldName)
                    firstFld = fld.FieldName
                    first = False
                Else
                    NameFld = "+','+inserted." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto1:  Next
            first = True

            objWriteOut.WriteLine(" AS CDC_AFTER_DATA")

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto2
                If first = True Then
                    NameFld = "    ,deleted." & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "+','+deleted." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto2:  Next
            first = True

            objWriteOut.WriteLine(" AS CDC_BEFORE_DATA")
            'wBracket(OpenClose.CLOSE, False)
            'wSemiLine()
            objWriteOut.WriteLine("FROM inserted")
            objWriteOut.WriteLine("INNER JOIN deleted on deleted." & firstFld & "= inserted." & firstFld)
            wSemiLine()
            objWriteOut.WriteLine("END")
            objWriteOut.WriteLine("go")
            objWriteOut.WriteLine("--------------------")
            objWriteOut.WriteLine("CREATE TRIGGER sqdaudit_D_" & RDash(TableName))
            objWriteOut.WriteLine("ON " & RDash(TableName))
            objWriteOut.WriteLine("AFTER UPDATE")
            objWriteOut.WriteLine("AS") '
            objWriteOut.WriteLine("BEGIN")
            objWriteOut.WriteLine("INSERT SQDENGCD")
            objWriteOut.WriteLine("SELECT")
            objWriteOut.WriteLine(" 'SQMS'    AS EYECATCHER")
            objWriteOut.WriteLine(",'D'       AS CHANGEOP")
            objWriteOut.WriteLine(",GETDATE() AS UPDATE_TSTMP")
            objWriteOut.WriteLine(",user      AS USER_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_UPDATED")
            objWriteOut.WriteLine(",'" & RDash(TableName) & "' AS TABLE_ALIAS")
            objWriteOut.WriteLine("," & ArrayNoGroupItemCnt.ToString)
            objWriteOut.WriteLine(",'' AS CDC_AFTER_DATA")

            For Each fld As clsField In FldArray
                '/// skip groupitems
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then GoTo nextgoto3
                If first = True Then
                    NameFld = "    ,deleted." & RDash(fld.FieldName)
                    first = False
                Else
                    NameFld = "+','+deleted." & RDash(fld.FieldName)
                End If
                objWriteOut.WriteLine(NameFld)
nextgoto3:  Next
            first = True

            objWriteOut.WriteLine(" AS CDC_BEFORE_DATA")
            'wBracket(OpenClose.CLOSE, False)
            'wSemiLine()
            objWriteOut.WriteLine("FROM deleted")
            wSemiLine()
            objWriteOut.WriteLine("END")
            objWriteOut.WriteLine("go")
            objWriteOut.WriteLine("--------------------")

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelSQL(MSSql Trigger)")
            Return False
        End Try

    End Function

#End Region

#Region "DB2MQ Trigger"

    Function objModelDB2Trig() As Boolean

        Try
            If SchemaName = "" Then
                SchemaName = "xxxxxxxx."
                objWriteOut.WriteLine("**** NOTE: Before Use, replace 'xxxxxxxx' with")
                objWriteOut.WriteLine("****       Schema Name or Qualifier")
                wBlankLine()
            Else
                SchemaName = SchemaName & "."
            End If
            Dim FORDropTrigIns As String = String.Format("{0}{1}", "    DROP TRIGGER " & SchemaName, RDash(TableName) & "_PAD_INS;")
            Dim FORCreateTrigIns As String = String.Format("{0}{1}", "    CREATE TRIGGER " & SchemaName, RDash(TableName) & "_PAD_INS")
            Dim FORDropTrigUPD As String = String.Format("{0}{1}", "    DROP TRIGGER " & SchemaName, RDash(TableName) & "_PAD_UPD;")
            Dim FORCreateTrigUPD As String = String.Format("{0}{1}", "    CREATE TRIGGER " & SchemaName, RDash(TableName) & "_PAD_UPD")
            Dim FORDropTrigDEL As String = String.Format("{0}{1}", "    DROP TRIGGER " & SchemaName, RDash(TableName) & "_PAD_DEL;")
            Dim FORCreateTrigDEL As String = String.Format("{0}{1}", "    CREATE TRIGGER " & SchemaName, RDash(TableName) & "_PAD_DEL")
            Dim FORAfterINS As String = String.Format("{0}{1}", "        AFTER INSERT ON " & SchemaName, RDash(TableName) & "_PAD_TBL")
            Dim FORAfterUPD As String = String.Format("{0}{1}", "        AFTER UPDATE ON " & SchemaName, RDash(TableName) & "_PAD_TBL")
            Dim FORAfterDEL As String = String.Format("{0}{1}", "        AFTER DELETE ON " & SchemaName, RDash(TableName) & "_PAD_TBL")
            Dim FORRefNEW As String = String.Format("{0}", "        REFERENCING NEW AS N")
            Dim FORRefOLD As String = String.Format("{0}", "        REFERENCING OLD AS D")
            Dim FORForEach As String = String.Format("{0}", "        FOR EACH ROW MODE DB2SQL")
            Dim FORValues As String = String.Format("{0}", "        VALUES DB2MQ.MQSEND")
            Dim FORBody1 As String = String.Format("{0}", "            CHAR(CURRENT TIMESTAMP)")
            Dim FORBodyComma As String = String.Format("{0}", "          ||','")
            Dim FORBody2 As String = String.Format("{0}", "          ||SUBSTR(USER,1,8)")
            Dim FORBody3 As String = String.Format("{0}", "          ||'" & RDash(TableName) & "DB'")
            Dim FORBody4 As String = String.Format("{0}", "          ||'" & RDash(TableName) & "_PAD_TBL'")
            Dim FORBodyA As String = String.Format("{0}", "          ||'A'")
            Dim FORBodyB As String = String.Format("{0}", "          ||'B'")
            Dim FORBodyI As String = String.Format("{0}", "          ||'I'")
            Dim FORBodyU As String = String.Format("{0}", "          ||'U'")
            Dim FORBodyD As String = String.Format("{0}", "          ||'D'")
            Dim FORBodyOpen As String = String.Format("{0}", "      (")
            Dim FORBodyClose As String = String.Format("{0}", "      );")

            '/// Insert
            objWriteOut.WriteLine(FORDropTrigIns)
            objWriteOut.WriteLine(FORCreateTrigIns)
            objWriteOut.WriteLine(FORAfterINS)
            objWriteOut.WriteLine(FORRefNEW)
            objWriteOut.WriteLine(FORForEach)
            objWriteOut.WriteLine(FORValues)
            objWriteOut.WriteLine(FORBodyOpen)
            objWriteOut.WriteLine(FORBody1)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody2)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody3)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody4)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyA)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyI)
            For Each fld As clsField In FldArray
                Dim FORBodyNdot As String = String.Format("{0}", "          ||'N." & fld.FieldName & "'")
                objWriteOut.WriteLine(FORBodyComma)
                objWriteOut.WriteLine(FORBodyNdot)
            Next
            objWriteOut.WriteLine(FORBodyClose)
            wBlankLine()

            '/// Update
            objWriteOut.WriteLine(FORDropTrigUPD)
            objWriteOut.WriteLine(FORCreateTrigUPD)
            objWriteOut.WriteLine(FORAfterUPD)
            objWriteOut.WriteLine(FORRefNEW)
            objWriteOut.WriteLine(FORForEach)
            objWriteOut.WriteLine(FORValues)
            objWriteOut.WriteLine(FORBodyOpen)
            objWriteOut.WriteLine(FORBody1)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody2)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody3)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody4)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyA)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyU)
            For Each fld As clsField In FldArray
                Dim FORBodyNdot As String = String.Format("{0}", "          ||'N." & fld.FieldName & "'")
                objWriteOut.WriteLine(FORBodyComma)
                objWriteOut.WriteLine(FORBodyNdot)
            Next
            objWriteOut.WriteLine(FORBodyClose)
            wBlankLine()

            '///Delete
            objWriteOut.WriteLine(FORDropTrigDEL)
            objWriteOut.WriteLine(FORCreateTrigDEL)
            objWriteOut.WriteLine(FORAfterDEL)
            objWriteOut.WriteLine(FORRefOLD)
            objWriteOut.WriteLine(FORForEach)
            objWriteOut.WriteLine(FORValues)
            objWriteOut.WriteLine(FORBodyOpen)
            objWriteOut.WriteLine(FORBody1)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody2)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody3)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBody4)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyB)
            objWriteOut.WriteLine(FORBodyComma)
            objWriteOut.WriteLine(FORBodyD)
            For Each fld As clsField In FldArray
                Dim FORBodyNdot As String = String.Format("{0}", "          ||'N." & fld.FieldName & "'")
                objWriteOut.WriteLine(FORBodyComma)
                objWriteOut.WriteLine(FORBodyNdot)
            Next
            objWriteOut.WriteLine(FORBodyClose)
            wBlankLine()

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler objModelMQ(DB2MQ Trigger)")
            Return False
        End Try

    End Function

#End Region

#End Region

#Region "Helper Functions and subs"

    Function wBlankLine() As Boolean

        Try
            objWriteOut.WriteLine()
            wBlankLine = True

        Catch ex As Exception
            LogError(ex, "modGenerate wBlankLine")
            wBlankLine = False
        End Try

    End Function

    Function wSemiLine() As Boolean

        Try
            objWriteOut.WriteLine(semi)
            wSemiLine = True

        Catch ex As Exception
            LogError(ex, "modGenerate wSemiLine")
            wSemiLine = False
        End Try

    End Function

    Function wBrace(ByVal OC As OpenClose, ByVal NewLine As Boolean) As Boolean

        Try
            If OC = OpenClose.OPEN Then
                If NewLine = True Then
                    objWriteOut.WriteLine("{")
                Else
                    objWriteOut.Write("{")
                End If
            Else
                If NewLine = True Then
                    objWriteOut.WriteLine("}")
                Else
                    objWriteOut.Write("}")
                End If
            End If
            wBrace = True

        Catch ex As Exception
            LogError(ex, "modModeler wBrace")
            wBrace = False
        End Try

    End Function

    Function wBracket(ByVal OC As OpenClose, ByVal NewLine As Boolean) As Boolean

        Try
            If OC = OpenClose.OPEN Then
                If NewLine = True Then
                    objWriteOut.WriteLine("(")
                Else
                    objWriteOut.Write("(")
                End If
            Else
                If NewLine = True Then
                    objWriteOut.WriteLine(")")
                Else
                    objWriteOut.Write(")")
                End If
            End If
            wBracket = True

        Catch ex As Exception
            LogError(ex, "modModeler wBracket")
            wBracket = False
        End Try

    End Function

    Function getFileExt() As String

        Try
            Select Case OutType
                Case "DTD"
                    getFileExt = ".dtd"
                Case "DB2DDL", "ORADDL", "SQDDDL"
                    getFileExt = ".ddl"
                Case "H"
                    getFileExt = ".h"
                Case "LOD"
                    getFileExt = ".lod"
                Case "SQL"
                    getFileExt = ".sql"
                Case "MSSQL", "SQLDDL"
                    getFileExt = ".sql"
                Case "DB2"
                    getFileExt = ".TRG"
                Case Else
                    getFileExt = ""
            End Select

        Catch ex As Exception
            LogError(ex, "modModeler getFileExt")
            Return ""
        End Try

    End Function

    Function GetOutFldType(ByVal InType As String, ByVal OutType As String, ByVal fldLen As Integer, ByVal fldScale As Integer) As String

        Try
            Dim OutSuffix As String
            If fldLen > 0 Then
                OutSuffix = "(" & fldLen.ToString & ")"
            Else
                OutSuffix = ""
            End If
            GetOutFldType = ""

            Select Case OutType

                Case "H"
                    Select Case InType.ToUpper()
                        Case "CHAR", "VARCHAR", "VARCHAR2", "NUMERIC", "DATE", "TIMESTAMP", "TIME", "TEXTNUM", "ZONE", "XMLCDATA"
                            GetOutFldType = "char"
                        Case "BINARY"
                            GetOutFldType = "unsigned char"
                        Case "INTEGER"
                            GetOutFldType = "int"
                        Case "SMALLINT"
                            GetOutFldType = "short"
                        Case "DECIMAL"
                            GetOutFldType = "double"
                        Case Else
                            GetOutFldType = "char"
                    End Select

                Case "DB2DDL"
                    ' binary()    varbinary()  Image()    BLOB()      BFILE()   char   nchar()  nvarchar()   varchar()
                    ' varchar()   NTEXT()      TEXT()     CLOB()    DBCLOB()   NCLOB()
                    ' DateTime()  smalldatetime()    timestamp()   DATE     DATETIME2()   DateTimeOffset()  TIME()
                    ' TIMESTAMP WITH TIME ZONE
                    ' bigint()    bit()    decimal      Int()    money()  number(p,s)  numeric()   smallint()   smallmoney()
                    ' tinyint()   FLOAT()   REAL()  DECFLOAT(16)   DECFLOAT(34)    DOUBLE
                    ' uniqueID()   DISTINCT (any DB2 data type)      GRAPHIC()   LONG VARGRAPHIC   ROWID()  VARGRAPHIC(n)   Xml()

                    Select Case InType.ToUpper()
                        Case "CHAR", "BINARY", "XMLCDATA", "NCHAR"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "VARCHAR2"
                            GetOutFldType = "VARCHAR" & OutSuffix
                        Case "VARCHAR", "NVARCHAR"
                            If fldLen > 2 Then
                                GetOutFldType = "VARCHAR(" & (fldLen - 2).ToString & ")"
                            ElseIf fldLen < 0 Then
                                GetOutFldType = "VARCHAR(8000)"
                            Else
                                GetOutFldType = "VARCHAR" & OutSuffix
                            End If
                        Case "VARBINARY"
                            GetOutFldType = "VARBINARY" & OutSuffix
                        Case "BLOB", "IMAGE", "BFILE"
                            GetOutFldType = "BLOB" & OutSuffix
                        Case "CLOB", "TEXT", "NTEXT"
                            GetOutFldType = "CLOB" & OutSuffix
                        Case "DBCLOB"
                            GetOutFldType = "DBCLOB" & OutSuffix
                        Case "TIMESTAMP"
                            GetOutFldType = "TIMESTAMP"
                        Case "DATE"
                            GetOutFldType = "DATE"
                        Case "DATETIME"
                            GetOutFldType = "DATETIME   <<< DataType unknown to DB2"
                        Case "SMALLDATETIME"
                            GetOutFldType = "SMALLDATETIME   <<< DataType unknown to DB2"
                        Case "DATETIME2"
                            GetOutFldType = "DATETIME2   <<< DataType unknown to DB2"
                        Case "DATETIMEOFFSET"
                            GetOutFldType = "DATETIMEOFFSET   <<< DataType unknown to DB2"
                        Case "TIME"
                            GetOutFldType = "TIME"
                        Case "TIMESTAMP WITH TIMEZONE"
                            GetOutFldType = "TIMESTAMP WITH TIMEZONE   <<< DataType unknown to DB2"
                        Case "BIGINT"
                            GetOutFldType = "BIGINT" '& OutLen
                        Case "INTEGER", "INT"
                            GetOutFldType = "INTEGER" '& OutLen
                        Case "SMALLINT", "TINYINT", "BIT"
                            GetOutFldType = "SMALLINT" '& OutLen
                        Case "SMALLMONEY"
                            GetOutFldType = "DECIMAL(11,4)"
                        Case "MONEY"
                            GetOutFldType = "DECIMAL(20,4)"
                        Case "DECIMAL", "TEXTNUM", "ZONE", "NUMBER"
                            If InStructureType = enumStructure.STRUCT_COBOL Or InStructureType = enumStructure.STRUCT_COBOL_IMS Or _
                            InStructureType = enumStructure.STRUCT_IMS Then
                                fldLen = (fldLen * 2) - 1
                            End If
                            OutSuffix = "(" & fldLen.ToString & "," & fldScale.ToString & ")"
                            GetOutFldType = "DECIMAL" & OutSuffix
                        Case "NUMERIC"
                            GetOutFldType = "DECIMAL(38)" ' & OutSuffix
                        Case "FLOAT"
                            GetOutFldType = "FLOAT" ' & OutSuffix
                        Case "REAL"
                            GetOutFldType = "REAL" ' & OutSuffix
                        Case "DECFLOAT(16)"
                            GetOutFldType = "DECFLOAT(16)"
                        Case "DECFLOAT(34)"
                            GetOutFldType = "DECFLOAT(34)"
                        Case "DOUBLE"
                            GetOutFldType = "DOUBLE"
                        Case "UNIQUEIDENTIFIER", "ROWID"
                            GetOutFldType = "ROWID"
                        Case "GRAPHIC"
                            GetOutFldType = "GRAPHIC" & OutSuffix
                        Case "LONG VARGRAPHIC"
                            GetOutFldType = "LONG VARGRAPHIC" '& OutSuffix
                        Case "VARGRAPHIC"
                            GetOutFldType = "VARGRAPHIC" & OutSuffix
                        Case "XML"
                            GetOutFldType = "XML" '& OutSuffix
                        Case Else
                            GetOutFldType = InType & OutSuffix
                    End Select

                Case "ORADDL"
                    Select Case InType.ToUpper()
                        Case "CHAR", "TEXTNUM", "ZONE", "XMLCDATA", "NCHAR"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "VARCHAR", "NVARCHAR", "VARBINARY"
                            If fldLen < 0 Then
                                GetOutFldType = "VARCHAR(8000)"
                            Else
                                GetOutFldType = "VARCHAR" & OutSuffix
                            End If
                        Case "VARCHAR2"
                            GetOutFldType = "VARCHAR2" & OutSuffix
                        Case "DATE", "DATETIME"
                            GetOutFldType = "DATE"
                        Case "TIMESTAMP"
                            GetOutFldType = "TIMESTAMP"
                        Case "SMALLDATETIME"
                            GetOutFldType = "SMALLDATETIME   <<< DataType unknown to Oracle"
                        Case "TIME"
                            GetOutFldType = "TIME   <<< DataType unknown to Oracle"
                        Case "BINARY"
                            GetOutFldType = "RAW" & OutSuffix
                        Case "NUMERIC", "DECIMAL", "NUMBER"
                            If InStructureType = enumStructure.STRUCT_COBOL Or InStructureType = enumStructure.STRUCT_COBOL_IMS Or _
                            InStructureType = enumStructure.STRUCT_IMS Then
                                fldLen = (fldLen * 2) - 1
                            End If
                            OutSuffix = "(" & fldLen.ToString & "," & fldScale.ToString & ")"
                            GetOutFldType = "NUMBER" & OutSuffix
                        Case "SMALLMONEY"
                            GetOutFldType = "NUMBER(11,4)"
                        Case "MONEY"
                            GetOutFldType = "NUMBER(20,4)"
                        Case "INTEGER", "INT", "SMALLINT", "TINYINT", "BIT", "BIGINT"
                            GetOutFldType = "NUMBER"
                        Case "UNIQUEIDENTIFIER"
                            GetOutFldType = "VARCHAR(128)" '& OutLen
                        Case "GRAPHIC"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "LONG VARGRAPHIC"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "VARGRAPHIC"
                            GetOutFldType = "VARCHAR" & OutSuffix
                        Case Else
                            GetOutFldType = InType & OutSuffix
                    End Select

                Case "SQLDDL"
                    Select Case InType.ToUpper().Trim
                        Case "CHAR", "XMLCDATA", "TEXTNUM", "ZONE"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "BINARY"
                            GetOutFldType = "BINARY" & OutSuffix
                        Case "VARCHAR2"
                            fldLen = (fldLen * 2)
                            OutSuffix = "(" & fldLen.ToString & ")"
                            GetOutFldType = "VARCHAR" & OutSuffix
                        Case "VARCHAR"
                            If fldLen = -1 Then
                                GetOutFldType = "VARCHAR(max)" '& OutLen
                            Else
                                GetOutFldType = "VARCHAR" & OutSuffix
                            End If
                        Case "TEXT"
                            GetOutFldType = InType  '"[" &  & "]"
                        Case "BLOB", "IMAGE", "BFILE", "CLOB", "DBCLOB", "NCLOB"
                            GetOutFldType = "IMAGE" '& OutSuffix
                        Case "DATE", "TIMESTAMP", "TIME", "DATETIME", "SMALLDATETIME", "DATETIME2", "DATETIMEOFFSET"
                            GetOutFldType = InType
                        Case "TIMESTAMP WITH TIMEZONE"
                            GetOutFldType = "TIMESTAMP WITH TIMEZONE  <<< DataType unknown to MSSQLserver"
                        Case "INTEGER", "INT"
                            GetOutFldType = "INTEGER"
                        Case "DECIMAL", "DOUBLE", "NUMERIC"
                            If InStructureType = enumStructure.STRUCT_COBOL Or InStructureType = enumStructure.STRUCT_COBOL_IMS Or _
                            InStructureType = enumStructure.STRUCT_IMS Then
                                fldLen = (fldLen * 2) - 1
                            End If
                            OutSuffix = "(" & fldLen.ToString & "," & fldScale.ToString & ")"
                            GetOutFldType = "DECIMAL" & OutSuffix
                        Case "SMALLINT"
                            GetOutFldType = "SMALLINT" '& OutLen
                        Case "TINYINT"
                            GetOutFldType = "TINYINT" '& OutLen
                        Case "BIT"
                            GetOutFldType = "BIT" '& OutLen
                        Case "BIGINT"
                            GetOutFldType = "BIGINT" '& OutLen
                        Case "MONEY"
                            GetOutFldType = "MONEY" '& OutLen
                        Case "SMALLMONEY"
                            GetOutFldType = "SMALLMONEY" '& OutLen
                        Case "REAL"
                            GetOutFldType = "REAL" '& OutLen
                        Case "FLOAT", "DECFLOAT(16)", "DECFLOAT(34)"
                            GetOutFldType = "FLOAT" '& OutLen
                        Case "UNIQUEIDENTIFIER"
                            GetOutFldType = "UNIQUEIDENTIFIER"
                        Case "GRAPHIC"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "LONG VARGRAPHIC"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "VARGRAPHIC"
                            GetOutFldType = "VARCHAR" & OutSuffix
                        Case "ROWID"
                            GetOutFldType = "VARCHAR(19)" ' & OutSuffix
                        Case Else
                            GetOutFldType = InType & OutSuffix  '"[" &  & "]"
                    End Select

                Case "SQDDDL"
                    Select Case InType.ToUpper()
                        Case "CHAR", "BINARY", "XMLCDATA", "NCHAR", "NCLOB", "CLOB", "DBCLOB", "TEXT", "NTEXT", "GRAPHIC", _
                        "LONG VARGRAPHIC", "TEXTNUM", "ZONE"
                            GetOutFldType = "CHAR" & OutSuffix
                        Case "VARCHAR", "VARCHAR2"
                            GetOutFldType = "VARCHAR" & OutSuffix
                        Case "DECIMAL", "NUMERIC", "NUMBER"
                            If InStructureType = enumStructure.STRUCT_COBOL Or InStructureType = enumStructure.STRUCT_COBOL_IMS Or _
                            InStructureType = enumStructure.STRUCT_IMS Then
                                fldLen = (fldLen * 2) - 1
                            End If
                            OutSuffix = "(" & fldLen.ToString & "," & fldScale.ToString & ")"
                            GetOutFldType = "DECIMAL" & OutSuffix
                        Case "MONEY"
                            GetOutFldType = "DECIMAL(20,4)" '& OutSuffix
                        Case "SMALLMONEY"
                            GetOutFldType = "DECIMAL(11,4)" '& OutSuffix
                        Case "BINARY"
                            GetOutFldType = "BINARY" & OutSuffix
                        Case "TIMESTAMP"
                            GetOutFldType = "TIMESTAMP" '& OutSuffix
                        Case "DATE", "TIME", "DATETIME"
                            GetOutFldType = "DATE" '& OutSuffix
                        Case "INTEGER", "INT"
                            GetOutFldType = "INTEGER" '& OutSuffix
                        Case "SMALLINT", "TINYINT", "BIT"
                            GetOutFldType = "SMALLINT" '& OutSuffix
                        Case "INTEGER", "INT"
                            GetOutFldType = "INTEGER"
                        Case "BIGINT"
                            GetOutFldType = "INTEGER(20)"
                        Case Else
                            GetOutFldType = InType & OutSuffix  '"[" &  & "]"
                    End Select
                Case Else
                    GetOutFldType = ""
            End Select

        Catch ex As Exception
            LogError(ex, "modModeler GetOutFldType")
            GetOutFldType = ""
        End Try

    End Function

    Function RDash(ByVal fldName As String) As String

        RDash = fldName.Replace("-", "_")

    End Function

#End Region

End Module
