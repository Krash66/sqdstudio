﻿Imports System.IO
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
    Dim InType As String
    Dim OutPath As String = ""
    Dim NameModl As String = ""
    Dim TableName As String = ""
    Dim OutType As String = ""
    Dim FileExt As String = ""
    Dim FullPathModl As String = ""
    Dim ErrFile As String = GetAppTemp() & "\" & "ModelErr.Log"


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
            InType = Obj.Type
            NameModl = ModName
            OutType = TypeToModl
            FileExt = getFileExt()
            OutPath = ModPath

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
                Select Case InType
                    Case NODE_STRUCT
                        ObjStr = CType(Obj, clsStructure)
                        TableName = ObjStr.StructureName
                        ObjStr.LoadMe()
                        ObjStr.LoadItems()
                    Case NODE_STRUCT_SEL
                        ObjSel = CType(Obj, clsStructureSelection)
                        TableName = ObjSel.SelectionName
                        ObjSel.LoadMe()
                        ObjSel.LoadItems()
                    Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                        ObjDSSel = CType(Obj, clsDSSelection)
                        TableName = ObjDSSel.SelectionName
                        ObjDSSel.LoadMe()
                        ObjDSSel.LoadItems()
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
                    Case Else
                        success = False
                        GoTo ErrorGoTo
                End Select
            End If

            If success Then
                Return FullPathModl
                Log("********* Modeler Finish *********")
            Else
ErrorGoTo:      '/// on errors
                Return ""
                Log("********* Modeler Finish with Errors *********")
            End If

        Catch ex As Exception
            LogError(ex, "modModeler ModelStructure")
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
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "DB2DDL", fldLen)


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
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "ORADDL", fldLen)


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
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "SQLDDL", fldLen)

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
            LogError(ex, "modModeler objModelSQLDDL")
            Return False
        End Try

    End Function

#End Region

#Region "C"

    Function objModelC() As Boolean

        Try
            Dim fldLen As String
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
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "H", fldLen)

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

                FORfld = String.Format("{0}{1,-30}{2}", "<!ELEMENT ", NameFld, " (#PCDATA)>")

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
                fldattr = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                OutAttr = GetOutFldType(fldattr, "ORADDL", 0)

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
            objWriteOut.WriteLine(");")
            objWriteOut.WriteLine("FROM inserted")
            objWriteOut.WriteLine(";")
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
            wBracket(OpenClose.CLOSE, False)
            wSemiLine()
            objWriteOut.WriteLine("FROM inserted")
            objWriteOut.WriteLine("INNER JOIN deleted on")
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
            wBracket(OpenClose.CLOSE, False)
            wSemiLine()
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
                Case "DB2DDL", "ORADDL", "SQLDDL"
                    getFileExt = ".ddl"
                Case "H"
                    getFileExt = ".h"
                Case "LOD"
                    getFileExt = ".lod"
                Case "SQL"
                    getFileExt = ".sql"
                Case "MSSQL"
                    getFileExt = ".sql"
                Case Else
                    getFileExt = ""
            End Select

        Catch ex As Exception
            LogError(ex, "modModeler getFileExt")
            Return ""
        End Try

    End Function

    Function GetOutFldType(ByVal InType As String, ByVal ModType As String, ByVal fldLen As Integer) As String

        Try
            Dim OutLen As String
            If fldLen > 0 Then
                OutLen = "(" & fldLen.ToString & ")"
            Else
                OutLen = ""
            End If
            Select Case ModType

                Case "H"
                    Select Case InType
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
                            GetOutFldType = InType
                    End Select

                Case "DB2DDL"
                    Select Case InType
                        Case "CHAR", "TEXTNUM", "ZONE", "BINARY", "XMLCDATA"
                            GetOutFldType = "CHAR" & OutLen
                        Case "VARCHAR2"
                            GetOutFldType = "VARCHAR" & OutLen
                        Case "VARCHAR"
                            If fldLen > 2 Then
                                GetOutFldType = "VARCHAR(" & (fldLen - 2).ToString & ")"
                            Else
                                GetOutFldType = "VARCHAR" & OutLen
                            End If
                        Case "DATE"
                            GetOutFldType = "DATE"
                        Case "TIMESTAMP"
                            GetOutFldType = "TIMESTAMP"
                        Case "TIME"
                            GetOutFldType = "TIME"
                            'Case "BINARY"
                            '    GetOutFldType = "BINARY" & OutLen
                        Case "INTEGER"
                            GetOutFldType = "INTEGER" & OutLen
                        Case "SMALLINT"
                            GetOutFldType = "SMALLINT"
                        Case "NUMERIC", "DECIMAL"
                            GetOutFldType = "DECIMAL" & OutLen
                        Case Else
                            GetOutFldType = InType & OutLen
                    End Select

                Case "ORADDL"
                    Select Case InType
                        Case "CHAR", "TEXTNUM", "ZONE", "XMLCDATA"
                            GetOutFldType = "CHAR" & OutLen
                        Case "VARCHAR"
                            GetOutFldType = "VARCHAR" & OutLen
                        Case "VARCHAR2"
                            GetOutFldType = "VARCHAR2" & OutLen
                        Case "DATE"
                            GetOutFldType = "DATE"
                        Case "TIMESTAMP", "TIME"
                            GetOutFldType = "TIMESTAMP"
                        Case "BINARY"
                            GetOutFldType = "RAW" & OutLen
                        Case "NUMERIC", "DECIMAL", "SMALLINT", "INTEGER"
                            GetOutFldType = "NUMBER" & OutLen
                        Case Else
                            GetOutFldType = InType & OutLen
                    End Select

                Case "SQLDDL"
                    Select Case InType
                        Case "CHAR", "TEXTNUM", "ZONE", "BINARY", "XMLCDATA"
                            GetOutFldType = "CHAR" & OutLen
                        Case "VARCHAR", "VARCHAR2"
                            GetOutFldType = "VARCHAR" & OutLen
                        Case "NUMERIC"
                            GetOutFldType = "NUMERIC" & OutLen
                        Case "DATE", "TIMESTAMP", "TIME"
                            GetOutFldType = "DATETIME"
                        Case "INTEGER"
                            GetOutFldType = "INT"
                            'Case "BINARY"
                            '    GetOutFldType = "BINARY" & OutLen
                        Case "DECIMAL"
                            GetOutFldType = "DECIMAL" & OutLen
                        Case "SMALLINT"
                            GetOutFldType = "SMALLINT" & OutLen
                        Case Else
                            GetOutFldType = InType & OutLen
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
