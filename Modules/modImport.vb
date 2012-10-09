Imports System.IO
Module modImport

    '/// Created July-Aug. 2012 by Tom Karasch ///
    '/// This module Imports Scripts and Descriptions ///
    '/// and replaces sqduiimp.exe ///
    Dim RC As clsImpRetCode
    Dim InitializeComplete As Boolean = False
   

#Region "General Processes"

    Function InitRetCode() As Boolean

        Try
            '/// initialize return code object

            RC = New clsImpRetCode
            RC.Eng = Nothing
            RC.Errline = 0
            RC.ErrorCount = 0
            RC.ErrorPath = ""
            RC.HasError = False
            RC.HasFatal = False
            RC.LocalErrorMsg = ""
            RC.ImportType = clsImpRetCode.enumImportType.None
            RC.Name = ""
            RC.ObjInode = Nothing
            RC.Path = ""
            RC.ReturnCode = clsImpRetCode.enumImpRetCode.Good
            InitializeComplete = True

            Return True

        Catch ex As Exception
            LogError(ex, "modImport Error Initializing Return code")
            Return False
        End Try

    End Function

#End Region

#Region " Script Input"

    Public Function InputScriptV3(ByVal InputPath As String) As clsImpRetCode

        Try
            '/// Initialize the Return Code Object
            If InitRetCode() = False Then
                GoTo ErrorGoTo
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

            '/// close file streams ******************************



        Catch ex As Exception
            LogError(ex, "modImport InputScriptV3")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
        Finally
            '/// send Return Code Object Back to Main Program
            If RC.HasError = False Then
                RC.ReturnCode = "Script generated Successfully !!"
            Else
                RC.HasError = True
            End If
            InputScriptV3 = RC

        End Try

    End Function

#End Region

#Region "Get Description"

    Public Function GetDescriptionFromFile(ByVal InFilePath As String) As clsStructure

        Try
            If InitializeComplete = False Then
                If InitRetCode() = False Then
                    GoTo ErrorGoTo
                End If
            End If

            If InFilePath <> "" Then
                Dim objreader As New System.IO.StreamReader(InFilePath)
                Dim LineNum As Integer = 0
                Dim LineList As New Collection
                Dim CurLine As String = ""

                LineList.Clear()

                While (objreader.Peek() > -1)
                    LineNum += 1
                    CurLine = objreader.ReadLine()
                    LineList.Add(CurLine, LineNum)
                End While

                objreader.Close()
            Else

            End If

            '/// Now Turn the collection into a clsStructure Object


ErrorGoTo:  '/// send returnPath or enumreturncode

          

        Catch ex As Exception
            LogError(ex, "modImport GetDescriptionFromFile")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
        Finally

        End Try

    End Function

    Public Function GetDescriptionFromString() As clsStructure


        Try



ErrorGoTo:  '/// send returnPath or enumreturncode

            '            '/// close file streams ******************************




        Catch ex As Exception
            LogError(ex, "modImport GetDescriptionFromString")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
        Finally

        End Try

    End Function

#End Region

#Region "Parser function"


   

#End Region

#Region "Create Output Files and Write Headers"



   
#End Region

#Region "Object Functions"




#End Region

#Region "Write to file Functions"

  

   
#End Region


End Module
