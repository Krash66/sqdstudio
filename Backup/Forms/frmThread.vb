Imports System.Windows.Forms

Public Class frmThread

    Public Event Generated(ByRef rc As clsRcode)

    Public Sub LoadMe()

        Dim i As Integer = 1

        Me.Show()

        Me.Focus()


doloop: Do
            Select Case i
                Case 1
                    Label2.Text = "."
                Case 2
                    Label2.Text = ".."
                Case 3
                    Label2.Text = "..."
                Case 4
                    Label2.Text = "...."
                Case 5
                    Label2.Text = "....."
            End Select
            For i = 1 To 100000
                'timer
            Next
        Loop Until i = 5
        i = 1
        GoTo doloop

    End Sub

    Public Sub OnGenerate()
        Me.SuspendLayout()

        Me.Close()

    End Sub

    Public Function ShowMe(Optional ByVal rename = False) As Boolean

        Try


            Me.Show()

            If rename = True Then
                Label2.Text = "Rename complete. Save Successful."
            Else
                Label2.Text = "Save Successful"
            End If

            Timer()

            Return True

        Catch ex As Exception
            LogError(ex, "frmThread ShowMe")
            Return True
        End Try

    End Function

    Function Timer() As Boolean

        Dim count As Integer = 0

        While count < 444444444
            count += 1
        End While

        Me.Close()

        Return True

    End Function

End Class
