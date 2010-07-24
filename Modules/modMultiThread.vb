Imports System
Imports System.Collections
Imports System.ComponentModel
Imports System.Threading
Module modMultiThread

    Private numberToCompute As Integer = 0
    Private highestPercentageReached As Integer = 0

    Private WithEvents backgroundWorker1 As System.ComponentModel.BackgroundWorker


    ' This event handler is where the actual work is done.
    Private Sub backgroundWorker1_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles backgroundWorker1.DoWork

        ' Get the BackgroundWorker object that raised this event.
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        ' Assign the result of the computation
        ' to the Result property of the DoWorkEventArgs 
        ' object. This is will be available to the 
        ' RunWorkerCompleted eventhandler.
        e.Result = ComputeFibonacci(e.Argument, worker, e)
    End Sub 'backgroundWorker1_DoWork

    ' This event handler deals with the results of the
    ' background operation.
    Private Sub backgroundWorker1_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles backgroundWorker1.RunWorkerCompleted

        ' First, handle the case where an exception was thrown.
        If (e.Error IsNot Nothing) Then
            MessageBox.Show(e.Error.Message)
        ElseIf e.Cancelled Then
            ' Next, handle the case where the user canceled the 
            ' operation.
            ' Note that due to a race condition in 
            ' the DoWork event handler, the Cancelled
            ' flag may not have been set, even though
            ' CancelAsync was called.
            'resultLabel.Text = "Canceled"
        Else
            ' Finally, handle the case where the operation succeeded.
            'resultLabel.Text = e.Result.ToString()
        End If

        ' Enable the UpDown control.
        'Me.numericUpDown1.Enabled = True

        ' Enable the Start button.
        'startAsyncButton.Enabled = True

        ' Disable the Cancel button.
        'cancelAsyncButton.Enabled = False
    End Sub 'backgroundWorker1_RunWorkerCompleted

    ' This event handler updates the progress bar.
    Private Sub backgroundWorker1_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles backgroundWorker1.ProgressChanged

        'Me.progressBar1.Value = e.ProgressPercentage

    End Sub

    ' This is the method that does the actual work. For this
    ' example, it computes a Fibonacci number and
    ' reports progress as it does its work.
    Function ComputeFibonacci(ByVal n As Integer, ByVal worker As BackgroundWorker, ByVal e As DoWorkEventArgs) As Long

        ' The parameter n must be >= 0 and <= 91.
        ' Fib(n), with n > 91, overflows a long.
        If n < 0 OrElse n > 91 Then
            Throw New ArgumentException( _
                "value must be >= 0 and <= 91", "n")
        End If

        Dim result As Long = 0

        ' Abort the operation if the user has canceled.
        ' Note that a call to CancelAsync may have set 
        ' CancellationPending to true just after the
        ' last invocation of this method exits, so this 
        ' code will not have the opportunity to set the 
        ' DoWorkEventArgs.Cancel flag to true. This means
        ' that RunWorkerCompletedEventArgs.Cancelled will
        ' not be set to true in your RunWorkerCompleted
        ' event handler. This is a race condition.
        If worker.CancellationPending Then
            e.Cancel = True
        Else
            If n < 2 Then
                result = 1
            Else
                result = ComputeFibonacci(n - 1, worker, e) + _
                         ComputeFibonacci(n - 2, worker, e)
            End If

            ' Report progress as a percentage of the total task.
            Dim percentComplete As Integer = _
                CSng(n) / CSng(numberToCompute) * 100
            If percentComplete > highestPercentageReached Then
                highestPercentageReached = percentComplete
                worker.ReportProgress(percentComplete)
            End If

        End If

        Return result

    End Function

End Module
