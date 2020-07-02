Imports System.ComponentModel
Public Class frmProgressBar
    Inherits System.Windows.Forms.Form
    'Public WithEvents btnCancel As System.Windows.Forms.Button
    'Public myProgressBar As System.Windows.Forms.ProgressBar
    'Public lblStatus As System.Windows.Forms.Label
    'Public WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        BackgroundWorker1.WorkerReportsProgress = True
        BackgroundWorker1.WorkerSupportsCancellation = True
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        amanda.isCancelled = True
        If BackgroundWorker1.WorkerSupportsCancellation = True Then
            'Cancel the asynchronus operation
            BackgroundWorker1.CancelAsync()
            btnCancel.Enabled = False
        End If
        Me.Close()
    End Sub

    ' This event handler is where the time-consuming work is done.
    Private Sub BackgroundWorker1_DoWork(ByVal sender As System.Object,
    ByVal e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim i As Integer

        For i = 1 To 10
            If (worker.CancellationPending = True) Then
                e.Cancel = True
                Exit For
            Else
                ' Perform a time consuming operation and report progress.
                System.Threading.Thread.Sleep(500)
                worker.ReportProgress(i * 10)
            End If
        Next
    End Sub

    ' This event handler updates the progress.
    Private Sub BackgroundWorker1_ProgressChanged(ByVal sender As System.Object, ByVal e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        lblStatus.Text = (e.ProgressPercentage.ToString() + "%")
        Me.myProgressBar.Value = e.ProgressPercentage
    End Sub


    ' This event handler deals with the results of the background operation.
    Private Sub BackgroundWorker1_RunWorkerCompleted(ByVal sender As System.Object,
    ByVal e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled = True Then
            lblStatus.Text = "Canceled!"
        ElseIf e.Error IsNot Nothing Then
            lblStatus.Text = "Error: " & e.Error.Message
        Else
            lblStatus.Text = "Done!"
        End If
    End Sub


End Class