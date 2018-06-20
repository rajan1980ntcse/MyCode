Imports System.IO
Imports System.ComponentModel
Imports Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices

Public Class frmIndexTransfer
    'Inherits System.Windows.Forms

    Public sEditedDocFullName As String
    Public sIndexDocFullName As String
    Const sMsgTitle As String = "Index transfer"

    Dim BW As BackgroundWorker = New BackgroundWorker

    Private Sub btnBrowseEditDoc_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowseEditDoc.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "Select edited document"
        openFileDialog1.Filter = "Word documents (*.Docx)|*.docx"
        ''openFileDialog1.FilterIndex = 2 
        openFileDialog1.RestoreDirectory = True
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            sEditedDocFullName = openFileDialog1.FileName
            Me.txtEditeddoc.Text = sEditedDocFullName ': Me.txtIndexdoc.Enabled = False
        End If
    End Sub

    Private Sub BtnBrowseIndexDoc_Click(sender As System.Object, e As System.EventArgs) Handles BtnBrowseIndexDoc.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Title = "Select index document"
        openFileDialog1.Filter = "Word documents (*.Docx)|*.docx"
        ''openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True
        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            sIndexDocFullName = openFileDialog1.FileName
            Me.txtIndexdoc.Text = sIndexDocFullName ': Me.txtIndexdoc.Enabled = False
        End If
    End Sub

    Private Sub txtEditeddoc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtEditeddoc.TextChanged
        If File.Exists(Me.txtEditeddoc.Text) = True AndAlso Path.GetExtension(Me.txtEditeddoc.Text).ToLower = ".docx" Then
            sEditedDocFullName = Me.txtEditeddoc.Text
        Else
            MessageBox.Show("Please select valid document file and try again", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If
    End Sub


    Private Sub txtIndexdoc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtIndexdoc.TextChanged
        If File.Exists(Me.txtIndexdoc.Text) = True AndAlso Path.GetExtension(Me.txtIndexdoc.Text).ToLower = ".docx" Then
            sIndexDocFullName = Me.txtIndexdoc.Text
        Else
            MessageBox.Show("Please select valid document file and try again", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

    End Sub

    Private Sub btncancel_Click(sender As System.Object, e As System.EventArgs) Handles btncancel.Click
        Me.Dispose(True)
    End Sub

    Private Sub btnProcess_Click(sender As System.Object, e As System.EventArgs) Handles btnProcess.Click
        If File.Exists(sEditedDocFullName) = True AndAlso File.Exists(sIndexDocFullName) = True Then
            Me.btnProcess.Enabled = False
            BW.WorkerSupportsCancellation = True
            BW.WorkerReportsProgress = True
            AddHandler BW.DoWork, AddressOf bw_DoWork
            AddHandler BW.ProgressChanged, AddressOf bw_ProgressChanged
            AddHandler BW.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
            Me.tbProgress.Text = "0%" : Me.tbProgress.Visible = True : Me.ProgressBar1.Visible = True : Me.Refresh()
            If Not BW.IsBusy = True Then
                BW.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Sub bw_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)
        If BW.CancellationPending = True Then
            e.Cancel = True
            Exit Sub
        Else
            Dim oclsIndex As New clsTransformIndexTerm
            If oclsIndex.ToTransformIndexTerm(sEditedDocFullName, sIndexDocFullName, BW) = True Then
                Call MessageBox.Show("Process was completed successfully.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub


    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        Me.tbProgress.Text = e.ProgressPercentage.ToString() & "%"
        Me.ProgressBar1.Value = e.ProgressPercentage
        Me.Refresh()
    End Sub

    Private Sub bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        If e.Cancelled = True Then
            Me.tbProgress.Text = "Canceled!"
            Me.tbProgress.Visible = False : Me.ProgressBar1.Visible = False : Me.btnProcess.Enabled = True
        ElseIf e.Error IsNot Nothing Then
            Me.tbProgress.Text = "Error: " & e.Error.Message
            Me.Dispose()
        Else
            Me.tbProgress.Text = "Done"
            Me.Dispose()
        End If

    End Sub

    Private Sub frmIndexTransfer_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.tbProgress.Visible = False : Me.ProgressBar1.Visible = False
        sEditedDocFullName = String.Empty : sIndexDocFullName = String.Empty
    End Sub

    Private Sub btnMerge_Click(sender As System.Object, e As System.EventArgs)
        'Dim ofrmMergeDoc As New frmMergeDocument
        'If ofrmMergeDoc.ShowDialog() = Windows.Forms.DialogResult.OK Then

        'End If
    End Sub

    Private Sub btnReview_Click(sender As System.Object, e As System.EventArgs) Handles btnReview.Click

        Try
            Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
            If WordApp Is Nothing = True Then
                MessageBox.Show("Word application not found", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            Dim oMDoc As Document = WordApp.ActiveDocument

            Dim oclsIndex As New clsTransformIndexTerm
            oclsIndex.ToReviewIndex(oMDoc)
        Catch ex As Exception
            MessageBox.Show("Document not found", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try
        
    End Sub
End Class