Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Public Class frmSortHarvardCitation
    Public wDoc As Word.Document
    Public wAPP As Word.Application
    Public htmlText As String
    Public Sub New(oDoc As Word.Document, oAPP As Word.Application)
        wDoc = oDoc
        wAPP = oAPP
        InitializeComponent()
    End Sub
    Private Sub frmSortHarvardCitation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            'For Each dPair As KeyValuePair(Of Word.Range, Dictionary(Of Word.Range, String)) In dictHavCitInfo
            '    Dim n As Integer = dgvCitation.Rows.Add()
            '    dgvCitation.Rows.Item(n).Cells(0).Value = dPair.Key.Text
            '    Dim CitVal As String = ""
            '    For Each pRair As KeyValuePair(Of Word.Range, String) In dPair.Value
            '        If dPair.Value.Last().Key Is pRair.Key Then
            '            CitVal = CitVal & pRair.Key.Text
            '        Else
            '            CitVal = CitVal & pRair.Key.Text & ModRefUtility.cCitationSep & " "
            '        End If
            '    Next
            '    dgvCitation.Rows.Item(n).Cells(1).Value = CitVal
            'Next
            ''dictHavCitInfo = dictHavCitInfo.OrderBy(Function(x) x.Value.OrderBy(Function(y) y.Value))
            Debug.Print("dfsfsf")


        Catch ex As Exception
            MessageBox.Show("Error:" + ex.Message)
        End Try
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click

    End Sub
End Class