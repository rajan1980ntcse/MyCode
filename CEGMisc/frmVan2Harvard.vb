Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports System.IO
Imports System.Text.RegularExpressions
Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices

Public Class frmVan2Harvard
    Public wDoc As Word.Document
    Public wordAPP As Word.Application
    Public bEtalItalic As Boolean

    Public Sub New(oDoc As Word.Document, oAPP As Word.Application, bItalic As Boolean)
        wDoc = oDoc
        wordAPP = oAPP
        bEtalItalic = bItalic
        InitializeComponent()
    End Sub

    Private Sub CBCitationPattern_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CBCitationPattern.SelectedIndexChanged
        If (CBBracketType.Text <> "") Then
            Select Case (CBBracketType.Text)
                Case "()"
                    If CBCitationPattern.Text.Contains("(") Then
                        Label3.Text = "" & CBCitationPattern.Text & ""
                    Else
                        Label3.Text = "(" & CBCitationPattern.Text & ")"
                    End If
                Case "[]"
                    Label3.Text = "[" & CBCitationPattern.Text & "]"
                Case "{};"
                    Label3.Text = "{" & CBCitationPattern.Text & "}"
            End Select

        End If
    End Sub

    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        Dim RefTextCitationList As New List(Of String)

        If RefSwapValidation() Then
            If DataGridView1.Rows.Count > 0 Then
                For i = 0 To DataGridView1.Rows.Count - 1
                    RefTextCitationList.Add(DataGridView1.Rows(i).Cells(0).Value)
                Next
                For Each dRow As DataGridViewRow In DataGridView1.Rows
                    If dRow.Cells(0).Value = "Untagged Reference" Then
                        dRow.Cells(0).Style.BackColor = Color.Aqua
                        dRow.Cells(0).ReadOnly = False
                    Else
                        dRow.Cells(0).Style.BackColor = Color.White
                        dRow.Cells(0).ReadOnly = True
                    End If
                Next
            End If

            If RefTextCitationList.Contains("Untagged Reference") Then
                MessageBox.Show("Untagged Reference found in the document. So unable to convert Havard reference.")
            Else
                ModRefUtility.ToConvertHavardCitation(RefTextCitationList)
                Me.Close()
                Dim frmCitViewer As New frmCitationViewer(wDoc, wordAPP, bEtalItalic) ''Jaisoft
                frmCitViewer.Show()
            End If
        Else
            MessageBox.Show("Sorry ..........not a valid file")
            Me.Close()
        End If
    End Sub

    Private Sub frmVan2Harvard_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        For Each pair As KeyValuePair(Of Word.Range, String) In ModRefUtility.dictHavCitaion
            Dim n As Integer = DataGridView1.Rows.Add()
            DataGridView1.Rows.Item(n).Cells(0).Value = pair.Value
            DataGridView1.Rows.Item(n).Cells(1).Value = pair.Key.Text
            If DataGridView1.Rows.Item(n).Cells(0).Value = "Untagged Reference" Then
                DataGridView1.Rows.Item(n).Cells(0).Style.BackColor = Color.Aqua
                DataGridView1.Rows.Item(n).Cells(0).ReadOnly = False
            Else
                DataGridView1.Rows.Item(n).Cells(0).ReadOnly = True
            End If
        Next
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Environment.Exit(0)
    End Sub
End Class