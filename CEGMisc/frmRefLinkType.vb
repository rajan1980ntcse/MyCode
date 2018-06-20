Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports CEGINI
Public Class frmRefLinkType
    Public oWordApp As Word.Application
    Public wListFiles As String
    Public wFilePath As String
    Public styleINI As String
    Public srefLinkType As String
    Public ImpStyleList As ArrayList
    Public Sub New(oWApp As Word.Application, wLstFiles As String, wFPath As String)
        oWordApp = oWApp
        wListFiles = wLstFiles
        wFilePath = wFPath
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Private Sub frmRefLinkType_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        srefLinkType = "book"
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles rbChapter.CheckedChanged
        If rbBook.Checked Then
            srefLinkType = "book"
        Else
            srefLinkType = "chapter"
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
        Dim curdocPath As String
        Dim curDoc As Word.Document
        If rbBook.Checked Then
            For Each fDoc As String In wListFiles.Split("||")
                If fDoc <> String.Empty Then
                    curdocPath = Path.Combine(wFilePath, fDoc)
                    oWordApp.Documents(curdocPath).Activate()
                    curDoc = oWordApp.ActiveDocument
                    If VariableExists(curDoc, "CEGRefLinkType") = False Then
                        curDoc.Variables.Add("CEGRefLinkType", "Book")
                    Else
                        curDoc.Variables("CEGRefLinkType").Value = "Book"
                    End If
                End If
            Next
        Else
            For Each fDoc As String In wListFiles.Split("||")
                If fDoc <> String.Empty Then
                    curdocPath = Path.Combine(wFilePath, fDoc)
                    oWordApp.Documents(curdocPath).Activate()
                    curDoc = oWordApp.ActiveDocument
                    If VariableExists(curDoc, "CEGRefLinkType") = False Then
                        curDoc.Variables.Add("CEGRefLinkType", "Chapter")
                    Else
                        curDoc.Variables("CEGRefLinkType").Value = "Chapter"
                    End If
                End If
            Next
        End If
        Me.Close()
    End Sub

    Private Sub rbBook_CheckedChanged(sender As Object, e As EventArgs) Handles rbBook.CheckedChanged
        If rbBook.Checked Then
            srefLinkType = "book"
        Else
            srefLinkType = "chapter"
        End If
    End Sub
End Class