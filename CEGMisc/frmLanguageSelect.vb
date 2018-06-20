Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports CEGINI
Public Class frmLanguageSelect

    Public oWordApp As Word.Application
    Public wListFiles As String
    Public wFilePath As String
    Public styleINI As String
    Public ImpStyleList As ArrayList
    Public Sub New(oWApp As Word.Application, wLstFiles As String, wFPath As String)
        oWordApp = oWApp
        wListFiles = wLstFiles
        wFilePath = wFPath
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Private Sub frmLanguageSelect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim cPathMisc = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
        Dim oReadINI As New CEGINI.clsINI(cPathMisc)
        Dim varLanguage = oReadINI.INIReadValue("LanguageVariable", "Language")
        If (String.IsNullOrEmpty(varLanguage)) Then
            MessageBox.Show("Please check the CEGMISC.ini file.. for Lanugage information.")
            Environment.Exit(0)
        Else
            ComboBox1.DataSource = varLanguage.Split("|").ToList()
            ComboBox1.Text = ""
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim curdocPath As String
        Dim curDoc As Word.Document

        If (ComboBox1.Text <> "") Then
            For Each fDoc As String In wListFiles.Split("||")
                If fDoc <> String.Empty Then
                    curdocPath = Path.Combine(wFilePath, fDoc)
                    oWordApp.Documents(curdocPath).Activate()
                    curDoc = oWordApp.ActiveDocument
                    If VariableExists(curDoc, "CEGLanguage") = False Then
                        curDoc.Variables.Add("CEGLanguage", ComboBox1.Text)
                    Else
                        curDoc.Variables("CEGLanguage").Value = ComboBox1.Text
                    End If
                    Call AddQCIteminCollection("CEGLanBook", curDoc)
                End If
            Next
            Me.Hide()
        Else
            MessageBox.Show("Please select the language...")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Environment.Exit(0)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class