Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq
Imports System.Windows
Imports Word = Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices
Imports CEGINI
Public Class frmReferenceStyler
    Public wDoc As Word.Document
    Public wAPP As Word.Application

    Public Sub New(aDoc As Word.Document, WAPPP As Word.Application)
        wDoc = aDoc
        wAPP = WAPPP
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If ComboBox1.Text <> "" Then
                RefStyleModifier.ReferenceStyleModifierMain(wDoc, wAPP, ComboBox1.Text.Trim())
                Me.Close()
            Else
                MessageBox.Show("Please select the Style from the list.")
            End If
        Catch ex As Exception
            MessageBox.Show("CEG Error: " & ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub frmReferenceStyler_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim pStyleStructureINI As String
        pStyleStructureINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGRefStyleStructure.ini")
        If Not File.Exists(pStyleStructureINI) Then
            MessageBox.Show("Unable to process due to configuration file missing in CEG", "CE Genius")
            Me.Close()
        Else
            Dim oReadINI As New CEGINI.clsINI(pStyleStructureINI)
            Dim sRefstyles() = oReadINI.INIReadValue("RefStyleStructure", "Styles").Split("|")
            ComboBox1.DataSource = sRefstyles
        End If
    End Sub
End Class