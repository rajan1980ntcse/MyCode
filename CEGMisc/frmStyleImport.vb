Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports CEGINI
Public Class frmStyleImport
    Public oWordApp As Word.Application
    Public wListFiles As String
    Public wFilePath As String
    Public styleINI As String
    Public ImpStyleList As ArrayList
    Public Sub New(oWApp As Word.Application, wLstFiles As String, wFPath As String, sINI As String)
        oWordApp = oWApp
        wListFiles = wLstFiles
        wFilePath = wFPath
        styleINI = sINI
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        Dim tPath As String
        Dim Templ As Word.Document
        Dim curDoc As Word.Document
        Dim curdocPath As String

        Dim oReadIni As New CEGINI.clsINI(styleINI, True)
        If lbStyleName.Items.Count > 0 Then
            ImpStyleList = New ArrayList()
            For I = 0 To lbStyleName.Items.Count - 1
                'ImpStyleList.Add(lbStyleName.Items(I).ToString() & " Open")
                'ImpStyleList.Add(lbStyleName.Items(I).ToString() & " Close")
                ImpStyleList.Add(lbStyleName.Items(I).ToString())
            Next
            tPath = oReadIni.INIReadValue("ImportTemplate", "Template")
            tPath = Path.Combine(Path.GetDirectoryName(styleINI), tPath)
            If File.Exists(tPath) Then
                oWordApp.Documents.Open(tPath)
                Templ = oWordApp.Documents(tPath)
                CheckStyleExistinTemplate(Templ, ImpStyleList)
                If ImpStyleList.Count > 0 Then
                    For Each fDoc As String In wListFiles.Split("||")
                        If fDoc <> String.Empty Then
                            curdocPath = Path.Combine(wFilePath, fDoc)
                            oWordApp.Documents(curdocPath).Activate()
                            curDoc = oWordApp.ActiveDocument
                            CopyStylesToandFrom(curDoc, Templ, ImpStyleList)
                        End If
                    Next
                End If
                Templ.Close()
            End If
        Else
            MessageBox.Show("style not found in the listbox")
        End If
        Me.Close()
    End Sub
    Function CheckStyleExistinTemplate(tempDoc As Word.Document, styleList As ArrayList)
        Dim sty As Word.Style
        Dim I As Integer
        For I = styleList.Count - 1 To 0 Step -1
            Dim t
            On Error Resume Next
            t = tempDoc.Styles(styleList(I).ToString())
            On Error GoTo 0
            If t Is Nothing Then
                MessageBox.Show("' " & styleList(I).ToString() & " '  :style not found in the template")
                styleList.RemoveAt(I)
            End If
            Err.Clear()
        Next
    End Function
    Function CopyStylesToandFrom(OrgDoc As Word.Document, CpyFile As Word.Document, styleList As ArrayList)
        Dim sty As Word.Style
        On Error Resume Next
        For Each sty In CpyFile.Styles
            If sty.BuiltIn = False Then
                If styleList.Contains(sty.NameLocal) Then
                    oWordApp.OrganizerCopy(CpyFile.FullName, OrgDoc.FullName, sty.NameLocal, Word.WdOrganizerObject.wdOrganizerObjectStyles)
                End If
            End If
        Next sty
        If Err.Number <> 0 Then
            Err.Clear()
        End If
    End Function
    Private Sub frmStyleImport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim sStyleList As String
        If File.Exists(styleINI) Then
            Dim oReadIni As New CEGINI.clsINI(styleINI, True)
            sStyleList = oReadIni.INIReadValue("ImportStyleList", "style")
            For Each itm As String In sStyleList.Split("||")
                If itm <> String.Empty Then cbStyleList.Items.Add(itm)
            Next
        End If
    End Sub



    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub btnAdd_Click(sender As System.Object, e As System.EventArgs) Handles btnAdd.Click
        If cbStyleList.Text <> String.Empty Then
            If Not (lbStyleName.Items.Contains(cbStyleList.Text)) Then
                lbStyleName.Items.Add(cbStyleList.Text & " Open")
                lbStyleName.Items.Add(cbStyleList.Text & " Close")
            End If
        End If
        'lbStyleName.Items.Add(Trim(txtStyleName.Text) & " Open")
        'lbStyleName.Items.Add(Trim(txtStyleName.Text) & " Close")
        'txtStyleName.Text = ""
        lbStyleName.Text = ""
    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        If lbStyleName.Items.Count > 0 Then
            lbStyleName.Items.Clear()
        End If
    End Sub
End Class