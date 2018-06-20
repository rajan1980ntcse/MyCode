Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports CEGINI
Public Class frmRefStyleType
    Public oWordApp As Word.Application
    Public wListFiles As String
    Public wFilePath As String
    Public styleINI As String
    Public sRefLinkStyle As String
    Public sPubName As String
    Public sCEGPath As String
    Public gRefStyleType As String
    Public ImpStyleList As ArrayList
    Public Sub New(oWApp As Word.Application, wLstFiles As String, wFPath As String, tRefLinkType As String, pName As String, CegPAth As String)
        oWordApp = oWApp
        wListFiles = wLstFiles
        wFilePath = wFPath
        sRefLinkStyle = tRefLinkType
        sPubName = pName
        sCEGPath = CegPAth
        ' This call is required by the designer.
        InitializeComponent()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        'If rbAPA.Checked Then
        '    If VariableExists(aDoc, "CEGRefStyleType") = False Then
        '        aDoc.Variables.Add("CEGRefStyleType", "APA")
        '    Else
        '        aDoc.Variables("CEGRefStyleType").Value = "APA"
        '    End If
        'Else
        '    If VariableExists(aDoc, "CEGRefStyleType") = False Then
        '        aDoc.Variables.Add("CEGRefStyleType", "AMA")
        '    Else
        '        aDoc.Variables("CEGRefStyleType").Value = "AMA"
        '    End If
        'End If
        Try
            Dim K As Integer
            Dim curdocPath As String
            Dim curDoc As Word.Document
            'If sRefLinkStyle.ToUpper() = "BOOK" Then
            '    If DGVRefStyle.Rows.Item(0).Cells(1).Value <> "" Then
            '        For Each fDoc As String In wListFiles.Split("||")
            '            If fDoc <> String.Empty Then
            '                curdocPath = Path.Combine(wFilePath, fDoc)
            '                oWordApp.Documents(curdocPath).Activate()
            '                curDoc = oWordApp.ActiveDocument
            '                If VariableExists(curDoc, "CEGRefStyleType") = False Then
            '                    curDoc.Variables.Add("CEGRefStyleType", DGVRefStyle.Rows.Item(0).Cells(1).Value)
            '                Else
            '                    curDoc.Variables("CEGRefStyleType").Value = DGVRefStyle.Rows.Item(0).Cells(1).Value
            '                End If
            '            End If
            '        Next
            '    End If
            'Else
            '    For K = 0 To DGVRefStyle.RowCount() - 1
            '        If DGVRefStyle.Rows.Item(K).Cells(1).Value <> "" Then
            '            curdocPath = Path.Combine(wFilePath, DGVRefStyle.Rows.Item(K).Cells(0).Value)
            '            oWordApp.Documents(curdocPath).Activate()
            '            curDoc = oWordApp.ActiveDocument
            '            If VariableExists(curDoc, "CEGRefStyleType") = False Then
            '                curDoc.Variables.Add("CEGRefStyleType", DGVRefStyle.Rows.Item(K).Cells(1).Value)
            '            Else
            '                curDoc.Variables("CEGRefStyleType").Value = DGVRefStyle.Rows.Item(K).Cells(1).Value
            '            End If
            '        Else
            '            MessageBox.Show("not set variable for " & DGVRefStyle.Rows.Item(K).Cells(0).Value)
            '        End If
            '    Next
            'End If
            If CBRefStyle.Text <> vbNullString Then
                For Each fDoc As String In wListFiles.Split("||")
                    If fDoc <> String.Empty Then
                        curdocPath = Path.Combine(wFilePath, fDoc)
                        oWordApp.Documents(curdocPath).Activate()
                        curDoc = oWordApp.ActiveDocument
                        If VariableExists(curDoc, "CEGRefStyleType") = False Then
                            curDoc.Variables.Add("CEGRefStyleType", CBRefStyle.Text)
                        Else
                            curDoc.Variables("CEGRefStyleType").Value = CBRefStyle.Text
                        End If
                    End If
                Next
            Else
                MessageBox.Show("Please select the reference type here....")
            End If
        Catch ex As Exception
            MsgBox("ERROR---")
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmRefStyleType_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            Dim pRefIni As String = sCEGPath & "\Main\Reference_Styles"
            Dim lstRefIni As New List(Of String)
            lstRefIni = Directory.GetFiles(pRefIni, sPubName & "_*").Select(Function(x) Replace(Replace(Path.GetFileNameWithoutExtension(x).ToUpper(), "_REFINI", ""), sPubName.ToUpper() & "_", "")).ToList()
            If lstRefIni.Count > 0 Then
                lstRefIni.Add(" ")
                CBRefStyle.DataSource = lstRefIni
            Else
                MessageBox.Show("Ref style not found in the configuration in the path of " & pRefIni)
            End If

            ''Dim N As Integer

            'If sRefLinkStyle.ToUpper() = "BOOK" Then
            '    N = DGVRefStyle.Rows.Add()
            '    Dim objCBC As New DataGridViewComboBoxCell
            '    objCBC.DataSource = lstRefIni
            '    DGVRefStyle.Rows.Item(N).Cells(0).Value = "All Chapters"
            '    DGVRefStyle.Rows.Item(N).Cells(0).ReadOnly = True
            '    DGVRefStyle.Rows.Item(N).Cells(1) = objCBC
            'Else
            '    For Each fDoc As String In wListFiles.Split("||")
            '        If fDoc <> String.Empty Then
            '            N = DGVRefStyle.Rows.Add()
            '            DGVRefStyle.Rows.Item(N).Cells(0).Value = fDoc
            '            DGVRefStyle.Rows.Item(N).Cells(0).ReadOnly = True
            '            Dim objCBC As New DataGridViewComboBoxCell
            '            objCBC.DataSource = lstRefIni
            '            DGVRefStyle.Rows.Item(N).Cells(1) = objCBC
            '        End If
            '    Next
            'End If
        Catch ex As Exception
            MsgBox("Error:--")
        End Try

    End Sub

    Private Sub rbAPA_CheckedChanged(sender As Object, e As EventArgs)
        gRefStyleType = "APA"
    End Sub

    Private Sub rbAMA_CheckedChanged(sender As Object, e As EventArgs)
        gRefStyleType = "AMA"
    End Sub

    Private Sub cbALL_CheckedChanged(sender As Object, e As EventArgs)
        'If cbALL.Checked Then
        '    gbRefStyle.Enabled = True
        'Else
        '    gbRefStyle.Enabled = False
        'End If
    End Sub
End Class