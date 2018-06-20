Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word
Imports System.Text.RegularExpressions
Public Class frmReviewIndex

    Public oBKList As Dictionary(Of String, String)
    Public oActDoc As Document
    Public oActApp As Application
    Dim I As Integer
    Dim sBKName As String
    Dim sIndexValue As String
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
    End Sub


    Private Function ToProcess(bDelete As Boolean, bReplace As Boolean, bReplaceAll As Boolean, bPrevious As Boolean, bNext As Boolean)
        Try
            lblinfo.Text = ""
            Select Case True
                Case bNext
                    I = I + 1 : sBKName = oBKList.ElementAt(I).Key : sIndexValue = oBKList.ElementAt(I).Value
                    If oActDoc.Bookmarks.Exists(sBKName) = True Then
                        oActDoc.Bookmarks(sBKName).Range.Select()
                        Call oActDoc.ActiveWindow.ScrollIntoView(oActApp.Selection.Range, True)
                        Me.txtIndexText.Text = sIndexValue.Replace(ChrW(21), "").Replace(ChrW(9), "")
                    End If
                    Me.btnreplace.Enabled = False : Me.btnreplaceall.Enabled = False

                Case bPrevious
                    I = I - 1 : sBKName = oBKList.ElementAt(I).Key : sIndexValue = oBKList.ElementAt(I).Value
                    If oActDoc.Bookmarks.Exists(sBKName) = True Then
                        oActDoc.Bookmarks(sBKName).Range.Select()
                        Call oActDoc.ActiveWindow.ScrollIntoView(oActApp.Selection.Range, True)
                        Me.txtIndexText.Text = sIndexValue.Replace(ChrW(21), "").Replace(ChrW(9), "")
                    End If
                    Me.btnreplace.Enabled = False : Me.btnreplaceall.Enabled = False

                Case bDelete
                    If oActDoc.Bookmarks.Exists(sBKName) = True Then
                        oActDoc.Bookmarks(sBKName).Range.Delete()
                        If oBKList.ContainsKey(sBKName) = True Then Call oBKList.Remove(sBKName)
                        If oActDoc.Bookmarks.Exists(sBKName) = True Then Call oActDoc.Bookmarks(sBKName).Delete()
                    Else
                        Call MessageBox.Show("Unable to find the specific index term within the document", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If
                    lblinfo.Text = "Index term was successfully deleted within document"

                Case bReplace
                    Dim oIndexTextRng As Range = oActDoc.Bookmarks(sBKName).Range.Duplicate
                    'If oIndexTextRng.Characters.First.Text = ChrW(19) OrElse Regex.IsMatch(oIndexTextRng.Characters.First.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                    '    oIndexTextRng.SetRange(oIndexTextRng.Start + 1, oIndexTextRng.End)
                    'End If
                    'If oIndexTextRng.Characters.Last.Text = ChrW(21) OrElse Regex.IsMatch(oIndexTextRng.Characters.Last.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                    '    oIndexTextRng.SetRange(oIndexTextRng.Start, oIndexTextRng.End - 1)
                    'End If
                    Dim sNewIndexVal As String = Me.txtIndexText.Text.Replace(ChrW(21), "").Replace(ChrW(9), "")
                    oBKList.Item(sBKName) = sNewIndexVal
                    oIndexTextRng.Select()
                    oActDoc.Activate()
                    oActApp.Selection.Find.ClearFormatting() : oActApp.Selection.Find.Replacement.ClearFormatting()
                    oActApp.Selection.Find.Replacement.ClearFormatting()
                    With oActApp.Selection.Find
                        .Text = sIndexValue
                        .Font.Shading.BackgroundPatternColorIndex = WdColorIndex.wdGray25
                        .Replacement.Text = sNewIndexVal
                        .Replacement.Font.Shading.BackgroundPatternColorIndex = WdColorIndex.wdGray25
                        .Forward = True : .Wrap = WdFindWrap.wdFindContinue
                        .Format = True : .MatchCase = False : .MatchWholeWord = False
                        .MatchWildcards = False : .MatchSoundsLike = False : .MatchAllWordForms = False
                    End With
                    oActApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceOne)
                    lblinfo.Text = "Index term was successfully modified within document"

                Case bReplaceAll
                    Dim oIndexTextRng As Range = oActDoc.Bookmarks(sBKName).Range.Duplicate
                    'If oIndexTextRng.Characters.First.Text = ChrW(19) OrElse Regex.IsMatch(oIndexTextRng.Characters.First.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                    '    oIndexTextRng.SetRange(oIndexTextRng.Start + 1, oIndexTextRng.End)
                    'End If
                    'If oIndexTextRng.Characters.Last.Text = ChrW(21) OrElse Regex.IsMatch(oIndexTextRng.Characters.Last.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                    '    Call oIndexTextRng.SetRange(oIndexTextRng.Start, oIndexTextRng.End - 1)
                    'End If
                    Dim sNewIndexVal As String = Me.txtIndexText.Text.Replace(ChrW(21), "").Replace(ChrW(9), "")
                    oIndexTextRng.Text = sNewIndexVal
                    oActDoc.Activate()
                    oActApp.Selection.HomeKey(Unit:=WdUnits.wdStory)
                    oActApp.Selection.Find.ClearFormatting() : oActApp.Selection.Find.Replacement.ClearFormatting()
                    oActApp.Selection.Find.Replacement.ClearFormatting()
                    With oActApp.Selection.Find
                        .Text = sIndexValue
                        .Font.Shading.BackgroundPatternColorIndex = WdColorIndex.wdGray25
                        .Replacement.Text = sNewIndexVal
                        .Replacement.Font.Shading.BackgroundPatternColorIndex = WdColorIndex.wdGray25
                        .Forward = True : .Wrap = WdFindWrap.wdFindContinue
                        .Format = True : .MatchCase = False : .MatchWholeWord = False
                        .MatchWildcards = False : .MatchSoundsLike = False : .MatchAllWordForms = False
                    End With
                    oActApp.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
                    For Each oKV As KeyValuePair(Of String, String) In oBKList
                        If oKV.Value.ToLower.Contains(sIndexValue.ToLower) = True Then
                            Call oBKList.TryGetValue(oKV.Value, oKV.Value.Replace(sIndexValue, sNewIndexVal))
                        End If
                    Next
                    If Not oIndexTextRng Is Nothing Then oIndexTextRng.Select()
                    lblinfo.Text = "Index terms are successfully modified within document"
            End Select
            If I + 1 = oBKList.Count Then Me.btnNext.Enabled = False
            If I <> 0 AndAlso I - 1 <= oBKList.Count Then Me.btnPrevious.Enabled = True
            If I = 0 Then Me.btnPrevious.Enabled = False
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle)
        End Try
    End Function
   

    Private Sub frmReviewIndex_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        sBKName = oBKList.First.Key : sIndexValue = oBKList.First.Value
        If oActDoc.Bookmarks.Exists(sBKName) = True Then
            oActDoc.Bookmarks(sBKName).Range.Select()
            Call oActDoc.ActiveWindow.ScrollIntoView(oActApp.Selection.Range, True)
        End If
        Me.txtIndexText.Text = sIndexValue.Replace(ChrW(21), "").Replace(ChrW(9), "") : I = 1
        Me.btndelete.Enabled = True : Me.btnreplace.Enabled = False : Me.btnreplaceall.Enabled = False
        Me.btnNext.Enabled = True : Me.btnPrevious.Enabled = False : Me.btnExit.Enabled = True

        lblinfo.Text = String.Empty
    End Sub

    Private Sub btndelete_Click(sender As System.Object, e As System.EventArgs) Handles btndelete.Click
        Call ToProcess(True, False, False, False, False)
    End Sub


    Private Sub btnreplace_Click(sender As System.Object, e As System.EventArgs) Handles btnreplace.Click
        Call ToProcess(False, True, False, False, False)
    End Sub

    Private Sub btnreplaceall_Click(sender As System.Object, e As System.EventArgs) Handles btnreplaceall.Click
        Call ToProcess(False, False, True, False, False)
    End Sub

    Private Sub btnPrevious_Click(sender As System.Object, e As System.EventArgs) Handles btnPrevious.Click
        Call ToProcess(False, False, False, True, False)
    End Sub

    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click
        Call ToProcess(False, False, False, False, True)
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        oActApp.Selection.HomeKey(Unit:=WdUnits.wdStory)
        Call Me.Dispose(True)
    End Sub

    Private Sub txtIndexText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIndexText.TextChanged
        If String.IsNullOrEmpty(sIndexValue) = False Then
            If sIndexValue <> Me.txtIndexText.Text Then
                Me.btnreplace.Enabled = True : Me.btnreplaceall.Enabled = True
            End If
        Else
            Me.btnreplace.Enabled = True : Me.btnreplaceall.Enabled = True
        End If

    End Sub
End Class