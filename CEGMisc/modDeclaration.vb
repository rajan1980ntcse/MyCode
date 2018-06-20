Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Word
Imports System.IO
Imports Word = Microsoft.Office.Interop.Word

Module modDeclaration
    Public Const sMsgTitle As String = "CEGenius"

    '###################### Common function to refer all class #######################
    Public Function AddQCIteminCollection(wString As String, wDoc As Word.Document)
        Dim oText As String
        If AnyVariableAdds(wDoc) = False Then wDoc.Variables.Add("Add")
        oText = wDoc.Variables("Add").Value
        If InStr(1, oText, "||" & wString & "||", vbTextCompare) = 0 Then
            wDoc.Variables("Add").Value = oText & "||" & wString & "||"
        End If
    End Function
    Public Function VariableExists(wDoc As Word.Document, vName As String) As Boolean
        Dim Desc As String
        On Error GoTo ErrJump
        Desc = wDoc.Variables(vName).Value
        VariableExists = True : Exit Function
ErrJump:
        Err.Clear()
        VariableExists = False
    End Function

    Private Function AnyVariableAdds(wDoc As Word.Document) As Boolean
        Dim Desc As String
        On Error GoTo ErrJump
        Desc = wDoc.Variables("Add").Value
        AnyVariableAdds = True : Exit Function
ErrJump:
        Err.Clear()
        AnyVariableAdds = False
    End Function

    Public Function GetStyleText(oTtlRngPara As Word.Range, sName As String, wDoc As Word.Document) As String
        Try
            Dim oTtlRngDup As Word.Range
            Dim oFRng As Word.Range
            Dim sTemp As String
            sTemp = ""
            oTtlRngDup = oTtlRngPara.Duplicate
            oTtlRngDup.Find.ClearFormatting()
            oTtlRngDup.Find.Style = sName
            oTtlRngDup.Find.Text = ""
            Do While oTtlRngDup.Find.Execute = True
                oTtlRngDup.Select()
                sTemp = sTemp & vbCrLf & oTtlRngDup.Text.Trim()
                oFRng = oTtlRngDup.Duplicate
                If oFRng.End >= oTtlRngPara.End Then Exit Do
                oTtlRngDup = wDoc.Range(oFRng.End, oTtlRngPara.End)
                oTtlRngDup.Find.Style = sName
                oTtlRngDup.Find.Text = ""
            Loop
            GetStyleText = sTemp
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Function



    Public Function FindandReplaceBasedOnStyle(sFind As String, sReplace As String, styleName As String, wDoc As Word.Document)
        Try
            Dim wAPP As Word.Application
            wAPP = wDoc.Application
            Dim ranDoc As Word.Range
            ranDoc = Nothing
            Dim I As Integer
            For I = 1 To 3
                Select Case (I)
                    Case 1
                        ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                    Case 2
                        If wDoc.Footnotes.Count > 0 Then
                            ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory)
                        Else
                            ranDoc = Nothing
                        End If
                    Case 3
                        If wDoc.Endnotes.Count > 0 Then
                            ranDoc = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory)
                        Else
                            ranDoc = Nothing
                        End If
                End Select
                If Not ranDoc Is Nothing Then
                    Dim ranFind As Word.Range
                    ranFind = ranDoc.Duplicate
                    ranFind.Find.ClearFormatting()
                    ranFind.Find.Style = wDoc.Styles(styleName)
                    ranFind.Find.Text = ""
                    'Do While (ranDoc.Find.Execute = True)
                    '    ranDoc.Text = sReplace
                    '    ranDoc.Collapse(Word.WdCollapseDirection.wdCollapseStart)
                    'Loop
                    Do While (ranFind.Find.Execute = True)
                        Dim ranSel As Word.Range
                        ranFind.Select()
                        ranSel = wAPP.Selection.Range
                        ranSel.Find.ClearFormatting()
                        ranSel.Find.Text = sFind
                        Do While ranSel.Find.Execute
                            ranSel.Select()
                            If ranSel.InRange(ranFind) Then
                                wAPP.Selection.Range.Text = sReplace
                                ranSel.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            End If
                        Loop

                        '' ranSel.Find.Execute(sFind, False, True, False, False, False, True, Word.WdFindWrap.wdFindContinue, False, sReplace, Word.WdReplace.wdReplaceAll, False, False, False, False)
                        'With ranSel.Find
                        '    .Text = sFind
                        '    .Replacement.Text = sReplace
                        '    ''.Replacement.Highlight = Word.WdColorIndex.wdBrightGreen
                        '    .Forward = True
                        '    .Wrap = Word.WdFindWrap.wdFindContinue
                        '    .Format = False
                        '    .MatchCase = False
                        '    .MatchWholeWord = True
                        '    .MatchWildcards = False
                        '    .MatchSoundsLike = False
                        '    .MatchAllWordForms = False
                        '    .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                        'End With
                        ''wAPP.Selection.Find.Execute()
                        ranSel = Nothing
                        ranFind.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Loop
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    '###################### Common function to refer all class #######################



    Public Function ToGetWordFiles() As Boolean
        Try
            Dim oFileList As New List(Of String)

            Dim sServerPath As String = "W:\CWWRIT\"
            Dim dirs As String() = Directory.GetDirectories(sServerPath, "vp*", SearchOption.TopDirectoryOnly)
            For Each sDirs As String In dirs
                Dim osubdir As String() = Directory.GetDirectories(sDirs, "word*", SearchOption.AllDirectories)
                For Each sSubDir As String In osubdir
                    Dim sfiles As String() = Directory.GetFiles(sSubDir, "*.docx", SearchOption.TopDirectoryOnly)
                    For Each sfile As String In sfiles
                        'MessageBox.Show(sfile)
                        oFileList.Add(sfile)
                    Next
                Next
            Next

            Dim WordApp As Application = Marshal.GetActiveObject("Word.Application")
            Dim WordDoc As Document = WordApp.ActiveDocument

            Dim oFindRng As Range : Dim lPrevRngEnd As Long
            Dim oActDocRng As Range : Dim oTempRng As Range
            Dim sFindstyle As String = "†Reference"

            Dim I As Integer = 0
            For Each sFilename As String In oFileList
                Dim oDoc As Document = WordApp.Documents.Add(sFilename)
                Dim oRng As Range = oDoc.StoryRanges(WdStoryType.wdMainTextStory)
                If Not oRng Is Nothing AndAlso ToCheckStyle(oDoc, sFindstyle) = True Then
                    oFindRng = oRng.Duplicate
                    With oFindRng.Find
                        .ClearFormatting() : .Replacement.ClearFormatting() : .ClearAllFuzzyOptions()
                        .Text = "" : .Replacement.Text = "" : .MatchWildcards = False : .MatchWholeWord = False : .MatchCase = False
                        .Style = sFindstyle
                    End With
                    Do While oFindRng.Find.Execute = True
                        If oFindRng.End <> lPrevRngEnd Then
                            'If oFindRng.Characters.Last.Text = vbCr Then oFindRng.SetRange(oFindRng.Start, oFindRng.End - 1)
                            'If oFindRng.Characters.First.Text = vbCr Then oFindRng.SetRange(oFindRng.Start + 1, oFindRng.End)

                            oTempRng = oFindRng.Duplicate : I = I + 1
                            oActDocRng = WordDoc.Range.Duplicate : oActDocRng.SetRange(oActDocRng.End, oActDocRng.End)
                            'Call oActDocRng.InsertAfter(vbCrLf)
                            WordDoc.Activate()
                            oActDocRng.FormattedText = oTempRng
                            oFindRng.Collapse(WdCollapseDirection.wdCollapseEnd) : lPrevRngEnd = oFindRng.End
                            If oFindRng.End = oDoc.StoryRanges(oFindRng.StoryType).End - 1 OrElse _
                                    oFindRng.End = oDoc.StoryRanges(oFindRng.StoryType).End OrElse _
                                    oFindRng.InRange(oDoc.Bookmarks("\EndofDoc").Range) = True Then
                                Exit Do
                            End If
                            If I > 500 Then
                                Exit Do
                            End If
                        Else
                            Call oFindRng.Collapse(WdCollapseDirection.wdCollapseEnd)
                        End If
                    Loop
                End If
                Call oDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try




    End Function


    Private Function ToCheckStyle(oDoc As Document, sStylename As String) As Boolean
        Try
            Dim sdes As Style = oDoc.Styles(sStylename)
            ToCheckStyle = True
        Catch ex As Exception
            ex.Data.Clear()
            ToCheckStyle = False
        End Try
    End Function
    Public Sub WriteHtmlFile(HPath As String, sWrite As String)
        Dim objW As New StreamWriter(HPath)
        objW.Write(sWrite)
        objW.Close()
    End Sub
End Module