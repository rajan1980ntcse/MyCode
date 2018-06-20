Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports NCalc
Imports CEGINI

Public Class clsAbsractKeyword
    Public dicQueryAbs As New Dictionary(Of String, String)

    Public Function OUPAbstractKeywordMain(oWordApp As Word.Application, wListFiles As String, wFilePath As String, sXMLPath As String)
        Dim pQueryINI As String : Dim whichFile As String = String.Empty
        Dim sTemp As String
        Dim sRule As String
        Dim I As Integer
        Dim oWordDoc As Word.Document
        'Dim oTestApp As Word.Application
        Dim ranDoc As Word.Range
        'oTestApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application")

        'Dim oXMLDoc As New XmlDocument() : Dim oXMLNodeList As XmlNodeList
        'Do While Left(wListFiles, 2) = "||"
        '    wListFiles = Mid(wListFiles, 3, Len(wListFiles))
        'Loop
        'Do While Right(wListFiles, 2) = "||"
        '    wListFiles = Mid(wListFiles, 1, Len(wListFiles) - 2)
        'Loop

        'Call oXMLDoc.Load(sXMLPath) : oXMLNodeList = oXMLDoc.SelectNodes("//File[@Filter='Abs']")
        'If oXMLNodeList.Count > 0 Then
        '    whichFile = oXMLNodeList(0).InnerText : whichFile = Replace(whichFile, "\\", "\")
        'End If


        Try
            pQueryINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            If File.Exists(pQueryINI) Then
                Dim oReadINI As New clsINI(pQueryINI)
                I = 1
                Do While (True)
                    sRule = oReadINI.INIReadValue("Query", "AQ" & I)
                    If sRule = String.Empty Then Exit Do
                    dicQueryAbs.Add("AQ" & I, sRule)
                    I = I + 1
                Loop
                For Each streachFile In wListFiles.Split("||".ToCharArray(), StringSplitOptions.RemoveEmptyEntries)

                    'If String.IsNullOrEmpty(whichFile) = True Then
                    '    MessageBox.Show("No abstract file in this book", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    '    GoTo AddDocVar
                    'End If

                    oWordApp.Documents(Path.Combine(wFilePath, streachFile)).Activate()
                    oWordDoc = oWordApp.ActiveDocument()

                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Call AbstractPreprocess(oWordDoc, oReadINI)                 'Abstract preprocess
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    oWordDoc.Activate()

                    I = 1
                    Do While (True)
                        sRule = oReadINI.INIReadValue("QueryCondition", "Rule" & I)
                        If sRule = String.Empty Then Exit Do
                        Dim eCont As String()
                        Dim sCont As String()
                        For Each sTemp In sRule.Split("#")
                            If Trim(sTemp) <> "" Then
                                If InStr(1, sTemp, "@") > 0 Then
                                    sCont = sTemp.Split("@")
                                Else
                                    eCont = sTemp.Split("|")
                                End If
                            End If
                        Next
                        For j = LBound(eCont) To UBound(eCont)
                            If eCont(j) <> "" Then
                                sRule = oReadINI.INIReadValue("QueryCondition", Trim(eCont(j)))
                                ranDoc = oWordDoc.Range
                                Do While (True)
                                    Dim ranGroup As Word.Range = GetRangeOfEachChapter(sCont(0), sCont(1), ranDoc, oWordDoc)
                                    If Not ranGroup Is Nothing Then
                                        ranGroup.Select()
                                        CheckConditionforQuery(ranGroup, sRule, oWordDoc, oReadINI)
                                    End If
                                    If ranDoc Is Nothing Then Exit Do
                                Loop

                            End If
                        Next
                        I = I + 1
                    Loop
                Next
                oWordApp.Selection.Collapse()
            Else
                MessageBox.Show("File not found : " & pQueryINI, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
AddDocVar:
            For Each xVrnt In Split(wListFiles, "||")
                If xVrnt <> "" Then
                    oWordApp.Documents(xVrnt).Activate()
                    oWordDoc = oWordApp.ActiveDocument
                    AddQCIteminCollection("AbstractkeyCheck", oWordDoc)
                End If
            Next
            oWordDoc.ActiveWindow.View.Type = Word.WdViewType.wdPrintView
            oWordDoc.ActiveWindow.View.ShowRevisionsAndComments = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Function GetRangeOfEachChapter(sStyle As String, eStyle As String, ByRef ranDoc As Word.Range, wDoc As Word.Document) As Word.Range
        Try
            Dim ranDup As Word.Range
            Dim ranEnd As Word.Range
            Dim ranSt As Word.Range
            Dim ranLt As Word.Range
            ranDup = ranDoc.Duplicate
            ranDup.Find.ClearFormatting()
            ranDup.Find.Style = sStyle
            ranDup.Find.Text = ""
            If ranDup.Find.Execute = True Then
                ranDup.Select()
                ranSt = ranDup.Duplicate
            End If
            If Not ranSt Is Nothing Then
                If ranSt.End + 1 <> ranDoc.End Then
                    ranDoc = wDoc.Range(ranSt.End + 1, ranDoc.End)
                Else
                    ranDoc = wDoc.Range(ranSt.End, ranDoc.End)
                End If
            End If

            ranEnd = ranDoc.Duplicate
            ranEnd.Find.ClearFormatting()
            ranEnd.Find.Style = eStyle
            ranEnd.Find.Text = ""
            If ranEnd.Find.Execute = True Then
                ranEnd.Select()
                ranLt = ranEnd.Duplicate
            End If
            If Not ranSt Is Nothing And Not ranLt Is Nothing Then
                ranDoc = wDoc.Range(ranLt.Start, ranDoc.End)
                If ranSt.Start <> ranLt.Start Then
                    GetRangeOfEachChapter = wDoc.Range(ranSt.Start, ranLt.Start - 1)
                Else
                    GetRangeOfEachChapter = Nothing
                End If
            ElseIf Not ranSt Is Nothing Then
                GetRangeOfEachChapter = wDoc.Range(ranSt.Start, ranDoc.End)
                ranDoc = Nothing
            Else
                ranDoc = Nothing
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Sub CheckConditionforQuery(ranCheck As Word.Range, sRule As String, wDoc As Word.Document, oReadINI As clsINI)
        Try
            Dim styleName As String
            Dim sCont As String
            Dim Query As String
            Dim Temp() As String
            Dim strTemp() As String
            Dim textStyle As String
            Temp = sRule.Split("@")
            styleName = Temp(0)
            textStyle = GetStyleText(ranCheck, styleName, wDoc)

            For Each sCont In Temp(1).Split("||")
                If sCont <> String.Empty Then
                    strTemp = sCont.Split("#")
                    sCont = strTemp(0)
                    dicQueryAbs.TryGetValue(strTemp(1), Query)
                    Dim qType As String = Regex.Match(sCont, "\[(.*?)\]", RegexOptions.IgnoreCase).Groups(1).ToString()
                    Call InsertAbractKeywordQuery(ranCheck, textStyle, styleName, Query, sCont, qType, wDoc, oReadINI)
                    Query = ""
                    sCont = ""
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    Private Sub InsertAbractKeywordQuery(oTtlRngPara As Word.Range, textStyle As String, sName As String, sQuery As String, sCont As String, qType As String, wDoc As Word.Document, oReadINI As clsINI)
        Try
            Dim Cnt As Integer
            Dim oTtlRngDup As Word.Range
            If textStyle = String.Empty Then
                Cnt = 0
            Else
                'pQueryINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
                Dim qParttern As String = oReadINI.IniReadValue("QueryCondition", qType)
                'Dim qParttern As String = oReadINI ReadINI(pQueryINI, "QueryCondition", qType, String.Empty, False)
                If qParttern <> String.Empty Then
                    Cnt = Regex.Matches(textStyle, qParttern).Count
                Else
                    MessageBox.Show("Please check query ini....", sMsgTitle)
                End If
            End If

            Dim expr As New Expression(sCont)
            expr.Parameters(qType) = Cnt
            If expr.Evaluate Then
                oTtlRngDup = oTtlRngPara.Duplicate
                oTtlRngDup.Find.ClearFormatting()
                oTtlRngDup.Find.Style = sName
                oTtlRngDup.Find.Text = String.Empty
                If oTtlRngDup.Find.Execute = True Then
                    oTtlRngDup.Select()
                    If sQuery.ToLower().Contains("ceg warning") Then
                        MessageBox.Show(sQuery, sMsgTitle)
                    Else
                        wDoc.Comments.Add(oTtlRngDup, sQuery)
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function AbstractPreprocess(wDoc As Word.Document, oReadINI As clsINI)
        Try
            Dim I As Integer
            Dim sRule As String
            Dim sTemp() As String
            Dim strFR() As String
            Dim strCont As String
            I = 1
            'pQueryINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Do While (True)
                'sRule = ReadINI(pQueryINI, "FindReplace", "Find" & I, String.Empty, False)
                sRule = oReadINI.IniReadValue("FindReplace", "Find" & I)
                If sRule = String.Empty Then Exit Do
                sTemp = sRule.Split("#")
                For Each strCont In sTemp(1).Split("||")
                    If strCont <> String.Empty Then
                        strFR = strCont.Split("=")
                        FindandReplaceBasedOnStyle(strFR(0), strFR(1), sTemp(0), wDoc)
                    End If
                Next
                I = I + 1
            Loop

            I = 1
            sTemp = Nothing
            strFR = Nothing
            sRule = String.Empty
            Do While (True)
                'sRule = ReadINI(pQueryINI, "FindReplace", "Format" & I, String.Empty, False)
                sRule = oReadINI.IniReadValue("FindReplace", "Format" & I)
                If sRule = String.Empty Then Exit Do
                sTemp = sRule.Split("#")
                For Each strCont In sTemp(1).Split("||")
                    If strCont <> String.Empty Then
                        strFR = strCont.Split("=")
                        FormatingBaseedOnStyle(strFR(0), strFR(1), sTemp(0), wDoc)
                    End If
                Next
                I = I + 1
            Loop

        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Function FormatingBaseedOnStyle(typeFormat As String, valueFormat As String, styleName As String, wDoc As Word.Document)
        Try
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
                    With ranDoc.Find
                        .Style = styleName
                        .Text = String.Empty
                        Select Case (UCase(typeFormat))
                            Case "BOLD"
                                .Font.Bold = Not Boolean.Parse(valueFormat)
                            Case "ITALIC"
                                .Font.Italic = Not Boolean.Parse(valueFormat)
                            Case "UNDERLINE"
                                .Font.Italic = Not Boolean.Parse(valueFormat)
                            Case "SMALLCAPS"
                                .Font.SmallCaps = Not Boolean.Parse(valueFormat)
                            Case "STRIKETHROUGH"
                                .Font.StrikeThrough = Not Boolean.Parse(valueFormat)
                        End Select
                        .Wrap = Word.WdFindWrap.wdFindContinue
                        .ClearFormatting()
                        .Replacement.ClearFormatting()
                        .Replacement.Highlight = False
                        Select Case (UCase(typeFormat))
                            Case "BOLD"
                                .Replacement.Font.Bold = Boolean.Parse(valueFormat)
                            Case "ITALIC"
                                .Replacement.Font.Italic = Boolean.Parse(valueFormat)
                            Case "UNDERLINE"
                                .Replacement.Font.Italic = Boolean.Parse(valueFormat)
                            Case "SMALLCAPS"
                                .Replacement.Font.SmallCaps = Boolean.Parse(valueFormat)
                            Case "STRIKETHROUGH"
                                .Replacement.Font.StrikeThrough = Boolean.Parse(valueFormat)
                        End Select
                        .Execute(Replace:=Word.WdReplace.wdReplaceAll)
                    End With
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
End Class
