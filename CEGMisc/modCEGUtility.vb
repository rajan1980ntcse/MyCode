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
Imports CEGINI
Module modCEGUtility
    Public wAPP As Word.Application
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Public Function CEGLanguageForBook(WordApp As Word.Application, wLstFiles As String, wFPath As String, sINI As String)
        Try
            Dim ocls As New frmLanguageSelect(WordApp, wLstFiles, wFPath)
            WordApp.ActiveDocument.Activate()
            ocls.ShowDialog()

        Catch ex As Exception
            MessageBox.Show("Error : CEGLanguageForBook", "CE-Genious")
        End Try

    End Function
    Public Function MissingFloatsObjectLogCreation(oWrdApp As Word.Application, wListFiles As String, wFilePath As String) As Boolean
        Try
            Dim wDoc As Word.Document
            Dim pQueryINI As String
            Dim floatStyle As String
            pQueryINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New clsINI(pQueryINI)
            floatStyle = oReadINI.INIReadValue("MissingFloats", "StyleName")
            If floatStyle = "" Then
                MessageBox.Show("Error : configuration missing  ", "Missing floats - CEGMISC")
                Exit Function
            End If
            For Each wordFileName In wListFiles.Split("||")
                If wordFileName <> String.Empty Then
                    oWrdApp.Documents(Path.Combine(wFilePath, wordFileName)).Activate() : wDoc = oWrdApp.ActiveDocument
                    CollectMissingFloatsObject(wDoc, floatStyle)
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error : MissingFloatsObjectLogCreation", "MISSING FLOATS")
            Exit Function
        End Try
    End Function
    Public Function CollectMissingFloatsObject(wDoc As Word.Document, ignoreStyleName As String)
        Dim rMatches As MatchCollection
        Dim oRng As Word.Range
        Dim lstMissingFloats As List(Of String)
        Dim fPattern() As String = {"(\bfig(ure)?(.)?(s)?\b|\btab(le)?(.)?(s)?\b|\bbox(.)?(s)?\b)([\u00A0\.\-\s*]+)(([,|,\s*|\s*and\s*|\s*&\s*|\-|\u2013]+)?([0-9\.]+(\()?([A-z])?(\))?))+", "(\bfig(ure)?(.)?(s)?\b|\btab(le)?(.)?(s)?\b|\bbox(.)?(s)?\b)([\u00A0\.\-\s*]+)(([,|,\s*|\s*and\s*|\s*&\s*|\-|\u2013]+)?\b([IVLXDCM\.]+\b(\()?([A-z])?(\))?))+", "(\bEquations?|\bEq(u|n)s?|\bEqs?)([\.\s]+)(\(\d+\.\d+[a-z]?\)|\(\d+[a-z]?\)|\d+\.\d+[a-z]?|\d+[a-z]?)(\s*(and|through|\,|or|to|&)\s*(\(\d+\.\d+[a-z]?\)|\(\d+[a-z]?\)|\d+\.\d+[a-z]?|\d+[a-z]?))*"}
        Try
            For I = 1 To 3
                Select Case I
                    Case 1 : oRng = wDoc.Content.Duplicate
                    Case 2 : If wDoc.Footnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory).Duplicate Else oRng = Nothing
                    Case 3 : If wDoc.Endnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory).Duplicate Else oRng = Nothing
                End Select
                If Not IsNothing(oRng) Then
                    Dim ranDoc As Word.Range
                    ranDoc = oRng.Duplicate
                    For k As Integer = LBound(fPattern) To UBound(fPattern)
                        rMatches = Regex.Matches(oRng.Text, fPattern(k))
                        For Each mat As Match In rMatches
                            ranDoc = oRng.Duplicate
                            ranDoc.Find.ClearFormatting()
                            ranDoc.Find.Text = mat.Value
                            Do While ranDoc.Find.Execute = True
                                If Not InStr(1, ranDoc.Style, ignoreStyleName) > 0 Then
                                    lstMissingFloats.Add(ranDoc.Text)
                                End If
                            Loop
                        Next
                    Next
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error : CollectMissingFloatsObject", "MISSING FLOATS")
        End Try
    End Function
    Public Function FigCaptionLogCreation(oWrdApp As Word.Application, wListFiles As String, wFilePath As String) As Boolean
        Try
            Dim wDoc As Word.Document
            Dim pQueryINI As String
            Dim sRule As String
            Dim sFigCapFileName As String
            pQueryINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New clsINI(pQueryINI)
            sRule = oReadINI.INIReadValue("FigCaption", "StyleName")
            If sRule = "" Then
                MessageBox.Show("Error : Config file not found ", "Fig caption log - CEG Misc")
                Exit Function
            End If

            Dim xMetaFilePath = Directory.GetFiles(wFilePath, "*.xml").Where(Function(fi) fi.ToLower().EndsWith("_metainfo.xml")).FirstOrDefault().ToString()
            If xMetaFilePath = "" Then
                MessageBox.Show("Error : metainfo.xml file not found ", "Fig caption log - CEG Misc")
                Exit Function
            End If
            Try
                Dim xDoc As XDocument = XDocument.Load(xMetaFilePath)
                Dim xFileNode As XElement
                xFileNode = xDoc.Descendants("Files").Descendants("File").LastOrDefault()
                ''01_Chap 01_ Weiss ''01_FigCaption_ Weiss - Copy
                'Dim FNumber As Integer = Convert.ToInt32(xFileNode.Attribute("Name").Value.Split("_")(0)) + 1
                'Dim AName As String = xFileNode.Attribute("Name").Value.Split("_")(2)
                'Dim ISBN As String = xDoc.Descendants("ISBN").Descendants("Hardback").FirstOrDefault().Value
                Dim FNumber As Integer = Convert.ToInt32(xFileNode.Attribute("Name").Value.Split("_")(0)) + 1
                Dim AName As String = xFileNode.Attribute("Name").Value.Split("_")(2)
                ''Dim ISBN As String = xDoc.Descendants("ISBN").Descendants("Hardback").FirstOrDefault().Value
                If AName = "" Or FNumber = 0 Then
                    sFigCapFileName = "OUP_FigCaption"
                Else
                    ''sFigCapFileName = FNumber.ToString("00") & "_" & ISBN & "_" & AName & "_FigCaption"
                    sFigCapFileName = FNumber.ToString("00") & "_FigCaption_" & AName
                End If
            Catch ex As Exception
                MessageBox.Show("Error : Metainfo.xml file", "Fig caption log - CEG Misc")
                Exit Function
            End Try


            'Name="02_52545544_JI_Chap 02.doc"

            '<last serial number>_<ISBN Number><Author surname><FigCap>
            Dim dicFigCap As New Dictionary(Of Word.Document, List(Of Word.Range))
            If Not (wListFiles.Contains("||")) Then String.Concat(wListFiles, "||")
            For Each wordFileName In wListFiles.Split("||")
                If wordFileName <> String.Empty Then
                    oWrdApp.Documents(Path.Combine(wFilePath, wordFileName)).Activate() : wDoc = oWrdApp.ActiveDocument
                    Dim figrangeList As New List(Of Word.Range)
                    'Dim figStyle() = {"†Figure_Caption", "FGC", "FGN", "FGS", "FFN"}
                    Dim figStyle() = sRule.Split("|")
                    For i = LBound(figStyle) To UBound(figStyle)
                        If AutoStyleExists(figStyle(i), wDoc) = True Then
                            Dim ranDoc As Word.Range
                            ranDoc = wDoc.Content
                            ranDoc.Find.ClearFormatting()
                            ranDoc.Find.Text = ""
                            ranDoc.Find.Style = figStyle(i)
                            Do While ranDoc.Find.Execute = True
                                ranDoc.Select()
                                figrangeList.Add(oWrdApp.Selection.Range)
                                ' oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                'oWrdApp.Selection.Range.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                'ranDoc = wDoc.Range(oWrdApp.Selection.Start, wDoc.Range.End)
                                'ranDoc.Find.Text = ""
                                'ranDoc.Find.Style = figStyle(i)
                            Loop
                        End If
                    Next
                    If figrangeList.Count > 0 Then
                        dicFigCap.Add(wDoc, figrangeList)
                    Else
                        dicFigCap.Add(wDoc, New List(Of Word.Range))
                    End If
                    wDoc = Nothing
                End If
            Next

            Dim oActDoc As Word.Document
            oActDoc = oWrdApp.Documents.Add()
            oWrdApp.Visible = True : oWrdApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
            oActDoc.Activate()
            If dicFigCap.Count > 0 Then
                For Each fRange As KeyValuePair(Of Word.Document, List(Of Word.Range)) In dicFigCap
                    oWrdApp.Selection.Range.Text = fRange.Key.Name
                    oWrdApp.Selection.Range.Style = "Heading 1"
                    oWrdApp.Selection.MoveEnd(Word.WdUnits.wdParagraph, 1)
                    oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    oWrdApp.Selection.InsertParagraphAfter()
                    oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Dim lstRange As List(Of Word.Range) = fRange.Value
                    If lstRange.Count > 0 And Not lstRange Is Nothing Then
                        For Each figRange In lstRange.OrderBy(Function(x) x.Start)
                            oWrdApp.Selection.Range.FormattedText = figRange
                            oWrdApp.Selection.MoveEnd(Word.WdUnits.wdParagraph, 1)
                            oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Next
                    Else
                        oWrdApp.Selection.Range.Text = "Not Found"
                        oWrdApp.Selection.MoveEnd(Word.WdUnits.wdParagraph, 1)
                        oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        oWrdApp.Selection.InsertParagraphAfter()
                        oWrdApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    End If
                Next
            End If
            oActDoc.SaveAs(Path.Combine(wFilePath, sFigCapFileName), Word.WdSaveFormat.wdFormatDocument, AddToRecentFiles:=False)
            oActDoc.Close(Word.WdSaveOptions.wdSaveChanges)
            wDoc = Nothing
            For Each fdict As KeyValuePair(Of Word.Document, List(Of Word.Range)) In dicFigCap
                oWrdApp.Documents(fdict.Key.Name).Activate()
                wDoc = oWrdApp.ActiveDocument
                'fdict.Key.Activate()
                Dim lstRange As List(Of Word.Range) = fdict.Value
                For Each figRange In lstRange.OrderBy(Function(x) x.Start)
                    wDoc.Range(figRange.Start, figRange.End).Delete()
                Next
            Next
            wDoc = Nothing
            For Each xVrnt In Split(wListFiles, "||")
                If xVrnt <> String.Empty Then
                    oWrdApp.Documents(xVrnt).Activate()
                    wDoc = oWrdApp.ActiveDocument
                    AddQCIteminCollection("FigCaptionLog", wDoc)
                End If
            Next
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, "Fig caption log - CEG Misc")
        End Try
    End Function
    Public Function FormatChangeAfterRSTinReference(jName As String, pName As String, oWrdapp As Word.Application, wDoc As Word.Document, pJourConfig As String)
        'OUP EXBTOJ journal client requirement Journal title as roman format
        'ReferenceFormat		=‡ref_titleJournal*roman
        Try
            Dim ranDoc As Word.Range
            ranDoc = wDoc.Content
            Dim oReadINI As New clsINI(pJourConfig)
            Dim strRefFomating = oReadINI.INIReadValue(pName + "@" + jName, "ReferenceFormat")
            If Not String.IsNullOrEmpty(strRefFomating) Then
                Select Case (LCase(strRefFomating.Split("*")(1)))
                    Case "roman"
                        ranDoc.Find.ClearFormatting()
                        ranDoc.Find.Style = strRefFomating.Split("*")(0)
                        ranDoc.Find.Text = ""
                        Do While (ranDoc.Find.Execute)
                            ranDoc.Font.Bold = False
                            ranDoc.Font.Italic = False
                        Loop
                End Select
            End If
        Catch ex As Exception

        End Try
    End Function
    Public Function ReportUsedFontFromListOfDocument(WordApp As Word.Application, wLstFiles As String, wFPath As String) As Boolean
        Dim wDoc As Word.Document
        Dim dictReport As New ArrayList
        Dim sHtmlText As String
        Dim strTemp As String = ""
        Try
            wAPP = WordApp
            For Each fDoc As String In wLstFiles.Split("||")
                If fDoc <> String.Empty Then
                    strTemp = ""
                    WordApp.Documents(Path.Combine(wFPath, fDoc)).Activate() : wDoc = WordApp.ActiveDocument
                    ''sHtmlText = sHtmlText + "<p class=MsoNormal style='mso-outline-level:2'><font size=3 color='Green'><b>" & fDoc & "</b></font></p>"

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
                            Dim C As Integer
                            For C = 1 To ranDoc.Characters.Count()
                                If Not dictReport.Contains(ranDoc.Characters(C).Font.Name) Then
                                    dictReport.Add(ranDoc.Characters(C).Font.Name)
                                End If
                            Next
                        End If
                    Next
                End If
            Next
            If dictReport.Count > 0 Then
                For Each stFont In dictReport
                    sHtmlText = sHtmlText + "<p>" + stFont + "</p>"
                Next

                Dim sR As New StreamWriter(Path.Combine(wFPath, "UsedFontReport.html"))
                sR.Write("<html><head><META http-equiv='Content-Type' content=text/html charset=utf-8></head><body>")
                sR.Write(sHtmlText)
                sR.Write("</body></html>")
                sR.Close()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function
    Public Function ReportUsedStylesFromListOfDocument(WordApp As Word.Application, wLstFiles As String, wFPath As String) As Boolean
        Dim wDoc As Word.Document
        Dim dictReport As New Dictionary(Of String, String)
        Dim sHtmlText As String
        Dim strTemp As String = ""
        Try
            wAPP = WordApp
            For Each fDoc As String In wLstFiles.Split("||")
                If fDoc <> String.Empty Then
                    strTemp = ""
                    WordApp.Documents(Path.Combine(wFPath, fDoc)).Activate() : wDoc = WordApp.ActiveDocument
                    sHtmlText = sHtmlText + "<p class=MsoNormal style='mso-outline-level:2'><font size=3 color='Green'><b>" & fDoc & "</b></font></p>"

                    strTemp = GetUsedStylesInTheDocument(wDoc)
                    If strTemp <> "" Then
                        sHtmlText = sHtmlText + strTemp
                    Else
                        sHtmlText = sHtmlText + "<p class=MsoNormal style='mso-outline-level:2'><font size=3 color='Red'><b>None</b></font></p>"
                    End If
                End If
            Next
            Dim sR As New StreamWriter(Path.Combine(wFPath, "ReportUsedStyle.html"))
            sR.Write("<html><body><head><META http-equiv='Content-Type' content=text/html charset=utf-8></head>")
            sR.Write(sHtmlText)
            sR.Write("</body></html>")
            sR.Close()

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function GetUsedStylesInTheDocument(fSDoc As Word.Document) As String
        Dim oVrnt
        Dim mnyCount As Integer
        Dim oHString As String
        Dim tmpTxt As String
        Dim tmpPath As String
        Dim kStyl As Word.Style
        Dim tmplt

        fSDoc.Activate()
        RemoveFormattinginParas()
        tmpTxt = ""
        For Each kStyl In fSDoc.Styles
            oVrnt = kStyl.NameLocal : fSDoc.UndoClear()

            If oVrnt <> "" Then
                mnyCount = 0 : oHString = ""
                GetStyleCountMis(oVrnt & "", mnyCount, oHString, "C", vbNull, vbNull, False)
                If mnyCount > 0 Then
                    tmpTxt = tmpTxt & "<p>" & oVrnt & "</p>"
                End If
            End If
        Next
        If tmpTxt = "" Then tmpTxt = "Not applicable"
        GetUsedStylesInTheDocument = tmpTxt
    End Function
    Function GetStyleCountMis(wStyle As String, ByRef sCount As Integer, ByRef htString As String, wMode As String, Optional aBook As Boolean = True, Optional subHEAD As Boolean = False, Optional wjStory As Boolean = False)
        Dim oFndRng As Word.Range
        Dim hRange As Word.Range
        Dim rRange As Word.Range

        GetStyleCountMis = 0 : sCount = 0
        If InStr(1, Microsoft.VisualBasic.Strings.Trim(wStyle), "char ", vbTextCompare) = 1 Or
      InStr(1, Microsoft.VisualBasic.Strings.Trim(wStyle), ",") > 0 Then Exit Function

        If AutoStyleExists(wStyle, wAPP.ActiveDocument) = False Then Exit Function
        rRange = Nothing
        For Each hRange In wAPP.ActiveDocument.StoryRanges
            If wjStory = True And hRange.StoryType <> Word.WdStoryType.wdMainTextStory Then GoTo ExtLoop : 
            oFndRng = hRange.Duplicate
            oFndRng.SetRange(oFndRng.Start, oFndRng.Start)
            oFndRng.Select()
            wAPP.Selection.Find.ClearFormatting()
            wAPP.Selection.Find.Replacement.ClearFormatting()
            wAPP.Selection.Find.Text = "" : wAPP.Selection.Find.Style = wStyle
            Do While wAPP.Selection.Find.Execute = True
                If Not rRange Is Nothing Then
                    If wAPP.Selection.InRange(rRange) = True Then Exit Do
                End If
                If IsNothing(wAPP.Selection.Style) Then
                    If wAPP.Selection.Style Is Nothing = False Then
                        If (wAPP.Selection.Style = wStyle Or wAPP.Selection.Paragraphs(1).Style = wStyle) Then
                            If wMode = "C" Then
                                sCount = sCount + 1
                            End If
                        Else
                            Exit Do
                        End If
                    End If
                End If
                If sCount > 0 Then Exit Do
                rRange = wAPP.Selection.Range.Duplicate
                wAPP.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Loop
ExtLoop:
            If sCount > 0 Then Exit For
        Next
        GetStyleCountMis = sCount
    End Function
    Function RemoveFormattinginParas()
        Dim oTbl As Word.Table
        Dim oCll As Word.Cell
        Dim oCllRng As Word.Range
        For Each oTbl In wAPP.ActiveDocument.Tables
            For Each oCll In oTbl.Range.Cells
                oCllRng = oCll.Range.Duplicate
                oCllRng.SetRange(oCllRng.End - 1, oCllRng.End - 1)
                oCllRng.Select()
                wAPP.Selection.Font.Reset()
            Next
        Next
        oCllRng = wAPP.ActiveDocument.StoryRanges(Word.WdStoryType.wdMainTextStory).Duplicate
        oCllRng.SetRange(oCllRng.Start, oCllRng.Start)
        oCllRng.Select
    End Function
    Public Function AutoStyleExists(ByVal sStyleName As String, ByVal StyleDoc As Word.Document) As Boolean
        Dim xDsc As String
        Try
            xDsc = StyleDoc.Styles.Item(sStyleName).Description
            AutoStyleExists = True
        Catch ex As Exception
            ex.Data.Clear()
            AutoStyleExists = False
        End Try
    End Function

    ''' <summary>
    ''' '
    ''' </summary>
    ''' <param name="urlString"></param>
    ''' <param name="valueString"></param>
    ''' <returns></returns>
    Function CEGWebQuery(urlString As String, valueString As String)
        Try
            Dim wbClient As New System.Net.WebClient()
            wbClient.Encoding = System.Text.Encoding.UTF8
            Dim response As String = wbClient.DownloadString(urlString + valueString)
            Return response
        Catch ex As Exception
            Return "ERROR: " + ex.Message
        End Try
        Return ""
    End Function
    Public Function UnAccent(ByVal aString As String) As String
        Dim toReplace() As Char = "àèìòùÀÈÌÒÙ äëïöüÄËÏÖÜ âêîôûÂÊÎÔÛ áéíóúÁÉÍÓÚðÐýÝ ãñõÃÑÕšŠžŽçÇåÅøØ".ToCharArray
        Dim replaceChars() As Char = "aeiouAEIOU aeiouAEIOU aeiouAEIOU aeiouAEIOUdDyY anoANOsSzZcCaAoO".ToCharArray
        For index As Integer = 0 To toReplace.GetUpperBound(0)
            aString = aString.Replace(toReplace(index), replaceChars(index))
        Next
        Return aString
    End Function


    Public Function CheckFMAuthorWithDiscloserAuthor(wDoc As Word.Document) As Boolean
        Dim version As String = wDoc.Application.Version.ToString()
        Try
            Dim BRTrack As Boolean = False
            If wDoc.Revisions.Count >= 1 Then
                BRTrack = True



                If Version = "12.0" Then
                    With wDoc.ActiveWindow.View
                        .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = False
                    End With
                Else
                    Try
                        With wDoc.ActiveWindow.View
                            .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone
                            .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                        End With
                    Catch ex As Exception
                    End Try
                End If
            End If


            Dim docComment As Word.Comment
            Dim FMAuthorList As Dictionary(Of Integer, List(Of clsAuthorInfo))
            Dim DiscloserAuthorList As Dictionary(Of Integer, List(Of clsAuthorInfo))
            FMAuthorList = CollectAuthorInformation(wDoc, "†FM_Authors")


            DiscloserAuthorList = CollectDisclosureAuthorInformation(wDoc, "†EM_Acknowledgments_Text")

            If IsNothing(DiscloserAuthorList) Then ''DISCLOSURES
                If FMAuthorList.Count > 0 Then
                    Dim lstdis As List(Of clsAuthorInfo) = FMAuthorList(1)
                    If lstdis.Count > 0 Then docComment = wDoc.Comments.Add(lstdis(1).tagRange, "Author count in byline:" & FMAuthorList.Count & "; " & "Disclosures Author not found in the document. Please correct this as needed.")
                End If
                Return True
            End If

            If FMAuthorList.Count <> DiscloserAuthorList.Count Then
                ''MessageBox.Show("FM author count not matched with discloser author count")
                If DiscloserAuthorList.Count > 0 Then
                    Dim lstdis As List(Of clsAuthorInfo) = DiscloserAuthorList(1)
                    '  If lstdis.Count > 0 Then docComment = wDoc.Comments.Add(lstdis(1).Paragraphs(1).Range, "Author count in byline : " & FMAuthorList.Count & "; " & "Author count in disclosure : " & DiscloserAuthorList.Count & "; " & "Please correct this as needed. ")
                    If lstdis.Count > 0 Then docComment = wDoc.Comments.Add(lstdis(1).tagRange, "Author count in byline: " & FMAuthorList.Count & "; " & "Author count in disclosure: " & DiscloserAuthorList.Count & "; " & "Please correct this as needed.")
                End If
            Else
                Dim L As Integer
                For L = 1 To FMAuthorList.Count
                    Dim K As Integer
                    Dim lstAuthor As List(Of clsAuthorInfo) = FMAuthorList(L)
                    Dim lstDisclose As List(Of clsAuthorInfo) = DiscloserAuthorList(L)
                    Dim stringAuthor As String = String.Join(" ", lstAuthor.Select(Function(x) x.tagRange.Text).ToArray())

                    Dim tagsPresentAut As List(Of String) = lstAuthor.Select(Function(x) x.tagName).ToArray().ToList()
                    Dim tagsPresentDis As List(Of String) = lstDisclose.Select(Function(x) x.tagName).ToArray().ToList()


                    '**************Check for the Missing Elements***************
                    If tagsPresentAut.Count > tagsPresentDis.Count Then
                        Dim tagsMismatch As List(Of String) = tagsPresentAut.Except(tagsPresentDis).ToList()
                        For loopcounter = 0 To tagsMismatch.Count - 1
                            Select Case tagsMismatch(loopcounter)

                                Case "‡fm_auSuffix", "‡fm_corrSuffix"

                                    docComment = wDoc.Comments.Add(lstAuthor(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the suffix between byline and disclosure. Set as per byline.")

                                Case "‡fm_auPrefix", "‡fm_corrPrefix"
                                    docComment = wDoc.Comments.Add(lstAuthor(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the prefix between byline and disclosure. Set as per byline.")

                                Case "‡fm_auDegree", "‡fm_corrDegree"
                                    docComment = wDoc.Comments.Add(lstAuthor(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the degree between byline and disclosure. Set as per byline.")

                            End Select

                        Next
                    ElseIf tagsPresentDis.Count > tagsPresentAut.Count Then
                        Dim tagsMismatch As List(Of String) = tagsPresentDis.Except(tagsPresentAut).ToList()
                        For loopcounter = 0 To tagsMismatch.Count - 1
                            Select Case tagsMismatch(loopcounter)
                                Case "‡fm_auSuffix", "‡fm_corrSuffix"
                                    docComment = wDoc.Comments.Add(lstDisclose(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the suffix between byline and disclosure. Set as per byline.")

                                Case "‡fm_auPrefix", "‡fm_corrPrefix"
                                    docComment = wDoc.Comments.Add(lstDisclose(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the prefix between byline and disclosure. Set as per byline.")

                                Case "‡fm_auDegree", "‡fm_corrDegree"
                                    docComment = wDoc.Comments.Add(lstDisclose(loopcounter + 1).tagRange, "CE: For the author """ & stringAuthor & """, check the degree between byline and disclosure. Set as per byline.")

                            End Select

                        Next
                    End If
                    '**************Check for the Missing Elements***************



                    If lstAuthor.Count = lstDisclose.Count Then
                        For K = 0 To lstAuthor.Count - 1
                            If Regex.Replace(lstAuthor(K).tagRange.Text, "([\.\,])", "") <> Regex.Replace(lstDisclose(K).tagRange.Text, "([\.\,])", "") Then
                                '************Added for Accented Characters************


                                If lstAuthor(K).tagRange.Text = UnAccent(lstDisclose(K).tagRange.Text) Or UnAccent(lstAuthor(K).tagRange.Text) = lstDisclose(K).tagRange.Text Then
                                    'Revised warning: For the author “Kárl A. Poterack,” check the accented character between byline and disclosure. Set as per byline.
                                    docComment = wDoc.Comments.Add(lstDisclose(K).tagRange, "CE: For the author """ & stringAuthor & """, check the accented character between byline and disclosure. Set as per byline.")
                                Else
                                    docComment = wDoc.Comments.Add(lstDisclose(K).tagRange, "CE: For the author " & stringAuthor & ", byline and disclosure is not matching. Set as per byline.")
                                End If
                                '************Added for Accented Characters************


                            End If

                        Next
                    Else
                        'If DiscloserAuthorList.Count < FMAuthorList.Count Then
                        '    Dim myList As List(Of Integer) = DiscloserAuthorList.Keys.ToList()

                        '    ''docComment = wDoc.Comments.Add(lstDisclose(1), "CE: For the author " & stringAuthor & ", byline and disclosure is not matching. Set as per byline.")
                        'End If
                    End If
                Next
            End If
            If BRTrack = True Then


                If Version = "12.0" Then
                    With wDoc.ActiveWindow.View
                        .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = True
                    End With
                Else
                    Try
                        With wDoc.ActiveWindow.View
                            .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                            .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                        End With
                    Catch ex As Exception
                    End Try
                End If

                BRTrack = False
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function



#Region "PublicFunctions"
    Public Function ReadINI(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal DefaultValue As String, Optional ByVal IsRegPattern As Boolean = False) As String
        Dim n As Int32
        Dim sData As New String(" ", 65536)
        If File.Exists(INIPath) = False Then Throw New Exception("File not found : " & INIPath)
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, sData, sData.Length, INIPath)
        If n > 0 Then
            Select Case IsRegPattern
                Case True
                    ReadINI = sData.Substring(0, n)
                    ReadINI = "(" & ReadINI.Replace(" ", "(\s*)") & ")"
                Case False
                    ReadINI = sData.Substring(0, n)
            End Select
        Else
            ReadINI = String.Empty
        End If
    End Function
#End Region

    '************************Adding Publisher Note for IOP Journals -- Mantis ID: 40684****************
    Public Function InsertPublisherNoteasComment(pubName As String, JName As String, wDoc As Word.Document, JConfigPath As String) As Boolean
        Try

            Dim sInsertPublisherNote As String = vbNullString
            Dim sMathWordsCheck As String = vbNullString

            sInsertPublisherNote = ReadINI(JConfigPath, pubName.ToUpper & "@" & JName.ToUpper, "sInsertPublisherNote", String.Empty, False)
            sMathWordsCheck = ReadINI(JConfigPath, pubName.ToUpper & "@" & JName.ToUpper, "MatchWordsforInsertPublisherNote", String.Empty, False)
            Dim sMatchWordsSplit() As String = sMathWordsCheck.Split("|")


            If sInsertPublisherNote = vbNullString Then
                MsgBox("Required Parameters Missing in JournalConfig. Kindly check.", MsgBoxStyle.Information, sMsgTitle)
                InsertPublisherNoteasComment = False
                Exit Function
            Else
                wDoc.Application.Selection.HomeKey(Unit:=Word.WdUnits.wdStory)
                wDoc.Application.Selection.Find.ClearFormatting()
                Try
                    wDoc.Application.Selection.Find.Style = wDoc.Styles("†FM_Affiliations")
                Catch ex As Exception
                    MsgBox("†FM_Affiliations Style is not present in the document. Kindly check" + ex.Message, MsgBoxStyle.Information, sMsgTitle)
                    InsertPublisherNoteasComment = False
                    Exit Function
                End Try

                wDoc.Application.Selection.Find.Replacement.ClearFormatting()
                With wDoc.Application.Selection.Find
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = True
                    .Wrap = Word.WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                Do While wDoc.Application.Selection.Find.Execute() = True
                    Dim mystring As String = wDoc.Application.Selection.Range.Text
#Region "LinQ"
                    ''  Dim result As String = sMatchWordsSplit.Where(P >= mystring.Contains(P))
                    ''  Dim elem As VariantType = From x In sMatchWordsSplit.AsEnumerable().Where(wDoc.Application.Selection.Range.Text.Contains(x.ToLower)) Select x
                    ''  Dim results As VariantType = query.Select(() => {  }, x.ToLower).Where(() => {  }, searchstrings.All(() => {  }, x.Contains(y)))
#End Region
                    Dim i As Integer
                    For i = LBound(sMatchWordsSplit) To UBound(sMatchWordsSplit) - 1
                        If (wDoc.Application.Selection.Range.Text.Trim().ToLower().Contains(sMatchWordsSplit(i).ToString().ToLower())) Then
                            Exit Do
                        End If
                    Next
                Loop
                wDoc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                wDoc.Application.Selection.Find.ClearFormatting()
                wDoc.Application.Selection.Find.Style = wDoc.Styles("†FM_Affiliations")
                wDoc.Application.Selection.Find.Replacement.ClearFormatting()
                With wDoc.Application.Selection.Find
                    .Text = ""
                    .Replacement.Text = ""
                    .Forward = False
                    .Wrap = Word.WdFindWrap.wdFindContinue
                    .Format = True
                    .MatchCase = False
                    .MatchWholeWord = False
                    .MatchWildcards = False
                    .MatchSoundsLike = False
                    .MatchAllWordForms = False
                End With
                wDoc.Application.Selection.Find.Execute()
                wDoc.Application.Selection.MoveRight()
                wDoc.Application.Selection.TypeText(sInsertPublisherNote.Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)(0))
                wDoc.Application.Selection.TypeParagraph()
                wDoc.Application.Selection.MoveUp(Unit:=Word.WdUnits.wdParagraph, Count:=1, Extend:=True)
                Try
                    wDoc.Application.Selection.Range.Style = sInsertPublisherNote.ToLower().Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)(1).ToString
                Catch ex As Exception
                    MsgBox("Unable to apply the style " & sInsertPublisherNote.ToLower().Split(New String() {"|"}, StringSplitOptions.RemoveEmptyEntries)(1).ToString & "" + ex.Message, MsgBoxStyle.Information, sMsgTitle)
                End Try
            End If
            InsertPublisherNoteasComment = True
        Catch ex As Exception
            MsgBox("Unable to insert Publisher Note Comment :" + ex.Message, MsgBoxStyle.Information, sMsgTitle)
        End Try
    End Function
    '************************Adding Publisher Note for IOP Journals -- Mantis ID: 40684****************


    Private Function CollectAuthorInformation(wDoc As Word.Document, sStyleName As String) As Dictionary(Of Integer, List(Of clsAuthorInfo))
        Try
            wDoc.Application.Selection.HomeKey(Word.WdUnits.wdStory)
            Dim dictFMAuthor As New Dictionary(Of Integer, List(Of clsAuthorInfo))
            Dim dictAuthor As New Dictionary(Of Integer, clsAuthorInfo)

            Dim boolDegreePresent As Boolean = False


            Dim auCount As Integer
            auCount = 0
            Dim fStyle() = {"‡fm_corrGivenName", "‡fm_corrSurname", "‡fm_corrPrefix", "‡fm_corrSuffix", "‡fm_corrDegree", "‡fm_auGivenName", "‡fm_auSurname", "‡fm_auPrefix", "‡fm_auSuffix", "‡fm_auDegree"}
            Dim ranDoc As Word.Range : Dim rCount As Integer
            ranDoc = wDoc.Content
            wAPP = wDoc.Application
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = sStyleName
            End With
            Do While ranDoc.Find.Execute = True
                Dim characterStylefound As Boolean = False
                rCount = rCount + 1
                Dim AuthorList As New List(Of clsAuthorInfo)
                ''Dim dictAuthor As New Dictionary(Of Integer, String)
                Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
                ranDoc.Select()
                ranDupRef = wAPP.Selection.Range
                Dim objAuthorInfor As clsAuthorInfo
                Dim lastStyleName As String
                For i = LBound(fStyle) To UBound(fStyle)
                    ranAuthor = ranDupRef.Duplicate
                    'With wDoc.Application.Selection.Find
                    '    .ClearFormatting() : .Replacement.ClearFormatting()
                    '    .Text = "" : .Style = fStyle(i)
                    'End With
                    With ranAuthor.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With

                    Do While ranAuthor.Find.Execute
                        ranAuthor.Select()
                        characterStylefound = True
                        If wAPP.Selection.Range.End > ranDupRef.End Then Exit Do
                        If wAPP.Selection.Range.Start = ranDupRef.End Then Exit Do
                        '***************Added on 10-Apr-2018, Based on Saritha's Feedback****************
                        If (wAPP.Selection.Range.Text Like vbCrLf Or wAPP.Selection.Range.Text Like vbCr Or wAPP.Selection.Range.Text Like vbLf) Then Exit Do
                        '***************Added on 10-Apr-2018, Based on Saritha's Feedback****************
                        If wAPP.Selection.Text = Nothing Then Exit Do
                        objAuthorInfor = New clsAuthorInfo
                        objAuthorInfor.tagName = fStyle(i)

                        '***************Added on 10-Apr-2018, Based on Saritha's Feedback****************
                        If wAPP.Selection.Style.NameLocal = wAPP.Selection.Range.Words.Last.Next.Words(1).Style.NameLocal Then
                            wAPP.Selection.Range.Words(1).Select()
                            wAPP.Selection.MoveRight(Word.WdUnits.wdWord, 1, True)
                        End If
                        Do While wAPP.Selection.Style.NameLocal = wAPP.Selection.Range.Words.Last.Next.Words(1).Style.NameLocal
                            wAPP.Selection.MoveRight(Word.WdUnits.wdWord, 1, True)
                        Loop
                        '***************Added on 10-Apr-2018, Based on Saritha's Feedback****************

                        objAuthorInfor.tagValue = wAPP.Selection.Text
                            objAuthorInfor.tagRange = wAPP.Selection.Range
                            If IsNothing(dictAuthor) Then
                                dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                            Else
                                If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
                                    dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                                End If
                            End If

                            If wAPP.Selection.Range.End >= ranDupRef.End Then Exit Do
                            ranAuthor = wDoc.Range(wAPP.Selection.Range.End + 1, ranDupRef.End)
                            ranAuthor.Find.Text = ""
                            ranAuthor.Find.Style = fStyle(i)
                            lastStyleName = fStyle(i)
                            ranAuthor.Select()
                        Loop

                Next
                If characterStylefound = False Then
                    rCount = rCount - 1
                    Continue Do
                End If
                If Not IsNothing(dictAuthor) Then

                    Dim MainLC As Integer = 0
                    For Each objAuthorInfor In dictAuthor.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
                        MainLC += 1
                        If (objAuthorInfor.tagName = "‡fm_corrDegree" Or objAuthorInfor.tagName = "‡fm_auDegree") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count > 0 Then
                                AuthorList.Add(objAuthorInfor)
                                auCount = auCount + 1
                                boolDegreePresent = True
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of clsAuthorInfo)
                                boolDegreePresent = False
                            Else
                                AuthorList.Add(objAuthorInfor)
                            End If

                        ElseIf (objAuthorInfor.tagName = "‡fm_corrPrefix" Or objAuthorInfor.tagName = "‡fm_auPrefix") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count = 0 Then
                                AuthorList.Add(objAuthorInfor)
                            ElseIf AuthorList.Count > 0 Then
                                auCount = auCount + 1
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of clsAuthorInfo)
                                AuthorList.Add(objAuthorInfor)
                            Else
                                MessageBox.Show("wrong code")
                            End If
                        ElseIf (objAuthorInfor.tagName = "‡fm_corrGivenName" Or objAuthorInfor.tagName = "‡fm_auGivenName") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            boolDegreePresent = True

                            If AuthorList.Count = 0 Then

                                AuthorList.Add(objAuthorInfor)
                            ElseIf (AuthorList.Count > 0) Then
                                If boolDegreePresent = False Then
                                    auCount = auCount + 1
                                    dictFMAuthor.Add(auCount, AuthorList)
                                    AuthorList = New List(Of clsAuthorInfo)
                                    AuthorList.Add(objAuthorInfor)

                                Else
                                    auCount = auCount + 1
                                    dictFMAuthor.Add(auCount, AuthorList)
                                    AuthorList = New List(Of clsAuthorInfo)
                                    AuthorList.Add(objAuthorInfor)

                                End If
                            End If
                            'ElseIf (objAuthorInfor.tagName = "‡fm_corrSurname" Or objAuthorInfor.tagName = "‡fm_auSurname") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            '    If AuthorList.Count = 0 Then
                            '        AuthorList.Add(objAuthorInfor)
                            '    ElseIf (AuthorList.Count > 0) Then
                            '        auCount = auCount + 1
                            '        AuthorList.Add(objAuthorInfor)
                            '        dictFMAuthor.Add(auCount, AuthorList)
                            '        AuthorList = New List(Of clsAuthorInfo)
                            '    End If
                        ElseIf Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            AuthorList.Add(objAuthorInfor)
                        End If
                        If dictAuthor.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values.Count = MainLC Then
                            If boolDegreePresent = True Then
                                auCount = auCount + 1
                                dictFMAuthor.Add(auCount, AuthorList)
                            End If
                        End If
                    Next
                End If
                If ranDoc.End >= wDoc.Range.End Then Exit Do
                ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
                ranDoc.Find.Text = ""
                ranDoc.Find.Style = sStyleName
                dictAuthor.Clear()
            Loop
            If dictFMAuthor.Count > 0 Then
                CollectAuthorInformation = dictFMAuthor : Exit Function
            Else
                CollectAuthorInformation = Nothing : Exit Function
            End If
            Exit Function
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Private Function CollectAuthorInformation_Old(wDoc As Word.Document, sStyleName As String) As Dictionary(Of Integer, List(Of Word.Range))
        Try
            wDoc.Application.Selection.HomeKey(Word.WdUnits.wdStory)
            Dim dictFMAuthor As New Dictionary(Of Integer, List(Of Word.Range))
            Dim dictAuthor As New Dictionary(Of Integer, clsAuthorInfo)

            Dim auCount As Integer
            auCount = 0
            Dim fStyle() = {"‡fm_corrGivenName", "‡fm_corrSurname", "‡fm_corrPrefix", "‡fm_corrSuffix", "‡fm_corrDegree", "‡fm_auGivenName", "‡fm_auSurname", "‡fm_auPrefix", "‡fm_auSuffix", "‡fm_auDegree"}
            Dim ranDoc As Word.Range : Dim rCount As Integer
            ranDoc = wDoc.Content
            wAPP = wDoc.Application
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = sStyleName
            End With
            Do While ranDoc.Find.Execute = True
                Dim characterStylefound As Boolean = False
                rCount = rCount + 1
                Dim AuthorList As New List(Of Word.Range)
                ''Dim dictAuthor As New Dictionary(Of Integer, String)
                Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
                ranDoc.Select()
                ranDupRef = wAPP.Selection.Range
                Dim objAuthorInfor As clsAuthorInfo
                Dim lastStyleName As String
                For i = LBound(fStyle) To UBound(fStyle)
                    ranAuthor = ranDupRef.Duplicate
                    With wDoc.Application.Selection.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With
                    With ranAuthor.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With

                    Do While ranAuthor.Find.Execute
                        ranAuthor.Select()
                        characterStylefound = True
                        If wAPP.Selection.Range.End > ranDupRef.End Then Exit Do
                        If wAPP.Selection.Range.Start = ranDupRef.End Then Exit Do
                        If wAPP.Selection.Text = Nothing Then Exit Do
                        objAuthorInfor = New clsAuthorInfo
                        objAuthorInfor.tagName = fStyle(i)
                        objAuthorInfor.tagValue = wAPP.Selection.Text
                        objAuthorInfor.tagRange = wAPP.Selection.Range
                        If IsNothing(dictAuthor) Then
                            dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                        Else
                            If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
                                dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                            End If
                        End If

                        If wAPP.Selection.Range.End >= ranDupRef.End Then Exit Do
                        ranAuthor = wDoc.Range(wAPP.Selection.Range.End + 1, ranDupRef.End)
                        ranAuthor.Find.Text = ""
                        ranAuthor.Find.Style = fStyle(i)
                        lastStyleName = fStyle(i)
                        ranAuthor.Select()
                    Loop

                Next
                If characterStylefound = False Then
                    rCount = rCount - 1
                    Continue Do
                End If
                If Not IsNothing(dictAuthor) Then
                    For Each objAuthorInfor In dictAuthor.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
                        If objAuthorInfor.tagName = "‡fm_corrDegree" Or objAuthorInfor.tagName = "‡fm_auDegree" Then
                            If AuthorList.Count > 0 Then
                                AuthorList.Add(objAuthorInfor.tagRange)
                                auCount = auCount + 1

                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of Word.Range)
                            Else
                                AuthorList.Add(objAuthorInfor.tagRange)
                            End If

                        ElseIf objAuthorInfor.tagName = "‡fm_corrPrefix" Or objAuthorInfor.tagName = "‡fm_auPrefix" Then
                            If AuthorList.Count = 0 Then
                                AuthorList.Add(objAuthorInfor.tagRange)
                            ElseIf AuthorList.Count > 0 Then
                                auCount = auCount + 1
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of Word.Range)
                                AuthorList.Add(objAuthorInfor.tagRange)
                            Else
                                MessageBox.Show("wrong code")
                            End If
                        ElseIf objAuthorInfor.tagName = "‡fm_corrGivenName" Or objAuthorInfor.tagName = "‡fm_auGivenName" Then
                            If AuthorList.Count > 1 Then
                                auCount = auCount + 1
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of Word.Range)
                            End If
                            AuthorList.Add(objAuthorInfor.tagRange)
                        Else
                            AuthorList.Add(objAuthorInfor.tagRange)
                        End If
                    Next
                End If
                If ranDoc.End >= wDoc.Range.End Then Exit Do
                ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
                ranDoc.Find.Text = ""
                ranDoc.Find.Style = sStyleName
                dictAuthor.Clear()
            Loop
            If dictFMAuthor.Count > 0 Then
                CollectAuthorInformation_Old = dictFMAuthor : Exit Function
            Else
                CollectAuthorInformation_Old = Nothing : Exit Function
            End If
            Exit Function
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    '*********Developed by Thiyagu for Disclosure Author Checking

    '  Private Function CollectDisclosureAuthorInformation(wDoc As Word.Document, sStyleName As String) As Dictionary(Of Integer, List(Of Word.Range))
    Private Function CollectDisclosureAuthorInformation(wDoc As Word.Document, sStyleName As String) As Dictionary(Of Integer, List(Of clsAuthorInfo))
        Try
            wDoc.Application.Selection.HomeKey(Word.WdUnits.wdStory)




            Dim dictFMAuthor As New Dictionary(Of Integer, List(Of clsAuthorInfo))
            Dim dictAuthor As New Dictionary(Of Integer, clsAuthorInfo)

            Dim auCount As Integer
            auCount = 0
            Dim fStyle() = {"‡fm_corrGivenName", "‡fm_corrSurname", "‡fm_corrPrefix", "‡fm_corrSuffix", "‡fm_corrDegree", "‡fm_auGivenName", "‡fm_auSurname", "‡fm_auPrefix", "‡fm_auSuffix", "‡fm_auDegree"}
            Dim ranDoc As Word.Range : Dim rCount As Integer
            ranDoc = wDoc.Content
            wAPP = wDoc.Application
            With wDoc.Application.Selection.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = sStyleName
            End With
            'Selection.MoveDown Unit:=wdParagraph, count:=2, Extend:=wdExtend
            If sStyleName = "†EM_Acknowledgments_Text" Then
                If wDoc.Application.Selection.Find.Execute = True Then
                    Try
                        If wDoc.Application.Selection.Bookmarks.Exists("\EndofDoc") = False Then
                            Try
Nextpara:
                                Do While (wDoc.Application.Selection.Paragraphs.Last.Next.Range.Style.NameLocal = "†EM_Acknowledgments_Text")
                                    wDoc.Application.Selection.MoveDown(Unit:=Word.WdUnits.wdParagraph, Count:=1, Extend:=True)
                                    If wDoc.Application.Selection.Bookmarks.Exists("\EndofDoc") = True Then
                                        Exit Do
                                    End If

                                Loop
                            Catch ex As Exception
                                If (wDoc.Application.Selection.Paragraphs.Last.Next.Range.Revisions(1).Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionDelete Or wDoc.Application.Selection.Paragraphs.Last.Next.Range.Revisions(1).Type = Microsoft.Office.Interop.Word.WdRevisionType.wdRevisionInsert) Then
                                    wDoc.Application.Selection.MoveDown(Unit:=Word.WdUnits.wdParagraph, Count:=1, Extend:=True)
                                    GoTo Nextpara
                                End If
                            End Try

                        End If
                    Catch ex As Exception

                    End Try
                End If
            End If




            ranDoc = wDoc.Application.Selection.Range

            Dim version As String = wDoc.Application.Version.ToString()

            If version = "12.0" Then
                With wDoc.ActiveWindow.View
                    .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                    .ShowRevisionsAndComments = True
                End With
            Else
                Try
                    With wDoc.ActiveWindow.View
                        .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                        .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                    End With
                Catch ex As Exception
                End Try
            End If

            For loopcounter = 1 To ranDoc.Paragraphs.Count
                Dim characterStylefound As Boolean = False
                rCount = rCount + 1
                Dim AuthorList As New List(Of clsAuthorInfo)
                ''Dim dictAuthor As New Dictionary(Of Integer, String)
                Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
                ranDoc.Paragraphs(loopcounter).Range.Select()

                If ranDoc.Paragraphs(loopcounter).Range.Style.namelocal <> "†EM_Acknowledgments_Text" Then
                    Continue For
                Else
                    If version = "12.0" Then
                        With wDoc.ActiveWindow.View
                            .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                            .ShowRevisionsAndComments = False
                        End With
                    Else
                        Try
                            With wDoc.ActiveWindow.View
                                .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupNone
                                .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                            End With
                        Catch ex As Exception
                        End Try
                    End If

                End If


                ranDupRef = wAPP.Selection.Range
                Dim startval As Integer = ranDupRef.Start
                Dim endval As Integer = ranDupRef.End
                Dim objAuthorInfor As clsAuthorInfo
                Dim lastStyleName As String
                For i = LBound(fStyle) To UBound(fStyle)
                    ranAuthor = ranDupRef.Duplicate
                    ranAuthor.Start = startval
                    ranAuthor.End = endval
                    ranAuthor.Select()
                    'ranDoc.Paragraphs(loopcounter).Range.Select()

                    With wDoc.Application.Selection.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With
                    'With ranAuthor.Find
                    '    .ClearFormatting() : .Replacement.ClearFormatting()
                    '    .Text = "" : .Style = fStyle(i)
                    'End With

                    Do While wDoc.Application.Selection.Find.Execute
                        ''    ranAuthor.Select()
                        characterStylefound = True
                        If wAPP.Selection.Range.End > ranDupRef.End Then Exit Do
                        If wAPP.Selection.Range.Start = ranDupRef.End Then Exit Do
                        If wAPP.Selection.Text = Nothing Then Exit Do
                        objAuthorInfor = New clsAuthorInfo
                        objAuthorInfor.tagName = fStyle(i)
                        objAuthorInfor.tagValue = wAPP.Selection.Text
                        objAuthorInfor.tagRange = wAPP.Selection.Range
                        If IsNothing(dictAuthor) Then
                            dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                        Else
                            If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
                                dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
                            End If
                        End If

                        If wAPP.Selection.Range.End >= ranDupRef.End Then Exit Do
                        ranAuthor = wDoc.Range(wAPP.Selection.Range.End + 1, ranDupRef.End)
                        ranAuthor.Find.Text = ""
                        ranAuthor.Find.Style = fStyle(i)
                        lastStyleName = fStyle(i)
                        ranAuthor.Select()
                    Loop
                    ranAuthor.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Next
                If characterStylefound = False Then
                    rCount = rCount - 1

                    If version = "12.0" Then
                        With wDoc.ActiveWindow.View
                            .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                            .ShowRevisionsAndComments = True
                        End With
                    Else
                        Try
                            With wDoc.ActiveWindow.View
                                .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                                .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                            End With
                        Catch ex As Exception
                        End Try
                    End If

                    ranDupRef.Select()
                    Continue For
                End If
                If Not IsNothing(dictAuthor) Then
                    auCount = auCount + 1
                    '   Dim bolAuthorElemPresent As Boolean = False
                    For Each objAuthorInfor In dictAuthor.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
                        If (objAuthorInfor.tagName = "‡fm_corrDegree" Or objAuthorInfor.tagName = "‡fm_auDegree") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count > 0 Then
                                AuthorList.Add(objAuthorInfor)
                                'If bolAuthorElemPresent = False Then
                                '    auCount = auCount + 1
                                '    bolAuthorElemPresent = True
                                'End If
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of clsAuthorInfo)
                            Else
                                '' AuthorList.Add(objAuthorInfor.tagRange)
                                AuthorList.Add(objAuthorInfor)
                            End If

                        ElseIf (objAuthorInfor.tagName = "‡fm_corrPrefix" Or objAuthorInfor.tagName = "‡fm_auPrefix") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count = 0 Then
                                AuthorList.Add(objAuthorInfor)
                            ElseIf AuthorList.Count > 0 Then

                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of clsAuthorInfo)
                                AuthorList.Add(objAuthorInfor)
                            Else
                                MessageBox.Show("wrong code")
                            End If
                        ElseIf (objAuthorInfor.tagName = "‡fm_corrGivenName" Or objAuthorInfor.tagName = "‡fm_auGivenName") And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count = 0 Then
                                AuthorList.Add(objAuthorInfor)
                            ElseIf (AuthorList.Count > 0) Then


                                AuthorList = New List(Of clsAuthorInfo)
                                AuthorList.Add(objAuthorInfor)
                                dictFMAuthor.Add(auCount, AuthorList)
                            End If
                            ' AuthorList.Add(objAuthorInfor)
                        ElseIf (objAuthorInfor.tagName = "‡fm_corrSurName" Or objAuthorInfor.tagName = "‡fm_auSurname") And (dictAuthor.Count = 2) And Not Microsoft.VisualBasic.Strings.Trim(Regex.Replace(objAuthorInfor.tagRange.Text, "([\,\.])", "")) = "" Then
                            If AuthorList.Count = 0 Then
                                AuthorList.Add(objAuthorInfor)
                            ElseIf (AuthorList.Count > 0) Then

                                ''dictFMAuthor.Add(auCount, AuthorList)

                                AuthorList.Add(objAuthorInfor)
                                dictFMAuthor.Add(auCount, AuthorList)
                                AuthorList = New List(Of clsAuthorInfo)
                            End If
                            '    AuthorList.Add(objAuthorInfor)

                        ElseIf Not Microsoft.VisualBasic.Strings.Trim(objAuthorInfor.tagRange.Text) = "" Then
                            AuthorList.Add(objAuthorInfor)
                        End If
                    Next
                End If
                'If ranDoc.End >= wDoc.Range.End Then Exit For
                'ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
                ranDoc.Find.Text = ""
                ranDoc.Find.Style = sStyleName
                dictAuthor.Clear()


                If version = "12.0" Then
                    With wDoc.ActiveWindow.View
                        .RevisionsView = Word.WdRevisionsView.wdRevisionsViewFinal
                        .ShowRevisionsAndComments = True
                    End With
                Else
                    Try
                        With wDoc.ActiveWindow.View
                            .RevisionsFilter.Markup = Word.WdRevisionsMarkup.wdRevisionsMarkupAll
                            .RevisionsFilter.View = Word.WdRevisionsView.wdRevisionsViewFinal
                        End With
                    Catch ex As Exception
                    End Try
                End If


                ranDupRef.Select()
            Next


            'Dim characterStylefound As Boolean = False

            'Dim AuthorList As New List(Of Word.Range)
            'Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
            'ranDupRef = wAPP.Selection.Range
            'Dim objAuthorInfor As clsAuthorInfo
            'Dim lastStyleName As String
            'For i = LBound(fStyle) To UBound(fStyle)
            '    ranAuthor = ranDupRef.Duplicate
            '    With wDoc.Application.Selection.Find
            '        .ClearFormatting() : .Replacement.ClearFormatting()
            '        .Text = "" : .Style = fStyle(i)
            '    End With

            '    Do While wDoc.Application.Selection.Find.Execute
            '        '  ranAuthor.Select()
            '        characterStylefound = True
            '        If wAPP.Selection.Range.End > ranDupRef.End Then Exit Do
            '        If wAPP.Selection.Range.Start = ranDupRef.End Then Exit Do
            '        If wAPP.Selection.Text = Nothing Then Exit Do
            '        objAuthorInfor = New clsAuthorInfo
            '        objAuthorInfor.tagName = fStyle(i)
            '        objAuthorInfor.tagValue = wAPP.Selection.Text
            '        objAuthorInfor.tagRange = wAPP.Selection.Range
            '        If IsNothing(dictAuthor) Then
            '            dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
            '        Else
            '            If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
            '                dictAuthor.Add(wAPP.Selection.Start, objAuthorInfor)
            '            End If
            '        End If

            '        If wAPP.Selection.Range.End >= ranDupRef.End Then Exit Do
            '        ranAuthor = wDoc.Range(wAPP.Selection.Range.End + 1, ranDupRef.End)
            '        ranAuthor.Find.Text = ""
            '        ranAuthor.Find.Style = fStyle(i)
            '        lastStyleName = fStyle(i)
            '        ranAuthor.Select()
            '    Loop

            'Next

            'If Not IsNothing(dictAuthor) Then
            '    For Each objAuthorInfor In dictAuthor.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
            '        If objAuthorInfor.tagName = "‡fm_corrDegree" Or objAuthorInfor.tagName = "‡fm_auDegree" Then
            '            If AuthorList.Count > 0 Then
            '                AuthorList.Add(objAuthorInfor.tagRange)
            '                auCount = auCount + 1

            '                dictFMAuthor.Add(auCount, AuthorList)
            '                AuthorList = New List(Of Word.Range)
            '            Else
            '                AuthorList.Add(objAuthorInfor.tagRange)
            '            End If

            '        ElseIf objAuthorInfor.tagName = "‡fm_corrPrefix" Or objAuthorInfor.tagName = "‡fm_auPrefix" Then
            '            If AuthorList.Count = 0 Then
            '                AuthorList.Add(objAuthorInfor.tagRange)
            '            ElseIf AuthorList.Count > 0 Then
            '                auCount = auCount + 1
            '                dictFMAuthor.Add(auCount, AuthorList)
            '                AuthorList = New List(Of Word.Range)
            '                AuthorList.Add(objAuthorInfor.tagRange)
            '            Else
            '                MessageBox.Show("wrong code")
            '            End If
            '        ElseIf objAuthorInfor.tagName = "‡fm_corrGivenName" Or objAuthorInfor.tagName = "‡fm_auGivenName" Then
            '            If AuthorList.Count > 1 Then
            '                auCount = auCount + 1
            '                dictFMAuthor.Add(auCount, AuthorList)
            '                AuthorList = New List(Of Word.Range)
            '            End If
            '            AuthorList.Add(objAuthorInfor.tagRange)
            '        Else
            '            AuthorList.Add(objAuthorInfor.tagRange)
            '        End If
            '    Next
            'End If
            'If ranDoc.End >= wDoc.Range.End Then Exit Function
            'ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
            'ranDoc.Find.Text = ""
            'ranDoc.Find.Style = sStyleName
            'dictAuthor.Clear()

            If dictFMAuthor.Count > 0 Then
                CollectDisclosureAuthorInformation = dictFMAuthor : Exit Function
            Else
                CollectDisclosureAuthorInformation = Nothing : Exit Function
            End If
            Exit Function
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    '''''Developed by RAJA ToShadingFootnoteText
    Public Function ToShadingFootnoteText(wDoc As Word.Document, WordApp As Word.Application) As Boolean
        Try
            Dim miscINI As String
            Dim sPattern As String
            Dim sShading As String
            Dim fn As Word.Footnote
            Dim rng1 As Word.Range

            wAPP = WordApp
            miscINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGMisc.ini")
            Dim oReadINI As New CEGINI.clsINI(miscINI)
            sPattern = oReadINI.INIReadValue("FootnoteRegex", "pattern")
            sShading = oReadINI.INIReadValue("FootnoteRegex", "highlight")

            Dim rcolor As Object = RGB(255, 255, 251)
            Dim matches As MatchCollection = Regex.Matches(sShading, "([0-9]+), ([0-9]+), ([0-9]+)", RegexOptions.IgnoreCase)
            For Each crgb As Match In matches
                rcolor = RGB(Integer.Parse(crgb.Groups(1).Value), Integer.Parse(crgb.Groups(2).Value), Integer.Parse(crgb.Groups(3).Value))
            Next
            If wDoc.Footnotes.Count < 1 Then
                MessageBox.Show("Error: Footnote text not available in this file")
            Else
                For Each fn In wDoc.Footnotes
                    Dim i As Integer
                    For i = 1 To fn.Range.Words.Count
                        'shadeText = shadeText + fn.Range.Words.Item(i)
                        Dim endPos As Integer
                        If fn.Range.Words.Item(i).Text <> "" Then
                            rng1 = fn.Range.Words.Item(i)
                            If i <> 1 And i < 4 Then
                                endPos = fn.Range.Words.Item(i).End
                                rng1.SetRange(fn.Range.Words.Item(1).Start, endPos)
                                rng1.Select()
                                Dim m As Match = Regex.Match(rng1.Text.Trim(), sPattern, RegexOptions.IgnoreCase)
                                If (m.Success) Then
                                    rng1.Shading.BackgroundPatternColor = rcolor
                                End If
                            End If
                            'rng1 = fn.Range.Words.Item(1)
                        End If
                    Next i
                Next fn
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try

    End Function

    Public Function ToRemoveAllBookmarks(wDoc As Word.Document, WordApp As Word.Application) As Boolean
        Try
            Dim objBookmark As Word.Bookmark

            For Each objBookmark In wDoc.Bookmarks
                objBookmark.Delete()
            Next
            Return True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function


End Module
