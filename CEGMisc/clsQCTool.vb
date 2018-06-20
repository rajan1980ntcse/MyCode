Imports Word = Microsoft.Office.Interop.Word
Imports System.Xml
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports NCalc
Imports CEGINI
Public Class clsQCTool
    Public Sub New()
        Type = String.Empty : DocName = String.Empty
        lText = String.Empty : rText = String.Empty
        dVal = String.Empty : tStr = String.Empty
    End Sub
    Friend Type As String = ""
    Friend DocName As String = ""
    Friend lText As String = ""
    Friend rText As String = ""
    Friend dVal As String = ""
    Friend tStr As String = ""

    Public prefixBookmark = "CEGMiscQC"
    Public cntBookmark As Integer
    Public QcReportDetails As Dictionary(Of String, List(Of clsQCTool))

    Public Function ToCallJournalBasedQC(oWordDoc As Word.Document, JName As String)
        Dim pQCINI As String
        Dim strRules As String
        Dim sRule As String
        Try
            pQCINI = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "CEGQCTOOL.ini")
            If File.Exists(pQCINI) Then
                Dim oReadINI As New CEGINI.clsINI(pQCINI)
                QcReportDetails = New Dictionary(Of String, List(Of clsQCTool))()
                strRules = oReadINI.INIReadValue("QCTOOL", UCase(JName))
                Dim bMark As Word.Bookmark
                For Each bMark In oWordDoc.Bookmarks
                    If bMark.Name.Contains(prefixBookmark) Then
                        bMark.Delete()
                    End If
                Next
                For Each sRule In strRules.Split("#")
                    If sRule.ToUpper.Contains("FRULE") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        FindRuleReport(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("WORDCNTRULE") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        WordCountRuleReport(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("ASCRULE") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        AscendingRuleReport(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("ABSTRACTHEADING") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        AbstractHeadingCheck(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("JOURNALHEADING") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        JournalHeadingCheck(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("TABLENOTE") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        TableNoteReport(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    ElseIf sRule.ToUpper.Contains("SPELLING") Then
                        oWordDoc.Application.StatusBar = sRule & "... Processing"
                        SpellCheck(sRule, oReadINI, oWordDoc)
                        oWordDoc.Application.StatusBar = sRule & "... Completed"
                    End If
                Next
                Call QCReportCreation(oWordDoc)
                Call AddQCIteminCollection("JournalQC", oWordDoc)
                Call MessageBox.Show("Journal based QC report was successfully created", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("File not found : " & pQCINI, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Function SpellCheck(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document) As Integer
        Try
            Dim strType As String = oReadINI.IniReadValue(sRule, "Type")
            Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
            Dim strIgnoreStyle As String = oReadINI.IniReadValue(sRule, "Ignore Styles")
            Dim SPErrorCnt As Integer
            Dim clsList As New List(Of clsQCTool)()
            Dim DocName As String = wDoc.Name
            Dim oRng As Word.Range

            'Dim oSpErrors As Word.ProofreadingErrors
            Dim oSp As Object
            Dim Sty As Word.Style
            Dim spList As New ArrayList
            For Each Sty In wDoc.Styles
                If Sty.InUse = True Then
                    Try
                        If strType.ToUpper.Contains("UK") Then
                            Sty.LanguageID = Word.WdLanguageID.wdEnglishUK
                        ElseIf strType.ToUpper.Contains("US") Then
                            Sty.LanguageID = Word.WdLanguageID.wdEnglishUS
                        End If
                    Catch ex As Exception
                    End Try
                End If
            Next

            For I = 1 To 3
                Select Case I
                    Case 1 : oRng = wDoc.Content.Duplicate
                    Case 2 : If wDoc.Footnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory).Duplicate Else oRng = Nothing
                    Case 3 : If wDoc.Endnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory).Duplicate Else oRng = Nothing
                End Select
                If Not IsNothing(oRng) Then
                    If strType.ToUpper.Contains("UK") Then
                        oRng.LanguageID = Word.WdLanguageID.wdEnglishUK
                    ElseIf strType.ToUpper.Contains("US") Then
                        oRng.LanguageID = Word.WdLanguageID.wdEnglishUS
                    End If
                    oRng.SpellingChecked = False
                End If
            Next
            Dim oSpErrors As Word.ProofreadingErrors = wDoc.SpellingErrors
            For Each oSp In oSpErrors
                Dim sName As Word.Style
                sName = oSp.Style
                If Not strIgnoreStyle.Contains(sName.NameLocal) Then
                    sName = oSp.Paragraphs(1).Style
                End If
                If Not strIgnoreStyle.Contains(sName.NameLocal) Then
                    If Not spList.Contains(oSp.Text) Then
                        spList.Add(oSp.Text)
                        Dim QcList As New clsQCTool()
                        QcList.DocName = DocName
                        QcList.lText = strDesc
                        QcList.rText = oSp.Text
                        clsList.Add(QcList)
                    End If
                End If
            Next

            If clsList.Count > 0 Then
                If Not QcReportDetails.ContainsKey(sRule) Then
                    QcReportDetails.Add(sRule.Replace("F", ""), clsList)
                Else
                    Dim kVal As List(Of clsQCTool)
                    QcReportDetails.TryGetValue(sRule, kVal)
                    For Each objcJob As clsQCTool In clsList
                        kVal.Add(objcJob)
                    Next
                    QcReportDetails(sRule) = kVal
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle)
        End Try
    End Function

    Private Function FindRuleReport(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document)
        'Dim strKeys As Dictionary(Of String, List(Of String))
        Dim strStyleName As String = oReadINI.IniReadValue(sRule, "StyleName")
        Dim strPattern As String = oReadINI.IniReadValue(sRule, "Pattern")
        Dim strContain As String = oReadINI.IniReadValue(sRule, "Contain")
        Dim strCase As String = oReadINI.IniReadValue(sRule, "Match case")
        Dim strFormat As String = oReadINI.IniReadValue(sRule, "Formating")
        Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
        Dim strIgnoreStyle As String = oReadINI.IniReadValue(sRule, "Ignore Styles")
        Dim clsList As New List(Of clsQCTool)()
        Dim DocName As String = wDoc.Name
        Dim matchList As New ArrayList
        Dim strTemp As New StringBuilder
        If wDoc.Footnotes.Count > 0 Then
            For i = 1 To wDoc.Footnotes.Count
                strTemp.Append(wDoc.Footnotes(i).Range.Text)
            Next
        End If
        If wDoc.Endnotes.Count > 0 Then
            For i = 1 To wDoc.Endnotes.Count
                strTemp.Append(wDoc.Endnotes(i).Range.Text)
            Next
        End If
        strTemp.Append(wDoc.Range.Text)
        Try
            Dim ranDoc As Word.Range
            If strPattern <> String.Empty Then
                Dim rx As Regex
                If strCase = "1" Then
                    rx = New Regex(strPattern)
                Else
                    rx = New Regex(strPattern, RegexOptions.IgnoreCase)
                End If
                Dim matches As MatchCollection = rx.Matches(strTemp.ToString())
                For Each mat As Match In matches
                    ranDoc = Nothing
                    If Not matchList.Contains(mat.Value) Then
                        matchList.Add(mat.Value)
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
                                ranDoc.Find.ClearFormatting()
                                If strStyleName <> String.Empty Then ranDoc.Find.Style = strStyleName 'StyleName
                                If Not strFormat.Contains("!") Then
                                    Select Case (UCase(strFormat))
                                        Case "B"
                                            ranDoc.Find.Font.Bold = True
                                        Case "I"
                                            ranDoc.Find.Font.Italic = True
                                        Case "U"
                                            ranDoc.Find.Font.Underline = True
                                        Case "SUP"
                                            ranDoc.Find.Font.Superscript = True
                                        Case "SUB"
                                            ranDoc.Find.Font.Subscript = True
                                    End Select
                                ElseIf strFormat.Contains("!") Then
                                    Select Case (UCase(strFormat))
                                        Case "!B"
                                            ranDoc.Find.Font.Bold = False
                                        Case "!I"
                                            ranDoc.Find.Font.Italic = False
                                        Case "!U"
                                            ranDoc.Find.Font.Underline = False
                                        Case "!SUP"
                                            ranDoc.Find.Font.Superscript = False
                                        Case "!SUB"
                                            ranDoc.Find.Font.Subscript = False
                                    End Select
                                End If
                                If strCase = "1" Then
                                    ranDoc.Find.MatchCase = True
                                End If
                                ranDoc.Find.Text = mat.Value
                                Do While ranDoc.Find.Execute = True
                                    If IsContainString(strContain, mat.Value) Then
                                        ranDoc.Select()
                                        Dim sName As Word.Style
                                        sName = ranDoc.Style
                                        If Not strIgnoreStyle.Contains(sName.NameLocal) Then
                                            sName = ranDoc.Paragraphs(1).Style
                                        End If
                                        If Not strIgnoreStyle.Contains(sName.NameLocal) Then
                                            cntBookmark += 1
                                            Dim QcList As New clsQCTool()
                                            QcList.DocName = DocName
                                            QcList.lText = strDesc
                                            wDoc.Bookmarks.Add(prefixBookmark & cntBookmark, ranDoc)
                                            QcList.rText = ranDoc.Paragraphs(1).Range.Text.Replace(mat.Value, "<a href=""" & DocName & "#" & cntBookmark & """>" & mat.Value & "</a>")
                                            clsList.Add(QcList)
                                        End If
                                    End If
                                Loop
                            End If
                        Next 'i=0
                    End If
                Next 'Regex match
            End If
            If clsList.Count > 0 Then
                If Not QcReportDetails.ContainsKey(sRule) Then
                    QcReportDetails.Add(sRule.Replace("F", ""), clsList)
                Else
                    Dim kVal As List(Of clsQCTool)
                    QcReportDetails.TryGetValue(sRule, kVal)
                    For Each objcJob As clsQCTool In clsList
                        kVal.Add(objcJob)
                    Next
                    QcReportDetails(sRule) = kVal
                End If
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle)
        End Try
    End Function
    Private Function IsContainString(strContain As String, sMatchValue As String) As Boolean
        Try
            If strContain.Contains("|") Then
                For Each sTemp As String In strContain.Split("|")
                    If sMatchValue.Contains(sTemp) Then
                        IsContainString = True
                        Exit Function
                    End If
                Next
            End If
        Catch ex As Exception

        End Try
    End Function

    Private Function WordCountRuleReport(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document)
        Try
            Dim strStyleName As String = oReadINI.IniReadValue(sRule, "StyleName")
            Dim strFormat As String = oReadINI.IniReadValue(sRule, "Formating")
            Dim strCondition As String = oReadINI.IniReadValue(sRule, "Condition")
            Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
            Dim strIgnoreStyle As String = oReadINI.IniReadValue(sRule, "Ignore Styles")
            Dim clsList As New List(Of clsQCTool)()
            Dim DocName As String = wDoc.Name
            Dim cnt As Integer
            Dim strTemp As String
            If strStyleName <> String.Empty Then
                strTemp = GetStyleText(wDoc.Range, strStyleName, wDoc)
                If strCondition <> String.Empty Then
                    Dim qType As String = Regex.Match(strCondition, "\[(.*?)\]", RegexOptions.IgnoreCase).Groups(1).ToString()
                    Dim qParttern As String = oReadINI.IniReadValue(sRule, qType)
                    cnt = Regex.Matches(strTemp, qParttern).Count
                    Dim expr As New Expression(strCondition)
                    expr.Parameters(qType) = cnt
                    If expr.Evaluate Then
                        Dim QcList As New clsQCTool()
                        QcList.DocName = DocName
                        QcList.lText = strDesc
                        QcList.rText = strTemp
                        clsList.Add(QcList)
                    End If
                End If
                If clsList.Count > 0 Then
                    If Not QcReportDetails.ContainsKey(sRule) Then
                        QcReportDetails.Add(sRule.Replace("WordCnt", "COUNT "), clsList)
                    Else
                        Dim kVal As List(Of clsQCTool)
                        QcReportDetails.TryGetValue(sRule, kVal)
                        For Each objcJob As clsQCTool In clsList
                            kVal.Add(objcJob)
                        Next
                        QcReportDetails(sRule) = kVal
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Function AscendingRuleReport(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document)
        Try
            Dim strStyleName As String = oReadINI.IniReadValue(sRule, "StyleName")
            Dim strFormat As String = oReadINI.IniReadValue(sRule, "Formating")
            Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
            Dim strIgnoreStyle As String = oReadINI.IniReadValue(sRule, "Ignore Styles")
            Dim clsList As New List(Of clsQCTool)()
            Dim DocName As String = wDoc.Name
            Dim strTemp As New StringBuilder
            Dim aList As New ArrayList
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
                    ranDoc.Find.ClearFormatting()
                    If strStyleName <> String.Empty Then ranDoc.Find.Style = strStyleName 'StyleName
                    If Not strFormat.Contains("!") Then
                        Select Case (UCase(strFormat))
                            Case "B"
                                ranDoc.Find.Font.Bold = True
                            Case "I"
                                ranDoc.Find.Font.Italic = True
                            Case "U"
                                ranDoc.Find.Font.Underline = True
                            Case "SUP"
                                ranDoc.Find.Font.Superscript = True
                            Case "SUB"
                                ranDoc.Find.Font.Subscript = True
                        End Select
                    ElseIf strFormat.Contains("!") Then
                        Select Case (UCase(strFormat))
                            Case "!B"
                                ranDoc.Find.Font.Bold = False
                            Case "!I"
                                ranDoc.Find.Font.Italic = False
                            Case "!U"
                                ranDoc.Find.Font.Underline = False
                            Case "!SUP"
                                ranDoc.Find.Font.Superscript = False
                            Case "!SUB"
                                ranDoc.Find.Font.Subscript = False
                        End Select
                    End If
                    ranDoc.Find.Text = ""
                    Do While ranDoc.Find.Execute = True
                        ranDoc.Select()
                        Dim sName As Word.Style
                        sName = ranDoc.Style
                        If Not strIgnoreStyle.Contains(sName.NameLocal) Then
                            aList.Add(ranDoc.Text)
                        End If
                    Loop
                End If
            Next 'i=0
            If aList.Count > 0 Then
                Dim orgList As New ArrayList
                For Each sTemp As String In aList
                    orgList.Add(sTemp)
                Next
                aList.Sort()
                If comparearlist(orgList, aList) = True Then                    '
                    Dim QcList As New clsQCTool()
                    QcList.DocName = DocName
                    QcList.lText = strDesc
                    QcList.rText = orgList.ToString()
                    clsList.Add(QcList)
                End If
            End If
            If clsList.Count > 0 Then
                If Not QcReportDetails.ContainsKey(sRule) Then
                    QcReportDetails.Add(sRule.Replace("ASC", "SORT "), clsList)
                Else
                    Dim kVal As List(Of clsQCTool)
                    QcReportDetails.TryGetValue(sRule, kVal)
                    For Each objcJob As clsQCTool In clsList
                        kVal.Add(objcJob)
                    Next
                    QcReportDetails(sRule) = kVal
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Shared Function comparearlist(a1 As ArrayList, a2 As ArrayList) As Boolean
        Dim Flag As Boolean
        Dim a3 As New ArrayList()
        For j As Integer = 0 To a1.Count - 1
            If Not a1.Item(j).Equals(a2(j)) Then
                a3.Add("test")
            End If
        Next
        If a3.Count > 0 Then Flag = True
        Return Flag
    End Function
    Private Function AbstractHeadingCheck(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document)
        Try
            Dim strStyleName As String = oReadINI.IniReadValue(sRule, "StyleName")
            Dim strContain As String = oReadINI.IniReadValue(sRule, "Contain")
            Dim strCondition As String = oReadINI.IniReadValue(sRule, "Condition")
            Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
            Dim clsList As New List(Of clsQCTool)()
            Dim ranDoc As Word.Range
            Dim DocName As String = wDoc.Name
            Dim cnt As Integer
            Dim headList As New ArrayList
            Dim strTemp As String
            ranDoc = wDoc.Range
            If strStyleName <> String.Empty Then
                strTemp = GetStyleText(wDoc.Range, strStyleName, wDoc)
                If strCondition <> String.Empty Then
                    Dim qType As String = Regex.Match(strCondition, "\[(.*?)\]", RegexOptions.IgnoreCase).Groups(1).ToString()
                    Dim qParttern As String = oReadINI.IniReadValue(sRule, qType)
                    cnt = Regex.Matches(strTemp, qParttern).Count
                    Dim expr As New Expression(strCondition)
                    expr.Parameters(qType) = cnt
                    If expr.Evaluate Then 'Condition 1
                        ''''''''''''''''''''''''''''
                        For Each strTemp In strContain.Split("|")
                            headList.Add(strTemp.ToUpper)
                        Next
                        ''''''''''''''''''''''''''''
                        ranDoc.Find.ClearFormatting()
                        ranDoc.Find.Style = strStyleName
                        Do While ranDoc.Find.Execute = True
                            ranDoc.Select()
                            Dim ranTemp As Word.Range
                            ranTemp = ranDoc.Duplicate
                            ranTemp.Find.ClearFormatting()
                            ranTemp.Find.Text = ""
                            ranTemp.Find.Font.ColorIndex = Word.WdColorIndex.wdGreen
                            If ranTemp.Find.Execute = True Then
                                If headList.Contains(ranTemp.Text.ToUpper) Then
                                    headList.Remove(ranTemp.Text.ToUpper)
                                End If
                            End If
                        Loop

                    End If
                End If
                If headList.Count > 0 Then
                    Dim QcList As New clsQCTool()
                    QcList.DocName = DocName
                    QcList.lText = strDesc
                    QcList.rText = "Please check the abstract heading"
                    clsList.Add(QcList)
                End If
                If clsList.Count > 0 Then
                    If Not QcReportDetails.ContainsKey(sRule) Then
                        QcReportDetails.Add(sRule, clsList)
                    Else
                        Dim kVal As List(Of clsQCTool)
                        QcReportDetails.TryGetValue(sRule, kVal)
                        For Each objcJob As clsQCTool In clsList
                            kVal.Add(objcJob)
                        Next
                        QcReportDetails(sRule) = kVal
                    End If
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Function JournalHeadingCheck(sRule As String, oReadINI As clsINI, wDoc As Word.Document)
        Try
            Dim strStyleName As String = oReadINI.INIReadValue(sRule, "StyleName")
            Dim strContain As String = oReadINI.INIReadValue(sRule, "Contain")
            Dim strCondition As String = oReadINI.INIReadValue(sRule, "Condition")
            Dim strDesc As String = oReadINI.INIReadValue(sRule, "Description")
            Dim clsList As New List(Of clsQCTool)()
            Dim ranDoc As Word.Range
            Dim DocName As String = wDoc.Name
            Dim headList As New ArrayList
            Dim strTemp As String
            ranDoc = wDoc.Range
            If strCondition <> String.Empty Then
                If strCondition.Contains("@") Then
                    Dim sTemp() As String = strCondition.Split("@")
                    ranDoc.Find.ClearFormatting()
                    ranDoc.Find.Style = sTemp(0)
                    ranDoc.Find.Text = ""
                    If ranDoc.Find.Execute = True Then
                        strTemp = ranDoc.Text.ToUpper.Trim
                    End If
                    If strTemp = sTemp(1).ToUpper Then 'condition 1
                        ''''''''''''''''''''''''''''
                        For Each strTemp In strContain.Split("|")
                            If Not headList.Contains(strTemp) Then
                                headList.Add(strTemp.ToUpper)
                            End If
                        Next
                        ''''''''''''''''''''''''''''
                        For Each stName As String In strStyleName.Split("|")
                            If stName <> String.Empty Then

                                ranDoc = wDoc.Range
                                ranDoc.Find.Style = stName
                                Do While ranDoc.Find.Execute = True
                                    ranDoc.Select()
                                    If headList.Contains(ranDoc.Text.Replace(":", "").ToUpper.Trim) Then
                                        headList.Remove(ranDoc.Text.Replace(":", "").ToUpper.Trim)
                                    End If
                                Loop
                            End If
                        Next
                    End If
                End If
            End If

            strTemp = String.Empty
            If headList.Count > 0 Then
                For Each sTem As String In headList
                    strTemp += sTem & ", "
                Next
                Dim QcList As New clsQCTool()
                QcList.DocName = DocName
                QcList.lText = strDesc
                QcList.rText = strTemp
                clsList.Add(QcList)
            End If
            If clsList.Count > 0 Then
                If Not QcReportDetails.ContainsKey(sRule) Then
                    QcReportDetails.Add(sRule.Replace("WordCnt", "COUNT "), clsList)
                Else
                    Dim kVal As List(Of clsQCTool)
                    QcReportDetails.TryGetValue(sRule, kVal)
                    For Each objcJob As clsQCTool In clsList
                        kVal.Add(objcJob)
                    Next
                    QcReportDetails(sRule) = kVal
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Function TableNoteReport(sRule As String, oReadINI As CEGINI.clsINI, wDoc As Word.Document)
        Dim ranPara As Word.Range
        Dim ranNote As Word.Range
        Dim strStyleName As String = oReadINI.IniReadValue(sRule, "StyleName")
        Dim strDesc As String = oReadINI.IniReadValue(sRule, "Description")
        Dim clsList As New List(Of clsQCTool)()
        Dim DocName As String = wDoc.Name
        Dim aList As New ArrayList
        Try
            Dim dTable As Word.Table
            For Each dTable In wDoc.Tables
                If dTable.Range.Next.Paragraphs.Count > 0 Then
                    ranPara = dTable.Range.Next.Paragraphs(1).Range
                    Dim sName As Word.Style
                    sName = ranPara.Style
                    Do While sName.NameLocal = strStyleName
                        ranPara.Select()
                        ranNote = ranPara.Duplicate
                        ranNote.Find.ClearFormatting()
                        ranNote.Find.Font.Superscript = True
                        ranNote.Find.Text = ""
                        Do While ranNote.Find.Execute = True
                            ranNote.Select()
                            If Not ranNote.InRange(ranPara) Then Exit Do
                            aList.Add(ranNote.Text)
                        Loop
                        If ranPara.End = wDoc.Range.End Then Exit Do
                        If ranPara.Next.Paragraphs.Count > 0 Then
                            ranPara = ranPara.Next.Paragraphs(1).Range
                            sName = ranPara.Style
                        End If
                    Loop
                End If
                If aList.Count > 0 Then
                    Dim orgList As New ArrayList
                    For Each sTemp As String In aList
                        orgList.Add(sTemp)
                    Next
                    aList.Sort()
                    If comparearlist(orgList, aList) = True Then                    '
                        Dim QcList As New clsQCTool()
                        QcList.DocName = DocName
                        QcList.lText = strDesc
                        If dTable.Range.Previous.Paragraphs.Count > 0 Then
                            QcList.rText = dTable.Range.Previous.Paragraphs(1).Range.Text.Trim
                        End If
                        clsList.Add(QcList)
                    End If
                End If
                aList.Clear()
            Next
            If clsList.Count > 0 Then
                If Not QcReportDetails.ContainsKey(sRule) Then
                    QcReportDetails.Add(sRule, clsList)
                Else
                    Dim kVal As List(Of clsQCTool)
                    QcReportDetails.TryGetValue(sRule, kVal)
                    For Each objcJob As clsQCTool In clsList
                        kVal.Add(objcJob)
                    Next
                    QcReportDetails(sRule) = kVal
                End If
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Function QCReportCreation(oWordDoc As Word.Document)
        Dim htmText As New StringBuilder()
        Dim HtmlRoot As String
        HtmlRoot = "<HTML><head><META http-equiv='Content-Type' content='text/html; charset=utf-8'>" +
                    "<H4 align=""center"" style='background-color:FF99CC;font-family:Verdana'>CE Genius QC Tool Report</H4><body bgcolor=""#FFFFFF"" style=""font-family:Verdana""><table border=""0"" align=""center""><tbody><tr><td><b>Date and time</b></td><td><b>: " + DateTime.Now + "</b></td></tr><tr><td><b> User name</b></td><td><b>: " + Environment.UserName + "</b></td></tr></tbody></table><hr color='#FF8C00'/>" +
                    "</head><body style='font-family:Times New Roman'>"
        Try
            For Each Temp As KeyValuePair(Of String, List(Of clsQCTool)) In QcReportDetails
                Dim sKey As String = Temp.Key
                Dim cList As List(Of clsQCTool) = Temp.Value
                'htmText.Append((Convert.ToString((Convert.ToString((Convert.ToString("<div class=""") & sKey) + """ id=""" + sKey.Replace(" ", "") + """><h3 class=""") & sKey) + """ style='background-color:00BFFF'>") & sKey) + "</h3><br/>")
                htmText.Append("<table border='1' width='100%'>")
                'htmText.Append("<tr><td style='background-color:#87CEEB'><b>Content In Document</b></td><td style='background-color:#87CEEB'><b>Description<b></td></tr>")
                For Each grp As Object In cList.GroupBy(Function(x1) x1.lText)
                    htmText.Append("<tr><td colspan=5 style='background-color:#87CEEB'><b>" + grp.key + "</b></td></tr>")
                    For Each gValue As clsQCTool In grp
                        htmText.Append("<tr><td>" + gValue.rText + "</td></tr>")
                    Next
                Next
                htmText.Append("<table>")
            Next
            If htmText.ToString <> String.Empty Then
                WriteHtmlFile(oWordDoc.Path & "\" & oWordDoc.Name.Replace(".docx", "").Replace(".doc", "") & "_QcReport.html", HtmlRoot & htmText.ToString & " </body></html>")
            Else
                MessageBox.Show("Empty Report", sMsgTitle)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, sMsgTitle)
        End Try
    End Function

    Private Sub WriteHtmlFile(HPath As String, sWrite As String)
        Dim objW As New StreamWriter(HPath)
        objW.Write(sWrite)
        objW.Close()
    End Sub
End Class



