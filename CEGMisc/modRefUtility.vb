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

Public Class clsRefInfo
    Public iRefNum As Integer
    Public oRefRng As Word.Range
    Public sRefYear As String
    Public sRefCollab As Boolean
    Public sRefEtal As Boolean
    Public olRefAuthors As List(Of String)
    Public sRefVol As String
    Public sRefIssue As String
    Public sRefJourTitle As String
    Public sRefPage As String
    Public oCitationRng As Word.Range
    Public oRefLabelRng As Word.Range
    Public rCondText As String
    Public cText As String
End Class
Public Class clsAuthorInfo
    Public oAuthorRng As Word.Range
    Public tagName As String
    Public tagValue As String
    Public tagRange As Word.Range
    'Public fm_corrGivenName As Word.Range
    'Public fm_corrSurname As Word.Range
    'Public fm_corrPrefix As Word.Range
    'Public fm_corrSuffix As Word.Range
    'Public fm_corrDegree As Word.Range
    'Public fm_auGivenName As Word.Range
    'Public fm_auSurname As Word.Range
    'Public fm_auPrefix As Word.Range
    'Public fm_auSuffix As Word.Range
    'Public fm_auDegree As Word.Range
    'Public corrGivenName As String
    'Public corrSurname As String
    'Public corrPrefix As String
    'Public corrSuffix As String
    'Public corrDegree As String
    'Public auGivenName As String
    'Public auSurname As String
    'Public auPrefix As String
    'Public auSuffix As String
    'Public auDegree As String
End Class
Public Class clsCitationInfo
    Public cTotalRange As Word.Range
    Public cIntRange As Word.Range
    Public sYear As String
    Public sComment As String
End Class
Module ModRefUtility
    Public dictHavCitInfo As Dictionary(Of Word.Range, Dictionary(Of Word.Range, List(Of String)))
    Public dictRefInfo As Dictionary(Of Integer, clsRefInfo)
    Public dictCitationInfo As Dictionary(Of Word.Range, String)
    Public dictHavCitaion As Dictionary(Of Word.Range, String)
    Public dictConvertedCitation As Dictionary(Of Word.Range, String)
    Public oActApp As Word.Application
    Public cBracketType As String
    Public cTextPattern As String
    Public cCitationOrder As String
    Public cNoOfEtal As String
    Public cEtalPattern As String
    Public cCitationSep As String

    Public Function ReferenceGranularFormating(wDoc As Word.Document, WAPP As Word.Application, pName As String, jName As String, sConfigPath As String) As Boolean
        Try
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim sRefFormat = oReadINI.INIReadValue(pName & "@" & jName, "ReferenceFormat")
            oActApp = WAPP
            Dim sStyleFormat = Split(sRefFormat, "|")
            Dim I As Integer
            For I = LBound(sStyleFormat) To UBound(sStyleFormat)
                If Not String.IsNullOrEmpty(sStyleFormat(I)) Then
                    ToApplyReferenceFormat(sStyleFormat(I), wDoc)
                End If
                oActApp.Selection.HomeKey(Word.WdUnits.wdStory)
            Next
            AddQCIteminCollection("Ref Formating", wDoc)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Public Function ReferenceCitationSort(wDoc As Word.Document, WAPP As Word.Application, pName As String, jName As String, sConfigPath As String) As Boolean
        Try
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim refSortType = oReadINI.INIReadValue(pName & "@" & jName, "RefCitationSort")
            Dim sCitStyle As String
            Dim lstBracketRange As New List(Of Word.Range)
            oActApp = WAPP
            dictHavCitInfo = New Dictionary(Of Word.Range, Dictionary(Of Word.Range, List(Of String)))
            ''RefCitationSort =()|,|Chronological|BibXref_online
            Dim PipeCount = refSortType.Count(Function(x) x = "|"c)
            If PipeCount = 3 Then
                cBracketType = refSortType.Split("|")(0)
                cCitationSep = refSortType.Split("|")(1)
                cCitationOrder = refSortType.Split("|")(2)
                sCitStyle = refSortType.Split("|")(3)
            Else
                MessageBox.Show("Please check the journal configuration : Unable sort citation text", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
            End If
            For Each sRng In wDoc.StoryRanges
                If sRng.StoryType = Word.WdStoryType.wdMainTextStory Or
                (wDoc.Footnotes.Count > 0 And sRng.StoryType = Word.WdStoryType.wdFootnotesStory) Or
                (wDoc.Endnotes.Count > 0 And sRng.StoryType = Word.WdStoryType.wdEndnotesStory) Then
                    ClrPair(sRng.Duplicate)
                    lstBracketRange.AddRange(PairingRoutine(wDoc, "(", ")", False, sRng.Duplicate))
                End If
            Next
            If lstBracketRange.Count > 0 Then
                GetCitationInfo(wDoc, sCitStyle, lstBracketRange)
            Else
                MessageBox.Show("() not found in the document.")
            End If


            If dictHavCitInfo.Count > 0 Then
                'Dim objSortCit As New frmSortHarvardCitation(wDoc, WAPP)
                'objSortCit.Show()

                For Each dPair As KeyValuePair(Of Word.Range, Dictionary(Of Word.Range, List(Of String))) In dictHavCitInfo
                    Dim oActDoc As Word.Document
                    oActDoc = oActApp.Documents.Add()
                    oActApp.Visible = True : oActApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                    Dim CitVal As String = ""
                    Dim strComment As String
                    For Each pRair As KeyValuePair(Of Word.Range, List(Of String)) In dPair.Value
                        If dPair.Value.Last().Key Is pRair.Key Then
                            ''CitVal = CitVal & pRair.Key.Text
                            oActApp.Selection.FormattedText = pRair.Key.FormattedText
                            oActApp.Selection.MoveEnd(Word.WdUnits.wdLine)
                            oActApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        Else
                            ''CitVal = CitVal & pRair.Key.Text & ModRefUtility.cCitationSep & " "
                            oActApp.Selection.FormattedText = pRair.Key.FormattedText
                            oActApp.Selection.MoveEnd(Word.WdUnits.wdParagraph)
                            oActApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            oActApp.Selection.Text = ModRefUtility.cCitationSep & " "
                            oActApp.Selection.ClearFormatting()
                            oActApp.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                        End If
                        If pRair.Value.Item(2).ToString() <> "" Then
                            strComment = pRair.Value.Item(2).ToString()
                        End If
                    Next
                    oActApp.Selection.EndKey(Word.WdUnits.wdStory)
                    oActApp.Selection.HomeKey(Word.WdUnits.wdStory, True)

                    oActDoc.Range.Copy()
                    If oActDoc.Range.Characters.Count() - 1 = dPair.Key.Characters.Count() And strComment = "" Then
                        wDoc.Activate()
                        dPair.Key.Select()
                        Dim selStyle As Word.Style
                        selStyle = oActApp.Selection.Paragraphs(1).Style
                        Debug.Print(dPair.Key.Text)
                        ''MsgBox(dPair.Key.Text)
                        ''dPair.Key.Delete()
                        oActApp.Selection.Delete()
                        oActApp.Selection.PasteSpecial()
                        ''oActApp.Selection.Style = selStyle.NameLocal
                        If oActApp.Selection.Previous.Characters(1).Text = Chr(13) Then
                            oActApp.Selection.Previous.Characters(1).Select()
                            oActApp.Selection.Delete()
                            oActApp.Selection.Style = selStyle.NameLocal
                        End If
                    Else
                        If strComment = "" Then
                            wDoc.Comments.Add(dPair.Key, "Unstyled content found in the citation. The tool not able to sorting. Please check manually")
                        Else
                            wDoc.Comments.Add(dPair.Key, strComment)
                        End If

                    End If

                    oActDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                    ''oActApp.Visible = True
                Next
                MessageBox.Show("Completed...")
            Else
                MessageBox.Show("Citation Not found in this document.")
            End If
        Catch ex As Exception
            MessageBox.Show("Error : " + ex.Message)
        End Try
    End Function

    Public Function RefArrageAlphabetical(wDoc As Word.Document, WAPP As Word.Application, pName As String, jName As String, sConfigPath As String) As Boolean
        Try
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim refSortType = oReadINI.INIReadValue(pName & "@" & jName, "RefSortType")
            If refSortType <> String.Empty Then
                oActApp = WAPP : dictRefInfo = New Dictionary(Of Integer, clsRefInfo)
                If CollectRefInfo(wDoc, "†Reference") = False Then
                    MessageBox.Show("ERROR : Unable to collect Reference information ")
                    Return False
                End If
                If dictRefInfo.Count > 0 Then
                    Dim dictRefSortInfo As Dictionary(Of Integer, clsRefInfo)
                    Select Case refSortType
                        Case "order by Surname"
                            dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.rCondText).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                        Case "order by character by character"
                            dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.oRefRng.Text.Replace(Chr(13), "")).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                        Case "order by number of authors"
                            dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.olRefAuthors.Count).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                        Case Else
                            MessageBox.Show("RefSortType not configured in " & sConfigPath & " for " & pName & " @ " & jName)
                            Return False
                    End Select
                    Dim oActDoc As Word.Document
                    oActDoc = WAPP.Documents.Add()
                    WAPP.Visible = True : WAPP.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                    For Each pair As KeyValuePair(Of Integer, clsRefInfo) In dictRefSortInfo
                        WAPP.Selection.Range.FormattedText = pair.Value.oRefRng
                        WAPP.Selection.MoveEnd(Word.WdUnits.wdParagraph, 1)
                        WAPP.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Next
                    oActDoc.Range.Copy()
                    wDoc.Activate()
                    ''wDoc.Application.Selection.Select()
                    Debug.Print(wDoc.Name)
                    WAPP.Selection.Select()
                    Debug.Print(wDoc.Name)
                    For Each pair As KeyValuePair(Of Integer, clsRefInfo) In dictRefInfo

                        pair.Value.oRefRng.Delete()
                    Next
                    WAPP.Selection.InsertParagraphAfter()
                    WAPP.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    WAPP.Selection.PasteSpecial()
                    oActDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
                    dictRefInfo.Clear()
                    If CollectRefInfo(wDoc, "†Reference") = True Then
                        Dim lstDuplicateRef As List(Of String) = dictRefInfo.Values.GroupBy(Function(x) x.rCondText).Where(Function(x)
                                                                                                                               Return x.Count() > 1
                                                                                                                           End Function).Select(Function(x) x.Key).ToList()

                        If lstDuplicateRef.Count > 0 Then
                            For Each dupRefStr As String In lstDuplicateRef
                                Dim lstDupRange As List(Of Word.Range) = dictRefInfo.Values.Where(Function(x) x.rCondText = dupRefStr).Select(Function(x) x.oRefRng).ToList()
                                If lstDupRange.Count > 0 Then
                                    For Each dRange As Word.Range In lstDupRange
                                        wDoc.Comments.Add(dRange, "Author names an year of publishing are same in two different references!")
                                    Next
                                End If
                            Next

                        End If
                    End If
                    MessageBox.Show("Reference sorting completed.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    ''MsgBox("Reference sorting completed.",,)
                Else
                    MessageBox.Show("ERROR: Reference character style not found in the document", "Reference Sort")
                End If
                'Dim objSHR As New frmSortHarvardReference()
                'objSHR.Show()
            Else
                MessageBox.Show("ERROR: RefSortType not configured in " & sConfigPath & " for " & pName & " @ " & jName)
            End If

        Catch ex As Exception
            MessageBox.Show("Error : " + ex.Message)
        End Try
    End Function


    Public Function Vancouver2HarvardCitationMain(wDoc As Word.Document, WAPP As Word.Application, pName As String, jName As String, sConfigPath As String)
        Try
            Dim sRefStyle As String
            Dim sCitStyle As String
            Dim oReadINI As New CEGINI.clsINI(sConfigPath)
            Dim refSortType = oReadINI.INIReadValue(pName & "@" & jName, "Vancouver2Harvard")
            ''MsgBox(pName & " " & jName & "  " & sConfigPath)
            Dim PipeCount = refSortType.Count(Function(x) x = "|"c)

            If PipeCount = 6 Then
                cBracketType = refSortType.Split("|")(0)
                cTextPattern = refSortType.Split("|")(1)
                cCitationOrder = refSortType.Split("|")(2)
                cNoOfEtal = refSortType.Split("|")(3)
                sRefStyle = refSortType.Split("|")(4)
                sCitStyle = refSortType.Split("|")(5)
                cEtalPattern = refSortType.Split("|")(6)
                oActApp = WAPP
                If modCEGUtility.AutoStyleExists(sRefStyle, wDoc) = False Then
                    MessageBox.Show(sRefStyle & " style not found in the document.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Function
                End If
                If modCEGUtility.AutoStyleExists(sCitStyle, wDoc) = False Then
                    MessageBox.Show(sCitStyle & " style not found in the document.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Exit Function
                End If
            Else
                MessageBox.Show("Please check the journal configuration : Unable convert vancouver to havard", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
            End If
            dictRefInfo = New Dictionary(Of Integer, clsRefInfo)
            dictCitationInfo = New Dictionary(Of Word.Range, String)
            dictHavCitaion = New Dictionary(Of Word.Range, String)
            dictConvertedCitation = New Dictionary(Of Word.Range, String)
            CollectRefInfo(wDoc, sRefStyle)

            If dictRefInfo.Count > 0 Then
                CollectCitaionInfo(wDoc, sCitStyle)
            Else
                MessageBox.Show("Reference not found in the document.", "Vancouver to Harvard", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exit Function
            End If
            If dictCitationInfo.Count > 0 Then
                ConvertVan2HavCitation()
                Dim sEtalText As String
                If Not String.IsNullOrEmpty(cEtalPattern) Then
                    If cEtalPattern.Contains("@") Then
                        sEtalText = cEtalPattern.Split("@")(1)
                    Else
                        sEtalText = Trim(cEtalPattern)
                    End If
                End If
                Dim bEtal As Boolean
                If (sEtalText.ToUpper() = "ITALIC") Then bEtal = True
                Dim objV2H As New frmVan2Harvard(wDoc, oActApp, bEtal)
                objV2H.Show()
            Else
                MessageBox.Show("Reference citation not found in the document.", "Vancouver to Harvard", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Catch ex As Exception
            MessageBox.Show("Error : " & ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Function

    Public Function ConvertVan2HavCitation()
        Try
            For Each pair As KeyValuePair(Of Integer, clsRefInfo) In ModRefUtility.dictRefInfo
                Dim sHavCitationText As String = ""
                sHavCitationText = GetHavCitationOfReference(pair.Value)
                dictHavCitaion.Add(pair.Value.oRefRng, sHavCitationText)
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Public Function GetHavCitationOfReference(objRefInfo As clsRefInfo) As String
        Dim sCitPattern As String = cTextPattern

        Try
            If objRefInfo.olRefAuthors.Count >= cNoOfEtal Then
                If objRefInfo.sRefEtal = True Then
                    ''MessageBox.Show("Reference contains etal informaion")
                Else
                    Dim sEtalText As String = ""
                    If Not String.IsNullOrEmpty(cEtalPattern) Then
                        If cEtalPattern.Contains("@") Then
                            sEtalText = cEtalPattern.Split("@")(0)
                        Else
                            sEtalText = Trim(cEtalPattern)
                        End If
                    End If
                    sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0) & " " & sEtalText)
                End If
            ElseIf objRefInfo.olRefAuthors.Count = 1 Then
                If objRefInfo.sRefEtal = True Then
                    Dim sEtText As String = ""
                    If Not String.IsNullOrEmpty(cEtalPattern) Then
                        If cEtalPattern.Contains("@") Then
                            sEtText = cEtalPattern.Split("@")(0)
                            sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0) & " " & sEtText)
                        Else
                            sEtText = Trim(cEtalPattern)
                            sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0) & " " & sEtText)
                        End If
                    Else
                        sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0))
                    End If
                Else
                    sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0))
                End If

            ElseIf objRefInfo.olRefAuthors.Count = 2 Then
                sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.olRefAuthors(0) & " and " & objRefInfo.olRefAuthors(1))
            Else
                Dim sTotAuthor As String
                For Each sAut As String In objRefInfo.olRefAuthors
                    If objRefInfo.olRefAuthors.LastIndexOf(sAut) Then
                        sTotAuthor = sTotAuthor & sAut
                    Else
                        sTotAuthor = sTotAuthor & sAut & ", "
                    End If
                Next
                sCitPattern = sCitPattern.Replace("<Author>", sTotAuthor)
            End If
            sCitPattern = sCitPattern.Replace("<Year>", objRefInfo.sRefYear)
            If objRefInfo.olRefAuthors.Count = 0 Or objRefInfo.sRefYear = "" Then
                GetHavCitationOfReference = "Untagged Reference"
            Else
                GetHavCitationOfReference = sCitPattern
            End If
        Catch ex As Exception
            MessageBox.Show("Error :" & ex.Message)
        End Try
    End Function

    Public Function ConvertHavardCitation(CitList As ArrayList) As String
        Dim dictempHavCitation As New Dictionary(Of Integer, clsRefInfo)
        Dim sCitPattern As String = cTextPattern
        Dim returnCitText As String
        Try
            For Each sCitation In CitList
                Dim objRefInfo As clsRefInfo
                If dictRefInfo.ContainsKey(Convert.ToInt32(sCitation)) Then
                    dictRefInfo.TryGetValue(Convert.ToInt32(sCitation), objRefInfo)
                    dictempHavCitation.Add(CitList.IndexOf(sCitation), objRefInfo)
                Else
                    ''MessageBox.Show("need to write code")
                    MessageBox.Show("Please contact to R&D team for this process.", sMsgTitle, MessageBoxButtons.OK)
                End If
            Next
            Dim dicTemp
            If dictempHavCitation.Count > 0 Then
                Select Case (cCitationOrder)
                    Case "Chronological"
                        dicTemp = dictempHavCitation.OrderBy(Function(item) item.Value.sRefYear)
                    Case "Alphabet"
                        dicTemp = dictempHavCitation.OrderBy(Function(item) item.Value.rCondText)
                    Case "None"
                        dicTemp = dictempHavCitation
                End Select
                For Each pair As KeyValuePair(Of Integer, clsRefInfo) In dicTemp
                    sCitPattern = cTextPattern
                    If pair.Value.olRefAuthors.Count >= cNoOfEtal Then
                        If pair.Value.sRefEtal = True Then
                            ''MessageBox.Show("1111111111111111111111111")
                        Else
                            Dim sEtalText As String = ""
                            If String.IsNullOrEmpty(cEtalPattern) Then
                                If cEtalPattern.Contains("@") Then
                                    sEtalText = cEtalPattern.Split("@")(0)
                                Else
                                    sEtalText = Trim(cEtalPattern)
                                End If
                            End If
                            sCitPattern = sCitPattern.Replace("<Author>", pair.Value.olRefAuthors(0) & " " & sEtalText)
                        End If
                    ElseIf pair.Value.olRefAuthors.Count = 1 Then
                        sCitPattern = sCitPattern.Replace("<Author>", pair.Value.olRefAuthors(0))
                    ElseIf pair.Value.olRefAuthors.Count = 2 Then
                        sCitPattern = sCitPattern.Replace("<Author>", pair.Value.olRefAuthors(0) & " and " & pair.Value.olRefAuthors(1))
                    Else
                        Dim sTotAuthor As String
                        For Each sAut As String In pair.Value.olRefAuthors
                            If pair.Value.olRefAuthors.LastIndexOf(sAut) Then
                                sTotAuthor = sTotAuthor & sAut
                            Else
                                sTotAuthor = sTotAuthor & sAut & ", "
                            End If
                        Next
                        sCitPattern = sCitPattern.Replace("<Author>", sTotAuthor)
                    End If
                    sCitPattern = sCitPattern.Replace("<Year>", pair.Value.sRefYear)
                    returnCitText = returnCitText & sCitPattern
                Next
            End If
            ConvertHavardCitation = returnCitText
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function GetHavardCitation(CitList As ArrayList, refTextCitationList As List(Of String)) As String
        Dim dictempHavCitation As New Dictionary(Of Integer, clsRefInfo)
        Dim sCitPattern As String = cTextPattern
        Dim returnCitText As String
        Dim tempCitationList As New List(Of String)
        Dim tempCitationDict As New Dictionary(Of String, String)
        Try
            For Each sCitation In CitList
                Dim sKey As String = refTextCitationList(Val(sCitation) - 1).ToString()
                Dim sYear As String = Regex.Match(refTextCitationList(Val(sCitation) - 1).ToString(), "(18|19|20)[0-9]{2}[a-z]?").Value
                If tempCitationDict.ContainsKey(sKey) = True Then
                    ''tempCitationDict.Add(refTextCitationList(Val(sCitation) - 1).ToString(), sYear)
                    MessageBox.Show("Same first author name and year present in the list. please check and confirm the citation" + Environment.NewLine + "Ref text : " + sKey + " " + sYear, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    tempCitationDict.Add(refTextCitationList(Val(sCitation) - 1).ToString(), sYear)
                End If
            Next
            Dim dicTemp
            If tempCitationDict.Count > 0 Then
                Select Case (cCitationOrder)
                    Case "Chronological"
                        dicTemp = tempCitationDict.OrderBy(Function(ite) ite.Value)
                    Case "Alphabet"
                        dicTemp = tempCitationDict.OrderBy(Function(ite) ite.Key)
                    Case "None"
                End Select
                For Each pair As KeyValuePair(Of String, String) In dicTemp
                    If pair.Key.LastIndexOf(pair.Key) Then
                        returnCitText = returnCitText & pair.Key
                    Else
                        returnCitText = returnCitText & pair.Key & " "
                    End If
                Next
            End If

            GetHavardCitation = returnCitText.Substring(0, returnCitText.LastIndexOf(", "))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function ExpandCitationText(cText As String) As ArrayList
        Dim CitList As New ArrayList
        Try
            cText = Regex.Replace(cText, "(-|–|—)", "-")
            cText = cText.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "")
            If cText.Contains(",") Then
                For Each st As String In cText.Split(",")
                    st = Trim(st)
                    If st.Contains("-") Then
                        For i As Integer = Convert.ToInt32(st.Split("-")(0)) To st.Split("-")(1)
                            CitList.Add(i.ToString())
                        Next
                    Else
                        CitList.Add(st)
                    End If
                Next
            Else
                If cText.Contains("-") Then
                    For i As Integer = Convert.ToInt32(cText.Split("-")(0)) To cText.Split("-")(1)
                        CitList.Add(i.ToString())
                    Next
                Else
                    CitList.Add(cText)
                End If
            End If
            ExpandCitationText = CitList
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Public Function CollectCitaionInfo(wDoc As Word.Document, refCitaionStyle As String)
        Try
            Dim ranDoc As Word.Range
            ranDoc = wDoc.Content
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Style = refCitaionStyle
            End With
            Do While ranDoc.Find.Execute = True
                ranDoc.Select()
                If dictCitationInfo.ContainsKey(ranDoc) Then
                Else
                    dictCitationInfo.Add(oActApp.Selection.Range, ranDoc.Text)
                End If
                ranDoc.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Loop
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function GetCitationInfo(wDoc As Word.Document, sCitStyle As String, lstBracketRange As List(Of Word.Range))
        Try
            Dim countGroup As Integer
            For Each ranCit As Word.Range In lstBracketRange
                Dim PrevCitationText As String
                Dim dictEachCitation As New Dictionary(Of Word.Range, List(Of String))
                Dim dictDupCitation As Dictionary(Of Word.Range, List(Of String))
                If ranCit.Start + 4 >= ranCit.End Then GoTo Loop1
                Dim ranCitation As Word.Range

                ranCitation = ranCit.Duplicate
                With ranCitation.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Text = "" : .Style = sCitStyle
                End With
                Dim bYearOnly As Boolean
                ranCitation.Select()
                Do While ranCitation.Find.Execute = True
                    Dim sComment As String = ""
                    ranCitation.Select()

                    Dim PreSTyle As Word.Style = oActApp.Selection.Previous.Characters(1).Style
                    If PreSTyle.NameLocal = "BibXref_online" Then Exit Do
                    ''oActApp.Selection.Range
                    If oActApp.Selection.Text = PrevCitationText Then
                        sComment = "Duplication Citation Found. Please check manually"
                    End If
                    If Regex.IsMatch(oActApp.Selection.Text, "^(18|19|20)[0-9]{2}[a-z]?") Then
                        bYearOnly = True
                        ''sComment = "Unable to sorting due to some citation didn't have author information. Please do manually."
                    End If
                    If oActApp.Selection.Characters.Count() = 1 And PrevCitationText <> "" Then
                        PrevCitationText = PrevCitationText
                        ''sComment = "Unable to sorting due to citaion didn't have author and year infromation. Please do manually."
                    ElseIf bYearOnly = True Then

                    Else
                        countGroup = countGroup + 1
                        If bYearOnly = False Then PrevCitationText = oActApp.Selection.Text
                    End If

                    Dim strYEar As String
                    If Regex.IsMatch(oActApp.Selection.Text, "(18|19|20)[0-9]{2}[a-z]?") Then
                        If oActApp.Selection.Characters.Count() = 1 Then
                            strYEar = Regex.Match(oActApp.Selection.Text, "((18|19|20)[0-9]{2}[a-z]?)").Groups(1).Value & oActApp.Selection.Text
                        Else
                            strYEar = Regex.Match(oActApp.Selection.Text, "((18|19|20)[0-9]{2}[a-z]?)").Groups(1).Value
                        End If
                    Else
                        strYEar = ""
                    End If
                    Dim lstCitStr As New List(Of String)
                    If bYearOnly = True Then
                        lstCitStr.Add(Regex.Replace(PrevCitationText, "(18|19|20)[0-9]{2}[a-z]?", strYEar))
                    Else
                        lstCitStr.Add(PrevCitationText)
                    End If
                    lstCitStr.Add(strYEar)
                    lstCitStr.Add(sComment)
                    lstCitStr.Add(countGroup)
                    If Not dictEachCitation.ContainsKey(oActApp.Selection.Range) Then
                        dictEachCitation.Add(oActApp.Selection.Range, lstCitStr)
                    End If
                    ranCitation.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    If ranCitation.End >= ranCit.End Then Exit Do
                    If oActApp.Selection.Characters.Count() = 1 Or bYearOnly = True Then
                    Else
                        PrevCitationText = oActApp.Selection.Text
                    End If
                    bYearOnly = False
                Loop
Loop1:
                If dictEachCitation.Count > 1 Then
                    Dim sorted = From pair In dictEachCitation Order By pair.Value.Item(0).ToString()
                    dictEachCitation = sorted.ToDictionary(Function(y) y.Key, Function(x) x.Value)

                    Dim dupCit = From pair In dictEachCitation Order By pair.Value.Item(1).ToString
                    dictEachCitation = dupCit.ToDictionary(Function(p) p.Key, Function(p) p.Value)
                    Dim groupCitaion = dictEachCitation.GroupBy(Function(x) x.Value.Item(3).ToString())

                    ''dictEachCitation.OrderBy(Function(x) x.Key.Text)
                    ''dictEachCitation.OrderBy(Function(x) x.Value.ToString())
                End If
                If (dictEachCitation.Count > 1 And Not dictHavCitInfo.ContainsKey(ranCit)) Then
                    dictHavCitInfo.Add(ranCit, dictEachCitation)
                End If

            Next

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Public Function CollectRefInfo(wDoc As Word.Document, refStyleName As String) As Boolean
        Try
            Const sTrimPattern As String = "([\[\]\(\)\t\s\.\:]+)"
            Dim fStyle() = {"‡ref_auSurname", "‡ref_auCollab"}
            Dim ranDoc As Word.Range : Dim rCount As Integer
            ranDoc = wDoc.Content
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = refStyleName
            End With
            Do While ranDoc.Find.Execute = True
                rCount = rCount + 1
                Dim AuthorList As New List(Of String)
                Dim dictAuthor As New Dictionary(Of Integer, String)
                Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range
                ranDoc.Select()
                ranDupRef = oActApp.Selection.Range
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''Validation
                Dim ranVal As Word.Range
                ranVal = ranDoc.Duplicate
                If Regex.IsMatch(ranVal.Text, "^(—)+") Then
                    MessageBox.Show(" Em dashes found instead of author names! Please change em dashes into author names and try again!!", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
                Dim skipStyle() = {"‡ref_auPrefix", "‡ref_edPrefix", "‡ref_transPrefix", "‡ref_transedPrefix", "‡ref_assigneePrefix", "‡ref_compilerPrefix", "‡ref_directorPrefix", "‡ref_guestedPrefix", "‡ref_inventorPrefix"}
                For i = LBound(skipStyle) To UBound(skipStyle)
                    If modCEGUtility.AutoStyleExists(skipStyle(i), wDoc) = True Then
                        ranVal = ranDoc.Duplicate
                        With ranVal.Find
                            .ClearFormatting() : .Replacement.ClearFormatting()
                            .Text = "" : .Style = skipStyle(i)
                        End With
                        If ranVal.Find.Execute = True Then
                            MessageBox.Show(skipStyle(i) & " style found! Please remove author prefix style and try again!!", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            Return False
                        End If
                    End If
                Next
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                For i = LBound(fStyle) To UBound(fStyle)
                    ranAuthor = ranDupRef.Duplicate
                    With ranAuthor.Find
                        .ClearFormatting() : .Replacement.ClearFormatting()
                        .Text = "" : .Style = fStyle(i)
                    End With
                    Do While ranAuthor.Find.Execute
                        isFoundSurname = True
                        'If ranAuthor.Previous.Words.Count > 0 Then
                        '    Dim ranSuffix As Word.Range
                        '    ranAuthor.Previous.Words(1).Select()
                        '    ranSuffix = oActApp.Selection.Range
                        '    With ranSuffix.Find
                        '        .ClearFormatting() : .Replacement.ClearFormatting()
                        '        .Text = "" : .Style = "‡ref_auPrefix"
                        '    End With
                        '    If ranSuffix.Find.Execute Then
                        '        ranSuffix.Select()
                        '    End If
                        'End If
                        ranAuthor.Select()
                        If Not dictAuthor.ContainsKey(ranAuthor.Start) Then
                            dictAuthor.Add(ranAuthor.Start, ranAuthor.Text)
                        End If
                        If ranAuthor.End > ranDupRef.End Then Exit Do
                        ranAuthor = wDoc.Range(ranAuthor.End + 1, ranDupRef.End)
                        ranAuthor.Find.Text = ""
                        ranAuthor.Find.Style = fStyle(i)
                    Loop
                Next

                Dim ranRef As Word.Range
                Dim sYear As String = ""
                Dim objRefInfo As New clsRefInfo
                ranRef = ranDoc.Duplicate
                With ranRef.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Text = "" : .Style = "‡ref_pubdateYear"
                End With
                If ranRef.Find.Execute = True Then
                    sYear = ranRef.Text
                End If
                '‡ref_number
                ranRef = ranDoc.Duplicate
                With ranRef.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Text = "" : .Style = "‡ref_number"
                End With
                If ranRef.Find.Execute = True Then
                    objRefInfo.oRefLabelRng = ranRef.Duplicate
                    objRefInfo.iRefNum = Regex.Replace(ranRef.Text, sTrimPattern, String.Empty, RegexOptions.IgnoreCase)
                    'MessageBox.Show("ERROR : Tool Not support vancouver style reference.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    'Return False
                    ''objRefInfo.iRefNum = ranRef.Text.Replace("[", "").Replace("]", "")
                Else
                    If modCEGUtility.AutoStyleExists("‡ref_label", wDoc) = True Then
                        ranRef.Find.Style = "‡ref_label"
                        If ranRef.Find.Execute = True Then
                            objRefInfo.oRefLabelRng = ranRef.Duplicate
                            objRefInfo.iRefNum = Regex.Replace(ranRef.Text, sTrimPattern, String.Empty, RegexOptions.IgnoreCase)
                            'MessageBox.Show("ERROR : Tool not support vancouver style reference.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                            'Return False
                        End If
                    End If
                End If
                ranRef = ranDoc.Duplicate
                With ranRef.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Text = "" : .Style = "‡ref_etal"
                End With
                If ranRef.Find.Execute = True Then
                    objRefInfo.sRefEtal = True
                Else
                    objRefInfo.sRefEtal = False
                End If



                If refStyleName = "‡ref_auCollab" Then
                    objRefInfo.sRefCollab = True
                Else
                    objRefInfo.sRefCollab = False
                End If


                If isFoundSurname = True Then
                    dictAuthor.OrderBy(Function(item) item.Key)
                    Dim val As String
                    Dim sCondText As String = ""
                    For Each val In dictAuthor.Values
                        AuthorList.Add(val)
                        sCondText = sCondText & Trim(val)
                    Next
                    objRefInfo.rCondText = sCondText & sYear
                Else
                    objRefInfo.rCondText = ""
                End If
                objRefInfo.oRefRng = ranDoc.Duplicate
                objRefInfo.cText = Regex.Replace(ranDoc.Text, "\n[\n\s\r\t]*", String.Empty, RegexOptions.IgnoreCase) ''ranDoc.Text.Replace(ChrW(13), "")
                objRefInfo.sRefYear = sYear
                objRefInfo.olRefAuthors = AuthorList
                If (dictRefInfo.ContainsKey(rCount)) = False Then
                    dictRefInfo.Add(rCount, objRefInfo)
                End If
                If ranDoc.End >= wDoc.Range.End Then Exit Do
                ranDoc = wDoc.Range(ranDupRef.End + 1, wDoc.Range.End)
                ranDoc.Find.Text = ""
                ranDoc.Find.Style = refStyleName
                dictAuthor = Nothing
            Loop
            Return True

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
    End Function
    Public Function RefSwapValidation() As Boolean
        Try
            If dictRefInfo.Count > 0 Then
                Dim refNumber As Integer
                If dictRefInfo.Select(Function(x) x.Value.iRefNum).Count = 0 Then
                    MessageBox.Show("Reference numbers not found in the document.")
                    RefSwapValidation = False
                End If
                For Each pair As KeyValuePair(Of Integer, clsRefInfo) In dictRefInfo
                    If refNumber <> 0 Then
                        If refNumber + 1 <> pair.Value.iRefNum Then
                            MessageBox.Show("Reference numbers are not sequence.")
                            RefSwapValidation = False
                        End If
                    End If
                    refNumber = pair.Value.iRefNum
                Next
                Dim disRef = dictRefInfo.GroupBy(Function(x) x.Value.oRefRng.Text).Where(Function(x) x.Count() > 1).ToList()
                If disRef.Count > 0 Then
                    For Each rm In disRef
                        MessageBox.Show("Duplicate reference :" + rm.Key)
                    Next
                    RefSwapValidation = False
                End If
                'Dim ts = dictRefInfo.GroupBy(Function(x) x.Value.rCondText).Where(Function(x) x.Count() > 1).ToList()
                'If ts.Count > 0 Then

                '    RefSwapValidation = False
                'End If
            Else
                MessageBox.Show("Reference Not found in the document.")
                RefSwapValidation = False
            End If
            If dictHavCitaion.Count > 0 Then

            End If
            RefSwapValidation = True
        Catch ex As Exception
            MessageBox.Show("Error: RefSwapValidation:")
        End Try
    End Function
    Public Function ToConvertHavardCitation(refTextCitationList As List(Of String))
        Try

            For Each pair As KeyValuePair(Of Word.Range, String) In ModRefUtility.dictCitationInfo
                Dim sHavCitationText As String = ""
                If Not pair.Value.Contains(".") And Not Regex.IsMatch("\[-", pair.Value) Then
                    Dim citList As ArrayList = ExpandCitationText(pair.Value)
                    If citList.Count > 0 Then
                        sHavCitationText = GetHavardCitation(citList, refTextCitationList)
                        'For Each sCitation In citList
                        '    sHavCitationText = sHavCitationText & ConvertHavardCitation(sCitation) & " "
                        'Next
                    End If
                    Select Case (cBracketType)
                        Case "()"
                            If sHavCitationText.Contains("(") Then
                            Else
                                sHavCitationText = "(" & sHavCitationText & ")"
                            End If
                        Case "[]"
                            sHavCitationText = "[" & sHavCitationText & "]"
                        Case "{};"
                            sHavCitationText = "{" & sHavCitationText & "}"
                    End Select
                    'dictHavCitaion.Add(pair.Key, sHavCitationText)
                    dictConvertedCitation.Add(pair.Key, sHavCitationText)
                    sHavCitationText = ""
                Else
                    dictHavCitaion.Add(pair.Key, pair.Value)
                End If
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function

    Private Function ToApplyReferenceFormat(strStyleformat As String, wDoc As Word.Document)
        Try
            Dim I As Integer
            Dim oRng As Word.Range
            Dim strRefStyle As String
            Dim strRefFormat As String
            strRefStyle = Split(strStyleformat, "@")(0)
            strRefFormat = Split(strStyleformat, "@")(1)
            If Not String.IsNullOrEmpty(strRefStyle) And Not String.IsNullOrEmpty(strRefFormat) Then
                If modCEGUtility.AutoStyleExists(strRefStyle, wDoc) = True Then
                    For I = 1 To 3
                        Select Case I
                            Case 1 : oRng = wDoc.Content.Duplicate
                            Case 2 : If wDoc.Footnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdFootnotesStory).Duplicate Else oRng = Nothing
                            Case 3 : If wDoc.Endnotes.Count > 0 Then oRng = wDoc.StoryRanges(Word.WdStoryType.wdEndnotesStory).Duplicate Else oRng = Nothing
                        End Select
                        If Not IsNothing(oRng) Then
                            Dim ranDoc As Word.Range
                            ranDoc = oRng.Duplicate
                            ranDoc.Find.ClearFormatting()
                            ranDoc.Find.Style = strRefStyle
                            ranDoc.Find.Text = ""
                            Do While ranDoc.Find.Execute
                                ranDoc.Select()
                                Select Case UCase(strRefFormat)
                                    Case "ROMAN"
                                        oActApp.Selection.Font.Bold = False
                                        oActApp.Selection.Font.Italic = False
                                    Case "ITALIC"
                                        oActApp.Selection.Font.Italic = True
                                        oActApp.Selection.Font.Bold = False
                                    Case "BOLD"
                                        oActApp.Selection.Font.Italic = False
                                        oActApp.Selection.Font.Bold = True
                                    Case "BOLDITALIC"
                                        oActApp.Selection.Font.Italic = True
                                        oActApp.Selection.Font.Bold = True
                                End Select
                                ranDoc.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                            Loop
                        End If
                    Next
                Else
                    MessageBox.Show("Style not found in the document.")
                End If
            Else
                    MessageBox.Show("Please check the configuration.")
            End If


        Catch ex As Exception

        End Try
    End Function

    'Private Function ToCollectRefInfo(wDoc As Word.Document, refStyleName As String) As Boolean
    '    Try
    '        Dim oNewDoc As Word.Document
    '        Dim ranDoc As Word.Range : Dim rCount As Integer
    '        ranDoc = wDoc.Content
    '        With ranDoc.Find
    '            .ClearFormatting() : .Replacement.ClearFormatting()
    '            .Text = "" : .Replacement.Text = "" : .Style = refStyleName
    '        End With
    '        oNewDoc = wDoc.Application.Documents.Add()
    '        Do While ranDoc.Find.Execute = True
    '            rCount = rCount + 1 : Dim AuthorList As New List(Of String) : Dim dictAuthor As New Dictionary(Of Integer, String)
    '            Dim isFoundSurname As Boolean : Dim ranAuthor As Word.Range : Dim ranDupRef As Word.Range


    '            If String.IsNullOrEmpty(sLabelStyle) = False Then
    '                With oFindlblRng.Find
    '                    .ClearFormatting() : .Replacement.ClearFormatting() : .ClearAllFuzzyOptions()
    '                    .Text = "" : .Replacement.Text = "" : .MatchWildcards = False : .MatchWholeWord = False : .MatchCase = False
    '                    .Style = sLabelStyle
    '                End With
    '                If oFindlblRng.Find.Execute = True AndAlso oFindlblRng.InRange(oFindRng) Then
    '                    Call oTempRng.SetRange(oFindlblRng.Start, oFindlblRng.End)
    '                    Call oFindRng.SetRange(oFindlblRng.End, oFindRng.End)
    '                    oclsRef.sRefLabel = oTempRng.Text
    '                Else
    '                    oclsRef.sRefLabel = "#"
    '                End If
    '            Else
    '                oclsRef.sRefLabel = "#"
    '            End If
    '            oTempRng = oFindRng.Duplicate : Dim oclsRef As New clsRef
    '        Loop
    '    Catch ex As Exception
    '    End Try
    'End Function
    'Public Function ConvertHavardCitation(sCitation As String) As String
    '    Dim sCitPattern As String = cTextPattern
    '    Try
    '        Dim objRefInfo As clsRefInfo
    '        If dictRefInfo.ContainsKey(Convert.ToInt32(sCitation)) Then
    '            dictRefInfo.TryGetValue(Convert.ToInt32(sCitation), objRefInfo)
    '            If objRefInfo.rAuthor.Count > 0 Then
    '                If objRefInfo.rCollab = True And cTextPattern.Contains("et al") Then
    '                    MessageBox.Show("write code here")
    '                Else
    '                    sCitPattern = sCitPattern.Replace("<Author>", objRefInfo.rAuthor.Item(0).ToString())
    '                End If
    '                sCitPattern = sCitPattern.Replace("<Year>", objRefInfo.rYear)
    '                'Author year,
    '            Else
    '                MessageBox.Show("djfslfjsklfjslkfjlskdfj")
    '            End If
    '        Else
    '            MessageBox.Show("need to write code")
    '        End If
    '        ConvertHavardCitation = sCitPattern
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Function
End Module
