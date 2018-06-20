Imports Microsoft.Office.Interop.Word
Imports Word = Microsoft.Office.Interop.Word
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.ComponentModel
<ComClass(clsTransformIndexTerm.ClassId, clsTransformIndexTerm.InterfaceId, clsTransformIndexTerm.EventsId)> _
Public Class clsTransformIndexTerm

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class and its COM interfaces. If you change them, existing clients will no longer be able to access the class.
    Public Const ClassId As String = "6DB79AF2-F662-44AC-8458-62B06BFDD9E4"
    Public Const InterfaceId As String = "EDED909C-9371-4670-BA32-109AE917B1D7"
    Public Const EventsId As String = "17C731B8-CE61-5B5F-B114-10F3E46153AC"
#End Region



    Private w As XNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    Private r As XName = w + "r"
    Private ins As XName = w + "ins"
    Const sMsgTitle As String = "Index Transform"

    Public Function ToReviewIndex(oMasterDoc As Word.Document) As Boolean
        Try
            Dim oRevView As WdRevisionsView : Dim oViewType As WdViewType
            Dim bTrack As Boolean = oMasterDoc.TrackRevisions
            Dim bShowFieldCodes As Boolean = oMasterDoc.ActiveWindow.View.ShowFieldCodes
            oRevView = oMasterDoc.ActiveWindow.View.RevisionsView : oViewType = oMasterDoc.ActiveWindow.View.Type
            oMasterDoc.TrackRevisions = False
            With oMasterDoc.ActiveWindow.View
                .Type = WdViewType.wdPrintView
                .ShowRevisionsAndComments = True : .RevisionsView = WdRevisionsView.wdRevisionsViewFinal
                .ShowAll = False : .ShowFieldCodes = True
                .ShowHiddenText = True
            End With
            oMasterDoc.Bookmarks.ShowHidden = False : oMasterDoc.Bookmarks.ShowHidden = True
            Dim B As Integer = oMasterDoc.Bookmarks.Count : Dim I As Integer = 0
            Dim oBKList As New Dictionary(Of String, String) : Dim sBKName As String = String.Empty
            For Each oBK As Bookmark In oMasterDoc.Bookmarks
                sBKName = oBK.Name
                If sBKName.ToLower.Contains("cegindex") = True Then
                    If oBKList.ContainsKey(sBKName) = False AndAlso Not oBK.Range.Text Is Nothing AndAlso oBK.Range.Text.Length > 2 Then
                        Dim oIndexTextRng As Range = oBK.Range.Duplicate
                        If oIndexTextRng.Characters.First.Text = ChrW(19) OrElse Regex.IsMatch(oIndexTextRng.Characters.First.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                            oIndexTextRng.SetRange(oIndexTextRng.Start + 1, oIndexTextRng.End)
                        End If
                        If oIndexTextRng.Characters.Last.Text = ChrW(21) OrElse Regex.IsMatch(oIndexTextRng.Characters.Last.Text, "[a-z0-9\s]", RegexOptions.IgnoreCase) = False Then
                            Call oIndexTextRng.SetRange(oIndexTextRng.Start, oIndexTextRng.End - 1)
                        End If
                        oIndexTextRng.Font.Shading.BackgroundPatternColorIndex = WdColorIndex.wdGray25
                        oBKList.Add(sBKName, oIndexTextRng.Text.Trim.Replace(ChrW(21), "").Replace(ChrW(9), "").Trim)
                    End If

                End If
                I += 1 : oMasterDoc.Application.StatusBar = "Collecting index fields in changed context (" & I & "/" & B & ")"
            Next
            Dim ofrmReview As New frmReviewIndex
            ofrmReview.oBKList = oBKList
            ofrmReview.oActDoc = oMasterDoc : ofrmReview.oActApp = oMasterDoc.Application
            If ofrmReview.ShowDialog = DialogResult.Cancel Then
                oMasterDoc.TrackRevisions = bTrack
                With oMasterDoc.ActiveWindow.View
                    .Type = oViewType
                    .ShowRevisionsAndComments = False : .RevisionsView = WdRevisionsView.wdRevisionsViewFinal
                    .ShowAll = False : .ShowFieldCodes = True
                    .ShowHiddenText = True
                End With
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Public Function ToTransformIndex()
        Try
            Dim ofrm As New frmIndexTransfer
            ofrm.Show()
        Catch ex As Exception

        End Try
    End Function

    Public Function ToTransformIndexTerm(sMasterDocPath As String, sIndexDocPath As String, Optional BW As BackgroundWorker = Nothing) As Boolean
        Dim oMasterDoc As WordprocessingDocument = WordprocessingDocument.Open(sMasterDocPath, True)
        Dim oIndexDoc As WordprocessingDocument = WordprocessingDocument.Open(sIndexDocPath, False)

        Try

            Dim sMasterDocXML As String = oMasterDoc.MainDocumentPart.Document.Body.InnerXml
            sMasterDocXML = Regex.Replace(sMasterDocXML, "(<w:moveFrom [^>]*>.*?<\/w:moveFrom>|<w:instrText[^>]*>.*?<\/w:instrText>|<w:delInstrText[^>]*>.*?<\/w:delInstrText>|\<w:ins [^>]*\>.*?\<\/w:ins\>)", "<!-- $1 -->", RegexOptions.IgnoreCase)
            ''sMasterDocXML = Regex.Replace(sMasterDocXML, "(<w:ins[^>]+?/>|\<w:ins [^>]*\>.*?\<\/w:ins[^>]*\>)", "<!-- $1 -->", RegexOptions.IgnoreCase)
            sMasterDocXML = Regex.Replace(sMasterDocXML, "(</w:p>)", "<w:r><w:t>|</w:t></w:r>$1", RegexOptions.IgnoreCase)
            oMasterDoc.MainDocumentPart.Document.Body.InnerXml = sMasterDocXML


            Dim sIndexDocXML As String = oIndexDoc.MainDocumentPart.Document.Body.InnerXml
            If sIndexDocXML.Contains("<w:ins[^>]+?") = True OrElse sIndexDocXML.Contains("<w:del[^>]+?") = True Then
                MessageBox.Show("Index manuscript contains few track changes information." & Environment.NewLine & "We can't able to continue the process with trackchanges in index manuscript." & Environment.NewLine & "So, could you please accept the revision in index manuscript and try again.", sMsgTitle, MessageBoxButtons.OK)
                Exit Function
            End If

            sIndexDocXML = Regex.Replace(sIndexDocXML, "(<w:instrText[^>]*>.*?<\/w:instrText>)", "<!-- $1 -->", RegexOptions.IgnoreCase)
            sIndexDocXML = Regex.Replace(sIndexDocXML, "(</w:p>)", "<w:r><w:t>|</w:t></w:r>$1", RegexOptions.IgnoreCase)
            oIndexDoc.MainDocumentPart.Document.Body.InnerXml = sIndexDocXML

            'oMasterDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oMasterDoc.MainDocumentPart.Document.Body.InnerXml, "(\<w:ins[^>]*\>.*?\<\/w:ins[^>]*\>)", "<!-- $1 -->", RegexOptions.IgnoreCase)
            'oIndexDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oIndexDoc.MainDocumentPart.Document.Body.InnerXml, "(\<w:ins[^>]*\>.*?\<\/w:ins[^>]*\>)", "<!-- $1 -->", RegexOptions.IgnoreCase)


            ''ToRemoveEmptyParagraphs(oMasterDoc.MainDocumentPart.Document.Body)
            ''ToRemoveEmptyParagraphs(oIndexDoc.MainDocumentPart.Document.Body)
            Dim oMasterParas As IEnumerable(Of OpenXmlElement) = oMasterDoc.MainDocumentPart.Document.Body.Descendants(Of Wordprocessing.Paragraph)()
            Dim oIndexParas As IEnumerable(Of OpenXmlElement) = oIndexDoc.MainDocumentPart.Document.Body.Descendants(Of Wordprocessing.Paragraph)()
            Dim iBKID As Integer = 0
            Dim oSW As New StreamWriter(Path.ChangeExtension(sIndexDocPath, ".txt")) : oSW.Write(oIndexDoc.MainDocumentPart.Document.Body.InnerText.ToLower.Replace("|", Environment.NewLine)) : oSW.Flush() : oSW.Close() : oSW.Dispose()
            oSW = New StreamWriter(Path.ChangeExtension(sMasterDocPath, ".txt")) : oSW.Write(oMasterDoc.MainDocumentPart.Document.Body.InnerText.ToLower.Replace("|", Environment.NewLine)) : oSW.Flush() : oSW.Close() : oSW.Dispose()

            sIndexDocXML = Regex.Replace(sIndexDocXML, "(<w:r><w:t>\|</w:t></w:r>)(</w:p>)", "$2", RegexOptions.IgnoreCase) : oIndexDoc.MainDocumentPart.Document.Body.InnerXml = sIndexDocXML
            sMasterDocXML = Regex.Replace(sMasterDocXML, "(<w:r><w:t>\|</w:t></w:r>)(</w:p>)", "$2", RegexOptions.IgnoreCase) : oMasterDoc.MainDocumentPart.Document.Body.InnerXml = sMasterDocXML
            If oIndexDoc.MainDocumentPart.Document.Body.InnerText.Length = oMasterDoc.MainDocumentPart.Document.Body.InnerText.Length AndAlso oMasterParas.Count = oIndexParas.Count Then
                Dim oSDIndexList As New Dictionary(Of String, List(Of OpenXmlElement))
                Dim oSDParaIDList As New Dictionary(Of Integer, Dictionary(Of String, List(Of OpenXmlElement)))
                Dim Z As Integer = oIndexParas.Count - 1
                If BW Is Nothing = False Then BW.ReportProgress(10)
                For P As Long = 0 To oIndexParas.Count - 1
                    Dim oIndexP As OpenXmlElement = oIndexParas(P)
                    If oIndexP.InnerXml.Contains("w:fldCharType=""begin""") = True Then
                        oSDIndexList = New Dictionary(Of String, List(Of OpenXmlElement))
                        Call ToCollectIndexTerm(oIndexP, oSDIndexList)
                        Call oSDParaIDList.Add(P, oSDIndexList)
                    End If
                Next
                If BW Is Nothing = False Then BW.ReportProgress(40)
                For Each oKV As KeyValuePair(Of Integer, Dictionary(Of String, List(Of OpenXmlElement))) In oSDParaIDList
                    Dim oMasterPara As OpenXmlElement = oMasterParas(oKV.Key)
                    Call ToTransferIndexTermWithinPara(oMasterPara, iBKID, oKV.Value)
                Next
                If BW Is Nothing = False Then BW.ReportProgress(70)
                '############# Transform the bookmark ############
                oMasterDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oMasterDoc.MainDocumentPart.Document.Body.InnerXml, "(\<!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                oIndexDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oIndexDoc.MainDocumentPart.Document.Body.InnerXml, "(\<!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                oMasterDoc.Close() : oMasterDoc = Nothing
                oIndexDoc.Close() : oIndexDoc = Nothing
                Dim oWordApp As New Word.Application
                Dim oIndexWordDoc As Word.Document = oWordApp.Documents.Open(sIndexDocPath, False, False, False, Visible:=False)
                Dim oOrgWordDoc As Word.Document = oWordApp.Documents.Open(sMasterDocPath, False, False, False, Visible:=False)
                Try
                    Call ToTransferBookmarks(oOrgWordDoc, oIndexWordDoc, False, True)
                    If BW Is Nothing = False Then BW.ReportProgress(90) : ToTransformIndexTerm = True
                Catch ex As Exception
                Finally
                    oIndexWordDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                    oOrgWordDoc.Close(WdSaveOptions.wdSaveChanges)
                    Call oWordApp.Quit()
                End Try
                If BW Is Nothing = False Then BW.ReportProgress(100)
            Else
                MessageBox.Show("Edited and Index manuscript characters count was not matched." & Environment.NewLine & "Please check and try again.", sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Information)
                oMasterDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oMasterDoc.MainDocumentPart.Document.Body.InnerXml, "(\<!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                oIndexDoc.MainDocumentPart.Document.Body.InnerXml = Regex.Replace(oIndexDoc.MainDocumentPart.Document.Body.InnerXml, "(\<!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                oMasterDoc.Close() : oMasterDoc = Nothing
                oIndexDoc.Close() : oIndexDoc = Nothing
                Exit Function
            End If
            Return True
        Catch ex As Exception
            Call MessageBox.Show("Unable to complete the process" & Environment.NewLine & ex.Message, sMsgTitle, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        Finally
            Try
                If Not (IsNothing(oMasterDoc)) Then oMasterDoc.Close()
                If Not (IsNothing(oIndexDoc)) Then oIndexDoc.Close()
            Catch ex As Exception
                ex.Data.Clear()
            End Try

        End Try
    End Function

    Protected Sub ToRemoveEmptyParagraphs(ByRef body As DocumentFormat.OpenXml.Wordprocessing.Body)
        Dim colP As IEnumerable(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph) = body.Descendants(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph)()

        Dim count As Integer = colP.Count
        For Each p As DocumentFormat.OpenXml.Wordprocessing.Paragraph In colP
            If (p.InnerText.Trim() = String.Empty) Then
                body.RemoveChild(Of DocumentFormat.OpenXml.Wordprocessing.Paragraph)(p)
            End If
        Next
    End Sub

    Private Function ToTransferIndexTermWithinPara(oMasterPara As OpenXmlElement, ByRef iBKID As Integer, oSDIndexList As Dictionary(Of String, List(Of OpenXmlElement))) As Boolean
        Dim oWRList As IEnumerable(Of OpenXmlElement) = oMasterPara.Descendants(Of Run)()
        Dim bCollectField As Boolean : Dim sFieldInfo As String = String.Empty : Dim sTextInfo As String = String.Empty
        Dim oSpAttr As New OpenXmlAttribute("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve")
        Dim L As Long : Dim lIndexStart As Long : Dim K As Long
        Try
            For Each oKV As KeyValuePair(Of String, List(Of OpenXmlElement)) In oSDIndexList.Reverse
                L = 0 : K = CInt(oKV.Key.Split(":")(0))
                Dim bIsTrackContext As Boolean = ToCheckIsTrackContext(oMasterPara, K)
                Dim oFldWRList As List(Of OpenXmlElement) = oKV.Value : Dim bInsertElement As Boolean = False
                For Each oWR As OpenXmlElement In oWRList
                    L = oWR.InnerText.Length
                    Dim oBKStart As New BookmarkStart() : Dim oBKEnd As New BookmarkEnd()
                    If K = 0 Then
                        iBKID += 1 : oBKStart.Name = "CEGIndex" & iBKID : oBKStart.Id = iBKID
                        If bIsTrackContext = True Then ToInsertStartBookMark(oMasterPara, oWR, oBKStart)
                        For Each oFldWR In oFldWRList
                            If oWR.Parent.LocalName = "del" Then oWR = oWR.Parent
                            Call oMasterPara.InsertBefore(Of OpenXmlElement)(oFldWR, oWR)
                        Next
                        If bIsTrackContext = True Then ToInsertEndBookMark(oMasterPara, oWR, oBKStart.Id, oBKEnd)
                        bInsertElement = True
                        Exit For
                    ElseIf K < L Then
                        iBKID += 1 : oBKStart.Name = "CEGIndex" & iBKID : oBKStart.Id = iBKID
                        Dim sRun1Text As String = String.Empty : Dim sRun2Text As String = String.Empty
                        For Each oWT As OpenXmlElement In oWR.Descendants  'oWR.Descendants(Of Text)()

                            Dim bIsDelText As Boolean = False : Dim bIsText As Boolean = False
                            Dim oDel As OpenXmlElement
                            Select Case True
                                Case oWT.LocalName = "t" : bIsText = True
                                Case oWT.LocalName = "delText" : bIsDelText = True
                                Case Else : bIsText = False : bIsDelText = False
                            End Select
                            If bIsText = True OrElse bIsDelText = True Then
                                If K < oWT.InnerText.Length Then
                                    Dim oNewWT1 As OpenXmlElement : Dim oNewWT2 As OpenXmlElement : Dim oDupRun1 As OpenXmlElement
                                    If bIsText = True Then
                                        sRun1Text = oWT.InnerText.Substring(0, K) : sRun2Text = oWT.InnerText.Substring(K)
                                        oWT.Remove() : oNewWT1 = New Text(sRun1Text) : oDupRun1 = oWR.Clone()
                                        oNewWT1.SetAttribute(oSpAttr) : oWR.AppendChild(Of OpenXmlElement)(oNewWT1)
                                        If bIsTrackContext = True Then ToInsertStartBookMark(oMasterPara, oWR, oBKStart)
                                        For Each oFldWR In oFldWRList
                                            Call oMasterPara.InsertAfter(Of OpenXmlElement)(oFldWR, oWR)
                                            oWR = oFldWR
                                        Next
                                        If bIsTrackContext = True Then ToInsertEndBookMark(oMasterPara, oWR, oBKStart.Id, oBKEnd)
                                        oNewWT2 = New Text(sRun2Text) : oNewWT2.SetAttribute(oSpAttr) : oDupRun1.AppendChild(Of OpenXmlElement)(oNewWT2)
                                        Call oMasterPara.InsertAfter(Of OpenXmlElement)(oDupRun1, oWR)
                                        bInsertElement = True
                                        Exit For
                                    ElseIf bIsDelText = True Then
                                        sRun1Text = oWT.InnerText.Substring(0, K) : sRun2Text = oWT.InnerText.Substring(K)
                                        oWT.Remove() : oNewWT1 = New DeletedText(sRun1Text) : oDupRun1 = oWR.Parent.Clone()
                                        oNewWT1.SetAttribute(oSpAttr) : oWR.AppendChild(Of OpenXmlElement)(oNewWT1)
                                        Dim bIsNotFirst As Boolean = False
                                        If bIsTrackContext = True Then ToInsertStartBookMark(oMasterPara, oWR, oBKStart)
                                        For Each oFldWR In oFldWRList
                                            Dim oRPr As New RunProperties(New Highlight() With {.Val = HighlightColorValues.Red})
                                            Call oFldWR.PrependChild(Of OpenXmlElement)(oRPr)
                                            Select Case bIsNotFirst
                                                Case True 'Do nothing
                                                Case False
                                                    If oWR.Parent.LocalName <> "del" AndAlso oWR.Parent.LocalName <> "p" Then
                                                        oWR = oWR.Parent
                                                    End If
                                                    bIsNotFirst = True
                                            End Select
                                            Call oMasterPara.InsertAfter(Of OpenXmlElement)(oFldWR, oWR)
                                            oWR = oFldWR
                                        Next
                                        If bIsTrackContext = True Then ToInsertEndBookMark(oMasterPara, oWR, oBKStart.Id, oBKEnd)
                                        oNewWT2 = New DeletedText(sRun2Text)
                                        oNewWT2.SetAttribute(oSpAttr)
                                        oDupRun1.ChildElements(0).AppendChild(Of OpenXmlElement)(oNewWT2)
                                        Call oMasterPara.InsertAfter(Of OpenXmlElement)(oDupRun1, oWR)
                                        bInsertElement = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    ElseIf K >= L Then
                        K = K - L
                    End If
                    If bInsertElement = True Then Exit For
                Next
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Function

    Private Function ToCheckIsTrackContext(oMasterPara As OpenXmlElement, K As Integer) As Boolean
        Try
            Dim sTestTrackText As String = Regex.Replace(oMasterPara.InnerXml, "(\<!--\s|\s--\>)", "", RegexOptions.IgnoreCase)
            sTestTrackText = Regex.Replace(sTestTrackText, "(\<w:ins\s[^>]*\>|<w:delText[^>]*\>)", "<w:t>@@", RegexOptions.IgnoreCase)
            sTestTrackText = Regex.Replace(sTestTrackText, "(\<\/w:ins\>|<\/w:delText[^>]*\>)", "@@</w:t>", RegexOptions.IgnoreCase)
            sTestTrackText = Regex.Replace(sTestTrackText, "(<\/?[^>]*>)", String.Empty, RegexOptions.IgnoreCase)
            Dim sContextText As String = String.Empty
            Dim L As Long = sTestTrackText.Length
            If L > 50 Then
                Select Case True
                    Case K = 0 : ToCheckIsTrackContext = sTestTrackText.Substring(0, 25).Contains("@@") ' Starting of para index
                    Case K < 30 AndAlso L > 50 : ToCheckIsTrackContext = sTestTrackText.Substring(0, 50).Contains("@@")
                    Case K + 30 < L : ToCheckIsTrackContext = sTestTrackText.Substring(K - 25, 50).Contains("@@")
                    Case K < L : ToCheckIsTrackContext = sTestTrackText.Substring(K - 25).Contains("@@")
                    Case Else : ToCheckIsTrackContext = sTestTrackText.Contains("@@")
                End Select
            Else
                ToCheckIsTrackContext = sTestTrackText.Contains("@@")
            End If
            If ToCheckIsTrackContext = True Then
                'MessageBox.Show("check")
            End If
        Catch ex As Exception
            ex.Data.Clear()
        End Try
    End Function

    Private Function ToInsertStartBookMark(oMasterPara As OpenXmlElement, ByRef oWR As OpenXmlElement, oBKStart As BookmarkStart) As OpenXmlElement
        Try
            If oWR.Parent.LocalName = "del" Then oWR = oWR.Parent
            Call oMasterPara.InsertAfter(Of BookmarkStart)(oBKStart, oWR)
            oWR = oBKStart
        Catch ex As Exception
            ex.Data.Clear()
        End Try
    End Function

    Private Function ToInsertEndBookMark(oMasterPara As OpenXmlElement, ByRef oWR As OpenXmlElement, sBKID As String, oBKEnd As BookmarkEnd) As OpenXmlElement
        Try
            oBKEnd.Id = sBKID
            If oWR.Parent.LocalName = "del" Then oWR = oWR.Parent
            Call oMasterPara.InsertAfter(Of BookmarkEnd)(oBKEnd, oWR)
            oWR = oBKEnd
        Catch ex As Exception
            ex.Data.Clear()
        End Try
    End Function

    Private Function ToCollectIndexTerm(oIndexPara As OpenXmlElement, oSDIndexList As Dictionary(Of String, List(Of OpenXmlElement))) As Boolean
        Try
            Dim oWRList As IEnumerable(Of OpenXmlElement) = oIndexPara.Descendants(Of Run)()
            Dim bCollectField As Boolean : Dim sFieldInfo As String : Dim sTextInfo As String : Dim L As Long : Dim lIndexStart As Long
            Dim oElementList As New List(Of OpenXmlElement)
            For Each oWR As OpenXmlElement In oWRList
                Select Case True
                    Case oWR.InnerXml.Contains("w:fldChar w:fldCharType=""begin""")
                        bCollectField = True : oElementList = New List(Of OpenXmlElement) : oElementList.Add(oWR.Clone)
                    Case oWR.InnerXml.Contains("w:fldChar w:fldCharType=""end""")
                        bCollectField = False : sFieldInfo = sFieldInfo & "<end>"
                        Call oElementList.Add(oWR.Clone)
                        Call oSDIndexList.Add(String.Concat(lIndexStart, ":", CStr(Rnd() * 100), ":", sFieldInfo), oElementList)
                        sTextInfo = sTextInfo & "###" : sFieldInfo = String.Empty : lIndexStart = 0
                    Case bCollectField
                        If String.IsNullOrEmpty(sFieldInfo) = True Then
                            lIndexStart = L : oWR.InnerXml = Regex.Replace(oWR.InnerXml, "(\<\!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                            Call oElementList.Add(oWR.Clone)
                            sFieldInfo = "<begin>" & oWR.InnerText
                        Else
                            oWR.InnerXml = Regex.Replace(oWR.InnerXml, "(\<\!--\s|\s--\>)", String.Empty, RegexOptions.IgnoreCase)
                            Call oElementList.Add(oWR.Clone)
                            sFieldInfo = sFieldInfo & oWR.InnerText
                        End If
                    Case Else
                        L = L + oWR.InnerText.Length
                        sTextInfo = sTextInfo & oWR.InnerText
                End Select
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function


    Private Function ToTransferBookmarks(oOrgDoc As Word.Document, oDupDoc As Word.Document, bInserFormattedText As Boolean, bIgnoreCount As Boolean) As Boolean
        On Error Resume Next
        Dim oRevView As WdRevisionsView : Dim oViewType As WdViewType
        Dim bShowFieldCodes As Boolean = oOrgDoc.ActiveWindow.View.ShowFieldCodes
        oRevView = oOrgDoc.ActiveWindow.View.RevisionsView : oViewType = oOrgDoc.ActiveWindow.View.Type
        '############ Original log file ##############
        With oOrgDoc.ActiveWindow.View
            .ShowRevisionsAndComments = False : .RevisionsView = WdRevisionsView.wdRevisionsViewOriginal
            .ShowAll = False : .ShowFieldCodes = False
            .ShowHiddenText = False
        End With

        '############ Duplicate log file ##############
        With oDupDoc.ActiveWindow.View
            .ShowRevisionsAndComments = False : .RevisionsView = WdRevisionsView.wdRevisionsViewOriginal
            .ShowAll = False : .ShowFieldCodes = False
            .ShowHiddenText = False
        End With

        If bIgnoreCount = True OrElse (oDupDoc.Characters.Count = oOrgDoc.Characters.Count AndAlso oDupDoc.Paragraphs.Count = oOrgDoc.Paragraphs.Count) OrElse (oDupDoc.Characters.Count - 1 = oOrgDoc.Characters.Count AndAlso oDupDoc.Paragraphs.Count - 1 = oOrgDoc.Paragraphs.Count) Then
            If bInserFormattedText = True AndAlso _
               oOrgDoc.Paragraphs.Count = oDupDoc.Paragraphs.Count AndAlso _
               oOrgDoc.Tables.Count = oDupDoc.Tables.Count AndAlso _
               oOrgDoc.Footnotes.Count = oDupDoc.Footnotes.Count AndAlso _
               oOrgDoc.Endnotes.Count = oDupDoc.Endnotes.Count Then
                oOrgDoc.Range.FormattedText = oDupDoc.Range.Duplicate
            ElseIf bIgnoreCount = True OrElse oDupDoc.Range.End = oOrgDoc.Range.End OrElse oDupDoc.Range.End - 1 = oOrgDoc.Range.End OrElse oDupDoc.Paragraphs.Count = oOrgDoc.Paragraphs.Count Then
                '####### Transfer bookmarks #############
                Dim oBK As Bookmark : Dim iI As Integer : Dim oBKRng As Range = Nothing : Dim oOrgBKRng As Range = Nothing
                oDupDoc.Bookmarks.ShowHidden = False : oDupDoc.Bookmarks.ShowHidden = True
                For iI = 1 To 3
                    Select Case iI
                        Case 1 : If oDupDoc.Footnotes.Count > 0 Then oBKRng = oDupDoc.StoryRanges(WdStoryType.wdFootnotesStory).Duplicate Else oBKRng = Nothing
                        Case 2 : If oDupDoc.Endnotes.Count > 0 Then oBKRng = oDupDoc.StoryRanges(WdStoryType.wdEndnotesStory).Duplicate Else oBKRng = Nothing
                        Case 3 : oBKRng = oDupDoc.Range.Duplicate
                    End Select
                    If Not oBKRng Is Nothing Then
                        For Each oBK In oBKRng.Bookmarks
                            Select Case iI
                                Case 1 : oOrgBKRng = oOrgDoc.StoryRanges(WdStoryType.wdFootnotesStory).Duplicate
                                Case 2 : oOrgBKRng = oOrgDoc.StoryRanges(WdStoryType.wdEndnotesStory).Duplicate
                                Case 3 : oOrgBKRng = oOrgDoc.Range.Duplicate
                            End Select
                            oOrgBKRng.SetRange(oBK.Range.Start, oBK.Range.End) : oOrgDoc.Bookmarks.Add(oBK.Name, oOrgBKRng) : oOrgDoc.UndoClear()
                        Next
                    End If
                Next
            Else
                Call TransferBookmarksUsingCharCount(oOrgDoc, oDupDoc)
            End If
        Else
            MessageBox.Show("Unable to hypherlink the consistency terms in the below mentioned original document." & vbCr & "File name : " & oOrgDoc.FullName & vbCr & vbCr & "Do you want to continue the process?", sMsgTitle, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        End If
        With oOrgDoc.ActiveWindow.View
            .ShowRevisionsAndComments = False : .RevisionsView = oRevView : .Type = oViewType : .ShowFieldCodes = bShowFieldCodes
        End With
    End Function


    Private Function TransferBookmarksUsingCharCount(oORgDoc As Word.Document, oDupDoc As Word.Document)
        On Error Resume Next
        Dim oDupStyRng As Range = Nothing : Dim oOrgStyRng As Range = Nothing : Dim I As Short
        For I = 1 To 3
            Select Case I
                Case 1
                    oOrgStyRng = oORgDoc.Range.Duplicate
                    oDupStyRng = oDupDoc.Range.Duplicate
                Case 2
                    If oORgDoc.Footnotes.Count > 0 AndAlso oDupDoc.Footnotes.Count > 0 Then
                        oOrgStyRng = oORgDoc.StoryRanges(WdStoryType.wdFootnotesStory).Duplicate : oDupStyRng = oDupDoc.StoryRanges(WdStoryType.wdFootnotesStory).Duplicate
                    Else
                        oOrgStyRng = Nothing : oDupStyRng = Nothing
                    End If
                Case 3
                    If oORgDoc.Endnotes.Count > 0 Then
                        oOrgStyRng = oORgDoc.StoryRanges(WdStoryType.wdEndnotesStory).Duplicate : oDupStyRng = oDupDoc.StoryRanges(WdStoryType.wdEndnotesStory).Duplicate
                    Else
                        oOrgStyRng = Nothing : oDupStyRng = Nothing
                    End If
            End Select
            If Not oOrgStyRng Is Nothing And Not oDupStyRng Is Nothing Then
                Dim oBK As Bookmark : Dim opRng As Word.Paragraph : Dim oOrgParaRng As Range : Dim oDupParaRng As Range : Dim iParaCnt As Long = 0 : Dim iCharCnt As Long = 0 : Dim lStChar As Long = 0 : Dim lEndChar As Long = 0
                For Each opRng In oDupStyRng.Paragraphs
                    iParaCnt = iParaCnt + 1
                    For Each oBK In opRng.Range.Bookmarks
                        If oBK.Name.ToLower.Contains("cbml") = False Then
                            If oBK.Range.InRange(opRng.Range.Duplicate) = True Then
                                oOrgParaRng = oOrgStyRng.Paragraphs(iParaCnt).Range.Duplicate : oDupParaRng = opRng.Range.Duplicate
                                Call oDupParaRng.SetRange(oDupParaRng.Start, oBK.Range.Start)
                                iCharCnt = oDupParaRng.Characters.Count
                                If iCharCnt = 1 Then iCharCnt = 0
                                lStChar = iCharCnt + 1 : lEndChar = iCharCnt + oBK.Range.Characters.Count
                                Call oOrgParaRng.SetRange(oOrgParaRng.Characters(lStChar).Start, oOrgParaRng.Characters(lEndChar).End)
                                oOrgParaRng.Bookmarks.Add(oBK.Name, oOrgParaRng.Duplicate) : oORgDoc.UndoClear() : oDupDoc.UndoClear()
                            End If
                        End If
                    Next oBK
                Next opRng
            End If
        Next I
    End Function

    Private Function ToTransferBookMarkUsingOXML(oMasterDoc As WordprocessingDocument, oIndexDoc As WordprocessingDocument)
        oIndexDoc.MainDocumentPart.Document.Body.Descendants()
        Dim oBKStart As IEnumerable(Of BookmarkStart) = oIndexDoc.MainDocumentPart.Document.Body.Descendants(Of BookmarkStart)()
        Dim oBKEnd As IEnumerable(Of BookmarkEnd) = oIndexDoc.MainDocumentPart.Document.Body.Descendants(Of BookmarkEnd)()
        Dim oNewBKStart As New Dictionary(Of String, BookmarkStart)
        Dim oNewBKEnd As New Dictionary(Of String, OpenXmlElement)
        If oBKStart.Count = oBKEnd.Count Then
            For Each oBK As OpenXmlElement In oBKStart
                Dim T As Short = 0 : Dim P As Long = 0 : Dim C As Long = 0
                Dim oOldBKStart As BookmarkStart = oBK.Clone
                Do
                    Select Case oBK.LocalName
                        Case "p" : P = P + 1
                        Case "t" : C = C + oBK.InnerText.Length
                        Case "body" : T = 1 : Exit Do
                        Case "endnote" : T = 2 : Exit Do
                        Case "footnote" : T = 3 : Exit Do
                    End Select
                Loop Until True
                Call oNewBKStart.Add(String.Concat(T, ":", P, ":", C, ":", CStr(Rnd() * 100)), oOldBKStart)
            Next
            For Each oBK As OpenXmlElement In oBKEnd
                Dim T As Short = 0 : Dim P As Long = 0 : Dim C As Long = 0
                Dim oOldBKEnd As BookmarkEnd = oBK.Clone
                Do
                    Select Case oBK.LocalName
                        Case "p" : P = P + 1
                        Case "t" : C = C + oBK.InnerText.Length
                        Case "body" : T = 1 : Exit Do
                        Case "endnote" : T = 2 : Exit Do
                        Case "footnote" : T = 3 : Exit Do
                    End Select
                Loop Until True
                Call oNewBKEnd.Add(String.Concat(T, ":", P, ":", C, ":", CStr(Rnd() * 100)), oOldBKEnd)
            Next
        End If
    End Function
End Class
