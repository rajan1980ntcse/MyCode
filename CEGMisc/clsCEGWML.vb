Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Xml
Imports System.Xml.Linq
Imports <xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
Imports System.IO.Packaging
Imports System.Xml.Serialization

Public Class clsDocInfo
    Public oParaUsedStyleList As New SortedDictionary(Of String, Integer)
    Public oCharUsedStyleList As New SortedDictionary(Of String, Integer)
    Public oParaTextWithOrder As New List(Of Tuple(Of String, String))
    Public oAllStyle As New SortedDictionary(Of String, String)
End Class

Public Class clsCEGWML
    Const documentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Const stylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Const FNRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    Const ENRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    'Const wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


    Public Function ToGetDocXUsedStyleList(sDocxFullName As String, ByRef sStyleMessage As String) As clsDocInfo
        Try
            Dim xDoc As XDocument = Nothing : Dim styleDoc As XDocument = Nothing
            Dim xFNDoc As XDocument = Nothing : Dim xENDoc As XDocument = Nothing
            Dim sDocName As String = Path.GetFileName(sDocxFullName)
            Dim T As Short
            Dim sTempDirFullPath As String = Path.Combine(Path.GetTempPath, DateTime.Now.ToString("yyyyMMdd_HH_mm_ss"))
            If Directory.Exists(sTempDirFullPath) = False Then Directory.CreateDirectory(sTempDirFullPath)
            Dim sTempFileFullPath As String = Path.Combine(sTempDirFullPath, sDocName)
            Call File.Copy(sDocxFullName, sTempFileFullPath, True)

            Using wdPackage As Package = Package.Open(sTempFileFullPath, FileMode.Open, FileAccess.Read)
                Dim docPackageRelationship As PackageRelationship = wdPackage.GetRelationshipsByType(documentRelationshipType).FirstOrDefault()
                If (docPackageRelationship IsNot Nothing) Then
                    Dim documentUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), docPackageRelationship.TargetUri)
                    Dim FNUri As New Uri("/word/footnotes.xml", UriKind.Relative)
                    Dim ENUri As New Uri("/word/endnotes.xml", UriKind.Relative)
                    Dim documentPart As PackagePart '= wdPackage.GetPart(documentUri)
                    For T = 1 To 3 'Load the document XML in the part into an XDocument instance.
                        Select Case T
                            Case 1
                                documentPart = wdPackage.GetPart(documentUri) : xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                            Case 2
                                If wdPackage.PartExists(FNUri) = True Then
                                    documentPart = wdPackage.GetPart(FNUri) : xFNDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                                End If
                            Case 3
                                If wdPackage.PartExists(ENUri) = True Then
                                    documentPart = wdPackage.GetPart(ENUri) : xENDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                                End If
                        End Select
                    Next
                    'xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream())) 'Load the document XML in the part into an XDocument instance.
                    documentPart = wdPackage.GetPart(documentUri) 'Find the styles part. There will only be one.
                    Dim styleRelation As PackageRelationship = documentPart.GetRelationshipsByType(stylesRelationshipType).FirstOrDefault()
                    If (styleRelation IsNot Nothing) Then
                        Dim styleUri As Uri = PackUriHelper.ResolvePartUri(documentUri, styleRelation.TargetUri)
                        Dim stylePart As PackagePart = wdPackage.GetPart(styleUri)
                        styleDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream())) '  Load the style XML in the part into an XDocument instance.
                    End If
                End If
            End Using

            Dim oParaStyleList As New SortedDictionary(Of String, String)
            Dim oCharStyleList As New SortedDictionary(Of String, String)

            styleDoc.Root.<w:style>.ToList().ForEach(Function(x)
                                                         If (x.@w:type.ToString() = "paragraph" AndAlso oParaStyleList.ContainsKey(x.@w:styleId.ToString()) = False) Then
                                                             oParaStyleList.Add(x.@w:styleId.ToString(), x.<w:name>.@w:val.ToString())
                                                         End If
                                                     End Function)

            styleDoc.Root.<w:style>.ToList().ForEach(Function(x)
                                                         If (x.@w:type.ToString() = "character" AndAlso oCharStyleList.ContainsKey(x.@w:styleId.ToString()) = False) Then
                                                             oCharStyleList.Add(x.@w:styleId.ToString(), x.<w:name>.@w:val.ToString())
                                                         End If
                                                     End Function)

            '############ Getting default style from document ###########
            Dim defaultParaStyle As String = _
                ( _
                    From style In styleDoc.Root.<w:style> _
                    Where style.@w:type = "paragraph" And _
                          style.@w:default = "1" _
                    Select style _
                ).First().@w:styleId

            Dim defaultCharStyle As String = _
                ( _
                    From style In styleDoc.Root.<w:style> _
                    Where style.@w:type = "character" And _
                          style.@w:default = "1" _
                    Select style _
                ).First().@w:styleId



            Dim oParaUsedStyleList As New SortedDictionary(Of String, Integer)
            Dim oCharUsedStyleList As New SortedDictionary(Of String, Integer)
            Dim oParaTextWithOrder As New List(Of Tuple(Of String, String))
            Dim oAllStyle As New SortedDictionary(Of String, String)
            Dim I As Integer ': Dim sStyleMessage As String = String.Empty
            Dim paragraphs = Nothing : Dim character = Nothing : Dim oParaText = Nothing
            For I = 1 To 3
                Dim oTempParaTextWithOrder As New List(Of Tuple(Of String, String))
                Select Case I
                    Case 1
                        paragraphs = (From para In xDoc.Root.<w:body>...<w:p>.<w:pPr>.<w:pStyle> _
                                     Where para.@w:val <> defaultParaStyle Select para.@w:val)
                        character = (From run In xDoc.Root.<w:body>...<w:r>.<w:rPr>.<w:rStyle> _
                                     Where run.@w:val <> defaultCharStyle Select run.@w:val)
                        oTempParaTextWithOrder = (From para In xDoc.Root.<w:body>...<w:p> _
                                     Let sStyleId = para.<w:pPr>.<w:pStyle>.@w:val _
                                     Where sStyleId <> defaultParaStyle AndAlso String.IsNullOrEmpty(sStyleId) = False
                                     Select New Tuple(Of String, String)(oParaStyleList(sStyleId), para.Value)).ToList()
                    Case 2
                        If xFNDoc IsNot Nothing Then
                            paragraphs = (From para In xFNDoc.Root.<w:footnote>...<w:p>.<w:pPr>.<w:pStyle> _
                                         Where para.@w:val <> defaultParaStyle Select para.@w:val)
                            character = (From run In xFNDoc.Root.<w:footnote>...<w:r>.<w:rPr>.<w:rStyle> _
                                         Where run.@w:val <> defaultCharStyle Select run.@w:val)
                            oTempParaTextWithOrder = (From para In xFNDoc.Root.<w:footnote>...<w:p> _
                                         Let sStyleId = para.<w:pPr>.<w:pStyle>.@w:val _
                                         Where sStyleId <> defaultParaStyle AndAlso String.IsNullOrEmpty(sStyleId) = False
                                         Select New Tuple(Of String, String)(oParaStyleList(sStyleId), para.Value)).ToList()
                        Else
                            paragraphs = Nothing : character = Nothing : oTempParaTextWithOrder = Nothing
                        End If
                    Case 3
                        If xENDoc IsNot Nothing Then
                            paragraphs = (From para In xENDoc.Root.<w:endnote>...<w:p>.<w:pPr>.<w:pStyle> _
                                         Where para.@w:val <> defaultParaStyle Select para.@w:val)
                            character = (From run In xENDoc.Root.<w:endnote>...<w:r>.<w:rPr>.<w:rStyle> _
                                         Where run.@w:val <> defaultCharStyle Select run.@w:val)
                            oTempParaTextWithOrder = (From para In xENDoc.Root.<w:endnote>...<w:p> _
                                         Let sStyleId = para.<w:pPr>.<w:pStyle>.@w:val _
                                         Where sStyleId <> defaultParaStyle AndAlso String.IsNullOrEmpty(sStyleId) = False
                                         Select New Tuple(Of String, String)(oParaStyleList(sStyleId), para.Value)).ToList()
                        Else
                            paragraphs = Nothing : character = Nothing : oTempParaTextWithOrder = Nothing
                        End If
                End Select






                If paragraphs IsNot Nothing Then
                    If Not oTempParaTextWithOrder Is Nothing Then oParaTextWithOrder.AddRange(oTempParaTextWithOrder)
                    Dim sParaStyleId As String = String.Empty : Dim sParaStyleVal As String = String.Empty
                    For Each sParaStyleId In paragraphs
                        Select Case oParaStyleList.ContainsKey(sParaStyleId)
                            Case True : sParaStyleVal = oParaStyleList(sParaStyleId).ToString()
                            Case False : I = I + 1 : sParaStyleVal = "Unknown style " & I
                        End Select
                        If oParaUsedStyleList.ContainsKey(sParaStyleVal) = False Then
                            oParaUsedStyleList.Add(sParaStyleVal, 1)
                        Else
                            oParaUsedStyleList(sParaStyleVal) = oParaUsedStyleList(sParaStyleVal) + 1
                        End If
                        '############# collect all styles and doc name ############
                        If oAllStyle.ContainsKey(sParaStyleVal) = False Then
                            oAllStyle.Add(sParaStyleVal, sDocName)
                        Else
                            Dim sDocList As String = oAllStyle(sParaStyleVal)
                            If sDocList.Contains(sDocName) = False Then
                                oAllStyle(sParaStyleVal) = String.Concat(oAllStyle(sParaStyleVal), ", ", sDocName)
                            End If
                        End If
                    Next
                End If



                If character IsNot Nothing Then
                    Dim sCharStyleId As String = String.Empty : Dim sCharStyleVal As String = String.Empty
                    For Each sCharStyleId In character
                        ''sCharStyleId = R.StyleName
                        Select Case oCharStyleList.ContainsKey(sCharStyleId)
                            Case True : sCharStyleVal = oCharStyleList(sCharStyleId).ToString()
                            Case False : I = I + 1 : sCharStyleVal = "Unknown style " & I
                        End Select
                        If oCharUsedStyleList.ContainsKey(sCharStyleVal) = False Then
                            sStyleMessage = sStyleMessage & sCharStyleId & "=" & sCharStyleVal & vbCrLf
                            oCharUsedStyleList.Add(sCharStyleVal, 1)
                        Else
                            oCharUsedStyleList(sCharStyleVal) = oCharUsedStyleList(sCharStyleVal) + 1
                        End If
                        '############# collect all styles and doc name ############
                        If oAllStyle.ContainsKey(sCharStyleVal) = False Then
                            oAllStyle.Add(sCharStyleVal, sDocName)
                        Else
                            Dim sDocList As String = oAllStyle(sCharStyleVal)
                            If sDocList.Contains(sDocName) = False Then
                                oAllStyle(sCharStyleVal) = String.Concat(oAllStyle(sCharStyleVal), ", ", sDocName)
                            End If
                        End If
                    Next
                End If
                character = Nothing
            Next
            Dim oclsDoc As New clsDocInfo()
            oclsDoc.oCharUsedStyleList = oCharUsedStyleList : oclsDoc.oParaUsedStyleList = oParaUsedStyleList
            oclsDoc.oParaTextWithOrder = oParaTextWithOrder : oclsDoc.oAllStyle = oAllStyle
            Return oclsDoc
        Catch ex As Exception
            sStyleMessage = ex.Message
        End Try
    End Function

    Public Function ToGetDocXParaList(sDocxFullName As String, ByRef oclsWMLInfo As clsCEGWMLInfo, ByRef sStyleMessage As String) As Boolean
        Dim xDoc As XDocument = Nothing : Dim xStylesDoc As XDocument = Nothing : Dim xFNDoc As XDocument = Nothing : Dim xENDoc As XDocument = Nothing
        Dim T As Short : Dim sDocName As String = Path.GetFileName(sDocxFullName)
        Dim sTempDirFullPath As String = Path.Combine(Path.GetTempPath, DateTime.Now.ToString("yyyyMMdd_HH_mm_ss"))
        If Directory.Exists(sTempDirFullPath) = False Then Directory.CreateDirectory(sTempDirFullPath)
        Dim sTempFileFullPath As String = Path.Combine(sTempDirFullPath, sDocName)
        Call File.Copy(sDocxFullName, sTempFileFullPath, True)

        Try
            oclsWMLInfo = New clsCEGWMLInfo
            Using wdPackage As Package = Package.Open(sTempFileFullPath, FileMode.Open, FileAccess.Read)
                Dim docPackageRelationship As PackageRelationship = wdPackage.GetRelationshipsByType(documentRelationshipType).FirstOrDefault()
                If (docPackageRelationship IsNot Nothing) Then
                    Dim oMainDocUri As Uri = PackUriHelper.ResolvePartUri(New Uri("/", UriKind.Relative), docPackageRelationship.TargetUri)
                    Dim oFNUri As New Uri("/word/footnotes.xml", UriKind.Relative) : Dim oENUri As New Uri("/word/endnotes.xml", UriKind.Relative)
                    Dim documentPart As PackagePart
                    For T = 1 To 3 'Load the document XML in the part into an XDocument instance.
                        Select Case T
                            Case 1
                                documentPart = wdPackage.GetPart(oMainDocUri) : xDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                            Case 2
                                If wdPackage.PartExists(oFNUri) = True Then
                                    documentPart = wdPackage.GetPart(oFNUri) : xFNDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                                End If
                            Case 3
                                If wdPackage.PartExists(oENUri) = True Then
                                    documentPart = wdPackage.GetPart(oENUri) : xENDoc = XDocument.Load(XmlReader.Create(documentPart.GetStream()))
                                End If
                        End Select
                    Next

                    documentPart = wdPackage.GetPart(oMainDocUri) 'Find the styles part. There will only be one.
                    Dim styleRelation As PackageRelationship = documentPart.GetRelationshipsByType(stylesRelationshipType).FirstOrDefault()
                    If (styleRelation IsNot Nothing) Then
                        Dim oStyleUri As Uri = PackUriHelper.ResolvePartUri(oMainDocUri, styleRelation.TargetUri)
                        Dim stylePart As PackagePart = wdPackage.GetPart(oStyleUri)
                        xStylesDoc = XDocument.Load(XmlReader.Create(stylePart.GetStream())) '  Load the style XML in the part into an XDocument instance.
                    End If
                End If
            End Using

            Dim oDocParaStylesList As New SortedDictionary(Of String, String)
            Dim oUsedParaStylesList As New SortedDictionary(Of String, Integer)
            Dim oUsedStyleAndParaInfo As New SortedDictionary(Of String, List(Of clsCEGParaWMLInfo))

            xStylesDoc.Root.<w:style>.ToList().ForEach(Function(x)
                                                           If (x.@w:type.ToString() = "paragraph" AndAlso oDocParaStylesList.ContainsKey(x.@w:styleId.ToString()) = False) Then
                                                               oDocParaStylesList.Add(x.@w:styleId.ToString(), x.<w:name>.@w:val.ToString())
                                                           End If
                                                       End Function)

            '############ Replace Tab and Symbols ######################
            'Call xDoc.Root.<w:p>...<w:r>...<w:tab>.ToList().ReplaceWith(<w:t></w:t>)
            '<w:tab/>
            xDoc.Root...<w:tab>.ToList().ForEach(Function(x)
                                                     'If x.Name = "w:tab" Then
                                                     x.ReplaceWith(<w:t>&#x9;</w:t>)
                                                     'End If
                                                 End Function)

            xDoc.Root...<w:sym>.ToList().ForEach(Function(x)
                                                     'If x.Name = "sym" Then
                                                     x.ReplaceWith(<w:t>&#x220e;</w:t>)
                                                     'End If
                                                 End Function)



            '############ Getting default style from document ###########
            Dim defaultStyle As String = _
                (From style In xStylesDoc.Root.<w:style> _
                    Where style.@w:type = "paragraph" And _
                          style.@w:default = "1" _
                    Select style).First().@w:styleId



            'Following is the new query that finds all paragraphs in the document and their styles.
            Dim oParaList As Object
            For T = 1 To 3
                Select Case T
                    Case 1
                        oParaList = From para In xDoc.Root.<w:body>...<w:p> _
                        Let styleNode As XElement = para.<w:pPr>.<w:pStyle>.FirstOrDefault() _
                        Select New With {.ParagraphNode = para, .StyleName = GetStyleOfParagraph(styleNode, defaultStyle), .Text = para.Value}
                    Case 2
                        If xFNDoc IsNot Nothing Then
                            oParaList = From para In xFNDoc.Root.<w:footnote>...<w:p> _
                            Let styleNode As XElement = para.<w:pPr>.<w:pStyle>.FirstOrDefault() _
                            Select New With {.ParagraphNode = para, .StyleName = GetStyleOfParagraph(styleNode, defaultStyle), .Text = para.Value}
                        Else
                            oParaList = Nothing
                        End If
                    Case 3
                        If xENDoc IsNot Nothing Then
                            oParaList = From para In xENDoc.Root.<w:endnote>...<w:p> _
                            Let styleNode As XElement = para.<w:pPr>.<w:pStyle>.FirstOrDefault() _
                            Select New With {.ParagraphNode = para, .StyleName = GetStyleOfParagraph(styleNode, defaultStyle), .Text = para.Value}
                        Else
                            oParaList = Nothing
                        End If
                End Select
                Dim sStyleId As String = String.Empty : Dim sStyleVal As String = String.Empty
                If oParaList IsNot Nothing Then
                    For Each P In oParaList
                        sStyleId = P.StyleName  '##### Get Style ID from the paragraph
                        Dim oclsPInfo As New clsCEGParaWMLInfo '########### Collecting Para Info ###########
                        oclsPInfo.sParaText = P.Text
                        'oclsPInfo.sParaWML = P.xml  'SerializeObjectToXmlNode(P)
                        Select Case T
                            Case 1 : oclsPInfo.eParaType = enumParaType.MainText
                            Case 2 : oclsPInfo.eParaType = enumParaType.FNText
                            Case 3 : oclsPInfo.eParaType = enumParaType.ENText
                            Case Else : oclsPInfo.eParaType = enumParaType.None
                        End Select
                        Select Case oDocParaStylesList.ContainsKey(sStyleId)  '##### checking Style ID exists in the docpara style list
                            Case True : sStyleVal = oDocParaStylesList(sStyleId).ToString()
                            Case False : T = T + 1 : sStyleVal = "Unknown style " & T
                        End Select
                        If oUsedParaStylesList.ContainsKey(sStyleVal) = False Then '##### checking Style name exists in the usedpara style list
                            sStyleMessage = sStyleMessage & sStyleId & "=" & sStyleVal & vbCrLf
                            oUsedParaStylesList.Add(sStyleVal, 1)
                            Dim olCEGParaInfo As New List(Of clsCEGParaWMLInfo) : olCEGParaInfo.Add(oclsPInfo)
                            Call oUsedStyleAndParaInfo.Add(sStyleVal, olCEGParaInfo)
                        Else
                            oUsedParaStylesList(sStyleVal) = oUsedParaStylesList(sStyleVal) + 1
                            oUsedStyleAndParaInfo(sStyleVal).Add(oclsPInfo)
                        End If
                    Next
                End If
                oParaList = Nothing
            Next
            oclsWMLInfo.oDocParaStylesList = oDocParaStylesList
            oclsWMLInfo.oUsedParaStylesList = oUsedParaStylesList
            oclsWMLInfo.oUsedStyleAndParaInfo = oUsedStyleAndParaInfo
            Return True
        Catch ex As Exception
            MessageBox.Show("WML Error : " & ex.Message, Application.ProductName, MessageBoxButtons.OK)
            ex.Data.Clear()
        End Try
    End Function

    Private Function ToGetXmlNode(element As XElement) As XmlNode
        Using oXMLRead As XmlReader = element.CreateReader
            Dim xmldoc As XmlDocument = New XmlDocument
            xmldoc.Load(oXMLRead)
            Return xmldoc
        End Using
    End Function

    Public Function SerializeObjectToXmlNode(obj As Object) As String
        If obj Is Nothing Then Throw New ArgumentNullException("Argument cannot be null")
        Dim resultNode As XmlNode
        Dim xmlSerializer As XmlSerializer = New XmlSerializer(obj.GetType())
        Using memoryStream As MemoryStream = New MemoryStream()
            Try
                xmlSerializer.Serialize(memoryStream, obj)
            Catch ex As Exception
                Return String.Empty
            End Try
            memoryStream.Position = 0
            Dim doc As XmlDocument = New XmlDocument()
            doc.Load(memoryStream)
            resultNode = doc.DocumentElement
        End Using
        Return resultNode.InnerXml
    End Function




    Private Function GetStyleOfParagraph(ByVal styleNode As XElement, ByVal defaultStyle As String) As String
        If (styleNode Is Nothing) Then
            Return defaultStyle
        Else
            Return styleNode.@w:val
        End If
    End Function


End Class

Public Class clsCEGWMLInfo
    Public oDocParaStylesList As SortedDictionary(Of String, String)
    Public oUsedParaStylesList As SortedDictionary(Of String, Integer)
    Public oUsedStyleAndParaInfo As SortedDictionary(Of String, List(Of clsCEGParaWMLInfo))
End Class


Public Class clsCEGParaWMLInfo
    Public sParaText As String
    Public sParaWML As String
    Public eParaType As enumParaType
End Class

Public Enum enumParaType
    MainText = 1
    FNText = 2
    ENText = 3
    None = 0
End Enum
