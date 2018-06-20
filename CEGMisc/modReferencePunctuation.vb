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
Module modReferencePunctuation

    Public wAPP As Word.Application
    Public aDoc As Word.Document


    Function ReferencePunctuationMain(oactApp As Word.Application, wDoc As Word.Document)
        Try
            wAPP = oactApp
            aDoc = wDoc
            Dim ranDoc As Word.Range : Dim rCount As Integer
            Dim dicRefInfo As New Dictionary(Of Integer, Dictionary(Of Integer, clsAuthorInfo))
            ranDoc = wDoc.Content
            With ranDoc.Find
                .ClearFormatting() : .Replacement.ClearFormatting()
                .Text = "" : .Replacement.Text = "" : .Style = "†Reference"
            End With
            Do While ranDoc.Find.Execute = True
                rCount = rCount + 1
                ranDoc.Select()
                Dim UsedStyleList As New List(Of String)
                UsedStyleList = GetUsedStylesIntheRange(oactApp.Selection.Range)
                dicRefInfo.Add(rCount, GetReferenceInformation(UsedStyleList, oactApp.Selection.Range))
            Loop
            Dim dictRefSortInfo As Dictionary(Of Integer, clsAuthorInfo)
            ''dictRefSortInfo = dicRefInfo.OrderBy(Function(x) x.Value.tagRange)
            StoringReferenceInfoasXML(dicRefInfo)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function StoringReferenceInfoasXML(dictrefInfo As Dictionary(Of Integer, Dictionary(Of Integer, clsAuthorInfo)))
        Try
            Dim objFile As File
            Dim K As Integer
            If Directory.Exists(wAPP.ActiveDocument.Path & "\cTemp") Then Directory.Delete(wAPP.ActiveDocument.Path & "\cTemp", True)
            Directory.CreateDirectory(wAPP.ActiveDocument.Path & "\cTemp")
            Dim totalAuthor As New List(Of Dictionary(Of Integer, Dictionary(Of String, String)))
            totalAuthor = GetAuthorInformationfromDictionary(dictrefInfo)
            For Each dicVal In dictrefInfo.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
                K = K + 1
                Dim xDoc As New XDocument(New XDeclaration("1.0", "utf-8", "true"), New XElement("CEGPunc"))
                For Each objAuthorInfor As KeyValuePair(Of Integer, clsAuthorInfo) In dicVal.OrderBy(Function(x) x.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                    If Not (objAuthorInfor.Value.tagName.ToLower().Contains("prefix") Or objAuthorInfor.Value.tagName.ToLower().Contains("suffix") Or objAuthorInfor.Value.tagName.ToLower().Contains("surname") Or objAuthorInfor.Value.tagName.ToLower().Contains("givenname")) Then
                        Dim xEle As New XElement(objAuthorInfor.Value.tagName.Replace("‡", "").ToLower(), objAuthorInfor.Value.tagValue)
                        xDoc.Root.Add(xEle)
                    End If
                Next
                Dim yEle As XElement = GetAuthorAndEditortoXML(totalAuthor.Item(K - 1))
                xDoc.Root.Add(yEle)
                xDoc.Save(wAPP.ActiveDocument.Path & "\cTemp" & "\Ref" & K & ".xml")
            Next


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function GetAuthorAndEditortoXML(authorInfo As Dictionary(Of Integer, Dictionary(Of String, String))) As XElement
        Try
            Dim AuthorElement As New XElement("AuthorList")
            Dim EditorElement As New XElement("EditorList")

            For Each aList As KeyValuePair(Of Integer, Dictionary(Of String, String)) In authorInfo
                Dim fEditor As Boolean
                For Each kVal As String In aList.Value.Keys
                    If InStr(1, kVal, "ed") > 0 Then
                        fEditor = True : Exit For
                    End If
                Next
                If fEditor = True Then
                    Dim eElement As New XElement("Editor")
                    For Each lVal As KeyValuePair(Of String, String) In aList.Value
                        eElement.Add(New XElement(lVal.Key.Replace("‡", ""), lVal.Value))
                    Next
                    EditorElement.Add(eElement)
                Else
                    Dim aElement As New XElement("Author")
                    For Each lVal As KeyValuePair(Of String, String) In aList.Value
                        aElement.Add(New XElement(lVal.Key.Replace("‡", ""), lVal.Value))
                    Next
                    AuthorElement.Add(aElement)
                End If
            Next
            Dim cEle As New XElement("AuthorInformation")
            cEle.Add(AuthorElement)
            cEle.Add(EditorElement)
            GetAuthorAndEditortoXML = cEle

        Catch ex As Exception

        End Try
    End Function
    Function GetAuthorInformationfromDictionary(dictrefInfo As Dictionary(Of Integer, Dictionary(Of Integer, clsAuthorInfo))) As List(Of Dictionary(Of Integer, Dictionary(Of String, String)))
        Try
            Dim eachRefAuthor As New List(Of Dictionary(Of Integer, Dictionary(Of String, String)))
            Dim totalAuthor As New Dictionary(Of Integer, Dictionary(Of String, String))
            Dim eachAuthor As Dictionary(Of String, String)
            Dim totAuthor As New Dictionary(Of List(Of String), List(Of Word.Range))
            Dim listTagName As New List(Of String)
            Dim listAuthorRange As New List(Of Word.Range)
            Dim grpTagName As String
            Dim K As Integer
            K = 1
            If Not IsNothing(dictrefInfo) Then
                For Each dicVal In dictrefInfo.OrderBy(Function(item) item.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value).Values
                    totalAuthor = New Dictionary(Of Integer, Dictionary(Of String, String))
                    listTagName.Clear() : grpTagName = "|" : listAuthorRange = New List(Of Word.Range) : eachAuthor = New Dictionary(Of String, String)
                    For Each objAuthorInfor As KeyValuePair(Of Integer, clsAuthorInfo) In dicVal.OrderBy(Function(x) x.Key).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                        If objAuthorInfor.Value.tagName.ToLower().Contains("prefix") Or objAuthorInfor.Value.tagName.ToLower().Contains("suffix") Or objAuthorInfor.Value.tagName.ToLower().Contains("surname") Or objAuthorInfor.Value.tagName.ToLower().Contains("givenname") Then
                            If (grpTagName.Contains("surname") And grpTagName.Contains("givenname")) And (objAuthorInfor.Value.tagName.ToLower().Contains("givenname") Or objAuthorInfor.Value.tagName.ToLower().Contains("surname")) Then
                                If listAuthorRange.Count > 0 Then
                                    totalAuthor.Add(K, eachAuthor) : eachAuthor = New Dictionary(Of String, String) : K = K + 1 : grpTagName = "|" : listTagName.Clear()
                                    '' totAuthor.Add(listTagName, listAuthorRange) : listTagName.Clear() : listAuthorRange = New List(Of Word.Range)
                                End If
                                listAuthorRange.Add(objAuthorInfor.Value.tagRange)
                            End If
                            eachAuthor.Add(objAuthorInfor.Value.tagName.ToLower(), objAuthorInfor.Value.tagValue)
                            grpTagName = grpTagName & objAuthorInfor.Value.tagName.ToLower() & "|"
                            listTagName.Add(objAuthorInfor.Value.tagName.ToLower())
                            listAuthorRange.Add(objAuthorInfor.Value.tagRange)
                            If grpTagName.Contains("prefix") And grpTagName.Contains("suffix") And grpTagName.Contains("surname") And grpTagName.Contains("givenname") Then
                                If listAuthorRange.Count > 0 Then
                                    totalAuthor.Add(K, eachAuthor) : eachAuthor = New Dictionary(Of String, String) : K = K + 1 : grpTagName = "|" : listTagName.Clear()
                                    '' totAuthor.Add(listTagName, listAuthorRange) : listTagName.Clear() : grpTagName = "|" : listAuthorRange = New List(Of Word.Range)
                                End If
                            ElseIf grpTagName.Contains("suffix") Then
                                listAuthorRange.Add(objAuthorInfor.Value.tagRange)
                                If listAuthorRange.Count > 0 Then
                                    totalAuthor.Add(K, eachAuthor) : eachAuthor = New Dictionary(Of String, String) : K = K + 1 : grpTagName = "|" : listTagName.Clear()
                                    totAuthor.Add(listTagName, listAuthorRange) : listTagName.Clear() : grpTagName = "|" : listAuthorRange = New List(Of Word.Range)
                                End If
                            ElseIf grpTagName.Contains("prefix") Then
                                If listAuthorRange.Count > 0 Then
                                    totalAuthor.Add(K = K + 1, eachAuthor) : eachAuthor = New Dictionary(Of String, String) : grpTagName = "|" : listTagName.Clear()
                                    ''totAuthor.Add(listTagName, listAuthorRange) : listTagName.Clear() : grpTagName = "|" : listAuthorRange = New List(Of Word.Range)
                                End If
                                listAuthorRange.Add(objAuthorInfor.Value.tagRange)
                            End If
                        End If
                    Next
                    If listAuthorRange.Count > 0 Then
                        totalAuthor.Add(K, eachAuthor) : eachAuthor = New Dictionary(Of String, String) : K = K + 1 : listTagName.Clear() : grpTagName = "|"
                        '' totAuthor.Add(listTagName, listAuthorRange) : listTagName.Clear() : grpTagName = "|" : listAuthorRange = New List(Of Word.Range)
                    End If
                    If totalAuthor.Count > 0 Then
                        eachRefAuthor.Add(totalAuthor)
                    Else
                        eachRefAuthor.Add(Nothing)
                    End If
                    totalAuthor = New Dictionary(Of Integer, Dictionary(Of String, String))
                Next
                GetAuthorInformationfromDictionary = eachRefAuthor
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function GetUsedStylesIntheRange(ranRef As Word.Range) As List(Of String)
        Try
            Dim I As Integer : Dim lstStyle As New List(Of String)
            For I = 1 To ranRef.Words.Count
                Dim stName As Word.Style
                stName = ranRef.Words(I).CharacterStyle
                If Not IsNothing(stName) Then
                    If Not lstStyle.Contains(stName.NameLocal) And aDoc.Styles(stName).BuiltIn = False Then
                        lstStyle.Add(stName.NameLocal)
                    End If
                End If
            Next
            GetUsedStylesIntheRange = lstStyle
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Function
    Function GetReferenceInformation(usedStyleList As List(Of String), ranRef As Word.Range) As Dictionary(Of Integer, clsAuthorInfo)
        Try
            Dim lstAuthorInfo As New Dictionary(Of Integer, clsAuthorInfo)

            If ranRef Is Nothing Then Exit Function
            For Each sStyle As String In usedStyleList

                Dim ranDoc As Word.Range : Dim rCount As Integer
                ranDoc = ranRef.Duplicate
                With ranDoc.Find
                    .ClearFormatting() : .Replacement.ClearFormatting()
                    .Text = "" : .Replacement.Text = "" : .Style = sStyle
                End With
                Do While ranDoc.Find.Execute = True
                    rCount = rCount + 1
                    ranDoc.Select()
                    Dim objRefInfo As New clsAuthorInfo()
                    If Not wAPP.Selection.InRange(ranRef) Then Exit Do
                    objRefInfo.oAuthorRng = ranRef.Duplicate
                    objRefInfo.tagName = sStyle
                    objRefInfo.tagRange = wAPP.Selection.Range
                    objRefInfo.tagValue = wAPP.Selection.Range.Text
                    lstAuthorInfo.Add(wAPP.Selection.Range.Start, objRefInfo)
                Loop
            Next

            GetReferenceInformation = lstAuthorInfo
        Catch ex As Exception

        End Try
    End Function
End Module
