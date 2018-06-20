Imports Word = Microsoft.Office.Interop.Word
Imports System.Diagnostics
Imports Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports CEGINI

Public Class clsPair
    Public OpenBrace As Long
    Public CloseBrace As Long
    Public OpenBraceRange As Word.Range
    Public CloseBraceRange As Word.Range
    Public MBoolean As Boolean
End Class
Module modPairing
    Public VryBkc As String
    Public AutoPair As Boolean
    Public Function PairingRoutine(wDoc As Word.Document, FrstStr As String, ScndStr As String, Optional HMatch As Boolean = False, Optional StryRng As Word.Range = Nothing) As List(Of Word.Range)
        Dim OpColln As New Collection
        Dim ClColln As New Collection
        Dim OpCls As clsPair
        Dim Para As Word.Paragraph
        Dim FndRng As Word.Range
        Dim I As Integer
        Dim J As Integer
        Dim LsBln As Boolean
        Dim PrvI As Integer
        Dim Fstr As String
        Dim Sstr As String
        Dim Hrng As Word.Range
        Dim dmColln As New Collection
        Dim xFormatStr As String
        Dim lstBraketRange As New List(Of Word.Range)
        If StryRng.Information(Word.WdInformation.wdInCommentPane) = True Then
            Exit Function
        End If

        xFormatStr = Replace(Space(Len(CStr(Len(wDoc.Content.Text)))), " ", "0")
        For Each Para In StryRng.Paragraphs
            If IsNumeric(FrstStr) = True Then
                Fstr = ChrW(FrstStr)
            Else
                Fstr = FrstStr
            End If
            If IsNumeric(ScndStr) = True Then
                Sstr = ChrW(ScndStr)
            Else
                Sstr = ScndStr
            End If
            OpColln = New Collection : ClColln = New Collection '###### Initializing the collection #####
            '######### Finding Open Item ############
            FndRng = Para.Range.Duplicate
            FndRng.Find.ClearFormatting()
            FndRng.Find.Text = Fstr
            ''MsgBox(FndRng.Text)
            Do While FndRng.Find.Execute = True
                If FndRng.Text <> Fstr Then Exit Do
                If FndRng.InRange(Para.Range.Duplicate) = False Then Exit Do
                OpCls = New clsPair
                OpCls.OpenBrace = FndRng.Start + 1
                OpCls.OpenBraceRange = FndRng.Duplicate
                OpColln.Add(OpCls)
                OpCls = Nothing
                FndRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Loop
            '######### Ends here ####################

            '######### Finding Close Item ###########
            FndRng = Para.Range.Duplicate
            FndRng.Find.ClearFormatting()
            FndRng.Find.Text = Sstr
            Do While FndRng.Find.Execute = True
                If FndRng.Text <> Sstr Then Exit Do
                If FndRng.InRange(Para.Range.Duplicate) = False Then Exit Do
                OpCls = New clsPair
                OpCls.CloseBrace = FndRng.Start + 1
                OpCls.CloseBraceRange = FndRng.Duplicate
                ClColln.Add(OpCls)
                OpCls = Nothing
                FndRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
            Loop
            '########### Ends here ################

            If OpColln.Count <> ClColln.Count Then
                If OpColln.Count > ClColln.Count Then
                    For I = 1 To ClColln.Count
                        LsBln = False : PrvI = 0
                        For J = 1 To OpColln.Count
                            If OpColln(J).OpenBrace > ClColln(I).CloseBrace Then
                                Exit For
                            Else
                                PrvI = J : LsBln = True
                            End If
                        Next J
                        If LsBln = True Then
                            If OpColln(PrvI).MBoolean = False Then
                                OpColln(PrvI).CloseBrace = ClColln(I).CloseBrace
                                OpColln(PrvI).CloseBraceRange = ClColln(I).CloseBraceRange
                                ClColln(I).CloseBrace = 0
                                OpColln(PrvI).MBoolean = True
                            End If
                        End If
                    Next I

                    For I = 1 To ClColln.Count
                        LsBln = False : PrvI = 0
                        If ClColln(I).CloseBrace <> "0" Then
                            For J = 1 To OpColln.Count
                                If OpColln(J).CloseBrace = "0" Then
                                    If OpColln(J).OpenBrace > ClColln(I).CloseBrace Then
                                        Exit For
                                    End If
                                    PrvI = J : LsBln = True
                                End If
                            Next J
                            If LsBln = True Then
                                OpColln(PrvI).CloseBrace = ClColln(I).CloseBrace
                                ClColln(I).CloseBrace = 0
                            End If
                        End If
                    Next I

                    '########### Highlight the Item ###################
                    For I = 1 To OpColln.Count
                        Fstr = OpColln(I).OpenBrace : Sstr = OpColln(I).CloseBrace
                        If Sstr = "0" Then
                            Hrng = Nothing : Hrng = OpColln(I).OpenBraceRange.Duplicate
                            VryBkc = Format(CStr(Hrng.Start), xFormatStr) 'VryBkc + 1
                            If Hrng Is Nothing = False Then
                                Do While wDoc.Bookmarks.Exists("Pr_Mt_" & VryBkc) = True
                                    VryBkc = VryBkc + 1
                                    VryBkc = Format(VryBkc, xFormatStr) 'VryBkc + 1
                                Loop
                                wDoc.Bookmarks.Add("Pr_Mt_" & VryBkc, Hrng.Duplicate)
                            End If
                        ElseIf Fstr <> "0" And Sstr <> "0" Then
                            If HMatch = True Then
                                Hrng = OpColln(I).OpenBraceRange.Duplicate
                                Hrng.SetRange(CLng(Fstr) - 1, CLng(Sstr))
                                'Hrng.HighlightColorIndex = wdTurquoise
                                Hrng.Shading.BackgroundPatternColor = RGB(102, 102, 153)
                            End If
                        End If
                    Next I

                ElseIf OpColln.Count < ClColln.Count Then
                    For I = 1 To OpColln.Count
                        LsBln = False : PrvI = 0
                        For J = ClColln.Count To 1 Step -1
                            If OpColln(I).OpenBrace < ClColln(J).CloseBrace Then
                                If ClColln(J).MBoolean = True Then
                                    Exit For
                                End If
                                PrvI = J
                                LsBln = True
                            Else
                                Exit For
                            End If
                        Next J
                        If LsBln = True Then
                            If ClColln(PrvI).MBoolean = False Then
                                ClColln(PrvI).OpenBrace = OpColln(I).OpenBrace
                                ClColln(PrvI).OpenBraceRange = OpColln(I).OpenBraceRange
                                OpColln(I).OpenBrace = "0"
                                ClColln(PrvI).MBoolean = True
                            End If
                        End If
                    Next I
                    For I = 1 To OpColln.Count
                        LsBln = False : PrvI = 0
                        If OpColln(I).OpenBrace <> "0" Then
                            For J = 1 To ClColln.Count
                                If ClColln(J).OpenBrace = "0" Then
                                    LsBln = True : PrvI = J
                                End If
                            Next J
                            If LsBln = True Then
                                ClColln(PrvI).OpenBrace = OpColln(I).OpenBrace
                            End If
                        End If
                    Next I

                    '############### Highlight the Item ###################
                    For I = 1 To ClColln.Count
                        Fstr = ClColln(I).OpenBrace : Sstr = ClColln(I).CloseBrace
                        If Fstr = "0" Then
                            Hrng = Nothing : Hrng = ClColln(I).CloseBraceRange.Duplicate
                            VryBkc = Format(CStr(Hrng.Start), xFormatStr) 'VryBkc + 1
                            If Hrng Is Nothing = False Then
                                Do While wDoc.Bookmarks.Exists("Pr_Mt_" & VryBkc) = True
                                    VryBkc = VryBkc + 1
                                    VryBkc = Format(VryBkc, xFormatStr) 'VryBkc + 1
                                Loop
                                wDoc.Bookmarks.Add("Pr_Mt_" & VryBkc, Hrng.Duplicate)
                            End If
                        ElseIf Fstr <> "0" And Sstr <> "0" Then
                            If HMatch = True Then
                                Hrng = ClColln(I).CloseBraceRange.Duplicate
                                Hrng.SetRange(CLng(Fstr) - 1, CLng(Sstr))
                                'Hrng.HighlightColorIndex = wdTurquoise
                                Hrng.Shading.BackgroundPatternColor = RGB(102, 102, 153)
                            End If
                        End If
                    Next I
                    '################### Ends here ########################
                End If
            Else

                dmColln = OpColln

                For I = 1 To ClColln.Count
                    LsBln = False
                    For J = 1 To OpColln.Count
                        If OpColln(J).MBoolean = False Then
                            If OpColln(J).OpenBrace < ClColln(I).CloseBrace Then
                                PrvI = J : LsBln = True
                            End If
                        End If
                    Next J
                    If LsBln = True Then
                        ClColln(I).OpenBrace = OpColln(PrvI).OpenBrace
                        ClColln(I).OpenBraceRange = OpColln(PrvI).OpenBraceRange
                        OpColln(PrvI).MBoolean = True
                    Else
                        ClColln(I).OpenBrace = 0
                    End If
                Next I
                OpColln = ClColln

                For I = 1 To dmColln.Count
                    LsBln = False
                    For J = 1 To OpColln.Count
                        If OpColln(J).OpenBrace = dmColln(I).OpenBrace Then
                            LsBln = True : Exit For
                        End If
                    Next J
                    If LsBln = False Then
                        Hrng = dmColln(I).OpenBraceRange.Duplicate
                        VryBkc = Microsoft.VisualBasic.Strings.Format(CStr(Hrng.Start), xFormatStr) 'VryBkc + 1
                        If Hrng Is Nothing = False Then
                            Do While wDoc.Bookmarks.Exists("Pr_Mt_" & VryBkc) = True
                                VryBkc = VryBkc + 1
                                VryBkc = Format(VryBkc, xFormatStr) 'VryBkc + 1
                            Loop
                            wDoc.Bookmarks.Add("Pr_Mt_" & VryBkc, Hrng.Duplicate)
                        End If
                    End If
                Next I


                For I = 1 To OpColln.Count
                    If OpColln(I).OpenBrace > ClColln(I).CloseBrace Then
                        Hrng = OpColln(I).OpenBraceRange.Duplicate
                        VryBkc = Format(CStr(Hrng.Start), xFormatStr) 'VryBkc + 1
                        If Hrng Is Nothing = False Then
                            Do While wDoc.Bookmarks.Exists("Pr_Mt_" & VryBkc) = True
                                VryBkc = VryBkc + 1
                                VryBkc = Format(VryBkc, xFormatStr) 'VryBkc + 1
                            Loop
                            wDoc.Bookmarks.Add("Pr_Mt_" & VryBkc, Hrng.Duplicate)
                        End If
                    Else
                        Fstr = OpColln(I).OpenBrace : Sstr = ClColln(I).CloseBrace
                        If Fstr <> "0" Then lstBraketRange.Add(wDoc.Range(Fstr, Sstr - 1)) '''''Jaisoft
                        If Fstr <> "0" And Sstr <> "0" Then
                            If HMatch = True Then
                                Hrng = OpColln(I).OpenBraceRange.Duplicate
                                Hrng.SetRange(CLng(Fstr) - 1, CLng(Sstr))
                                'Hrng.HighlightColorIndex = wdTurquoise
                                Hrng.Shading.BackgroundPatternColor = RGB(102, 102, 153)
                            End If
                        ElseIf Fstr = "0" Then
                            ''comment by jasoft
                            'Hrng = OpColln(I).CloseBraceRange.Duplicate
                            'VryBkc = Microsoft.VisualBasic.Strings.Format(CStr(Hrng.Start), xFormatStr) 'VryBkc + 1
                            'If Hrng Is Nothing = False Then
                            '    Do While wDoc.Bookmarks.Exists("Pr_Mt_" & VryBkc) = True
                            '        VryBkc = VryBkc + 1
                            '        VryBkc = Microsoft.VisualBasic.Strings.Format(VryBkc, xFormatStr) 'VryBkc + 1
                            '    Loop
                            '    wDoc.Bookmarks.Add("Pr_Mt_" & VryBkc, Hrng.Duplicate)
                            'End If
                        End If
                    End If
                Next I
            End If
        Next
        PairingRoutine = lstBraketRange
    End Function
    Public Function FindPairing(wDoc As Word.Document, Optional Gvnrng As Word.Range = Nothing)
        Dim sRng As Word.Range

        If Gvnrng Is Nothing = False Then
            ClrPair(Gvnrng.Duplicate)
            '######### Pairing Routine ############
            PairingRoutine(wDoc, "(", ")", False, Gvnrng.Duplicate)
            PairingRoutine(wDoc, "[", "]", False, Gvnrng.Duplicate)
            PairingRoutine(wDoc, "{", "}", False, Gvnrng.Duplicate)
            PairingRoutine(wDoc, "8220", "8221", True, Gvnrng.Duplicate)
            PairingRoutine(wDoc, "8216", "8217", True, Gvnrng.Duplicate)
            '#########  Ends here    #################
        Else
            For Each sRng In wDoc.StoryRanges
                If sRng.StoryType = Word.WdStoryType.wdMainTextStory Or
                    (wDoc.Footnotes.Count > 0 And sRng.StoryType = Word.WdStoryType.wdFootnotesStory) Or
                    (wDoc.Endnotes.Count > 0 And sRng.StoryType = Word.WdStoryType.wdEndnotesStory) Then
                    ClrPair(sRng.Duplicate)
                    '######### Pairing Routine ###########
                    PairingRoutine(wDoc, "(", ")", False, sRng.Duplicate)
                    PairingRoutine(wDoc, "[", "]", False, sRng.Duplicate)
                    PairingRoutine(wDoc, "{", "}", False, sRng.Duplicate)
                    PairingRoutine(wDoc, "8220", "8221", True, sRng.Duplicate)
                    PairingRoutine(wDoc, "8216", "8217", True, sRng.Duplicate)
                    '#########  Ends here ################
                End If
            Next
        End If
        '    On Error Resume Next
        '    wDoc.Variables.Add "PCheck", "Nl"
        'Err.Clear()

        wDoc.StoryRanges(Word.WdStoryType.wdMainTextStory).Select()
        wDoc.Application.Selection.HomeKey(Word.WdUnits.wdStory)
        wDoc.Application.Selection.SetRange(wDoc.Application.Selection.Start, wDoc.Application.Selection.Start)
        ''GothruMismatch
    End Function

    Function ClrPair(GetRng As Word.Range)
        Dim bk As Integer
        If GetRng.Bookmarks.Count > 0 Then
            For bk = GetRng.Bookmarks.Count To 1 Step -1
                If InStr(1, GetRng.Bookmarks(bk).Name, "pr_mt_", vbTextCompare) > 0 Or
                          InStr(1, GetRng.Bookmarks(bk).Name, "frmt_", vbTextCompare) > 0 Then
                    GetRng.Bookmarks(bk).Delete()
                End If
            Next bk
        End If
    End Function
End Module
