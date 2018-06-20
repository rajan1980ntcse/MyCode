Public Class frmSortHarvardReference
    Public dictRefSortInfo As Dictionary(Of Integer, clsRefInfo)
    Private Sub frmSortHarvardReference_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            If dictRefInfo.Count > 0 Then
                For Each pair As KeyValuePair(Of Integer, clsRefInfo) In ModRefUtility.dictRefInfo
                    Dim n As Integer = dgvRefList.Rows.Add()
                    dgvRefList.Rows.Item(n).Cells(0).Value = pair.Value.oRefRng.Text
                Next
            End If


        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub cbSortType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cbSortType.SelectedIndexChanged
        If dgvSortedList.RowCount > 0 Then
            dgvSortedList.Rows.Clear()
        End If
        If cbSortType.Text <> "" Then
            Select Case cbSortType.Text
                Case "order of Surname"
                    dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.rCondText).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                Case "order by character by character"
                    dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.oRefRng.Text.Replace(Chr(13), "")).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                Case "order by number of authors"
                    dictRefSortInfo = dictRefInfo.OrderBy(Function(x) x.Value.olRefAuthors.Count).ToDictionary(Function(x) x.Key, Function(x) x.Value)
            End Select
            If dictRefInfo.Count > 0 Then
                For Each pair As KeyValuePair(Of Integer, clsRefInfo) In dictRefSortInfo
                    ' lvSortedRefList.Items.Add(pair.Value.ranRef.Text)
                    Dim n As Integer = dgvSortedList.Rows.Add()
                    dgvSortedList.Rows.Item(n).Cells(0).Value = pair.Value.oRefRng.Text
                Next
            End If
        End If
    End Sub
End Class