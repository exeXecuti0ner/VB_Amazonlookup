Public Class Find

    Private Sub Find_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Searchbox.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        IC.DataGridView1.ClearSelection()
        Dim temp As Integer = 0
        For i As Integer = 0 To IC.DataGridView1.RowCount - 1
            For j As Integer = 0 To IC.DataGridView1.ColumnCount - 1
                If IC.DataGridView1.Rows(i).Cells(j).Value.ToString = Searchbox.Text Then
                    IC.DataGridView1.FirstDisplayedScrollingRowIndex = i
                    IC.DataGridView1.Rows(i).Cells(j).Selected = True
                    GoTo 22
                    temp = 1
                    MsgBox("Item found")
                End If
            Next
        Next
        If temp = 0 Then
            MsgBox("Item not found")
        End If
22:
        Me.Close()
    End Sub

End Class