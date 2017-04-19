Imports System.Data.SqlClient
Imports System.Drawing.Text
Imports FlexFlow

Public Class Multi_Return
    Dim pw As String = "reports"
    Dim constr As String = "Data Source=lount013;Initial Catalog= SDW" &
                                         ";User ID=dsreports;Password='" & pw & "'" &
                                         ";Persist Security Info=True;Pooling=True"
    Dim cn As New SqlConnection(constr)
    Dim encryptedPW As String = "usgLc58WBSnsZCQvTm0JOg=="
    Dim pw2 As String = UtilityMgr.DecryptString(encryptedPW)
    Dim constr2 As String = "Data Source=lount013;Initial Catalog=p_AmazonEng;" &
                                "User ID=AmazonEng;Password='" & pw2 & "';" &
                                "Persist Security Info=True;Pooling=True;"
    Dim cn2 As New SqlConnection(constr2)
    Dim dt As New DataTable
    Private Sub Multi_Return_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AMZmani()
        Me.DataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        For i = 0 To DataGridView1.Rows.Count - 1
            Dim r As DataGridViewRow = DataGridView1.Rows(i)
            r.Height = 100
        Next
        Me.DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
    End Sub
    Private Sub AMZmani()
        dt.Clear()
        Dim sql5 As New SqlCommand("sp_getAMZmanifest", cn2)
        sql5.CommandType = CommandType.StoredProcedure
        sql5.Parameters.AddWithValue("@LPN", Form1.Label2.Text)
        sql5.Connection = cn2
        cn2.Open()
        sql5.ExecuteNonQuery()
        Dim da5 As New SqlDataAdapter(sql5)
        da5.Fill(dt)
        DataGridView1.DataSource = dt
        cn2.Close()
    End Sub
End Class