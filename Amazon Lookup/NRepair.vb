Imports System.Data.SqlClient
Imports FlexFlow
Public Class NRepair
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
    Dim dt, dt2, dt3 As New DataTable
    Dim ASIN As String
    Dim Prtfam As String
    Dim nid As Integer = "0"
    Private Sub textbox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        AMZgetpartneeded()
        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
            MsgBox("Flex ID has already been used, you can update the original and add comments as needed")
        Else
            AMZpartDetails()
        End If
    End Sub
    Private Sub NRepair_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lbldate.Text = Now()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            TextBox2.Select()
        End If
        If nid = 0 Then
            updateprtneeded1()
        Else
            updateprtneeded2()
        End If
        lbllpn.Text = "LPN"
        lblpn.Text = "PN"
        RichTextBox1.Text = ""
        RichTextBox2.Text = ""
        RichTextBox3.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox2.Select()
    End Sub
    Private Sub AMZpartDetails()
        dt.Clear()
        Dim sql As New SqlCommand("sp_AMZpartDetails", cn)
        sql.CommandType = CommandType.StoredProcedure
        sql.Parameters.AddWithValue("@flexID", TextBox2.Text)
        sql.Connection = cn
        cn.Open()
        sql.ExecuteNonQuery()
        Dim da As New SqlDataAdapter(sql)
        da.Fill(dt)
        cn.Close()
        If dt.Rows.Count < 1 Then
            MsgBox("This Flexid is not found in the system. Make sure you scanned the correct id and if you did notify your supervisor")
            TextBox2.Select()
            TextBox2.SelectAll()
            Return
        Else
            lbllpn.Text = dt.Rows(0).Item(1)
            lblpn.Text = dt.Rows(0).Item(2).ToString
            RichTextBox1.Text = dt.Rows(0).Item(3).ToString
            Prtfam = dt.Rows(0).Item(4).ToString
            nid = 0
        End If
    End Sub
    Private Sub updateprtneeded1()
        Dim subdt As String = Format(Now, "yyyyMMddhhmmss")
        Dim sql2 As New SqlCommand("sp_updateprtneeded1", cn2)
        sql2.CommandType = CommandType.StoredProcedure
        sql2.Parameters.AddWithValue("@flxid", TextBox2.Text)
        sql2.Parameters.AddWithValue("@LPN", lbllpn.Text)
        sql2.Parameters.AddWithValue("@Main_PN", lblpn.Text)
        sql2.Parameters.AddWithValue("@Main_Desc", RichTextBox1.Text)
        sql2.Parameters.AddWithValue("@prtfam", Prtfam)
        sql2.Parameters.AddWithValue("@Tstcom", RichTextBox3.Text)
        sql2.Parameters.AddWithValue("@tst_dt", Now)
        sql2.Parameters.AddWithValue("@sub_pn", TextBox1.Text)
        sql2.Parameters.AddWithValue("@sub_description", RichTextBox2.Text)
        sql2.Connection = cn2
        cn2.Open()
        sql2.ExecuteNonQuery()
        cn2.Close()
    End Sub
    Private Sub AMZgetpartneeded()
        dt2.Clear()
        Dim sql As New SqlCommand("sp_getpartneeded", cn2)
        sql.CommandType = CommandType.StoredProcedure
        sql.Parameters.AddWithValue("@flxID", TextBox2.Text)
        sql.Connection = cn2
        cn2.Open()
        sql.ExecuteNonQuery()
        Dim da As New SqlDataAdapter(sql)
        da.Fill(dt2)
        cn2.Close()
        If dt2 IsNot Nothing AndAlso dt2.Rows.Count > 0 Then
            lbllpn.Text = dt2.Rows(0).Item(1)
            lblpn.Text = dt2.Rows(0).Item(2)
            RichTextBox1.Text = dt2.Rows(0).Item(3)
            Prtfam = dt2.Rows(0).Item(4)
            RichTextBox2.Text = dt2.Rows(0).Item(8)
            RichTextBox3.Text = dt2.Rows(0).Item(5)
            TextBox1.Text = dt2.Rows(0).Item(7)
            nid = 1
        Else
            Return
        End If
    End Sub
    Private Sub updateprtneeded2()
        Dim subdt As String = Format(Now, "yyyyMMddhhmmss")
        Dim sql2 As New SqlCommand("sp_updateprtneeded2", cn2)
        sql2.CommandType = CommandType.StoredProcedure
        sql2.Parameters.AddWithValue("@flxid", TextBox2.Text)
        sql2.Parameters.AddWithValue("@Main_Desc", RichTextBox1.Text)
        sql2.Parameters.AddWithValue("@Tstcom", RichTextBox3.Text)
        sql2.Parameters.AddWithValue("@sub_pn", TextBox1.Text)
        sql2.Parameters.AddWithValue("@sub_description", RichTextBox2.Text)
        sql2.Parameters.AddWithValue("@pic1", DBNull.Value)
        sql2.Parameters.AddWithValue("@pic2", DBNull.Value)
        sql2.Parameters.AddWithValue("@pic3", DBNull.Value)
        sql2.Parameters.AddWithValue("@Website", DBNull.Value)
        sql2.Parameters.AddWithValue("@auditdate", DBNull.Value)
        sql2.Parameters.AddWithValue("@auditid", DBNull.Value)

        sql2.Connection = cn2
        cn2.Open()
        sql2.ExecuteNonQuery()
        cn2.Close()
    End Sub

End Class