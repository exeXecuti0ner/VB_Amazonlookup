Imports System.Data.SqlClient
Imports FlexFlow
Imports System.Runtime.InteropServices
Imports System.DirectoryServices

Public Class IC
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
    Dim usename As String

    Private Sub IC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        usename = GetRealNameFromAD(UsernameToFind:=Environment.UserName)
        AMZgetengpartneeded()
    End Sub
    Private Sub DataGridView1_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) _
     Handles DataGridView1.RowHeaderMouseDoubleClick
        lblflxid.Text = DataGridView1.SelectedRows(0).Cells(0).Value.ToString
        lbllpn.Text = DataGridView1.SelectedRows(0).Cells(1).Value.ToString
        lblpn.Text = DataGridView1.SelectedRows(0).Cells(2).Value.ToString
        RTB1.Text = DataGridView1.SelectedRows(0).Cells(3).Value.ToString
        RichTextBox1.Text = DataGridView1.SelectedRows(0).Cells(5).Value.ToString
        TextBox1.Text = DataGridView1.SelectedRows(0).Cells(7).Value.ToString
        RTB2.Text = DataGridView1.SelectedRows(0).Cells(8).Value.ToString
    End Sub
    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If DataGridView1.GetCellCount(DataGridViewElementStates.Selected) > 0 Then
            Try
                ' Add the selection to the clipboard.
                Clipboard.SetDataObject(DataGridView1.GetClipboardContent())
            Catch ex As ExternalException
            End Try
        End If
    End Sub
    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Find.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        updateprtneeded2()
        lblflxid.Text = ""
        lbllpn.Text = ""
        lblpn.Text = ""
        RTB1.Text = ""
        RTB2.Text = ""
        RichTextBox1.Text = ""
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        AMZgetengpartneeded()
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub
    Private Sub AMZgetengpartneeded()
        dt.Clear()
        Dim sql As New SqlCommand("sp_getengpartneeded", cn2)
        sql.CommandType = CommandType.StoredProcedure
        sql.Connection = cn2
        cn2.Open()
        sql.ExecuteNonQuery()
        Dim da As New SqlDataAdapter(sql)
        da.Fill(dt)
        DataGridView1.DataSource = dt
        '9-14
        DataGridView1.Columns(9).Visible = False
        DataGridView1.Columns(10).Visible = False
        DataGridView1.Columns(11).Visible = False
        DataGridView1.Columns(12).Visible = False
        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(14).Visible = False
        cn2.Close()
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        Printing_note.Show()
    End Sub

    Private Sub updateprtneeded2()
        Dim subdt As String = Format(Now, "yyyyMMddhhmmss")
        Dim sql2 As New SqlCommand("sp_updateprtneeded2", cn2)
        sql2.CommandType = CommandType.StoredProcedure
        sql2.Parameters.AddWithValue("@flxid", lblflxid)
        sql2.Parameters.AddWithValue("@Main_Desc", RTB1.Text)
        sql2.Parameters.AddWithValue("@Tstcom", RichTextBox1.Text)
        sql2.Parameters.AddWithValue("@sub_pn", TextBox1.Text)
        sql2.Parameters.AddWithValue("@sub_description", RTB2.Text)
        sql2.Parameters.AddWithValue("@pic1", TextBox3.Text)
        sql2.Parameters.AddWithValue("@pic2", TextBox4.Text)
        sql2.Parameters.AddWithValue("@pic3", TextBox5.Text)
        sql2.Parameters.AddWithValue("@Website", TextBox2.Text)
        sql2.Parameters.AddWithValue("@auditdate", Now)
        sql2.Parameters.AddWithValue("@auditid", usename)
        sql2.Connection = cn2
        cn2.Open()
        sql2.ExecuteNonQuery()
        cn2.Close()
    End Sub
    Public Function GetRealNameFromAD(ByVal UsernameToFind As String) As String
        Using searcher As New DirectorySearcher(New DirectoryEntry())
            searcher.PageSize = 1000
            searcher.SearchScope = SearchScope.Subtree
            searcher.Filter = "(&(samAccountType=805306368)(sAMAccountName=" & UsernameToFind & "))"
            Using Results As SearchResultCollection = searcher.FindAll
                If Results Is Nothing OrElse Results.Count <> 1 Then
                    Throw New ApplicationException("Invalid number of results returned – either no users were found or more than one user account was found")
                End If
                Using UserDE As DirectoryEntry = Results(0).GetDirectoryEntry
                    Return CStr(UserDE.Properties("displayname").Value)
                End Using
            End Using
        End Using
    End Function
End Class