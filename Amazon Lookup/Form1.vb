Imports System.Data.SqlClient
Imports System.Drawing.Text
Imports FlexFlow
Imports System.DirectoryServices

Public Class Form1
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
    Dim dt, dt2, dt3, dt4, dt5 As New DataTable
    Dim ASIN, Prtfam, shtdwn, login, acct, ct, multi As String

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'get user name and title of associate logged into computer and check for engineer to be in the title
        '(note: two functions below are used)
        ToolStripButton3.Visible = False
        ToolStripSeparator3.Visible = False
        login = GetRealloginFromAD(UsernameToFind:=Environment.UserName)
        AMZacct()
        Try
            If dt4 IsNot Nothing AndAlso dt4.Rows.Count > 0 Then
                acct = dt4.Rows(0).Item(1)
                Select Case acct
                    Case "Engineering"
                    Case "Operations"
                    Case "Buyers"
                End Select
            Else
                InfoCollectionToolStripMenuItem.Visible = False
                ToolStripButton1.Visible = False
                ToolStripButton2.Visible = False
            End If
        Catch
            InfoCollectionToolStripMenuItem.Visible = False
            ToolStripButton1.Visible = False
            ToolStripButton2.Visible = False
        End Try
        AMZshutdwn()
        Timer1.Start()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Text = "FLXID"
        Label2.Text = "LPN"
        Label3.Text = "NA"
        Dim pfc As PrivateFontCollection = New PrivateFontCollection
        pfc.AddFontFile("\\10.162.102.14\common\AMAZON\AMZ39.ttf")
        'pfc.AddFontFile(".\AMZ39.ttf")
        AMZprtView()
        AMZmani()
        If Label1.Text = "FLXID" Then
        Else
            AMZprtget()
        End If
        Try
            WebBrowser1.Navigate(New Uri("https://www.amazon.com/gp/product/" + ASIN))
            If Err.Number <> 0 Then
                'MsgBox("Error :" & Err.Description)  'Display error message
            End If

            Label10.Text = Prtfam

            TextBox2.Text = "*" & ASIN & "*"
            TextBox2.Font = New Font(pfc.Families(0), 24)
            TextBox1.Clear()
            TextBox1.Focus()
        Catch
            Return
        End Try
        If multi = 1 Then
            multi = 0
            Messagebx.Show()
        End If
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        AMZshutdwn()
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Multi_Return.Show()
    End Sub

    Private Sub NeedsRepairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NeedsRepairToolStripMenuItem.Click
        NRepair.Show()
    End Sub
    Private Sub InfoCollectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoCollectionToolStripMenuItem.Click
        IC.Show()
    End Sub
    Private Sub AMZprtView()
        dt.Clear()
        Dim sql As New SqlCommand("sp_AMZprtView", cn)
        sql.CommandType = CommandType.StoredProcedure
        sql.Parameters.AddWithValue("@FLXLPN", TextBox1.Text)
        sql.Connection = cn
        cn.Open()
        sql.ExecuteNonQuery()
        Dim da As New SqlDataAdapter(sql)
        da.Fill(dt)
        cn.Close()
        If dt.Rows.Count < 1 Then
            Return
        Else
            Label1.Text = dt.Rows(0).Item(0)
            Label2.Text = dt.Rows(0).Item(1)
            ASIN = dt.Rows(0).Item(2).ToString
            Prtfam = dt.Rows(0).Item(3).ToString
            TextBox2.Clear()
        End If
    End Sub
    Private Sub AMZprtget()
        dt2.Clear()
        Dim sql2 As New SqlCommand("sp_getPartNumber", cn2)
        sql2.CommandType = CommandType.StoredProcedure
        sql2.Parameters.AddWithValue("@PN", ASIN)
        sql2.Connection = cn2
        cn2.Open()
        sql2.ExecuteNonQuery()
        Dim da2 As New SqlDataAdapter(sql2)
        da2.Fill(dt2)
        cn2.Close()
        If dt2.Rows.Count < 1 Then
            If Prtfam = "25200" Then
                Label3.Text = "Contact Engineering"
            Else
                Label3.Text = ""
            End If
            Return
        Else
            Label3.Text = dt2.Rows(0).Item(4)
        End If
    End Sub
    Private Sub AMZshutdwn()
        dt3.Clear()
        Dim sql3 As New SqlCommand("sp_shutdwn", cn2)
        sql3.CommandType = CommandType.StoredProcedure
        sql3.Connection = cn2
        cn2.Open()
        sql3.ExecuteNonQuery()
        Dim da3 As New SqlDataAdapter(sql3)
        da3.Fill(dt3)
        cn2.Close()
        shtdwn = dt3.Rows(0).Item(0)
        If shtdwn <> 0 Then
            Me.Close()
        Else
        End If
        For Each f As Form In Me.OwnedForms
            If f.Modal = True Then
                'your code here
            End If
        Next
    End Sub
    Private Sub AMZacct()
        dt4.Clear()
        Dim sql4 As New SqlCommand("sp_getAMZacct", cn2)
        sql4.CommandType = CommandType.StoredProcedure
        sql4.Parameters.AddWithValue("@uname", login)
        sql4.Connection = cn2
        cn2.Open()
        sql4.ExecuteNonQuery()
        Dim da4 As New SqlDataAdapter(sql4)
        da4.Fill(dt4)
        cn2.Close()
    End Sub
    Private Sub AMZmani()
        dt5.Clear()
        Dim sql5 As New SqlCommand("sp_getAMZmanifest", cn2)
        sql5.CommandType = CommandType.StoredProcedure
        sql5.Parameters.AddWithValue("@LPN", Label2.Text)
        sql5.Connection = cn2
        cn2.Open()
        sql5.ExecuteNonQuery()
        Dim da5 As New SqlDataAdapter(sql5)
        da5.Fill(dt5)
        cn2.Close()
        Try
            Label5.Text = dt5.Rows(0).Item(4).ToString
            Label7.Text = dt5.Rows(0).Item(5).ToString
            Label9.Text = dt5.Rows(0).Item(6).ToString
            If dt5 IsNot Nothing AndAlso dt5.Rows.Count > 1 Then
                ct = dt5.Rows.Count
                Messagebx.lblnumber.Text = ct
                multi = 1
                ToolStripButton3.Visible = True
                ToolStripSeparator3.Visible = True
            Else
                ToolStripButton3.Visible = False
                ToolStripSeparator3.Visible = False
            End If
        Catch
            Return
        End Try
    End Sub
    Public Function GetRealloginFromAD(ByVal UsernameToFind As String) As String
        Using searcher As New DirectorySearcher(New DirectoryEntry())
            searcher.PageSize = 1000
            searcher.SearchScope = SearchScope.Subtree
            searcher.Filter = "(&(samAccountType=805306368)(sAMAccountName=" & UsernameToFind & "))"
            Using Results As SearchResultCollection = searcher.FindAll
                If Results Is Nothing OrElse Results.Count <> 1 Then
                    Throw New ApplicationException("Invalid number of results returned – either no users were found or more than one user account was found")
                End If
                Using UserDE As DirectoryEntry = Results(0).GetDirectoryEntry
                    Return CStr(UserDE.Properties("samaccountname").Value)
                End Using
            End Using
        End Using
    End Function
    'Private Shared Sub Main()
    '    'Try
    '    '    ' Open the file using a stream reader.
    '    '    Using sr As New StreamReader("\\10.162.102.14\common\AMAZON\AMZlookup.txt")
    '    '        Dim line As String
    '    '        ' Read the stream to a string and write the string to the console.
    '    '        line = sr.ReadToEnd()
    '    '        If line = "0" Then
    '    '        Else
    '    '            Form1.Close()
    '    '        End If
    '    '    End Using
    '    'Catch e As Exception
    '    'End Try
    'End Sub
End Class
