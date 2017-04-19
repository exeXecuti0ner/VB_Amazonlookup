Public Class Printing_note

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintDocument()
    End Sub

#Region " PrintDocument "
    Public Sub PrintDocument()
        'Create an instance of our printer class
        Dim printer As New PCPrint()
        'Set the font we want to use
        printer.PrinterFont = New Font("Verdana", 24)
        'Set the TextToPrint property
        printer.TextToPrint = RichTextBox1.Text
        'Issue print command
        printer.Print()
    End Sub

#End Region
End Class
