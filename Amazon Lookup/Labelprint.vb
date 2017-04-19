'Imports System.Drawing
'Imports System.Drawing.Printing
'Module Labelprint

'    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
'        Static currentChar As Integer
'        Static currentLine As Integer
'        Dim textfont As Font = Richtextbox1.Font
'        Dim h, w As Integer
'        Dim left, top As Integer
'        With PrintDocument1.DefaultPageSettings
'            h = .PaperSize.Height - .Margins.Top - .Margins.Bottom
'            w = .PaperSize.Width - .Margins.Left - .Margins.Right
'            left = PrintDocument1.DefaultPageSettings.Margins.Left
'            top = PrintDocument1.DefaultPageSettings.Margins.Top
'        End With
'        e.Graphics.DrawRectangle(Pens.Blue, New Rectangle(left, top, w, h))
'        If PrintDocument1.DefaultPageSettings.Landscape Then
'            Dim a As Integer
'            a = h
'            h = w
'            w = a
'        End If
'        Dim lines As Integer = CInt(Math.Round(h / textfont.Height))
'        Dim b As New Rectangle(left, top, w, h)
'        Dim format As StringFormat
'        If Not Richtextbox1.WordWrap Then
'            format = New StringFormat(StringFormatFlags.NoWrap)
'            format.Trimming = StringTrimming.EllipsisWord
'            Dim i As Integer
'            For i = currentLine To Math.Min(currentLine + lines, Richtextbox1.Lines.Length - 1)
'                e.Graphics.DrawString(Richtextbox1.Lines(i), textfont, Brushes.Black, New RectangleF(left, top + textfont.Height * (i - currentLine), w, textfont.Height), format)
'            Next
'            currentLine += lines
'            If currentLine >= Richtextbox1.Lines.Length Then
'                e.HasMorePages = False
'                currentLine = 0
'            Else
'                e.HasMorePages = True
'            End If
'            Exit Sub
'        End If
'        format = New StringFormat(StringFormatFlags.LineLimit)
'        Dim line, chars As Integer
'        e.Graphics.MeasureString(Mid(Richtextbox1.Text, currentChar + 1), textfont, New SizeF(w, h), format, chars, line)
'        If currentChar + chars < Richtextbox1.Text.Length Then
'            If Richtextbox1.Text.Substring(currentChar + chars, 1) <> " " And Richtextbox1.Text.Substring(currentChar + chars, 1) <> vbLf Then
'                While chars > 0
'                    Richtextbox1.Text.Substring(currentChar + chars, 1)
'                    Richtextbox1.Text.Substring(currentChar + chars, 1)
'                    chars -= 1
'                End While
'                chars += 1
'            End If
'        End If
'        e.Graphics.DrawString(Richtextbox1.Text.Substring(currentChar, chars), textfont, Brushes.Black, b, format)
'        currentChar = currentChar + chars
'        If currentChar < Richtextbox1.Text.Length Then
'            e.HasMorePages = True
'        Else
'            e.HasMorePages = False
'            currentChar = 0
'        End If
'    End Sub
'End Module
