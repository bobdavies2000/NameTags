Imports cv = OpenCvSharp
Imports cvext = OpenCvSharp.Extensions
Imports System.Drawing
Imports System.IO
Public Class NameTags
    Dim picPage As PictureBox
    Dim cols As Integer = 790, rows As Integer = 1010
    Dim printFont As Font
    Dim nametagMat As cv.Mat
    Dim r1 = New cv.Rect(58, 48, 320, 220)
    Private Sub clearNameTag()
        nametagMat = New cv.Mat(rows, cols, cv.MatType.CV_8UC3, cv.Scalar.White)
        For i = 0 To 4 - 1
            Dim r = New cv.Rect(r1.X, r1.Y + i * (r1.Height + 23), r1.Width, r1.Height)
            nametagMat.Rectangle(r, cv.Scalar.Black, 1)
            Dim r2 = New cv.Rect(r1.X + r1.Width + 40, r1.Y + i * (r1.Height + 23), r1.Width, r1.Height)
            nametagMat.Rectangle(r2, cv.Scalar.Black, 1)
        Next
    End Sub
    Private Sub displayPage()
        printFont = New Font("Arial", 10)

        Dim redDevil = cv.Cv2.ImRead("../../../RedDevil.jpg")
        redDevil = redDevil.Resize(New cv.Size(redDevil.Width * 0.4, redDevil.Height * 0.4))
        For i = 0 To 10
            Dim picDir As New DirectoryInfo("../../../pics" + CStr(i))
            If picDir.Exists = False Then Exit For
            Dim index As Integer = 0, xIncr As Integer = 0, yIncr As Integer = 0, tweak As Integer = 0
            For Each fn In picDir.GetFiles("*.jpg")
                Dim pic = cv.Cv2.ImRead(fn.FullName)
                If index Mod 2 = 0 Then xIncr = 0 Else xIncr = r1.width + 40
                Dim newSize = New cv.Size(CInt(pic.Width / pic.Height * r1.height), CInt(r1.height))
                If pic.Height > r1.height Then pic = pic.Resize(newSize)
                pic = pic.Resize(New cv.Size(pic.Width * 0.8, pic.Height * 0.8))
                If index Mod 2 = 0 Then tweak = 0 Else tweak = 80
                Dim r = New cv.Rect(r1.x + xIncr + r1.width - pic.Width, r1.y + yIncr, pic.Width, pic.Height)
                pic.CopyTo(nametagMat(r))
                Dim personName = fn.Name.Substring(0, Len(fn.Name) - 4)
                Dim split = personName.Split(" ")
                Dim rdRect = New cv.Rect(r1.x + 20 + xIncr, r1.y + 20 + yIncr, redDevil.Width, redDevil.Height)
                redDevil.CopyTo(nametagMat(rdRect))
                If split(0).Length > 5 Then
                    cv.Cv2.PutText(nametagMat, split(0), New cv.Point(CInt(r1.X + xIncr), r.Y + pic.Height - 15), cv.HersheyFonts.HersheyPlain, 2.8,
                                   cv.Scalar.Black, 2, cv.LineTypes.AntiAlias)
                Else
                    cv.Cv2.PutText(nametagMat, split(0), New cv.Point(CInt(r1.X + xIncr), r.Y + pic.Height - 15), cv.HersheyFonts.HersheyPlain, 3.5,
                                   cv.Scalar.Black, 2, cv.LineTypes.AntiAlias)
                End If
                cv.Cv2.PutText(nametagMat, split(1), New cv.Point(CInt(r1.X + xIncr), r.Y + pic.Height + 25), cv.HersheyFonts.HersheyPlain, 3.0,
                                   cv.Scalar.Black, 2, cv.LineTypes.AntiAlias)
                If index Mod 2 = 1 Then yIncr += r1.height + 23
                index += 1
            Next
            cv.Cv2.ImShow("Nametagmat", nametagMat)
            cv.Cv2.WaitKey(5000)
            clearNameTag()
        Next
    End Sub
    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        displayPage()
        pd.Print()
    End Sub
    Private Sub pd_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles pd.PrintPage
        Dim g As Graphics = e.Graphics
        e.PageSettings.Landscape = False
        g.DrawImage(picPage.Image, 0, 0)
    End Sub
    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        displayPage()
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PrintPreviewDialog1 = New PrintPreviewDialog()

        PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        PrintPreviewDialog1.Location = New System.Drawing.Point(29, 29)
        PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        PrintPreviewDialog1.UseAntiAlias = True

        Me.Location = New Point(10, 10)
        picPage = New PictureBox()
        picPage.Size = New Size(cols, rows)
        picPage.Location = New Point(10, 50)
        picPage.Image = New Bitmap(cols, rows, Imaging.PixelFormat.Format24bppRgb)
        clearNameTag()
        cvext.BitmapConverter.ToBitmap(nametagMat, picPage.Image)
    End Sub
End Class
