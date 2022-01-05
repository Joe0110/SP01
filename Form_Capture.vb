Public Class Form_Capture

    '==============================
    '   20200807 - 已經不使用這個方法了
    '==============================

    Dim img As Image
    Dim imgThumb As Image = Nothing

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ms As IO.MemoryStream = New IO.MemoryStream
        Dim imgPicBox As New Bitmap(PictureBox1.Image)
        ' imgPicBox.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        imgPicBox.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Dim bytBLOBData(CInt(ms.Length - 1)) As Byte
        ms.Position = 0
        ms.Read(bytBLOBData, 0, CInt(ms.Length))
        ms.Close()

        'My.Computer.FileSystem.WriteAllBytes(Application.StartupPath + "\MyImage.jpg", bytBLOBData, False)
        My.Computer.FileSystem.WriteAllBytes(Form_Spec.Get_Full_Small_Name, bytBLOBData, False)
    End Sub
    Public Sub Save_file()
        Dim ms As IO.MemoryStream = New IO.MemoryStream
        Dim imgPicBox As New Bitmap(PictureBox1.Image)
        ' imgPicBox.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
        imgPicBox.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        Dim bytBLOBData(CInt(ms.Length - 1)) As Byte
        ms.Position = 0
        ms.Read(bytBLOBData, 0, CInt(ms.Length))
        ms.Close()

        'My.Computer.FileSystem.WriteAllBytes(Application.StartupPath + "\MyImage.jpg", bytBLOBData, False)
        My.Computer.FileSystem.WriteAllBytes(Form_Spec.Get_Full_Small_Name, bytBLOBData, False)
    End Sub

    Public Sub dw_bmp()

        Dim mBmpImg As Bitmap
        Dim g As Graphics
        Dim b As New SolidBrush(Color.Gray)

        mBmpImg = New Bitmap(PictureBox1.Width, PictureBox1.Height)
        g = Graphics.FromImage(mBmpImg)
        b.Color = Color.White
        g.FillRectangle(b, 5, 5, PictureBox1.Width - 10, PictureBox1.Height - 10)
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
        b.Color = Color.Red
        g.DrawString("My Image", New Font("Arial", 24), b, 10, 40)

        PictureBox1.Image = mBmpImg

    End Sub


    'Public Class Form1
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)  '點擊執行
        Opacity = 0 '視窗透明(不然會照到自己)

        Dim ScrnPB As PictureBox = PictureBox1

        Dim ScrnSize As Size = My.Computer.Screen.Bounds.Size
        Dim ScrnImage As New Bitmap(ScrnSize.Width, ScrnSize.Height)

        Dim g As Graphics = Graphics.FromImage(ScrnImage)
        g.CopyFromScreen(New Point(Me.Left + 8, Me.Top + 30), New Point(0, 0), ScrnSize) '以自己視窗為始點修正偏移

        Dim dc As IntPtr = g.GetHdc
        g.ReleaseHdc(dc)

        With ScrnPB '使大小相同
            .Size = ScrnSize
            .Image = ScrnImage
        End With

        Opacity = 70 '視窗透明度70%
    End Sub
    'End Class


    Private Sub DrawinPbToolStripMenuItem_Click(sender As Object, e As EventArgs)
        dw_bmp()
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Save_file()
    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Capture_IMG()

    End Sub
    Public Sub Capture_IMG()
        Opacity = 0 '視窗透明(不然會照到自己)

        Dim ScrnPB As PictureBox = PictureBox1

        Dim ScrnSize As Size = PictureBox1.Size
        Dim ScrnImage As New Bitmap(ScrnSize.Width, ScrnSize.Height)

        Dim g As Graphics = Graphics.FromImage(ScrnImage)
        'g.CopyFromScreen(New Point(PictureBox1.Location.X, PictureBox1.Location.Y), New Point(0, 0), ScrnSize) '以自己視窗為始點修正偏移
        'g.CopyFromScreen(New Point(Me.Left + 8, Me.Top + 30), New Point(0, 0), ScrnSize) '以自己視窗為始點修正偏移
        g.CopyFromScreen(New Point(Me.Left + 38, Me.Top + 45), New Point(0, 0), ScrnSize) '以自己視窗為始點修正偏移

        Dim dc As IntPtr = g.GetHdc
        g.ReleaseHdc(dc)

        With ScrnPB '使大小相同
            .Size = ScrnSize
            .Image = ScrnImage
        End With

        Opacity = 70 '視窗透明度70%
    End Sub

    Private Sub Form_Capture_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Location = Form_Spec_CHK_Ctl.Location

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim xx As Size
        xx.Height = 100
        xx.Width = 50

        PictureBox1.Size = xx
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Picture_save("d:\TEST_123.png")
        'Picture_ReSize_save(100, 200, Form_Spec.Get_Full_Small_Name)
    End Sub
    Public Sub Picture_ReSize_save(ByVal xx As Integer, ByVal yy As Integer, ByVal zz As String)

        'picturebox => img
        Dim img As Image = Me.PictureBox1.Image

        Dim bmp As New Bitmap(xx, yy)

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(bmp)
        gfx.DrawImage(img, 0, 0, xx, yy)

        gfx.Dispose()

        bmp.Save(zz, System.Drawing.Imaging.ImageFormat.Png)

        bmp.Dispose()
    End Sub
    Public Sub Picture_save(ByVal zz As String)

        'picturebox => img
        Dim img As Image = Me.PictureBox1.Image

        Dim bmp As New Bitmap(PictureBox1.Width, PictureBox1.Height)

        ' img => bmp
        Dim gfx As Graphics = Graphics.FromImage(bmp)
        gfx.DrawImage(img, 0, 0, PictureBox1.Width, PictureBox1.Height)

        gfx.Dispose()

        bmp.Save(zz, System.Drawing.Imaging.ImageFormat.Png)

        bmp.Dispose()
    End Sub
End Class