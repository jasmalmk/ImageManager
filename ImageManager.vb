Imports Microsoft.VisualBasic
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Drawing.Imaging

Public Class ImageManager
    Sub New()

    End Sub
    Sub New(PostedFile As HttpPostedFile, DestinationFolderPath As String, newfileName As String, newWidth As Integer, newHeight As Integer, newFormat As ImageFromat)
        Dim bmpPostedImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PostedFile.InputStream)
        Dim newBmpImage As System.Drawing.Bitmap = SizeImage(bmpPostedImage, newWidth, newHeight)
        newBmpImage.Save(DestinationFolderPath & "/" & newfileName & retImageExt(newFormat), switchImageFormat(newFormat))
    End Sub
    Sub New(PostedFile As HttpPostedFile, DestinationFolderPath As String, newfileName As String, newFormat As ImageFromat)
        Dim bmpPostedImage As System.Drawing.Bitmap = New System.Drawing.Bitmap(PostedFile.InputStream)
        Dim newBmpImage As System.Drawing.Bitmap = SizeImage(bmpPostedImage, bmpPostedImage.Width, bmpPostedImage.Height)
        newBmpImage.Save(DestinationFolderPath & "/" & newfileName & retImageExt(newFormat), switchImageFormat(newFormat))
    End Sub

    Private Function SizeImage(ByVal img As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap
        Dim newImage As Bitmap = New Bitmap(width, height)
        Dim gr As Graphics = Graphics.FromImage(newImage)
        Using gr
            gr.SmoothingMode = SmoothingMode.HighQuality
            gr.InterpolationMode = InterpolationMode.HighQualityBicubic
            gr.PixelOffsetMode = PixelOffsetMode.HighQuality
            gr.DrawImage(img, New Rectangle(0, 0, width, height))
        End Using
        Return newImage
    End Function

    Public Function checkFileType(postedFile As HttpPostedFile) As Boolean
        If (postedFile.ContentType.ToLower() <> "image/jpg" And
                   postedFile.ContentType.ToLower() <> "image/jpeg" And
                   postedFile.ContentType.ToLower() <> "image/pjpeg" And
                   postedFile.ContentType.ToLower() <> "image/gif" And
                   postedFile.ContentType.ToLower() <> "image/x-png" And
                   postedFile.ContentType.ToLower() <> "image/png") Then

            Return False
        Else
            Return True
        End If
    End Function


    Private Function switchImageFormat(Format As ImageFromat) As System.Drawing.Imaging.ImageFormat
        Select Case Format
            Case ImageFromat.JPEG
                Return System.Drawing.Imaging.ImageFormat.Jpeg
            Case ImageFromat.BMP
                Return System.Drawing.Imaging.ImageFormat.Bmp
            Case ImageFromat.PNG
                Return System.Drawing.Imaging.ImageFormat.Png
            Case ImageFromat.Icon
                Return System.Drawing.Imaging.ImageFormat.Icon
            Case ImageFromat.GIF
                Return System.Drawing.Imaging.ImageFormat.Gif
            Case Else
                Return Nothing
        End Select
    End Function
    Private Function retImageExt(Format As ImageFromat) As String
        Select Case Format
            Case ImageFromat.JPEG
                Return ".jpeg"
            Case ImageFromat.BMP
                Return ".bmp"
            Case ImageFromat.PNG
                Return ".png"
            Case ImageFromat.Icon
                Return ".ico"
            Case ImageFromat.GIF
                Return ".gif"
            Case Else
                Return Nothing
        End Select
    End Function
    Public Enum ImageFromat
        JPEG
        BMP
        Icon
        GIF
        PNG
        MEMBMP
    End Enum

    Protected Sub GenerateThumbnail(path As String)
        Dim image As System.Drawing.Image = System.Drawing.Image.FromFile(path)
        Using thumbnail As System.Drawing.Image = image.GetThumbnailImage(100, 100, New System.Drawing.Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback), IntPtr.Zero)
            Using memoryStream As New MemoryStream()
                thumbnail.Save(memoryStream, ImageFormat.Png)
                Dim bytes As [Byte]() = New [Byte](memoryStream.Length - 1) {}
                memoryStream.Position = 0
                memoryStream.Read(bytes, 0, CInt(bytes.Length))
                Dim base64String As String = Convert.ToBase64String(bytes, 0, bytes.Length)
                'Image2.ImageUrl = "data:image/png;base64," & base64String
                'Image2.Visible = True
                'Dim ThumbImg As Image = "data:image/png;base64," & base64String
            End Using
        End Using
    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

End Class
