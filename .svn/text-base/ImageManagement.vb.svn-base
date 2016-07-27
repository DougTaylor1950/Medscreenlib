Imports ImageMagick
Namespace ImageManagement
    Public Class ImageManagement
        Private Shared Sub SavePDFAsJPGLocal(ByVal Filename As String, ByVal DestinationPath As String, ByVal settings As MagickReadSettings)
            Using images As New MagickImageCollection()
                Dim OutputFilename As String = IO.Path.GetFileNameWithoutExtension(Filename)
                images.Read(Filename, settings)
                Dim page As Integer = 1

                For Each image As MagickImage In images
                    ' Write page to file that contains the page number
                    image.Format = MagickFormat.Jpg
                    image.Write(DestinationPath & page & "-" & OutputFilename & ".jpg")
                    ' Writing to a specific format works the same as for a single image
                    page += 1
                Next
            End Using
        End Sub

        ''' <developer></developer>
        ''' <summary>
        ''' Save a PDF to images as a set of page images
        ''' </summary>
        ''' <param name="Filename"></param>
        ''' <param name="DestinationPath"></param>
        ''' <param name="settings"></param>
        ''' <remarks></remarks>
        ''' <revisionHistory></revisionHistory>
        Public Overloads Shared Sub SavePDFAsJPGPages(ByVal Filename As String, ByVal DestinationPath As String, ByVal settings As MagickReadSettings)
            SavePDFAsJPGLocal(Filename, DestinationPath, settings)
        End Sub

        Public Overloads Shared Sub SavePDFAsJPGPages(ByVal Filename As String, ByVal DestinationPath As String)
            Dim settings As New MagickReadSettings()
            ' Settings the density to 300 dpi will create an image with a better quality
            settings.Density = New MagickGeometry(300, 300)
            SavePDFAsJPGLocal(Filename, DestinationPath, settings)
        End Sub
    End Class
End Namespace