Imports System.IO
Imports DevExpress.Office.Services
Imports DevExpress.Office.Utils
Imports DevExpress.Drawing

Namespace ExportOnlyBodyContent
    Public Class MyUriProvider
        Implements IUriProvider

        Private rootDirectory As String

        Public Sub New(rootDirectory As String)
            If String.IsNullOrEmpty(rootDirectory) Then
                Exceptions.ThrowArgumentException("rootDirectory", rootDirectory)
            End If
            Me.rootDirectory = rootDirectory
        End Sub

        Public Function CreateCssUri(rootUri As String, styleText As String, relativeUri As String) As String Implements IUriProvider.CreateCssUri
            Dim cssDir As String = String.Format("{0}\{1}", Me.rootDirectory, rootUri.Trim("/"c))
            If Not Directory.Exists(cssDir) Then
                Directory.CreateDirectory(cssDir)
            End If
            Dim cssFileName As String = String.Format("{0}\style.css", cssDir)
            File.AppendAllText(cssFileName, styleText)
            Return GetRelativePath(cssFileName)
        End Function

        Public Function CreateImageUri(rootUri As String, image As OfficeImage, relativeUri As String) As String Implements IUriProvider.CreateImageUri
            Dim imagesDir As String = String.Format("{0}\{1}", Me.rootDirectory, rootUri.Trim("/"c))
            If Not Directory.Exists(imagesDir) Then
                Directory.CreateDirectory(imagesDir)
            End If
            Dim imageName As String = String.Format("{0}\{1}.png", imagesDir, Guid.NewGuid())
            image.DXImage.Save(imageName, DXImageFormat.Png)
            Return GetRelativePath(imageName)
        End Function

        Private Function GetRelativePath(path As String) As String
            Dim substring As String = path.Substring(Me.rootDirectory.Length)
            Return substring.Replace("\", "/").Trim("/"c)
        End Function
    End Class
End Namespace
