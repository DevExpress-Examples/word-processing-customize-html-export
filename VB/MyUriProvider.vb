Imports System.Text
Imports System.Drawing
Imports System.Drawing.Imaging
Imports DevExpress.Office.Services
Imports DevExpress.Office.Utils
Imports DevExpress.Drawing

Namespace ExportOnlyBodyContent

    Public Class MyUriProvider
        Implements IUriProvider

        Private rootDirecory As String

        Public Sub New(ByVal rootDirectory As String)
            If [String].IsNullOrEmpty(rootDirectory) Then Exceptions.ThrowArgumentException("rootDirectory", rootDirectory)
            rootDirecory = rootDirectory
        End Sub

        Public Function CreateCssUri(ByVal rootUri As String, ByVal styleText As String, ByVal relativeUri As String) As String Implements IUriProvider.CreateCssUri
            Dim cssDir As String = [String].Format("{0}\{1}", rootDirecory, rootUri.Trim("/"c))
            If Not Directory.Exists(cssDir) Then Directory.CreateDirectory(cssDir)
            Dim cssFileName As String = [String].Format("{0}\style.css", cssDir)
            File.AppendAllText(cssFileName, styleText)
            Return Me.GetRelativePath(cssFileName)
        End Function

        Public Function CreateImageUri(ByVal rootUri As String, ByVal image As OfficeImage, ByVal relativeUri As String) As String Implements IUriProvider.CreateImageUri
            Dim imagesDir As String = [String].Format("{0}\{1}", rootDirecory, rootUri.Trim("/"c))
            If Not Directory.Exists(imagesDir) Then Directory.CreateDirectory(imagesDir)
            Dim imageName As String = [String].Format("{0}\{1}.png", imagesDir, Guid.NewGuid())
            image.DXImage.Save(imageName, DXImageFormat.Png)
            Return Me.GetRelativePath(imageName)
        End Function

        Private Function GetRelativePath(ByVal path As String) As String
            Dim substring As String = path.Substring(rootDirecory.Length)
            Return substring.Replace("\", "/").Trim("/"c)
        End Function
    End Class
End Namespace
