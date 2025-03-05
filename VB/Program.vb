Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Export.Html
Imports System.Windows.Forms
Imports System.IO
Imports DevExpress.XtraRichEdit.Export

Namespace ExportOnlyBodyContent
    Module Program
        Private cssExportType As CssPropertiesExportType
        Private htmlExportType As ExportRootTag

        Sub Main()
            Using wordProcessor As New RichEditDocumentServer()
                wordProcessor.LoadDocument("Document.docx")
                Dim document As Document = wordProcessor.Document

                Console.WriteLine("Do you want to export HTML body? y/n")
                Dim answer1 As String = Console.ReadLine()?.ToLower()
                If answer1 = "y" Then
                    htmlExportType = ExportRootTag.Body
                Else
                    htmlExportType = ExportRootTag.Html
                End If

                Console.WriteLine("Choose one of the CSS options:" & vbCrLf &
                                  "Include CSS in a <STYLE> tag. - 1" & vbCrLf &
                                  "Save style sheets in a separate CSS file - 2" & vbCrLf &
                                  "Place CSS as an attribute to an HTML tag - 3")
                Dim answer2 As String = Console.ReadLine()?.ToLower()
                Select Case answer2
                    Case "1"
                        cssExportType = CssPropertiesExportType.Style
                    Case "2"
                        cssExportType = CssPropertiesExportType.Link
                    Case "3"
                        cssExportType = CssPropertiesExportType.Inline
                End Select

                Dim fileName As String = "Exported.html"
                Dim stringHtml As String = String.Empty
                ExportHtml(stringHtml, fileName, wordProcessor)
                SaveFile(fileName, stringHtml)

                Dim p As New Process()
                p.StartInfo = New ProcessStartInfo("Exported.html") With {
                    .UseShellExecute = True
                }
                p.Start()
            End Using
        End Sub

        Private Sub SaveFile(fileName As String, value As String)
            Using stream As New FileStream(fileName, FileMode.Create, FileAccess.Write)
                Using writer As New StreamWriter(stream)
                    writer.Write(value)
                End Using
            End Using
        End Sub

        Private Sub ExportHtml(ByRef stringHtml As String, fileName As String, wordProcessor As RichEditDocumentServer)
            stringHtml = String.Empty
            Dim options As New HtmlDocumentExporterOptions()
            options.ExportRootTag = htmlExportType
            options.CssPropertiesExportType = cssExportType
            options.TargetUri = Path.GetFileNameWithoutExtension(fileName)
            Dim document As Document = wordProcessor.Document
            Dim uriProvider As New MyUriProvider(Path.GetDirectoryName(Application.StartupPath))
            stringHtml = document.GetHtmlText(document.Range, uriProvider, options)
        End Sub
    End Module
End Namespace
