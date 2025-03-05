Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit.Export.Html
Imports DevExpress.XtraRichEdit.Export
Imports System.Diagnostics
Imports System.Runtime.InteropServices

Namespace ExportOnlyBodyContent

    Friend Module Program

        ''' <summary>
        ''' The main entry point for the application.
        ''' </summary>
        Private cssExportType As CssPropertiesExportType

        Private htmlExportType As ExportRootTag

        Sub Main()
            'Application.EnableVisualStyles();
            'Application.SetCompatibleTextRenderingDefault(false);
            'Application.Run(new Form1());
            Using wordProcessor As RichEditDocumentServer = New RichEditDocumentServer()
                wordProcessor.LoadDocument("Document.docx")
                Dim document As Document = wordProcessor.Document
                Console.WriteLine("Do you want to export HTML body? y/n")
                Dim answer1 As String = Console.ReadLine()?.ToLower()
                If Equals(answer1, "y") Then
                    htmlExportType = ExportRootTag.Body
                Else
                    htmlExportType = ExportRootTag.Html
                End If

                Console.WriteLine("Choose one of the CSS options:" & Microsoft.VisualBasic.Constants.vbCrLf & "Include CSS in a <STYLE> tag. - 1" & Microsoft.VisualBasic.Constants.vbCrLf & "Save style sheets in a separate CSS file - 2" & Microsoft.VisualBasic.Constants.vbCrLf & "Place CSS as an attribute to an HTML tag - 3")
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
                Dim stringHtml As String = [String].Empty
                Program.ExportHtml(stringHtml, fileName, wordProcessor)
                Program.SaveFile(fileName, stringHtml)
                Dim p = New Process()
                p.StartInfo = New ProcessStartInfo("Exported.html") With {.UseShellExecute = True}
                p.Start()
            End Using
        End Sub

        Private Sub SaveFile(ByVal fileName As String, ByVal value As String)
            Using stream As FileStream = New FileStream(fileName, FileMode.Create, FileAccess.Write)
                Using writer As StreamWriter = New StreamWriter(stream)
                    writer.Write(value)
                End Using
            End Using
        End Sub

#Region "#exporting"
        Private Sub ExportHtml(<Out> ByRef stringHtml As String, ByVal fileName As String, ByVal wordProcessor As RichEditDocumentServer)
            stringHtml = [String].Empty
            Dim options As HtmlDocumentExporterOptions = New HtmlDocumentExporterOptions()
            options.ExportRootTag = htmlExportType
            options.CssPropertiesExportType = cssExportType
            options.TargetUri = Path.GetFileNameWithoutExtension(fileName)
            Dim document As Document = wordProcessor.Document
            Dim uriProvider = New MyUriProvider(Path.GetDirectoryName(Application.StartupPath))
            stringHtml = document.GetHtmlText(document.Range, uriProvider, options)
        End Sub
#End Region  ' #exporting
    End Module
End Namespace
