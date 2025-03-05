using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System;
using DevExpress.XtraRichEdit.Export.Html;
using System.Windows.Forms;
using System.IO;
using DevExpress.XtraRichEdit.Export;
using System.Diagnostics;

namespace ExportOnlyBodyContent {
    static class Program {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static CssPropertiesExportType cssExportType;
        static ExportRootTag htmlExportType;
        static void Main() {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            using (RichEditDocumentServer wordProcessor = new RichEditDocumentServer())
            {
                wordProcessor.LoadDocument("Document.docx");
                Document document = wordProcessor.Document;

                Console.WriteLine("Do you want to export HTML body? y/n");
                string answer1 = Console.ReadLine()?.ToLower();
                if (answer1 == "y") { htmlExportType = ExportRootTag.Body; }
                else { htmlExportType = ExportRootTag.Html; }
                Console.WriteLine("Choose one of the CSS options:\r\nInclude CSS in a <STYLE> tag. - 1\r\nSave style sheets in a separate CSS file - 2\r\nPlace CSS as an attribute to an HTML tag - 3");
                string answer2 = Console.ReadLine()?.ToLower();
                switch (answer2)
                {
                    case "1": cssExportType = CssPropertiesExportType.Style; break;
                    case "2": cssExportType = CssPropertiesExportType.Link; break;
                    case "3": cssExportType = CssPropertiesExportType.Inline; break;
                }
                string fileName = "Exported.html";
                string stringHtml = String.Empty;
                ExportHtml(out stringHtml, fileName, wordProcessor);
                SaveFile(fileName, stringHtml);

                var p = new Process();
                p.StartInfo = new ProcessStartInfo("Exported.html")
                {
                    UseShellExecute = true
                };
                p.Start();
            }
        }
        private static void SaveFile(string fileName, string value)
        {
            using (FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                using (StreamWriter writer = new StreamWriter(stream))
                {
                    writer.Write(value);
                }
            }
        }
        #region #exporting
        private static void ExportHtml(out string stringHtml, string fileName, RichEditDocumentServer wordProcessor)
        {
            stringHtml = String.Empty;
            HtmlDocumentExporterOptions options = new HtmlDocumentExporterOptions();
            options.ExportRootTag = htmlExportType;
            options.CssPropertiesExportType = cssExportType;
            options.TargetUri = Path.GetFileNameWithoutExtension(fileName);
            Document document = wordProcessor.Document;
            var uriProvider = new MyUriProvider(Path.GetDirectoryName(Application.StartupPath));
            stringHtml = document.GetHtmlText(document.Range, uriProvider, options);
        }
        #endregion #exporting
    }

}