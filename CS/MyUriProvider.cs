using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using DevExpress.Office.Services;
using DevExpress.XtraRichEdit.Utils;
using DevExpress.Office.Utils;
using DevExpress.Drawing;


namespace ExportOnlyBodyContent {
    public class MyUriProvider : IUriProvider {
        string rootDirecory;
        public MyUriProvider(string rootDirectory) {
            if(String.IsNullOrEmpty(rootDirectory))
                Exceptions.ThrowArgumentException("rootDirectory", rootDirectory);
            this.rootDirecory = rootDirectory;
        }

        public string CreateCssUri(string rootUri, string styleText, string relativeUri) {
            string cssDir = String.Format("{0}\\{1}", this.rootDirecory, rootUri.Trim('/'));
            if(!Directory.Exists(cssDir))
                Directory.CreateDirectory(cssDir);
            string cssFileName = String.Format("{0}\\style.css", cssDir);
            File.AppendAllText(cssFileName, styleText);
            return GetRelativePath(cssFileName);
        }
        public string CreateImageUri(string rootUri, OfficeImage image, string relativeUri) {
            string imagesDir = String.Format("{0}\\{1}", this.rootDirecory, rootUri.Trim('/'));
            if(!Directory.Exists(imagesDir))
                Directory.CreateDirectory(imagesDir);
            string imageName = String.Format("{0}\\{1}.png", imagesDir, Guid.NewGuid());
            image.DXImage.Save(imageName, DXImageFormat.Png);
            return GetRelativePath(imageName);
        }
        string GetRelativePath(string path) {
            string substring = path.Substring(this.rootDirecory.Length);
            return substring.Replace("\\", "/").Trim('/');
        }
    }
}
