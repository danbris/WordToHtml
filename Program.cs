using DevExpress.XtraRichEdit;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Encoder = System.Drawing.Imaging.Encoder;

namespace WordToHtml
{
    internal class Program
    {
        private static Dictionary<string, Bitmap> _imagesDictionary = new Dictionary<string, Bitmap>();

        private static void Main(string[] args)
        {
            //ConvertToHtml(
            //    @"c:\Work\DocuPerformer\Main\src\Application\AppData\Generated\BI2\Queries\REP-IMOPCA_M01_Q0001-20200701-Doc_EN-Com_EN-BI2.docx",
            //    @"c:\Work\DocuPerformer\test html");

            ConvertToHtmlUsingDevexpress(
                @"c:\Work\DocuPerformer\Main\src\Application\AppData\Generated\BI2\DesignStudio Reports BW\AZAP-0EPM_OPEN_ITEMS_ANALYTICS-20200702-Doc_EN-Com_EN-BI2.docx",
                @"c:\Work\DocuPerformer\test html\AZAP-0EPM_OPEN_ITEMS_ANALYTICS-20200702-Doc_EN-Com_EN-BI2.html");
        }

        public static void ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            Console.WriteLine(fi.Name);
            var byteArray = File.ReadAllBytes(fi.FullName);

            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (var wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                   var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                    if (!string.IsNullOrEmpty(outputDirectory))
                    {
                        var di = new DirectoryInfo(outputDirectory);
                        if (!di.Exists) throw new OpenXmlPowerToolsException("Output directory does not exist");

                        destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
                    }

                    var imageDirectoryName =
                        destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    var imageCounter = 0;

                    var pageTitle = fi.FullName;
                    var part = wDoc.CoreFilePropertiesPart;
                    if (part != null)
                        pageTitle = (string) part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;

                    // TODO: Determine max-width from size of content area.
                    var settings = new WmlToHtmlConverterSettings
                    {
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            var localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                           
                            var extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                            {
                                imageFormat = ImageFormat.Png;
                            }
                            else if (extension == "gif")
                            {
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "bmp")
                            {
                                imageFormat = ImageFormat.Bmp;
                            }
                            else if (extension == "jpeg")
                            {
                                imageFormat = ImageFormat.Jpeg;
                            }
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            var imageFileName = imageDirectoryName + "/image" + imageCounter + "." + extension;
                            var imageSource = $"/download/attachments/89790344/image{imageCounter}.{extension}";

                            try
                            {
                                //if (!_imagesDictionary.Values.Any(x=> imageInfo.Bitmap.Equals(x)))
                                //{
                                    imageInfo.Bitmap.Save(imageFileName, imageFormat);
                                    _imagesDictionary.Add(imageFileName, imageInfo.Bitmap);
                                    imageCounter++;
                                //}
                            }
                            catch (ExternalException)
                            {
                                return null;
                            }

                            //var imageSource = localDirInfo.Name + "/image" +
                            //                  imageCounter + "." + extension;

                           var img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageSource),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

                            return img;
                        }
                    };
                    var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    //var html = new XDocument(
                    //    new XDocumentType("html", null, null, null),
                    //    htmlElement);
                    var html = new XDocument(htmlElement);

                    var htmlString = html.ToString(SaveOptions.None);
                    File.WriteAllText(destFileName.FullName, htmlString);
                }
            }
        }

        public static void ConvertToHtmlUsingDevexpress(string file, string outputDirectory)
        {
            var richEditDocumentServer = new RichEditDocumentServer();
            richEditDocumentServer.LoadDocument(file, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
            using (var htmlFileStream = new FileStream(outputDirectory, FileMode.Create))
            {
                richEditDocumentServer.SaveDocument(htmlFileStream, DevExpress.XtraRichEdit.DocumentFormat.Html);
            }

            System.Diagnostics.Process.Start(outputDirectory);
        }
    }
}