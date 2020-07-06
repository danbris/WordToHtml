using DevExpress.XtraRichEdit;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using Encoder = System.Drawing.Imaging.Encoder;

namespace WordToHtml
{
    internal class Program
    {
       private static List<string> _images = new List<string>();

       private static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();
            //ConvertToHtml(
            //    @"c:\Work\DocuPerformer\Test tables.docx",
            //    @"c:\Work\DocuPerformer\test html");

            ConvertToHtml(
                @"C:\Work\DocuPerformer\Main\src\Application\AppData\Generated\BI2\Queries\REP-0CCA_C03_Q1001-20200703-Doc_EN-Com_EN-BI2.docx",
                @"c:\Work\DocuPerformer\test html");

            //var doc = new HtmlDocument();
            //doc.Load(@"C:\Work\DocuPerformer\Main\src\Application\AppData\Generated\BI2\Queries\REP-VMM_H001_QT001-20200703-Doc_EN-Com_EN-BI2.html");
            //doc.Save(@"c:\Work\DocuPerformer\test html\test with agility pack.html");

            
            Debug.WriteLine($"Conversion took  {sw.ElapsedMilliseconds}ms");
            Console.ReadLine();
        }

        public static void ConvertToHtml(string file, string outputDirectory)
        {
            var sw = Stopwatch.StartNew();
            var fi = new FileInfo(file);
            Console.WriteLine(fi.Name);
            var byteArray = File.ReadAllBytes(fi.FullName);

            using (var memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                Console.WriteLine($"Loaded file in memory at {sw.ElapsedMilliseconds}ms");
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

                            if (imageFormat == null)
                                return null;

                            var imageFileName = imageDirectoryName + "\\" + imageInfo.AltText + "." + extension;
                            var imageSource = $"/download/attachments/89790660/{imageInfo.AltText + "." + extension}";

                            try
                            {
                                if (!_images.Contains(imageFileName))
                                {
                                    imageInfo.Bitmap.Save(imageFileName, imageFormat);
                                    _images.Add(imageFileName);
                                }
                            }
                            catch (ExternalException)
                            {
                                return null;
                            }

                            var img = new XElement(Xhtml.img,
                                 new XAttribute(NoNamespace.src, imageSource),
                                 imageInfo.ImgStyleAttribute,
                                 imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);

                            return img;
                        }

                    };
                    var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    Console.WriteLine($"Converted to HTML at {sw.ElapsedMilliseconds}ms");
                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    //var html = new XDocument(
                    //    new XDocumentType("html", null, null, null),
                    //    htmlElement);
                    var html = new XDocument(htmlElement);
                   
                    var htmlString = html.ToString(SaveOptions.None);
                    File.WriteAllText(destFileName.FullName, htmlString);
                    
                    var htmlFixer = new HtmlFixer();
                    File.WriteAllText(destFileName.FullName + "fixer.html", htmlFixer.FormatHtmlForConfluence(htmlString));
                    Console.WriteLine($"Fixed HTML format at {sw.ElapsedMilliseconds}ms");
                    Process.Start(destFileName.FullName + "fixer.html");
                }
            }
        }
    }
}