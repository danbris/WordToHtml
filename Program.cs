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
using Encoder = System.Drawing.Imaging.Encoder;

namespace WordToHtml
{
    internal class Program
    {
        private static Dictionary<string, Bitmap> _imagesDictionary = new Dictionary<string, Bitmap>();

        private static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();
            ConvertToHtml(
                @"c:\Work\DocuPerformer\Main\src\Application\AppData\Generated\BI2\DesignStudio Reports BW\AZAP-0EPM_OPEN_ITEMS_ANALYTICS-20200702-Doc_EN-Com_EN-BI2.docx",
                @"c:\Work\DocuPerformer\test html");

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
                        
                    };
                    Console.WriteLine($"Prepared settings at {sw.ElapsedMilliseconds}ms");
                    var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    Console.WriteLine($"Converted to HTML at {sw.ElapsedMilliseconds}ms");
                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    //var html = new XDocument(
                    //    new XDocumentType("html", null, null, null),
                    //    htmlElement);
                    var html = new XDocument(htmlElement);
                    Console.WriteLine($"Created XDocument at {sw.ElapsedMilliseconds}ms");
                    var htmlString = html.ToString(SaveOptions.None);
                    File.WriteAllText(destFileName.FullName, htmlString);
                    Console.WriteLine($"Save to disk at {sw.ElapsedMilliseconds}ms");
                    
                    var htmlFixer = new HtmlFixer();
                    File.WriteAllText(destFileName.FullName + "fixer.html", htmlFixer.FormatHtmlForConfluence(htmlString));
                    Console.WriteLine($"Fixed HTML format at {sw.ElapsedMilliseconds}ms");
                    System.Diagnostics.Process.Start(destFileName.FullName + "fixer.html");
                }
            }
        }
    }
}