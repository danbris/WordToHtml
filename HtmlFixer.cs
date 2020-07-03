using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace WordToHtml
{
    public class HtmlFixer
    {
        private readonly Dictionary<string, string> _styles = new Dictionary<string, string>();

        public string FormatHtmlForConfluence(string htmlString)
        {
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(htmlString);
            var nodes = xmlDoc.ChildNodes.OfType<XmlNode>();

            var head = xmlDoc.DocumentElement.ChildNodes.OfType<XmlNode>().FirstOrDefault(x => x.Name == "head");
            var style = head.ChildNodes.OfType<XmlNode>().FirstOrDefault(x => x.Name == "style");
            var body = xmlDoc.DocumentElement.ChildNodes.OfType<XmlNode>().FirstOrDefault(x => x.Name == "body");

            SetupStyleDictionary(style.InnerText);
            head.RemoveChild(style);

            ReplaceClassWithActualStyle(body, xmlDoc);

            return xmlDoc.InnerXml;
            using (var sw = new StringWriter())
            {
                //using (var xw = new XmlTextWriter(sw))
                //{//saving the xml document like this will decrease the file size, but will result in an UNARRANGED file
                //    xmlDoc.Save(xw);
                //}

                xmlDoc.Save(sw);

                return sw.ToString();
            }
        }

        private void ReplaceClassWithActualStyle(XmlNode node, XmlDocument doc)
        {
            FixClassAttribute(node, doc);

            foreach (var nodeChildNode in node.ChildNodes.OfType<XmlNode>())
            {
                ReplaceClassWithActualStyle(nodeChildNode, doc);
            }
        }

        private void FixClassAttribute(XmlNode node, XmlDocument doc)
        {
            if (node.Attributes == null) return;

            var classAttribute = node.Attributes.OfType<XmlAttribute>().FirstOrDefault(x => x.Name == "class");
            var styleAttribute = node.Attributes?.OfType<XmlAttribute>().FirstOrDefault(x => x.Name == "style");
            if (classAttribute == null)
            {
                if (!_styles.ContainsKey(node.Name)) return;
                if (styleAttribute == null)
                {
                    styleAttribute = doc.CreateAttribute("style");
                    styleAttribute.Value = _styles[node.Name].Replace(Environment.NewLine, string.Empty);
                    node.Attributes.Append(styleAttribute);
                }
                else
                {
                    //check which properties are already set and avoid duplicates
                    styleAttribute.Value += _styles[node.Name].Replace(Environment.NewLine, string.Empty);
                }

                return;
            }
            
            if (styleAttribute == null)
            {
                styleAttribute = doc.CreateAttribute("style");
                styleAttribute.Value = _styles[node.Name+"."+classAttribute.Value].Replace(Environment.NewLine,string.Empty);
                node.Attributes.Append(styleAttribute);
            }
            else
            {
                //check which properties are already set and avoid duplicates
                styleAttribute.Value += _styles[node.Name + "." + classAttribute.Value]
                    .Replace(Environment.NewLine, string.Empty);
            }

           
            node.Attributes.Remove(classAttribute);
        }

      
        private void SetupStyleDictionary(string style)
        {
            var styles = style.Split('}');
            foreach (var s in styles)
            {
                var splits = s.Split('{');
                if (splits.Length != 2) continue;

                _styles.Add(splits.First().Trim(), FormatStyleDefinition(splits.Last().Trim()));
            }
        }

        private string FormatStyleDefinition(string definition)
        {
            var formatedDefinition = string.Empty;
            var splittedDefinitions = definition.Split(';');
           
            foreach (var style in splittedDefinitions)
            {
                if (!style.Contains(':')) continue;
                var styleParts = style.Split(':');
                formatedDefinition += styleParts[0].Trim().Replace(Environment.NewLine,string.Empty) + ":" + styleParts[1].Trim().Replace(Environment.NewLine, string.Empty) + ";";
            }

            return formatedDefinition;
        }
    }
}
