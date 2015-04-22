using System.Text.RegularExpressions;
using System.Xml;

namespace Maximis.Toolkit.Xml
{
    public class NamespacedXmlDocument
    {
        private static readonly Regex REGEX_NAMESPACE = new Regex("xmlns(:?)(.*?)=[\"'](.*?)[\"']",
            RegexOptions.Compiled);

        private NamespacedXmlDocument()
        {
        }

        public XmlDocument XmlDocument { get; set; }

        public XmlNamespaceManager XmlNamespaceManager { get; set; }

        public static NamespacedXmlDocument FromXmlString(string xml, string defaultPrefix = "dflt")
        {
            NamespacedXmlDocument result = new NamespacedXmlDocument();

            // Load the XML into the XmlDocument
            result.XmlDocument = new XmlDocument();
            result.XmlDocument.LoadXml(xml);

            // Get all the namespaces from the XML
            result.XmlNamespaceManager = new XmlNamespaceManager(result.XmlDocument.NameTable);
            foreach (Match m in REGEX_NAMESPACE.Matches(xml))
            {
                string prefix = m.Groups[2].Value;
                if (string.IsNullOrEmpty(prefix))
                {
                    result.XmlNamespaceManager.AddNamespace(defaultPrefix, m.Groups[3].Value);
                }
                else
                {
                    result.XmlNamespaceManager.AddNamespace(prefix, m.Groups[3].Value);
                }
            }

            return result;
        }

        public XmlNodeList SelectNodes(string xPath)
        {
            return XmlDocument.SelectNodes(xPath, XmlNamespaceManager);
        }

        public XmlNode SelectSingleNode(string xPath)
        {
            return XmlDocument.SelectSingleNode(xPath, XmlNamespaceManager);
        }
    }
}