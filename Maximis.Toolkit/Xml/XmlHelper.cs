using Maximis.Toolkit.IO;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Xsl;

namespace Maximis.Toolkit.Xml
{
    public static class XmlHelper
    {
        private static readonly Regex REGEX_DEFAULT_NAMESPACE = new Regex(" xmlns=[\"'].*?[\"']", RegexOptions.Compiled);

        private static readonly Regex REGEX_EMPTY_DEFAULT_NAMESPACE = new Regex(" xmlns=[\"']{2}", RegexOptions.Compiled);

        public static XmlDocument CombineXml(string rootElement, params string[] xmlChunks)
        {
            XmlDocument result = new XmlDocument();
            result.LoadXml(string.Format("<{0}/>", rootElement));

            foreach (string xmlChunk in xmlChunks)
            {
                if (string.IsNullOrEmpty(xmlChunk)) continue;
                XmlDocument chunkDoc = new XmlDocument();
                chunkDoc.LoadXml(xmlChunk);

                XmlElement el = result.CreateElement(chunkDoc.DocumentElement.Name);
                result.DocumentElement.AppendChild(el);
                el.InnerXml = chunkDoc.DocumentElement.InnerXml;
            }

            return result;
        }

        public static string GetAttributeValue(XmlNode node, string attrName)
        {
            XmlAttribute attr = node.Attributes[attrName];
            return attr == null ? null : attr.InnerText;
        }

        public static string GetNodeText(XmlNode xmlNode)
        {
            return xmlNode == null ? null : xmlNode.InnerText;
        }

        public static XslCompiledTransform GetXsltFromAssembly(Assembly assembly, string xsltFile, bool trustedSettings = false)
        {
            XslCompiledTransform xct = new XslCompiledTransform();
            using (Stream stream = assembly.GetManifestResourceStream(xsltFile))
            {
                using (StreamReader sr = new StreamReader(stream))
                {
                    XmlDocument xtDoc = new XmlDocument();
                    xtDoc.LoadXml(sr.ReadToEnd());
                    xct.Load(xtDoc, trustedSettings ? XsltSettings.TrustedXslt : XsltSettings.Default, null);
                }
            }
            return xct;
        }

        public static XslCompiledTransform GetXsltFromFile(string path, bool trustedSettings = false)
        {
            XslCompiledTransform xct = new XslCompiledTransform();
            using (StreamReader sr = File.OpenText(path))
            {
                XmlDocument xtDoc = new XmlDocument();
                xtDoc.LoadXml(sr.ReadToEnd());
                xct.Load(xtDoc, trustedSettings ? XsltSettings.TrustedXslt : XsltSettings.Default, null);
            }
            return xct;
        }

        public static XslCompiledTransform GetXsltFromString(string xsltXml, bool trustedSettings = false)
        {
            XmlDocument xtDoc = new XmlDocument();
            xtDoc.LoadXml(xsltXml);
            XslCompiledTransform xct = new XslCompiledTransform();
            xct.Load(xtDoc, trustedSettings ? XsltSettings.TrustedXslt : XsltSettings.Default, null);
            return xct;
        }

        public static string RemoveDefaultNamespace(string xml)
        {
            return REGEX_DEFAULT_NAMESPACE.Replace(xml, string.Empty);
        }

        public static string RemoveEmptyDefaultNamespace(string xml)
        {
            return REGEX_EMPTY_DEFAULT_NAMESPACE.Replace(xml, string.Empty);
        }

        public static string XslTransform(XslCompiledTransform xslCompiledTransform, string sourceXml,
            XsltArgumentList xsltArgumentList = null, XmlResolver xmlResolver = null)
        {
            XmlDocument xd = new XmlDocument();
            xd.LoadXml(sourceXml);
            return XslTransform(xslCompiledTransform, xd, xsltArgumentList, xmlResolver);
        }

        public static string XslTransform(XslCompiledTransform xslCompiledTransform, XmlDocument xmlDocument,
            XsltArgumentList xsltArgumentList = null, XmlResolver xmlResolver = null)
        {
            StringBuilder outXmlSb = new StringBuilder();
            using (StringWriterUtf8 outXml = new StringWriterUtf8(outXmlSb))
            using (XmlWriter outXw = new XmlTextWriter(outXml))
            {
                xslCompiledTransform.Transform(xmlDocument,
                    xsltArgumentList == null ? default(XsltArgumentList) : xsltArgumentList, outXw, xmlResolver);
            }
            return outXmlSb.ToString();
        }
    }
}