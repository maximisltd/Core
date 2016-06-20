using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Maximis.Toolkit.Office
{
    public static class WordDocHelper
    {
        private static readonly XDeclaration dec = new XDeclaration("1.0", "UTF-8", "yes");
        private static readonly XProcessingInstruction pi = new XProcessingInstruction("mso-application", "progid=\"Word.Document\"");
        private static readonly XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly string relContentType = "application/vnd.openxmlformats-package.relationships+xml";

        public static void ConvertDocXToFlatXml(Stream input, Stream output)
        {
            // Based on code from http://blogs.msdn.com/b/ericwhite/archive/2008/09/29/transforming-open-xml-documents-to-flat-opc-format.aspx
            using (Package package = Package.Open(input))
            {
                XDocument xd = new XDocument(dec, pi, new XElement(pkg + "package", new XAttribute(XNamespace.Xmlns + "pkg", pkg.ToString()),
                        package.GetParts().Select(part =>
                      {
                          if (part.ContentType.EndsWith("xml"))
                          {
                              using (Stream s = (part.GetStream()))
                                  return new XElement(pkg + "part", new XAttribute(pkg + "name", part.Uri),
                                      new XAttribute(pkg + "contentType", part.ContentType),
                                      new XElement(pkg + "xmlData", XElement.Load(s)));
                          }
                          else
                          {
                              using (BinaryReader binaryReader = new BinaryReader(part.GetStream()))
                              {
                                  byte[] byteArray = binaryReader.ReadBytes((int)binaryReader.BaseStream.Length);
                                  string base64String = (System.Convert.ToBase64String(byteArray)).Select
                                  ((c, i) => new { Character = c, Chunk = i / 76 }).GroupBy(c => c.Chunk)
                                  .Aggregate(new StringBuilder(), (s, i) => s.Append(i.Aggregate(new StringBuilder(),
                                  (seed, it) => seed.Append(it.Character), sb => sb.ToString())).Append(Environment.NewLine), s => s.ToString());
                                  return new XElement(pkg + "part",
                                  new XAttribute(pkg + "name", part.Uri),
                                  new XAttribute(pkg + "contentType", part.ContentType),
                                  new XAttribute(pkg + "compression", "store"),
                                  new XElement(pkg + "binaryData", base64String));
                              }
                          }
                      })));
                xd.Save(output);
            }
        }

        public static void ConvertFlatXmlToDocX(Stream input, Stream output)
        {
            // Based on code from http://blogs.msdn.com/b/ericwhite/archive/2008/09/29/transforming-flat-opc-format-to-open-xml-documents.aspx

            using (Package package = Package.Open(output, FileMode.Create))
            {
                XDocument xdWordXml = XDocument.Load(input);

                // add all parts (but not relationships)
                foreach (XElement xmlPart in xdWordXml.Root.Elements().Where(p => (string)p.Attribute(pkg + "contentType") != relContentType))
                {
                    Uri uri = new Uri((string)xmlPart.Attribute(pkg + "name"), UriKind.Relative);
                    string contentType = (string)xmlPart.Attribute(pkg + "contentType");
                    PackagePart part = package.CreatePart(uri, contentType, CompressionOption.SuperFast);

                    if (contentType.EndsWith("xml"))
                    {
                        using (XmlWriter xmlWriter = XmlWriter.Create(part.GetStream(FileMode.Create)))
                        {
                            xmlPart.Element(pkg + "xmlData").Elements().First().WriteTo(xmlWriter);
                        }
                    }
                    else
                    {
                        using (BinaryWriter binaryWriter = new BinaryWriter(part.GetStream(FileMode.Create)))
                        {
                            string base64StringInChunks = (string)xmlPart.Element(pkg + "binaryData");
                            char[] base64CharArray = base64StringInChunks.Where(c => c != '\r' && c != '\n').ToArray();
                            byte[] byteArray = Convert.FromBase64CharArray(base64CharArray, 0, base64CharArray.Length);
                            binaryWriter.Write(byteArray);
                        }
                    }
                }

                foreach (XElement xmlPart in xdWordXml.Root.Elements().Where(p => (string)p.Attribute(pkg + "contentType") == relContentType))
                {
                    string name = (string)xmlPart.Attribute(pkg + "name");

                    if (name == "/_rels/.rels")
                    {
                        // Add package-level relationships
                        foreach (XElement xmlRel in xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode = (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External")
                                package.CreateRelationship(new Uri(target, UriKind.Absolute), TargetMode.External, type, id);
                            else
                                package.CreateRelationship(new Uri(target, UriKind.Relative), TargetMode.Internal, type, id);
                        }
                    }
                    else
                    {
                        // Add part-level relationships
                        string directory = name.Substring(0, name.IndexOf("/_rels"));
                        string relsFilename = name.Substring(name.LastIndexOf('/'));
                        string filename = relsFilename.Substring(0, relsFilename.IndexOf(".rels"));
                        PackagePart fromPart = package.GetPart(new Uri(directory + filename, UriKind.Relative));
                        foreach (XElement xmlRel in xmlPart.Descendants(rel + "Relationship"))
                        {
                            string id = (string)xmlRel.Attribute("Id");
                            string type = (string)xmlRel.Attribute("Type");
                            string target = (string)xmlRel.Attribute("Target");
                            string targetMode = (string)xmlRel.Attribute("TargetMode");
                            if (targetMode == "External") fromPart.CreateRelationship(new Uri(target, UriKind.Absolute), TargetMode.External, type, id);
                            else
                                fromPart.CreateRelationship(new Uri(target, UriKind.Relative), TargetMode.Internal, type, id);
                        }
                    }
                }
            }
        }
    }
}