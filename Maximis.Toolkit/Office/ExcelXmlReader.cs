using ICSharpCode.SharpZipLib.Zip;
using Maximis.Toolkit.IO;
using Maximis.Toolkit.Xml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Maximis.Toolkit.Office
{
    public class ExcelXmlReader
    {
        private static readonly Regex reLetters = new Regex("[A-Z]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public ExcelXmlReader(string xlsxFilePath)
        {
            XlsxFilePath = xlsxFilePath;

            Dictionary<string, string> worksheetNames = new Dictionary<string, string>();
            Dictionary<string, string> worksheetContent = new Dictionary<string, string>();

            using (FileStream fileStream = File.OpenRead(xlsxFilePath))
            {
                ZipFile zipFile = new ZipFile(fileStream);
                IEnumerator enumerator = zipFile.GetEnumerator();
                while (enumerator.MoveNext())
                {
                    ZipEntry zipEntry = (ZipEntry)enumerator.Current;
                    if (zipEntry.Name == "docProps/app.xml")
                    {
                        PopulateWorksheetNames(worksheetNames,
                            StreamHelper.ReadStringFromStream(zipFile.GetInputStream(zipEntry)));
                    }
                    else if (zipEntry.Name == "xl/sharedStrings.xml")
                    {
                        SharedStrings =
                            NamespacedXmlDocument.FromXmlString(
                                StreamHelper.ReadStringFromStream(zipFile.GetInputStream(zipEntry)));
                    }
                    else if (zipEntry.Name.Contains("xl/worksheets/sheet"))
                    {
                        worksheetContent.Add(zipEntry.Name,
                            StreamHelper.ReadStringFromStream(zipFile.GetInputStream(zipEntry)));
                    }
                }
            }

            Worksheets = new Dictionary<string, NamespacedXmlDocument>();
            foreach (string key in worksheetNames.Keys)
            {
                Worksheets.Add(worksheetNames[key], NamespacedXmlDocument.FromXmlString(worksheetContent[key]));
            }
        }

        public NamespacedXmlDocument SharedStrings { get; set; }

        public Dictionary<string, NamespacedXmlDocument> Worksheets { get; set; }

        public string XlsxFilePath { get; set; }

        public string GetFriendlyXml()
        {
            StringBuilder sb = new StringBuilder();
            using (StringWriterUtf8 sw = new StringWriterUtf8(sb))
            using (XmlTextWriter xtw = new XmlTextWriter(sw))
            {
                GetFriendlyXml(xtw);
            }
            return sb.ToString();
        }

        public void GetFriendlyXml(XmlTextWriter xtw)
        {
            Dictionary<int, string> sharedStringLookup = new Dictionary<int, string>();
            int index = 0;
            foreach (XmlNode ss in SharedStrings.SelectNodes("//dflt:si"))
            {
                sharedStringLookup.Add(index++, ss.InnerText);
            }

            xtw.WriteStartElement("wb");
            xtw.WriteAttributeString("path", XlsxFilePath);

            foreach (string worksheetName in Worksheets.Keys)
            {
                NamespacedXmlDocument xmlDoc = Worksheets[worksheetName];
                xtw.WriteStartElement("ws");
                xtw.WriteAttributeString("name", worksheetName);

                foreach (XmlNode row in xmlDoc.SelectNodes("//dflt:sheetData/dflt:row"))
                {
                    int currentCellIndex = 1;
                    xtw.WriteStartElement("r");
                    foreach (XmlNode cell in row.SelectNodes("dflt:c", xmlDoc.XmlNamespaceManager))
                    {
                        int blankCells = CalculateBlankCells(currentCellIndex++, cell.Attributes["r"].InnerText);
                        if (blankCells > 0)
                        {
                            for (int i = 0; i < blankCells; i++)
                            {
                                xtw.WriteElementString("c", string.Empty);
                                currentCellIndex++;
                            }
                        }

                        xtw.WriteStartElement("c");
                        if (!string.IsNullOrEmpty(cell.InnerText))
                        {
                            if (XmlHelper.GetAttributeValue(cell, "t") == "s")
                            {
                                // Shared String
                                xtw.WriteString(sharedStringLookup[int.Parse(cell.InnerText)].Replace((char)160, (char)32));
                            }
                            else if (XmlHelper.GetAttributeValue(cell, "s") == "1")
                            {
                                // Date
                                xtw.WriteString(DateTime.FromOADate(int.Parse(cell.InnerText)).ToString());
                            }
                            else
                            {
                                xtw.WriteString(cell.InnerText);
                            }
                        }
                        xtw.WriteEndElement(); // Cell
                    }
                    xtw.WriteEndElement(); // Row
                }

                xtw.WriteEndElement(); // Worksheet
            }
            xtw.WriteEndElement(); // Workbook
        }

        private static int CalculateBlankCells(int currentCellIndex, string cellRef)
        {
            return ExcelColumnNameToNumber(reLetters.Match(cellRef).Value) - currentCellIndex;
        }

        private static int ExcelColumnNameToNumber(string columnName)
        {
            columnName = columnName.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }
            return sum;
        }

        private static void PopulateWorksheetNames(Dictionary<string, string> worksheetNames, string appXml)
        {
            int i = 0;
            NamespacedXmlDocument appXmlDoc = NamespacedXmlDocument.FromXmlString(appXml);
            foreach (XmlNode xmlNode in appXmlDoc.SelectNodes("//dflt:TitlesOfParts/vt:vector/vt:lpstr"))
            {
                worksheetNames.Add(string.Format("xl/worksheets/sheet{0}.xml", ++i), xmlNode.InnerText);
            }
        }
    }
}