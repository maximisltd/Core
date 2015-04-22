using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Csv
{
    public class CsvWriter
    {
        private static Regex REGEX_NEEDSQUOTES = new Regex("[,\\s\\n]", RegexOptions.Compiled);

        private bool alwaysEncloseInQuotes = true;

        public CsvWriter(TextWriter writer, string[] columnHeadings)
        {
            this.Writer = writer;

            this.HeadingIndex = new Dictionary<string, int>();
            for (int i = 0; i < columnHeadings.Length; i++)
            {
                this.HeadingIndex.Add(columnHeadings[i], i);
                this.WriteValue(i, columnHeadings[i], true);
            }

            this.ColumnCount = columnHeadings.Length;
        }

        public bool AlwaysEncloseInQuotes { get { return this.alwaysEncloseInQuotes; } set { this.alwaysEncloseInQuotes = value; } }

        public int ColumnCount { get; set; }

        public Dictionary<string, int> HeadingIndex { get; set; }

        public TextWriter Writer { get; set; }

        public void WriteValue(string columnHeading, string val)
        {
            WriteValue(this.HeadingIndex[columnHeading], val);
        }

        public void WriteValue(int index, string val)
        {
            WriteValue(index, val, false);
        }

        private void WriteValue(int index, string val, bool isHeaderRow)
        {
            if (!isHeaderRow && index == 0) this.Writer.WriteLine();
            if (index > 0) this.Writer.Write(",");
            if (string.IsNullOrWhiteSpace(val))
            {
                this.Writer.Write(string.Empty);
            }
            else
            {
                val = val.Trim().Replace("\"", "\\\"");
                this.Writer.Write(this.AlwaysEncloseInQuotes || REGEX_NEEDSQUOTES.IsMatch(val) ? string.Format("\"{0}\"", val) : val);
            }
        }
    }
}