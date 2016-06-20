using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Csv
{
    public class CsvWriter
    {
        /// <summary>
        /// A Regular Expression to determine if a value needs to be surrounded in quotation marks
        /// </summary>
        private static Regex REGEX_NEEDSQUOTES = new Regex("[,\\s\\n]", RegexOptions.Compiled);

        private bool allValuesInQuotes;
        private int columnCount;
        private Dictionary<string, int> headingIndex;
        private int lastIndex = -1;

        private TextWriter writer;

        public CsvWriter(TextWriter writer, bool allValuesInQuotes, params string[] columnHeadings)
        {
            this.writer = writer;
            this.allValuesInQuotes = allValuesInQuotes;

            // Build Heading Index
            this.headingIndex = new Dictionary<string, int>();
            for (int i = 0; i < columnHeadings.Length; i++)
            {
                this.HeadingIndex.Add(columnHeadings[i], i);
                this.WriteValue(i, columnHeadings[i], true);
            }

            this.columnCount = columnHeadings.Length;
        }

        /// <summary>
        /// If true, all values are enclosed in Quotation Marks, otherwise, just values containing
        /// commas, spaces and newline characters
        /// </summary>
        public bool AllValuesInQuotes { get { return allValuesInQuotes; } }

        /// <summary>
        /// The number of columns in the CSV output
        /// </summary>
        public int ColumnCount { get { return columnCount; } }

        /// <summary>
        /// Lookup of Column Headers and their positions
        /// </summary>
        public Dictionary<string, int> HeadingIndex { get { return headingIndex; } }

        /// <summary>
        /// The underlying TextWriter
        /// </summary>
        public TextWriter Writer { get { return writer; } }

        public void NewRow()
        {
            WriteValue(0, string.Empty);
        }

        /// <summary>
        /// Write a value for a given column. If it is the first column, a new line will be started
        /// </summary>
        public void WriteValue(string columnHeading, string val)
        {
            WriteValue(this.HeadingIndex[columnHeading], val);
        }

        /// <summary>
        /// Write a value for a given column. If index is zero, a new line will be started
        /// </summary>
        public void WriteValue(int index, string val)
        {
            WriteValue(index, val, false);
        }

        /// <summary>
        /// Internal WriteValue method
        /// </summary>
        private void WriteValue(int index, string val, bool isHeaderRow)
        {
            // Handle new line
            if (!isHeaderRow && index == 0)
            {
                this.Writer.WriteLine();
                lastIndex = -1;
            }

            // Basic validation - values must be added in order
            if (!isHeaderRow)
            {
                if (index <= lastIndex) throw new InvalidOperationException("Column Index must be greater than " + lastIndex);
                if (index > ColumnCount - 1) throw new InvalidOperationException("Column Index must be no more than " + ColumnCount);
            }

            // Add the required number of commas
            if (index > 0)
            {
                for (int i = lastIndex; i < index; i++)
                {
                    this.Writer.Write(",");
                }
            }

            // Add the value
            if (string.IsNullOrWhiteSpace(val))
            {
                this.Writer.Write(string.Empty);
            }
            else
            {
                val = val.Trim().Replace("\"", "\\\"");
                this.Writer.Write(this.AllValuesInQuotes || REGEX_NEEDSQUOTES.IsMatch(val) ? string.Format("\"{0}\"", val) : val);
            }

            lastIndex = index;
        }
    }
}