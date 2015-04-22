using Microsoft.VisualBasic.FileIO;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Maximis.Toolkit.Csv
{
    public class CsvReader : TextFieldParser
    {
        private static readonly Regex REGEX_SQUAREBRACKETS = new Regex(@"\[.*?\]", RegexOptions.Compiled);

        public CsvReader(Stream stream, bool hasHeadings = true)
            : base(stream)
        {
            Configure(hasHeadings);
        }

        public CsvReader(string path, bool hasHeadings = true)
            : base(path)
        {
            Configure(hasHeadings);
        }

        public CsvReader(TextReader reader, bool hasHeadings = true)
            : base(reader)
        {
            Configure(hasHeadings);
        }

        public CsvReader(Stream stream, Encoding defaultEncoding, bool hasHeadings = true)
            : base(stream, defaultEncoding)
        {
            Configure(hasHeadings);
        }

        public CsvReader(string path, Encoding defaultEncoding, bool hasHeadings = true)
            : base(path, defaultEncoding)
        {
            Configure(hasHeadings);
        }

        public CsvReader(Stream stream, Encoding defaultEncoding, bool detectEncoding, bool hasHeadings = true)
            : base(stream, defaultEncoding, detectEncoding)
        {
            Configure(hasHeadings);
        }

        public CsvReader(string path, Encoding defaultEncoding, bool detectEncoding, bool hasHeadings = true)
            : base(path, defaultEncoding, detectEncoding)
        {
            Configure(hasHeadings);
        }

        public CsvReader(Stream stream, Encoding defaultEncoding, bool detectEncoding, bool leaveOpen, bool hasHeadings = true)
            : base(stream, defaultEncoding, detectEncoding, leaveOpen)
        {
            Configure(hasHeadings);
        }

        public Dictionary<string, int> HeadingIndex { get; set; }

        private void Configure(bool hasHeadings)
        {
            this.Delimiters = new string[] { "," };
            this.HasFieldsEnclosedInQuotes = true;
            this.TextFieldType = FieldType.Delimited;
            this.TrimWhiteSpace = true;
            if (hasHeadings && !this.EndOfData)
            {
                int i = 0;
                this.HeadingIndex = new Dictionary<string, int>();
                foreach (string heading in this.ReadFields())
                {
                    this.HeadingIndex.Add(REGEX_SQUAREBRACKETS.Replace(heading, string.Empty).Trim(), i++);
                }
            }
        }
    }
}