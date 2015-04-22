using System;
using System.IO;
using System.Text;

namespace Maximis.Toolkit.IO
{
    public class StringWriterUtf8 : StringWriter
    {
        public StringWriterUtf8()
        {
        }

        public StringWriterUtf8(IFormatProvider formatProvider)
            : base(formatProvider)
        {
        }

        public StringWriterUtf8(StringBuilder sb)
            : base(sb)
        {
        }

        public StringWriterUtf8(StringBuilder sb, IFormatProvider formatProvider)
            : base(sb, formatProvider)
        {
        }

        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }
    }
}