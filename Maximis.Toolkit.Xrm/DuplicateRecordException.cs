using System;
using System.Runtime.Serialization;

namespace Maximis.Toolkit.Xrm
{
    public class DuplicateRecordException : Exception
    {
        public DuplicateRecordException()
        {
        }

        public DuplicateRecordException(string message)
            : base(message)
        {
        }

        public DuplicateRecordException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public DuplicateRecordException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}