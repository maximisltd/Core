using Maximis.Toolkit.IO;
using System;
using System.Diagnostics;

namespace Maximis.Toolkit.Logging
{
    public class LogFileTraceListener : TraceListener
    {
        private string filePath;

        public LogFileTraceListener(string filePath)
        {
            this.filePath = filePath;
        }

        public override void Write(string message)
        {
            FileHelper.AppendToFile(this.filePath, message);
        }

        public override void WriteLine(string message)
        {
            FileHelper.AppendToFile(this.filePath, message + Environment.NewLine);
        }
    }
}