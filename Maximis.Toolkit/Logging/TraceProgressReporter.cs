using System;
using System.Diagnostics;

namespace Maximis.Toolkit.Logging
{
    public class TraceProgressReporter : IDisposable
    {
        private bool operationComplete;
        private int traceProgressCount = 0;
        private int traceProgressPoint = 0;

        public TraceProgressReporter(string message, int traceProgressPoint = 0)
        {
            Trace.Write(message + "...");
            this.traceProgressPoint = traceProgressPoint;
        }

        public int TotalProcessed { get; set; }

        public void Dispose()
        {
            OperationComplete();
        }

        public void IterationComplete(int processed = 1)
        {
            TotalProcessed += processed;

            if (traceProgressPoint > 0)
            {
                traceProgressCount += processed;
                if (traceProgressCount >= traceProgressPoint)
                {
                    traceProgressCount = 0;
                    Trace.Write(".");
                }
            }
        }

        public void IterationFailed(string message)
        {
            Trace.Write(" [" + message + "] ");
        }

        public void OperationComplete()
        {
            if (!operationComplete)
            {
                if (traceProgressPoint > 0)
                {
                    Trace.WriteLine(string.Format("Done [{0}].", TotalProcessed));
                }
                else
                {
                    Trace.WriteLine("Done.");
                }
                operationComplete = true;
            }
        }

        public void OperationFailed(string message)
        {
            Trace.WriteLine("Failed: " + message);
            operationComplete = true;
        }

        public void WriteInfo(string message)
        {
            Trace.Write(message + "...");
        }
    }
}