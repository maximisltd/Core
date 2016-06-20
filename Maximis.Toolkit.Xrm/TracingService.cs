using Microsoft.Xrm.Sdk;
using System.Text;

namespace Maximis.Toolkit.Xrm
{
    public class TracingService : ITracingService
    {
        private ITracingService crmTrace;
        private StringBuilder traceBuffer;

        public TracingService(ITracingService crmTrace)
        {
            this.crmTrace = crmTrace;
            this.traceBuffer = new StringBuilder();
        }

        public override string ToString()
        {
            return traceBuffer.ToString();
        }

        public void Trace(string format, params object[] args)
        {
            traceBuffer.AppendFormat(format, args);
            traceBuffer.AppendLine();
            crmTrace.Trace(format, args);
        }
    }
}