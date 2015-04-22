using Microsoft.Xrm.Sdk.Client;

namespace Maximis.Toolkit.Xrm
{
    public class CrmContext
    {
        public AuthenticationProviderType AuthenticationProviderType { get; set; }

        public string Domain { get; set; }

        public string HostName { get; set; }

        public string Organization { get; set; }

        public string Password { get; set; }

        public int Port { get; set; }

        public bool Secure { get; set; }

        public double TimeoutSeconds { get; set; }

        public string Username { get; set; }

        public string Version { get; set; }
    }
}