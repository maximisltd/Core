using Maximis.Toolkit.Logging;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Deployment;
using Microsoft.Xrm.Sdk.Deployment.Proxy;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Threading;

namespace Maximis.Toolkit.Xrm
{
    public static class ServiceHelper
    {
        private enum ServiceType { Organization, Deployment };

        public static void ExecuteAsyncRequest(IOrganizationService orgService, OrganizationRequest request, TraceProgressReporter progress)
        {
            ColumnSet colSet = new ColumnSet("statuscode", "message");
            Guid asyncJobId = ((ExecuteAsyncResponse)orgService.Execute(new ExecuteAsyncRequest { Request = request })).AsyncJobId;
            int prevStatus = int.MinValue;
            int status = int.MinValue;
            Entity asyncEntity = null;
            while (status < 30)
            {
                asyncEntity = orgService.Retrieve("asyncoperation", asyncJobId, colSet);
                status = asyncEntity.GetAttributeValue<OptionSetValue>("statuscode").Value;
                if (status != prevStatus)
                {
                    switch (status)
                    {
                        case 0:
                            WriteAsyncStatus(progress, "Waiting for Resources", asyncEntity);
                            break;

                        case 10:
                            WriteAsyncStatus(progress, "Waiting", asyncEntity);
                            break;

                        case 20:
                            WriteAsyncStatus(progress, "In Progress", asyncEntity);
                            break;

                        case 21:
                            WriteAsyncStatus(progress, "Pausing", asyncEntity);
                            break;

                        case 22:
                            WriteAsyncStatus(progress, "Canceling", asyncEntity);
                            break;
                    }
                    prevStatus = status;
                }
                Thread.Sleep(5000);
                progress.IterationComplete();
            }

            switch (status)
            {
                case 31:
                    WriteAsyncStatus(progress, "Failed", asyncEntity);
                    break;

                case 32:
                    WriteAsyncStatus(progress, "Cancelled", asyncEntity);

                    break;
            }
        }

        public static DeploymentServiceClient GetDeploymentServiceClient(CrmContext crmContext)
        {
            string serviceUrl = BuildUrl(crmContext, ServiceType.Deployment);
            DeploymentServiceClient client = ProxyClientHelper.CreateClient(new Uri(serviceUrl));
            client.ClientCredentials.Windows.ClientCredential = new NetworkCredential(crmContext.Username, crmContext.Password, crmContext.Domain);
            client.Endpoint.Binding.SendTimeout = TimeSpan.FromMinutes(10);
            return client;
        }

        public static OrganizationServiceProxy GetOrganizationServiceProxy(CrmContext crmConfiguration)
        {
            ClientCredentials credentials = new ClientCredentials();

            switch (crmConfiguration.AuthenticationProviderType)
            {
                case AuthenticationProviderType.ActiveDirectory:
                    credentials.Windows.ClientCredential = string.IsNullOrEmpty(crmConfiguration.Username) ? CredentialCache.DefaultNetworkCredentials :
                    new NetworkCredential(crmConfiguration.Username, crmConfiguration.Password, crmConfiguration.Domain);
                    break;

                case AuthenticationProviderType.LiveId:
                    throw new NotImplementedException("LiveId is not supported.");

                default:
                    credentials.UserName.UserName = crmConfiguration.Username;
                    credentials.UserName.Password = crmConfiguration.Password;
                    break;
            }

            string serviceUrl = BuildUrl(crmConfiguration, ServiceType.Organization);

            return new OrganizationServiceProxy(new Uri(serviceUrl.ToString()), null, credentials, null)
            {
                Timeout = (crmConfiguration.TimeoutSeconds > 0) ? TimeSpan.FromSeconds(crmConfiguration.TimeoutSeconds) : TimeSpan.FromMinutes(10)
            };
        }

        public static void RenewTokenIfRequired(OrganizationServiceProxy orgServiceProxy)
        {
            if (orgServiceProxy == null || orgServiceProxy.SecurityTokenResponse == null) return;
            if (DateTime.UtcNow.AddMinutes(15) >= orgServiceProxy.SecurityTokenResponse.Response.Lifetime.Expires)
            {
                orgServiceProxy.Authenticate();
            }
        }

        private static string BuildUrl(CrmContext crmContext, ServiceType serviceType)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(crmContext.Secure ? "https://" : "http://");
            sb.Append(crmContext.HostName);
            if (crmContext.Port > 0) sb.AppendFormat(":{0}", crmContext.Port);
            if (serviceType == ServiceType.Organization)
            {
                if (crmContext.AuthenticationProviderType == AuthenticationProviderType.ActiveDirectory)
                    sb.AppendFormat("/{0}", crmContext.Organization);

                sb.Append("/XRMServices/2011/Organization.svc");
            }
            else
            {
                sb.Append("/XRMDeployment/2011/Deployment.svc");
            }
            return sb.ToString();
        }

        private static void WriteAsyncStatus(TraceProgressReporter progress, string status, Entity asyncEntity)
        {
            string message = asyncEntity.GetAttributeValue<string>("message");
            if (string.IsNullOrEmpty(message))
            {
                progress.WriteInfo(status);
            }
            else
            {
                progress.WriteInfo(string.Format("{0} ({1})", status, message));
            }
        }
    }
}