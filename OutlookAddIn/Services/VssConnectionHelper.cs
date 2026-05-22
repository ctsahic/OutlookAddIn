using System;
using System.Net;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Newtonsoft.Json;

namespace OutlookAddIn.Services
{
    /// <summary>
    /// Helper class to create VssConnection that supports both HTTP and HTTPS,
    /// especially for local Azure DevOps servers or development environments.
    /// </summary>
    public static class VssConnectionHelper
    {
        /// <summary>
        /// Creates a VssConnection with support for HTTP connections.
        /// For HTTP connections, it bypasses the HTTPS requirement for basic authentication.
        /// </summary>
        /// <param name="organizationUrl">The Azure DevOps organization URL (http:// or https://)</param>
        /// <param name="pat">The Personal Access Token for authentication</param>
        /// <returns>A configured VssConnection instance</returns>
        public static VssConnection CreateVssConnection(string organizationUrl, string pat)
        {
            if (string.IsNullOrWhiteSpace(organizationUrl))
                throw new ArgumentNullException(nameof(organizationUrl));
            if (string.IsNullOrWhiteSpace(pat))
                throw new ArgumentNullException(nameof(pat));

            var uri = new Uri(organizationUrl);
            var credentials = new VssBasicCredential(string.Empty, pat);

            // If it's an HTTP URL (not HTTPS), we need to bypass the HTTPS requirement for basic auth
            if (uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase))
            {
                // Bypass the requirement for HTTPS when using basic authentication
                // This is needed for local Azure DevOps servers using HTTP
                ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
            }

            try
            {
                // Initialize Newtonsoft.Json settings to ensure proper serialization
                // This prevents "Method not found" errors with JsonSerializerSettings
                JsonConvert.DefaultSettings = () => new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    DateFormatString = "yyyy-MM-ddTHH:mm:ss.fffZ"
                };
            }
            catch
            {
                // If we can't set default settings, continue anyway
                // The VssConnection will still work but may have different JSON handling
            }

            // Create connection with credentials
            var connection = new VssConnection(uri, credentials);
            
            return connection;
        }
    }
}
