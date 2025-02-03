using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.SharePoint.Client;
using Microsoft.AspNetCore.Http;

public class SharePointFunction
{
    private readonly SecretClient _secretClient;
    private readonly ILogger<SharePointFunction> _logger;

    public SharePointFunction(ILogger<SharePointFunction> logger)
    {
        _logger = logger;
        Console.WriteLine(Environment.GetEnvironmentVariable("AZURE_CLIENT_ID"));

        _secretClient = new SecretClient(new Uri("https://lennoxsharepointkv.vault.azure.net/"), new DefaultAzureCredential());
    }

    [Function("SharePointOperations")]
    public async Task<IActionResult> RunAsync([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        try
        {
            // Retrieve secrets from Azure Key Vault
            string tenantId = (await _secretClient.GetSecretAsync("SharePointTenantId")).Value.Value;
            string clientId = (await _secretClient.GetSecretAsync("SharePointClientId")).Value.Value;
            string certificateBase64 = (await _secretClient.GetSecretAsync("SharePointCertificateBase64")).Value.Value;
            string certificatePassword = (await _secretClient.GetSecretAsync("SharePointCertificatePassword")).Value.Value; // Fetch password
            string siteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";

            // Load certificate from Base64 with password
            byte[] certBytes = Convert.FromBase64String(certificateBase64);
            X509Certificate2 certificate = new X509Certificate2(certBytes, certificatePassword, X509KeyStorageFlags.EphemeralKeySet);


            // Get Access Token
            string accessToken = await GetAccessTokenAsync(tenantId, clientId, certificate);

            // Perform SharePoint operations
            using (var context = new ClientContext(siteUrl))
            {
                context.ExecutingWebRequest += (sender, e) =>
                {
                    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
                };

                Web web = context.Web;
                context.Load(web, w => w.Title);
                context.ExecuteQuery();

                return new OkObjectResult($"Connected to SharePoint Site: {web.Title}");
            }
        }
        catch (Exception ex)
        {
            _logger.LogError($"Error: {ex.Message}");
            return new BadRequestObjectResult($"Error: {ex.Message}");
        }
    }

    private async Task<string> GetAccessTokenAsync(string tenantId, string clientId, X509Certificate2 certificate)
    {
        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithCertificate(certificate)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();

        AuthenticationResult result = await app.AcquireTokenForClient(new[] { "https://lennoxfamily.sharepoint.com/.default" }).ExecuteAsync();
        return result.AccessToken;
    }
}
