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
using System.Collections.Generic;

public class SharePointFunction
{
    private readonly SecretClient _secretClient;
    private readonly ILogger<SharePointFunction> _logger;

    //https://your-function-url/api/SharePointOperations?libraryName=Documents

    //https://lennoxsharepointfunctions.azurewebsites.net/api/SharePointOperations?code=kHFYIZtN3KdB9lhn2K41ReliPCgmlS9aKm85rt2K0FwOAzFuzoiQUg%3D%3D&libraryName=Documents


    public SharePointFunction(ILogger<SharePointFunction> logger)
    {
        _logger = logger;
        _secretClient = new SecretClient(new Uri("https://lennoxsharepointkv.vault.azure.net/"), new DefaultAzureCredential());
    }

    [Function("SharePointOperations")]
    public async Task<IActionResult> RunAsync([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        try
        {
            string libraryName = req.Query["libraryName"];
            if (string.IsNullOrEmpty(libraryName))
            {
                return new BadRequestObjectResult("Please provide a SharePoint library name.");
            }

            // Retrieve secrets from Azure Key Vault
            string tenantId = (await _secretClient.GetSecretAsync("SharePointTenantId")).Value.Value;
            string clientId = (await _secretClient.GetSecretAsync("SharePointClientId")).Value.Value;
            string certificateBase64 = (await _secretClient.GetSecretAsync("SharePointCertificateBase64")).Value.Value;
            string certificatePassword = (await _secretClient.GetSecretAsync("SharePointCertificatePassword")).Value.Value;

            _logger.LogInformation($"Received tenantId: {tenantId}");
            _logger.LogInformation($"Received clientId: {clientId}");
            _logger.LogInformation($"Received certificateBase64: {certificateBase64.Substring(0, 10)}..."); // Log only the first 10 characters for security
            _logger.LogInformation($"Received certificatePassword: {new string('*', certificatePassword.Length)}"); // Log masked password

            string siteUrl = "https://lennoxfamily.sharepoint.com/sites/AuthDemoSite";

            // Load certificate from Base64 with password
            byte[] certBytes = Convert.FromBase64String(certificateBase64);
            X509Certificate2 certificate = new X509Certificate2(certBytes, certificatePassword, X509KeyStorageFlags.EphemeralKeySet);

            // Get Access Token
            string accessToken = await GetAccessTokenAsync(tenantId, clientId, certificate);

            // Fetch documents from SharePoint library
            List<string> documentList = GetDocumentListFromLibrary(siteUrl, accessToken, libraryName);
            return new OkObjectResult(documentList);
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

    private List<string> GetDocumentListFromLibrary(string siteUrl, string accessToken, string libraryName)
    {
        List<string> documents = new List<string>();

        using (var context = new ClientContext(siteUrl))
        {
            context.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
            };

            Microsoft.SharePoint.Client.List documentLibrary = context.Web.Lists.GetByTitle(libraryName);
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            ListItemCollection items = documentLibrary.GetItems(query);
            context.Load(items);
            context.ExecuteQuery();

            foreach (var item in items)
            {
                documents.Add(item["FileRef"].ToString());
            }
        }
        return documents;
    }
}
