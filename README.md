# SharePointFunctions

## Overview

This project contains Azure Functions to interact with SharePoint Online. The primary function, `SharePointFunction`, retrieves documents from a specified SharePoint library using secrets stored in Azure Key Vault.

## Function1.cs

### Purpose

The `SharePointFunction` class is designed to fetch documents from a SharePoint Online library. It uses Azure Key Vault to securely manage secrets such as the SharePoint tenant ID, client ID, and certificate details.

### How It Works

1. **Initialization**: The function initializes a `SecretClient` to interact with Azure Key Vault and an `ILogger` for logging.
2. **HTTP Trigger**: The function is triggered via HTTP requests. It expects a query parameter `libraryName` specifying the SharePoint library to fetch documents from.
3. **Retrieve Secrets**: The function retrieves necessary secrets (tenant ID, client ID, certificate) from Azure Key Vault.
4. **Access Token**: It obtains an access token using the Microsoft Identity Client library.
5. **Fetch Documents**: The function connects to the specified SharePoint library and retrieves a list of documents.
![image](https://github.com/user-attachments/assets/66cb369d-f358-4074-a07a-cffbe19efa9a)

### Example Usage

To call this function, use the following URL format:

Replace `your-function-url` with the actual URL of your deployed Azure Function and `Documents` with the name of the SharePoint library you want to access.

### Error Handling

The function includes error handling to log and return appropriate error messages if any step fails.

### Security

- **Secrets Management**: Secrets are securely managed using Azure Key Vault.
- **Logging**: Sensitive information is masked in logs to prevent exposure.

## Prerequisites

- Azure Key Vault with the necessary secrets stored.
- SharePoint Online site and library.
- Azure Function App deployed and configured.

## Configuration

Ensure the following settings are configured in `local.settings.json` or your Azure Function App settings:

- `AzureWebJobsStorage`
- `FUNCTIONS_WORKER_RUNTIME`
- `KeyVaultUri`

## Dependencies

- `Microsoft.Azure.Functions.Worker`
- `Microsoft.Extensions.Logging`
- `Microsoft.Graph`
- `Microsoft.Identity.Client`
- `Azure.Identity`
- `Azure.Security.KeyVault.Secrets`
- `Microsoft.SharePoint.Client`

## Conclusion

This function provides a secure and efficient way to interact with SharePoint Online libraries, leveraging Azure Key Vault for secrets management and Azure Functions for serverless execution.
