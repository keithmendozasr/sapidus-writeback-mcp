using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using OutlookWriteback.Auth;
using OutlookWriteback.Graph;
using OutlookWriteback.Graph.Auth;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

var tenantId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_TENANT_ID")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_TENANT_ID is not set.");

var clientId = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_CLIENT_ID")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_CLIENT_ID is not set.");

var keyVaultUri = Environment.GetEnvironmentVariable("OUTLOOK_WRITEBACK_KEY_VAULT_URI")
    ?? throw new InvalidOperationException("OUTLOOK_WRITEBACK_KEY_VAULT_URI is not set.");

builder.Services.AddSingleton<IRefreshTokenStore>(_ =>
    new KeyVaultRefreshTokenStore(new SecretClient(new Uri(keyVaultUri), new DefaultAzureCredential())));

builder.Services.AddSingleton(sp => OutlookGraphClient.CreateWithSilentRefreshAuth(
    tenantId,
    clientId,
    ["Mail.ReadWrite", "Calendars.ReadWrite"],
    sp.GetRequiredService<IRefreshTokenStore>()));

builder.Build().Run();
