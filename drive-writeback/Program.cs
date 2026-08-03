using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using DriveWriteback.Auth;
using DriveWriteback.Graph;
using DriveWriteback.Graph.Auth;
using DriveWriteback.Graph.Writes;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

builder.Services
    .AddApplicationInsightsTelemetryWorkerService()
    .ConfigureFunctionsApplicationInsights();

var tenantId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_TENANT_ID")
    ?? throw new InvalidOperationException("DRIVE_WRITEBACK_TENANT_ID is not set.");

var clientId = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_CLIENT_ID")
    ?? throw new InvalidOperationException("DRIVE_WRITEBACK_CLIENT_ID is not set.");

var keyVaultUri = Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_KEY_VAULT_URI")
    ?? throw new InvalidOperationException("DRIVE_WRITEBACK_KEY_VAULT_URI is not set.");

var dryRun = (Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_DRY_RUN") ?? "true")
    .Equals("true", StringComparison.OrdinalIgnoreCase);

var maxContentBytes = long.TryParse(Environment.GetEnvironmentVariable("DRIVE_WRITEBACK_MAX_CONTENT_BYTES"), out var configuredMaxContentBytes)
    ? configuredMaxContentBytes
    : 1_048_576;

builder.Services.AddSingleton<IRefreshTokenStore>(_ =>
    new KeyVaultRefreshTokenStore(new SecretClient(new Uri(keyVaultUri), new DefaultAzureCredential())));

builder.Services.AddSingleton(sp => DriveGraphClient.CreateWithSilentRefreshAuth(
    tenantId,
    clientId,
    ["Files.ReadWrite.All", "Sites.Read.All"],
    sp.GetRequiredService<IRefreshTokenStore>()));

builder.Services.AddSingleton(new DriveWriteOptions(dryRun, maxContentBytes));
builder.Services.AddSingleton<DriveWriteService>();

builder.Build().Run();
