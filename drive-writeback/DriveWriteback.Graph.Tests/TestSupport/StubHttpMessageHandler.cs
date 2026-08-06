using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DriveWriteback.Graph.Tests.TestSupport;

/// <summary>
/// Fakes the HTTP boundary Microsoft.Graph's Kiota-generated client sends requests through, so
/// tests exercise real request-building and response deserialization without a network call.
///
/// Answers GET /drives/{id} (a bare drive-metadata fetch, no /items segment) itself, with
/// ownDriveType ("business" by default), rather than forwarding it to respond - this is
/// DriveGraphClient.ResolveDriveIdAsync's SharePoint-rejection check, which every test that
/// passes an explicit driveId now triggers once. Defaulting it to "business" means every
/// pre-existing test that stubs its own driveId-scoped response (even a blanket one, like a
/// 404-for-everything respond) keeps working unmodified; only a test that specifically wants
/// to exercise rejection needs to override ownDriveType.
/// </summary>
internal sealed class StubHttpMessageHandler(
    Func<HttpRequestMessage, Task<HttpResponseMessage>> respond,
    string? ownDriveType = "business") : HttpMessageHandler
{
    private static readonly Regex DriveMetadataPath = new(@"^/v1\.0/drives/[^/]+$");

    public HttpRequestMessage? LastRequest { get; private set; }

    /// <summary>
    /// How many times the GET /drives/{id} intercept above fired - lets a test assert
    /// DriveGraphClient's per-drive-id validation cache (ResolveDriveIdAsync's
    /// validatedOneDriveIds) is actually deduplicating repeat calls against the same drive id
    /// within one operation, not re-querying Graph every time.
    /// </summary>
    public int DriveMetadataRequestCount { get; private set; }

    protected override Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request,
        CancellationToken cancellationToken)
    {
        LastRequest = request;

        if (request.Method == HttpMethod.Get && DriveMetadataPath.IsMatch(request.RequestUri!.AbsolutePath))
        {
            DriveMetadataRequestCount++;

            return Task.FromResult(DriveMetadataResponse());
        }

        return respond(request);
    }

    private HttpResponseMessage DriveMetadataResponse()
    {
        var body = JsonSerializer.Serialize(new { id = "drive-id", driveType = ownDriveType });

        return new HttpResponseMessage(System.Net.HttpStatusCode.OK)
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json"),
        };
    }
}
