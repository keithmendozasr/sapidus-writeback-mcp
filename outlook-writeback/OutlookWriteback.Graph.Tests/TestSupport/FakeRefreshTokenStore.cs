using OutlookWriteback.Graph.Auth;

namespace OutlookWriteback.Graph.Tests.TestSupport;

internal sealed class FakeRefreshTokenStore : IRefreshTokenStore
{
    private string _refreshToken;

    public FakeRefreshTokenStore(string initialRefreshToken)
    {
        _refreshToken = initialRefreshToken;
    }

    public int SaveCallCount { get; private set; }

    public Task<string> GetRefreshTokenAsync(CancellationToken cancellationToken = default) =>
        Task.FromResult(_refreshToken);

    public Task SaveRefreshTokenAsync(string refreshToken, CancellationToken cancellationToken = default)
    {
        _refreshToken = refreshToken;
        SaveCallCount++;

        return Task.CompletedTask;
    }
}
