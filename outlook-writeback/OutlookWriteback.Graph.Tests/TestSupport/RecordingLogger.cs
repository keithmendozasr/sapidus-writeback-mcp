using Microsoft.Extensions.Logging;

namespace OutlookWriteback.Graph.Tests.TestSupport;

internal sealed class RecordingLogger<T> : ILogger<T>
{
    public List<string> Messages { get; } = [];

    public List<IReadOnlyList<KeyValuePair<string, object?>>> Properties { get; } = [];

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter)
    {
        Messages.Add(formatter(state, exception));

        if (state is IReadOnlyList<KeyValuePair<string, object?>> properties)
            Properties.Add(properties);
    }
}
