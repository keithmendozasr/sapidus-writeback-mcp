using Microsoft.Extensions.Logging;

namespace DriveWriteback.Graph.Tests.TestSupport;

/// <summary>
/// Captures formatted log messages instead of writing anywhere, so tests can assert that
/// DriveWriteService's audit-log entry (PRD §9 item 7) fired without wiring up a real
/// logging provider.
/// </summary>
internal sealed class RecordingLogger<T> : ILogger<T>
{
    public List<string> Messages { get; } = [];

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

    public bool IsEnabled(LogLevel logLevel) => true;

    public void Log<TState>(
        LogLevel logLevel,
        EventId eventId,
        TState state,
        Exception? exception,
        Func<TState, Exception?, string> formatter) =>
        Messages.Add(formatter(state, exception));
}
