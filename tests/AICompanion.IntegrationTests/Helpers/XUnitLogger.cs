using System;
using Microsoft.Extensions.Logging;
using Xunit.Abstractions;

namespace AICompanion.IntegrationTests.Helpers
{
    /// <summary>
    /// Bridges Microsoft.Extensions.Logging to xUnit's ITestOutputHelper so that
    /// all ILogger calls from production code appear in the test runner output.
    /// </summary>
    public class XUnitLogger<T> : ILogger<T>
    {
        private readonly ITestOutputHelper _output;
        private readonly string _category;

        public XUnitLogger(ITestOutputHelper output)
        {
            _output = output;
            _category = typeof(T).Name;
        }

        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;

        public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Debug;

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
            Exception? exception, Func<TState, Exception?, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;
            var level = logLevel switch
            {
                LogLevel.Trace       => "TRC",
                LogLevel.Debug       => "DBG",
                LogLevel.Information => "INF",
                LogLevel.Warning     => "WRN",
                LogLevel.Error       => "ERR",
                LogLevel.Critical    => "CRT",
                _                    => "???"
            };
            var message = formatter(state, exception);
            try
            {
                _output.WriteLine($"  [LOG:{level}] [{_category}] {message}");
                if (exception != null)
                    _output.WriteLine($"  [EX] {exception.GetType().Name}: {exception.Message}");
            }
            catch (InvalidOperationException)
            {
                // xUnit throws if output is used after the test has ended — safe to ignore
            }
        }
    }
}
