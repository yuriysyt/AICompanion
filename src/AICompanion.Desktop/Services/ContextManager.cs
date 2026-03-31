using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services
{
    /// <summary>
    /// Maintains a sliding window of the last N commands and their results so the
    /// AI backend can plan future actions with full short-term context.
    ///
    /// Thread-safe: all mutations are protected by a lock.
    /// </summary>
    public class ContextManager
    {
        private readonly ILogger<ContextManager>? _logger;
        private readonly int _maxEntries;
        private readonly object _lock = new();

        private readonly List<ContextEntry> _history = new();

        // Tracks the currently open application/window so every prompt includes it
        private string _activeApp = "";
        private string _activeWindowTitle = "";

        public ContextManager(ILogger<ContextManager>? logger = null, int maxEntries = 10)
        {
            _logger = logger;
            _maxEntries = maxEntries;
        }

        // ── Public API ────────────────────────────────────────────────────────

        /// <summary>
        /// Records a command and its result after execution.
        /// Keeps the ring-buffer trimmed to <see cref="_maxEntries"/> entries.
        /// </summary>
        public void AddCommandResult(string command, string result)
        {
            lock (_lock)
            {
                _history.Add(new ContextEntry
                {
                    Command   = command,
                    Result    = result,
                    Timestamp = DateTime.UtcNow
                });

                while (_history.Count > _maxEntries)
                    _history.RemoveAt(0);
            }

            _logger?.LogDebug("[ContextMgr] Stored: '{Cmd}' → '{Res}'", command, result.Length > 60 ? result[..60] + "…" : result);
        }

        /// <summary>
        /// Updates the currently tracked application and window title.
        /// Called after every open_app / focus_window step.
        /// </summary>
        public void UpdateActiveApp(string appName, string windowTitle = "")
        {
            lock (_lock)
            {
                _activeApp         = appName;
                _activeWindowTitle = windowTitle;
            }
        }

        /// <summary>
        /// Builds a compact context string suitable for injection into the LLM prompt
        /// via the <c>session_context</c> field.  Format is intentionally concise.
        /// </summary>
        public string BuildContextString()
        {
            lock (_lock)
            {
                if (_history.Count == 0 && string.IsNullOrEmpty(_activeApp))
                    return "";

                var sb = new StringBuilder();

                if (!string.IsNullOrEmpty(_activeApp))
                    sb.AppendLine($"Active app: {_activeApp}" +
                                  (string.IsNullOrEmpty(_activeWindowTitle) ? "" : $" ({_activeWindowTitle})"));

                if (_history.Count > 0)
                {
                    sb.AppendLine("Recent commands:");
                    foreach (var entry in _history)
                    {
                        var shortResult = entry.Result.Length > 80
                            ? entry.Result[..80] + "…"
                            : entry.Result;
                        sb.AppendLine($"  [{entry.Timestamp:HH:mm:ss}] CMD: {entry.Command} → {shortResult}");
                    }
                }

                return sb.ToString().TrimEnd();
            }
        }

        /// <summary>
        /// Returns the most recently recorded active application name (e.g. "winword", "notepad").
        /// </summary>
        public string GetActiveApp()
        {
            lock (_lock) { return _activeApp; }
        }

        /// <summary>
        /// Clears all history (call on user logout or new session).
        /// </summary>
        public void Reset()
        {
            lock (_lock)
            {
                _history.Clear();
                _activeApp         = "";
                _activeWindowTitle = "";
            }
            _logger?.LogInformation("[ContextMgr] Context reset");
        }

        // ── Inner type ────────────────────────────────────────────────────────

        private sealed class ContextEntry
        {
            public string    Command   { get; init; } = "";
            public string    Result    { get; init; } = "";
            public DateTime  Timestamp { get; init; }
        }
    }
}
