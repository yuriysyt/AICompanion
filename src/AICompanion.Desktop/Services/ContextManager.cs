using System;
using System.Collections.Generic;
using System.Text;

namespace AICompanion.Desktop.Services
{
    /// <summary>
    /// Keeps a sliding window of the last N command/result pairs and
    /// serialises them into a prompt snippet for the LLM.
    /// Thread-safe via lock.
    /// </summary>
    public class ContextManager
    {
        private const int MaxItems = 10;

        private readonly Queue<ContextEntry> _history = new();
        private readonly object _lock = new();

        /// <summary>
        /// Adds a command/result pair to the context history.
        /// Oldest entry is removed once the limit is reached.
        /// </summary>
        public void AddContext(string command, string result)
        {
            if (string.IsNullOrWhiteSpace(command)) return;

            lock (_lock)
            {
                if (_history.Count >= MaxItems)
                    _history.Dequeue();

                _history.Enqueue(new ContextEntry
                {
                    Command   = command.Trim(),
                    Result    = result?.Trim() ?? string.Empty,
                    Timestamp = DateTime.UtcNow
                });
            }
        }

        /// <summary>
        /// Returns a formatted string with previous commands and results
        /// suitable for prepending to an LLM prompt.
        /// Returns an empty string when the history is empty.
        /// </summary>
        public string GetContextPrompt()
        {
            ContextEntry[] snapshot;
            lock (_lock)
            {
                if (_history.Count == 0) return string.Empty;
                snapshot = _history.ToArray();
            }

            var sb = new StringBuilder();
            sb.AppendLine("Recent conversation context:");
            for (int i = 0; i < snapshot.Length; i++)
            {
                var entry = snapshot[i];
                sb.AppendLine($"  [{i + 1}] User: {entry.Command}");
                if (!string.IsNullOrEmpty(entry.Result))
                    sb.AppendLine($"       Result: {entry.Result}");
            }
            sb.AppendLine("---");
            return sb.ToString();
        }

        /// <summary>Clears all stored context.</summary>
        public void Clear()
        {
            lock (_lock) { _history.Clear(); }
        }

        public int Count
        {
            get { lock (_lock) { return _history.Count; } }
        }

        private sealed class ContextEntry
        {
            public string Command   { get; init; } = "";
            public string Result    { get; init; } = "";
            public DateTime Timestamp { get; init; }
        }
    }
}
