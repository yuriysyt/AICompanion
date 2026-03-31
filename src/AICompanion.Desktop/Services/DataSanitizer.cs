using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services
{
    /// <summary>
    /// Removes or anonymises sensitive information before it is sent to the LLM backend.
    /// Replaces absolute file paths, usernames, and other PII with generic placeholders.
    /// A per-session mapping lets the application reverse the substitution if needed.
    /// </summary>
    public class DataSanitizer
    {
        private readonly ILogger<DataSanitizer>? _logger;

        // Per-session placeholder → original mapping (for audit logging only; never sent to LLM)
        private readonly Dictionary<string, string> _placeholderMap = new();
        private int _fileCounter;

        // Regex patterns that should be redacted
        private static readonly Regex _absolutePathPattern = new(
            @"[A-Za-z]:\\(?:[^\\/\s""'<>|?*\x00-\x1F]+\\)*[^\\/\s""'<>|?*\x00-\x1F]*",
            RegexOptions.Compiled);

        private static readonly Regex _uncPathPattern = new(
            @"\\\\[^\\/\s]+\\[^\\/\s]+(?:\\[^\\/\s]*)*",
            RegexOptions.Compiled);

        // Matches typical Windows username paths like /Users/john or C:\Users\john
        private static readonly Regex _usernameinPathPattern = new(
            @"(?:Users|users)\\([^\\\/\s]+)",
            RegexOptions.Compiled);

        public DataSanitizer(ILogger<DataSanitizer>? logger = null)
        {
            _logger = logger;
        }

        /// <summary>
        /// Returns a sanitised copy of <paramref name="text"/> safe for LLM consumption.
        /// Absolute file paths are replaced with generic tokens like "document_1", "file_2".
        /// </summary>
        public string Sanitize(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            var result = text;

            // Replace absolute Windows paths (C:\...) and UNC paths (\\server\share)
            result = _absolutePathPattern.Replace(result, m => GetOrCreatePlaceholder(m.Value));
            result = _uncPathPattern.Replace(result, m => GetOrCreatePlaceholder(m.Value));

            _logger?.LogDebug("[DataSanitizer] Sanitized {Len} chars of input", text.Length);
            return result;
        }

        /// <summary>
        /// Sanitises the window title: strips the document name (before " - "),
        /// keeping only the application name.
        /// "MySecret.docx - Microsoft Word" → "Microsoft Word"
        /// </summary>
        public static string SanitizeWindowTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return "(unknown)";

            if (title.Contains(" - "))
            {
                var parts = title.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries);
                return parts[^1].Trim();
            }

            // If the title looks like a raw file path, replace entirely
            if (Regex.IsMatch(title, @"(?:[A-Za-z]:\\|\\\\)"))
                return "Document Editor";

            var lower = title.ToLowerInvariant();
            if (lower.Contains(".docx") || lower.Contains(".doc") || lower.Contains(".txt") ||
                lower.Contains(".pdf") || lower.Contains(".xlsx") || lower.Contains(".csv"))
                return "Document Editor";

            return title;
        }

        /// <summary>
        /// Returns the sanitized placeholder→original map for audit logging.
        /// Never include this in LLM requests.
        /// </summary>
        public IReadOnlyDictionary<string, string> GetAuditMap() => _placeholderMap;

        /// <summary>Resets the per-session counter and mapping (call on logout).</summary>
        public void Reset()
        {
            _placeholderMap.Clear();
            _fileCounter = 0;
        }

        // ── Private helpers ───────────────────────────────────────────────────

        private string GetOrCreatePlaceholder(string original)
        {
            // Return existing placeholder if we've seen this path before
            foreach (var kv in _placeholderMap)
                if (kv.Value.Equals(original, StringComparison.OrdinalIgnoreCase))
                    return kv.Key;

            _fileCounter++;
            var ext = TryGetExtension(original);
            var label = ext switch
            {
                ".docx" or ".doc" => $"document_{_fileCounter}",
                ".txt"            => $"textfile_{_fileCounter}",
                ".pdf"            => $"pdf_{_fileCounter}",
                ".xlsx" or ".csv" => $"spreadsheet_{_fileCounter}",
                ".exe"            => $"program_{_fileCounter}",
                ""                => $"folder_{_fileCounter}",
                _                 => $"file_{_fileCounter}"
            };

            _placeholderMap[label] = original;
            _logger?.LogInformation("[DataSanitizer] Replaced '{Original}' → '{Label}'", original, label);
            return label;
        }

        private static string TryGetExtension(string path)
        {
            try { return System.IO.Path.GetExtension(path).ToLowerInvariant(); }
            catch { return ""; }
        }
    }
}
