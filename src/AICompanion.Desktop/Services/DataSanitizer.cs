using System;
using System.IO;
using System.Text.RegularExpressions;

namespace AICompanion.Desktop.Services
{
    /// <summary>
    /// Sanitizes user input before it is sent to the LLM.
    /// Removes full file system paths so the model does not see
    /// machine-specific paths (privacy + portability).
    /// </summary>
    public class DataSanitizer
    {
        private int _fileCounter;
        private int _docCounter;
        private int _imgCounter;

        // Matches absolute Windows paths: C:\...\file.ext or C:/...
        private static readonly Regex _winAbsPath = new Regex(
            @"[A-Za-z]:[\\\/](?:[^\\\/\r\n""<>|?*:]+[\\\/])*[^\\\/\r\n""<>|?*:]*",
            RegexOptions.Compiled);

        // Matches UNC paths: \\server\share\...
        private static readonly Regex _uncPath = new Regex(
            @"\\\\[^\\\/\r\n""<>|?*:]+(?:[\\\/][^\\\/\r\n""<>|?*:]*)*",
            RegexOptions.Compiled);

        /// <summary>
        /// Replaces all absolute file paths in <paramref name="input"/> with
        /// short generic placeholders (e.g. document_1, file_2) so the LLM
        /// never sees local directory structure.
        /// </summary>
        public string SanitizeForLlm(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            var result = _winAbsPath.Replace(input, m => GenericName(m.Value));
            result = _uncPath.Replace(result, m => GenericName(m.Value));
            return result;
        }

        /// <summary>
        /// Returns only the file name from a path (no directory components).
        /// Safe to pass to the LLM as context.
        /// </summary>
        public string SanitizeFilePath(string path)
        {
            if (string.IsNullOrEmpty(path)) return path;
            try { return Path.GetFileName(path); }
            catch { return path; }
        }

        private string GenericName(string path)
        {
            var ext = string.Empty;
            try { ext = Path.GetExtension(path).ToLowerInvariant(); } catch { }

            return ext switch
            {
                ".docx" or ".doc"             => $"document_{++_docCounter}",
                ".xlsx" or ".xls" or ".csv"   => $"spreadsheet_{++_docCounter}",
                ".pdf"                         => $"document_{++_docCounter}",
                ".txt" or ".md"               => $"text_{++_docCounter}",
                ".jpg" or ".jpeg" or ".png"
                    or ".gif" or ".bmp"       => $"image_{++_imgCounter}",
                _                             => $"file_{++_fileCounter}"
            };
        }

        /// <summary>Resets all internal counters (useful for unit tests).</summary>
        public void Reset()
        {
            _fileCounter = 0;
            _docCounter  = 0;
            _imgCounter  = 0;
        }
    }
}
