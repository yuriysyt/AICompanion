using System;
using System.Text.RegularExpressions;

namespace AICompanion.Desktop.Helpers
{
    /*
        String extension methods for common text processing operations.
        
        These utilities support voice command parsing and natural language
        processing tasks throughout the application. The methods handle
        edge cases like null strings and empty inputs gracefully.
    */
    public static class StringExtensions
    {
        /*
            Removes extra whitespace from a string, collapsing multiple
            spaces into single spaces and trimming the result.
            
            Useful for cleaning up speech recognition output which may
            contain irregular spacing due to pause detection.
        */
        public static string NormalizeWhitespace(this string input)
        {
            if (string.IsNullOrEmpty(input))
            {
                return string.Empty;
            }

            return Regex.Replace(input.Trim(), @"\s+", " ");
        }

        /*
            Extracts the first word from a string, useful for identifying
            command verbs like "open", "close", "click", "type".
        */
        public static string GetFirstWord(this string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var normalized = input.NormalizeWhitespace();
            var spaceIndex = normalized.IndexOf(' ');
            
            return spaceIndex > 0 
                ? normalized.Substring(0, spaceIndex) 
                : normalized;
        }

        /*
            Removes the first word from a string, returning the remainder.
            Used to extract parameters after a command verb.
        */
        public static string RemoveFirstWord(this string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            var normalized = input.NormalizeWhitespace();
            var spaceIndex = normalized.IndexOf(' ');
            
            return spaceIndex > 0 
                ? normalized.Substring(spaceIndex + 1) 
                : string.Empty;
        }

        /*
            Checks if the string contains any of the specified keywords,
            using case-insensitive comparison.
        */
        public static bool ContainsAny(this string input, params string[] keywords)
        {
            if (string.IsNullOrEmpty(input) || keywords == null || keywords.Length == 0)
            {
                return false;
            }

            var lowerInput = input.ToLowerInvariant();
            
            foreach (var keyword in keywords)
            {
                if (!string.IsNullOrEmpty(keyword) && 
                    lowerInput.Contains(keyword.ToLowerInvariant()))
                {
                    return true;
                }
            }
            
            return false;
        }

        /*
            Truncates a string to a maximum length, adding ellipsis if truncated.
            Useful for displaying long text in UI elements with limited space.
        */
        public static string Truncate(this string input, int maxLength)
        {
            if (string.IsNullOrEmpty(input) || input.Length <= maxLength)
            {
                return input ?? string.Empty;
            }

            return input.Substring(0, maxLength - 3) + "...";
        }
    }
}
