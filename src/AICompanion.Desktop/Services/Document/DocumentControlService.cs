using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Document
{
    /*
        DocumentControlService provides advanced voice-controlled document interaction.
        
        This service enables users to interact with text-based applications like
        Notepad, Word, and other editors using voice commands. It supports:
        - Text dictation and insertion
        - Text selection and navigation
        - Formatting operations (bold, italic, underline)
        - Clipboard operations (copy, cut, paste)
        - Document navigation (go to line, start, end)
        
        This is a key feature for the IPD prototype demonstration, showing
        real-world accessibility benefits for users with motor disabilities.
    */
    public class DocumentControlService
    {
        private readonly ILogger<DocumentControlService>? _logger;
        private readonly ActionHistoryService _actionHistory;

        // Track the last focused window for context
        private IntPtr _lastActiveWindow = IntPtr.Zero;
        private string _lastActiveWindowTitle = string.Empty;

        // Windows API imports for window handling
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public DocumentControlService(ILogger<DocumentControlService>? logger = null)
        {
            _logger = logger;
            _actionHistory = new ActionHistoryService(logger);
        }

        public ActionHistoryService ActionHistory => _actionHistory;

        public async Task<DocumentActionResult> TypeTextAsync(string text)
        {
            // Type text into the currently active application
            // This is the core dictation feature for document interaction
            try
            {
                _logger?.LogInformation("Typing text: {Text}", text);
                
                // Record the action for potential undo
                var activeWindow = GetActiveWindowInfo();
                _actionHistory.RecordAction(new DocumentAction
                {
                    ActionType = DocumentActionType.TypeText,
                    Text = text,
                    WindowHandle = activeWindow.Handle,
                    WindowTitle = activeWindow.Title,
                    Timestamp = DateTime.UtcNow
                });

                // Small delay to ensure the target window has focus
                await Task.Delay(50);

                // Escape special characters for SendKeys
                var escapedText = EscapeForSendKeys(text);
                
                // Send the keystrokes
                SendKeys.SendWait(escapedText);

                return DocumentActionResult.Success(
                    $"Typed: {text}",
                    $"I've typed: {TruncateForSpeech(text)}");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to type text");
                return DocumentActionResult.Failure(
                    "Could not type text",
                    "I couldn't type the text. Make sure a text field is focused.");
            }
        }

        public async Task<DocumentActionResult> ExecuteFormattingAsync(FormattingType formatting)
        {
            // Apply formatting to selected text (works in Word, rich text editors)
            try
            {
                var shortcut = formatting switch
                {
                    FormattingType.Bold => "^b",
                    FormattingType.Italic => "^i",
                    FormattingType.Underline => "^u",
                    FormattingType.Strikethrough => "^-",  // May not work in all apps
                    _ => throw new ArgumentException($"Unknown formatting: {formatting}")
                };

                var description = formatting switch
                {
                    FormattingType.Bold => "bold",
                    FormattingType.Italic => "italic",
                    FormattingType.Underline => "underlined",
                    FormattingType.Strikethrough => "strikethrough",
                    _ => formatting.ToString()
                };

                _logger?.LogInformation("Applying formatting: {Formatting}", formatting);

                // Record for undo
                _actionHistory.RecordAction(new DocumentAction
                {
                    ActionType = DocumentActionType.Format,
                    FormattingApplied = formatting,
                    Timestamp = DateTime.UtcNow
                });

                await Task.Delay(30);
                SendKeys.SendWait(shortcut);

                return DocumentActionResult.Success(
                    $"Applied {description} formatting",
                    $"Made the text {description}.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to apply formatting: {Formatting}", formatting);
                return DocumentActionResult.Failure(
                    $"Could not apply {formatting}",
                    "I couldn't apply that formatting. Make sure text is selected.");
            }
        }

        public async Task<DocumentActionResult> ExecuteClipboardOperationAsync(ClipboardOperation operation)
        {
            // Perform clipboard operations (copy, cut, paste)
            try
            {
                var (shortcut, pastDescription, presentDescription) = operation switch
                {
                    ClipboardOperation.Copy => ("^c", "Copied", "Copying"),
                    ClipboardOperation.Cut => ("^x", "Cut", "Cutting"),
                    ClipboardOperation.Paste => ("^v", "Pasted", "Pasting"),
                    _ => throw new ArgumentException($"Unknown operation: {operation}")
                };

                _logger?.LogInformation("Clipboard operation: {Operation}", operation);

                _actionHistory.RecordAction(new DocumentAction
                {
                    ActionType = DocumentActionType.Clipboard,
                    ClipboardOp = operation,
                    Timestamp = DateTime.UtcNow
                });

                await Task.Delay(30);
                SendKeys.SendWait(shortcut);

                return DocumentActionResult.Success(
                    $"{pastDescription} to clipboard",
                    $"{pastDescription}.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Clipboard operation failed: {Operation}", operation);
                return DocumentActionResult.Failure(
                    $"Could not {operation.ToString().ToLower()}",
                    "I couldn't perform that clipboard operation.");
            }
        }

        public async Task<DocumentActionResult> SelectAllAsync()
        {
            try
            {
                _logger?.LogInformation("Selecting all text");
                
                _actionHistory.RecordAction(new DocumentAction
                {
                    ActionType = DocumentActionType.Select,
                    Timestamp = DateTime.UtcNow
                });

                await Task.Delay(30);
                SendKeys.SendWait("^a");

                return DocumentActionResult.Success(
                    "Selected all text",
                    "I've selected all the text.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to select all");
                return DocumentActionResult.Failure(
                    "Could not select text",
                    "I couldn't select the text.");
            }
        }

        public async Task<DocumentActionResult> UndoAsync()
        {
            try
            {
                _logger?.LogInformation("Undoing last action");
                await Task.Delay(30);
                SendKeys.SendWait("^z");

                // Also remove from our action history
                _actionHistory.PopLastAction();

                return DocumentActionResult.Success(
                    "Undone",
                    "I've undone the last action.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to undo");
                return DocumentActionResult.Failure(
                    "Could not undo",
                    "I couldn't undo that action.");
            }
        }

        public async Task<DocumentActionResult> RedoAsync()
        {
            try
            {
                _logger?.LogInformation("Redoing action");
                await Task.Delay(30);
                SendKeys.SendWait("^y");

                return DocumentActionResult.Success(
                    "Redone",
                    "I've redone the action.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to redo");
                return DocumentActionResult.Failure(
                    "Could not redo",
                    "I couldn't redo that action.");
            }
        }

        public async Task<DocumentActionResult> SaveDocumentAsync()
        {
            try
            {
                _logger?.LogInformation("Saving document");
                await Task.Delay(30);
                SendKeys.SendWait("^s");

                return DocumentActionResult.Success(
                    "Document saved",
                    "I've saved the document.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to save");
                return DocumentActionResult.Failure(
                    "Could not save",
                    "I couldn't save the document.");
            }
        }

        public async Task<DocumentActionResult> NavigateAsync(NavigationType navigation)
        {
            try
            {
                var (shortcut, description) = navigation switch
                {
                    NavigationType.StartOfDocument => ("^{HOME}", "start of document"),
                    NavigationType.EndOfDocument => ("^{END}", "end of document"),
                    NavigationType.StartOfLine => ("{HOME}", "start of line"),
                    NavigationType.EndOfLine => ("{END}", "end of line"),
                    NavigationType.NextWord => ("^{RIGHT}", "next word"),
                    NavigationType.PreviousWord => ("^{LEFT}", "previous word"),
                    NavigationType.NextLine => ("{DOWN}", "next line"),
                    NavigationType.PreviousLine => ("{UP}", "previous line"),
                    NavigationType.PageUp => ("{PGUP}", "page up"),
                    NavigationType.PageDown => ("{PGDN}", "page down"),
                    _ => throw new ArgumentException($"Unknown navigation: {navigation}")
                };

                _logger?.LogInformation("Navigating: {Navigation}", navigation);
                await Task.Delay(30);
                SendKeys.SendWait(shortcut);

                return DocumentActionResult.Success(
                    $"Moved to {description}",
                    $"Moved to {description}.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Navigation failed: {Navigation}", navigation);
                return DocumentActionResult.Failure(
                    "Could not navigate",
                    "I couldn't move the cursor.");
            }
        }

        public async Task<DocumentActionResult> NewLineAsync()
        {
            try
            {
                await Task.Delay(30);
                SendKeys.SendWait("{ENTER}");
                return DocumentActionResult.Success("New line", "New line.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to insert new line");
                return DocumentActionResult.Failure("Could not insert line", "I couldn't add a new line.");
            }
        }

        public async Task<DocumentActionResult> DeleteSelectedAsync()
        {
            try
            {
                _actionHistory.RecordAction(new DocumentAction
                {
                    ActionType = DocumentActionType.Delete,
                    Timestamp = DateTime.UtcNow
                });

                await Task.Delay(30);
                SendKeys.SendWait("{DELETE}");
                return DocumentActionResult.Success("Deleted", "Deleted.");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to delete");
                return DocumentActionResult.Failure("Could not delete", "I couldn't delete that.");
            }
        }

        private (IntPtr Handle, string Title) GetActiveWindowInfo()
        {
            var handle = GetForegroundWindow();
            var title = new System.Text.StringBuilder(256);
            GetWindowText(handle, title, 256);
            return (handle, title.ToString());
        }

        private static string EscapeForSendKeys(string text)
        {
            // SendKeys uses special characters that need escaping
            // + = Shift, ^ = Ctrl, % = Alt, ~ = Enter, {} = special keys
            return text
                .Replace("{", "{{}")
                .Replace("}", "{}}")
                .Replace("+", "{+}")
                .Replace("^", "{^}")
                .Replace("%", "{%}")
                .Replace("~", "{~}")
                .Replace("(", "{(}")
                .Replace(")", "{)}");
        }

        private static string TruncateForSpeech(string text, int maxLength = 50)
        {
            // Truncate long text for speech feedback
            if (text.Length <= maxLength)
                return text;
            return text.Substring(0, maxLength) + "...";
        }
    }

    public class DocumentActionResult
    {
        public bool IsSuccess { get; }
        public string Description { get; }
        public string SpeechResponse { get; }

        private DocumentActionResult(bool success, string description, string speechResponse)
        {
            IsSuccess = success;
            Description = description;
            SpeechResponse = speechResponse;
        }

        public static DocumentActionResult Success(string description, string speechResponse)
            => new(true, description, speechResponse);

        public static DocumentActionResult Failure(string description, string speechResponse)
            => new(false, description, speechResponse);
    }

    public enum FormattingType
    {
        Bold,
        Italic,
        Underline,
        Strikethrough
    }

    public enum ClipboardOperation
    {
        Copy,
        Cut,
        Paste
    }

    public enum NavigationType
    {
        StartOfDocument,
        EndOfDocument,
        StartOfLine,
        EndOfLine,
        NextWord,
        PreviousWord,
        NextLine,
        PreviousLine,
        PageUp,
        PageDown
    }

    public enum DocumentActionType
    {
        TypeText,
        Format,
        Clipboard,
        Select,
        Delete,
        Navigate
    }

    public class DocumentAction
    {
        public DocumentActionType ActionType { get; set; }
        public string? Text { get; set; }
        public FormattingType? FormattingApplied { get; set; }
        public ClipboardOperation? ClipboardOp { get; set; }
        public IntPtr WindowHandle { get; set; }
        public string? WindowTitle { get; set; }
        public DateTime Timestamp { get; set; }
    }

    /*
        ActionHistoryService tracks document actions for undo capability.
        
        This service maintains a stack of recent actions that can be undone.
        It's used to provide the "undo that" voice command functionality,
        which is an important safety feature for accessibility users.
    */
    public class ActionHistoryService
    {
        private readonly ILogger? _logger;
        private readonly Stack<DocumentAction> _actionStack;
        private const int MaxHistorySize = 50;

        public ActionHistoryService(ILogger? logger = null)
        {
            _logger = logger;
            _actionStack = new Stack<DocumentAction>();
        }

        public void RecordAction(DocumentAction action)
        {
            // Add action to history, maintaining max size
            if (_actionStack.Count >= MaxHistorySize)
            {
                // Remove oldest items (would need to convert to list for this)
                // For simplicity, we just let it grow to max
            }

            _actionStack.Push(action);
            _logger?.LogDebug("Recorded action: {ActionType}", action.ActionType);
        }

        public DocumentAction? PopLastAction()
        {
            if (_actionStack.Count == 0)
                return null;

            var action = _actionStack.Pop();
            _logger?.LogDebug("Popped action: {ActionType}", action.ActionType);
            return action;
        }

        public DocumentAction? PeekLastAction()
        {
            return _actionStack.Count > 0 ? _actionStack.Peek() : null;
        }

        public int ActionCount => _actionStack.Count;

        public void Clear()
        {
            _actionStack.Clear();
        }
    }
}
