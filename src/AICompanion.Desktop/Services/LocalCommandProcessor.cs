using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using AICompanion.Desktop.Models;
using FuzzySharp;
using FuzzyProcess = FuzzySharp.Process;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services
{
    /*
        LocalCommandProcessor handles voice commands locally without AI.

        This service provides basic command recognition and execution for
        common tasks like opening applications, controlling windows, and
        system operations. It serves as a fallback when AI is unavailable
        and demonstrates the application's capabilities.
    */
    public class LocalCommandProcessor
    {
        private readonly ILogger<LocalCommandProcessor>? _logger;

        // Phase 6: Reuse a single HttpClient for /api/intent calls (performance)
        private static readonly System.Net.Http.HttpClient _sharedHttpClient = new()
        {
            Timeout = TimeSpan.FromSeconds(5)
        };

        // P/Invoke — only for CaptureTargetWindow (read-only window info)
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        // Store the last target window (before AICompanion took focus)
        private IntPtr _lastTargetWindow = IntPtr.Zero;
        private string _lastTargetWindowTitle = "";

        // SMART MEMORY: Remember the last text editor window for easier dictation
        private IntPtr _lastTextEditorWindow = IntPtr.Zero;
        private string _lastTextEditorTitle = "";
        private string _lastTextEditorProcess = ""; // "WINWORD", "notepad", etc.

        public void CaptureTargetWindow()
        {
            // Call this BEFORE starting voice recognition to capture the window user wants to control
            var currentWindow = GetForegroundWindow();
            var title = GetActiveWindowTitle(currentWindow);

            // Don't capture AICompanion itself
            if (!title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase))
            {
                _lastTargetWindow = currentWindow;
                _lastTargetWindowTitle = title;
                _logger?.LogInformation("[CAPTURE] Stored target window: '{Title}' (Handle: {Handle})", title, currentWindow);

                // SMART MEMORY: If this is a text editor, remember it for future dictation
                if (IsTextEditorWindow(title))
                {
                    _lastTextEditorWindow = currentWindow;
                    _lastTextEditorTitle = title;
                    _lastTextEditorProcess = DetectProcessType(title);
                    _logger?.LogInformation("[SMART] Remembered text editor: '{Title}' ({Process})", title, _lastTextEditorProcess);
                }
            }
        }

        private bool IsTextEditorWindow(string title)
        {
            var lower = title.ToLowerInvariant();
            return lower.Contains("word") ||
                   lower.Contains(".docx") ||
                   lower.Contains(".doc") ||
                   lower.Contains("notepad") ||
                   lower.Contains("блокнот") ||
                   lower.Contains(".txt") ||
                   lower.Contains("document") ||
                   lower.Contains("документ");
        }

        private string DetectProcessType(string title)
        {
            var lower = title.ToLowerInvariant();
            if (lower.Contains("word") || lower.Contains(".docx") || lower.Contains(".doc"))
                return "WINWORD";
            if (lower.Contains("notepad") || lower.Contains("блокнот") || lower.Contains(".txt"))
                return "notepad";
            return "unknown";
        }

        public IntPtr GetTargetWindow()
        {
            // Return the captured target window, or fall back to current foreground
            if (_lastTargetWindow != IntPtr.Zero)
            {
                // Verify window still exists
                var title = GetActiveWindowTitle(_lastTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                {
                    return _lastTargetWindow;
                }
            }
            return GetForegroundWindow();
        }

        public string GetTargetWindowTitle()
        {
            return !string.IsNullOrEmpty(_lastTargetWindowTitle) ? _lastTargetWindowTitle : "(none)";
        }

        private string GetActiveWindowTitle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return "(none)";
            var sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            return sb.ToString();
        }

        public string GetCurrentActiveWindowInfo()
        {
            var hWnd = GetForegroundWindow();
            var title = GetActiveWindowTitle(hWnd);
            return $"Window: '{title}' (Handle: {hWnd})";
        }

        /// <summary>
        /// Returns the last remembered text editor window title (Word, Notepad, etc.)
        /// This persists across voice commands for easier dictation.
        /// </summary>
        public string GetActiveTextEditorTitle()
        {
            if (_lastTextEditorWindow != IntPtr.Zero)
            {
                var title = GetActiveWindowTitle(_lastTextEditorWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                {
                    return title;
                }
            }
            return "(no text editor)";
        }

        /// <summary>
        /// Clears the remembered text editor (user can say "forget editor" or similar)
        /// </summary>
        public void ClearTextEditorMemory()
        {
            _lastTextEditorWindow = IntPtr.Zero;
            _lastTextEditorTitle = "";
            _lastTextEditorProcess = "";
            _logger?.LogInformation("[SMART] Text editor memory cleared");
        }

        // Common application mappings with synonyms for fuzzy matching
        private readonly Dictionary<string, string> _appMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            // Russian
            { "блокнот", "notepad" },
            { "калькулятор", "calc" },
            { "браузер", "msedge" },
            { "хром", "chrome" },
            { "проводник", "explorer" },
            { "ворд", "WINWORD" },
            { "эксель", "EXCEL" },
            { "пэинт", "mspaint" },
            { "пейнт", "mspaint" },
            { "командную строку", "cmd" },
            { "терминал", "wt" },
            { "настройки", "ms-settings:" },

            // English - primary names
            { "notepad", "notepad" },
            { "notebook", "notepad" },
            { "calculator", "calc" },
            { "browser", "msedge" },
            { "chrome", "chrome" },
            { "explorer", "explorer" },
            { "word", "WINWORD" },
            { "excel", "EXCEL" },
            { "paint", "mspaint" },
            { "command prompt", "cmd" },
            { "terminal", "wt" },
            { "settings", "ms-settings:" },
            { "edge", "msedge" },
            { "firefox", "firefox" },
            
            // Additional synonyms for better recognition
            { "note pad", "notepad" },
            { "note book", "notepad" },
            { "text editor", "notepad" },
            { "calc", "calc" },
            { "the browser", "msedge" },
            { "microsoft edge", "msedge" },
            { "google chrome", "chrome" },
            { "chrome browser", "chrome" },
            { "file explorer", "explorer" },
            { "files", "explorer" },
            { "microsoft word", "WINWORD" },
            { "ms word", "WINWORD" },
            { "microsoft excel", "EXCEL" },
            { "ms excel", "EXCEL" },
            { "spreadsheet", "EXCEL" },
            { "ms paint", "mspaint" },
            { "cmd", "cmd" },
            { "powershell", "powershell" },
            { "windows terminal", "wt" },
            
            // Additional synonyms for better Watson recognition
            { "the notepad", "notepad" },
            { "a notepad", "notepad" },
            { "my notepad", "notepad" },
            { "open notepad", "notepad" },
            { "the calculator", "calc" },
            { "a calculator", "calc" },
            { "the word", "WINWORD" },
            { "document", "WINWORD" },
            { "word document", "WINWORD" },
            { "the excel", "EXCEL" },
            { "spreadsheets", "EXCEL" },
            { "web browser", "msedge" },
            { "internet", "msedge" },
            { "the internet", "msedge" },
            { "my files", "explorer" },
            { "my documents", "explorer" },
            { "folder", "explorer" },
            { "folders", "explorer" },
        };

        // Phrases to strip from the beginning of commands for flexible matching
        private readonly string[] _prefixPhrases = new[]
        {
            "please", "can you", "could you", "would you", "i want to", "i'd like to",
            "i would like to", "kindly", "go ahead and", "just", "now",
            "пожалуйста", "можешь", "можешь ли ты", "хочу"
        };

        // Command patterns (Russian and English)
        private readonly List<(Regex Pattern, Func<Match, CommandResult> Handler)> _commandPatterns;

        public LocalCommandProcessor(ILogger<LocalCommandProcessor>? logger = null)
        {
            _logger = logger;
            _commandPatterns = InitializePatterns();
        }

        private List<(Regex Pattern, Func<Match, CommandResult> Handler)> InitializePatterns()
        {
            // PHASE 6 CLEANUP: Only instant/local commands kept here.
            // ALL document editing (type, copy, paste, bold, italic, save, undo, redo,
            // select, delete, format, window-aware typing) is handled by IBM Granite
            // via /api/intent (single action) or /api/plan (multi-step agentic).
            return new List<(Regex, Func<Match, CommandResult>)>
            {
                // Open application
                (new Regex(@"(?:открой|открыть|запусти|open|start|launch)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteOpenApp(m.Groups[1].Value.Trim())),

                // Close window
                (new Regex(@"(?:закрой|закрыть|close)\s*(?:это|this|окно|window)?", RegexOptions.IgnoreCase),
                    _ => ExecuteCloseWindow()),

                // Time
                (new Regex(@"(?:который час|сколько времени|what time|current time)", RegexOptions.IgnoreCase),
                    _ => ExecuteGetTime()),

                // Date
                (new Regex(@"(?:какой сегодня день|какая дата|what day|what date|today)", RegexOptions.IgnoreCase),
                    _ => ExecuteGetDate()),

                // Minimize
                (new Regex(@"(?:сверни|свернуть|minimize)", RegexOptions.IgnoreCase),
                    _ => ExecuteMinimize()),

                // Help
                (new Regex(@"(?:помощь|справка|help|что ты умеешь|what can you do)", RegexOptions.IgnoreCase),
                    _ => ExecuteHelp()),

                // Web search
                (new Regex(@"(?:найди|поиск|search|find)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteSearch(m.Groups[1].Value.Trim())),

                // Greeting
                (new Regex(@"^(?:привет|здравствуй|hello|hi|hey)$", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Hello! I'm your AI assistant.",
                        "Hello! I'm ready to assist you.")),

                // How do I
                (new Regex(@"(?:how do i|how can i|как мне|как я могу)\s+(.+)", RegexOptions.IgnoreCase),
                    m => ExecuteHowDoI(m.Groups[1].Value.Trim())),

                // What commands
                (new Regex(@"(?:what commands|list commands|show commands|какие команды|покажи команды)", RegexOptions.IgnoreCase),
                    _ => ExecuteShowCommands()),

                // Tutorial markers
                (new Regex(@"(?:start tutorial|begin tutorial|teach me|начать обучение)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Starting tutorial", "TUTORIAL_START")),
                (new Regex(@"(?:stop tutorial|end tutorial|exit tutorial|закончить обучение)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Stopping tutorial", "TUTORIAL_STOP")),
                (new Regex(@"(?:skip|skip this|next step|пропустить)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Skipping step", "TUTORIAL_SKIP")),
                (new Regex(@"(?:give me a hint|hint|подсказка)", RegexOptions.IgnoreCase),
                    _ => new CommandResult(true, "Hint requested", "TUTORIAL_HINT")),
            };
        }

        public CommandResult ProcessCommand(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return new CommandResult(false, "Empty command", "");
            }

            _logger?.LogInformation("Processing local command: {Text}", text);

            // === PHASE 5 FIX: Intercept complex Agentic Commands FIRST ===
            // If the user says 4+ words with multiple actions, don't let simple
            // "open" or "type" regexes hijack it. Send straight to Granite!
            if (IsComplexCommand(text))
            {
                _logger?.LogInformation("[AGENTIC] Complex command detected at top level, bypassing local regex: {Text}", text);
                return AgenticPlanRequired;
            }

            // Strip common prefix phrases for more flexible matching
            // This allows "please open notepad" to work the same as "open notepad"
            var normalizedText = StripPrefixPhrases(text);
            _logger?.LogDebug("Normalized command: {NormalizedText}", normalizedText);

            // Try to match against known patterns with normalized text
            foreach (var (pattern, handler) in _commandPatterns)
            {
                var match = pattern.Match(normalizedText);
                if (match.Success)
                {
                    try
                    {
                        return handler(match);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, "Error executing command");
                        return new CommandResult(false, $"Error: {ex.Message}", "");
                    }
                }
            }

            // Also try with original text in case normalization removed important parts
            if (normalizedText != text)
            {
                foreach (var (pattern, handler) in _commandPatterns)
                {
                    var match = pattern.Match(text);
                    if (match.Success)
                    {
                        try
                        {
                            return handler(match);
                        }
                        catch (Exception ex)
                        {
                            _logger?.LogError(ex, "Error executing command");
                            return new CommandResult(false, $"Error: {ex.Message}", "");
                        }
                    }
                }
            }

            // === BUG 4 FIX: Fuzzy matching fallback ===
            // If regex failed, try Levenshtein-based fuzzy matching
            var fuzzyResult = FuzzyMatchCommand(normalizedText);
            if (fuzzyResult != null)
            {
                _logger?.LogInformation("[FUZZY] Matched '{Input}' via fuzzy matching", text);
                return fuzzyResult;
            }

            // === PHASE 2: AI Backend Fallback ===
            // When both regex and fuzzy matching fail, try the Python AI backend
            var aiResult = FallbackToAIBackend(text);
            if (aiResult != null)
            {
                _logger?.LogInformation("[AI-FALLBACK] Matched '{Input}' via Python backend", text);
                return aiResult;
            }

            // Truly unknown command
            return new CommandResult(false, $"Unknown command: {text}", "");
        }

        /// <summary>
        /// BUG 4 FIX: Fuzzy command matching using Levenshtein distance ratio.
        /// When regex patterns fail, tries to match the input against known command
        /// verbs and application names using a similarity ratio threshold of 0.80.
        /// Ratio = 1 - (levenshteinDistance / max(len1, len2))
        /// </summary>
        private CommandResult? FuzzyMatchCommand(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;

            var words = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) return null;

            // PHASE 6: Only local-only verbs here. Everything else → Granite /api/intent
            var commandVerbs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "open", "open" }, { "launch", "open" }, { "start", "open" },
                { "close", "close" }, { "minimize", "minimize" },
                { "help", "help" }, { "search", "search" }, { "find", "search" },
                // Russian
                { "открой", "open" }, { "запусти", "open" },
                { "закрой", "close" }, { "сверни", "minimize" },
                { "помощь", "help" }, { "найди", "search" },
            };

            string? matchedVerb = null;
            int bestScore = 0;

            foreach (var verb in commandVerbs.Keys)
            {
                var score = Fuzz.Ratio(words[0].ToLowerInvariant(), verb.ToLowerInvariant());
                if (score > bestScore && score >= 80)
                {
                    bestScore = score;
                    matchedVerb = commandVerbs[verb];
                }
            }

            if (matchedVerb == null) return null;

            _logger?.LogInformation("[FUZZY] Verb '{Input}' matched to '{Verb}' (score: {Score})",
                words[0], matchedVerb, bestScore);

            return matchedVerb switch
            {
                "open" when words.Length > 1 => ExecuteOpenApp(string.Join(" ", words.Skip(1))),
                "close" => ExecuteCloseWindow(),
                "minimize" => ExecuteMinimize(),
                "help" => ExecuteHelp(),
                "search" when words.Length > 1 => ExecuteSearch(string.Join(" ", words.Skip(1))),
                _ => null
            };
        }

        /// <summary>
        /// Determines if a command is complex enough to require multi-step agentic planning.
        /// Commands with 4+ words that contain multiple action verbs or conjunctions ("and", "then")
        /// are routed to the Python /api/plan endpoint via AgenticExecutionService.
        /// </summary>
        public bool IsComplexCommand(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;
            var lower = input.ToLowerInvariant();
            var words = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            // Phase 7: Conversational prefixes always go to agentic planner
            // "Can you open notepad", "Please write hello", "Could you save"
            var conversationalPrefixes = new[] {
                "can you", "could you", "would you", "please",
                "i want", "i need", "i'd like", "go ahead",
                "можешь", "пожалуйста", "мне нужно", "я хочу"
            };
            foreach (var prefix in conversationalPrefixes)
            {
                if (lower.StartsWith(prefix))
                {
                    _logger?.LogInformation("[AGENTIC] Conversational prefix detected: '{Prefix}'", prefix);
                    return true;
                }
            }

            // Multi-step indicators (conjunctions)
            var hasConjunction = lower.Contains(" and ") || lower.Contains(" then ") ||
                                 lower.Contains(" after ") || lower.Contains(", ") ||
                                 lower.Contains(" и ") || lower.Contains(" потом ") ||
                                 lower.Contains(" затем ");
            if (hasConjunction) return true;

            // Multiple action verbs in one sentence
            var actionVerbs = new[] { "open", "type", "write", "save", "close", "bold", "italic",
                                      "copy", "paste", "search", "delete", "format", "select",
                                      "create", "new", "tab", "screenshot", "refresh",
                                      "открой", "напиши", "сохрани", "закрой", "найди", "выдели",
                                      "создай", "новый", "вкладк", "скриншот" };
            int verbCount = 0;
            foreach (var verb in actionVerbs)
            {
                if (lower.Contains(verb)) verbCount++;
            }
            if (verbCount >= 2) return true;

            // 4+ words → likely needs planning (but simple 2-3 word commands stay local)
            if (words.Length >= 4) return true;

            return false;
        }

        /// <summary>
        /// Special return value that signals the caller to route this command
        /// through AgenticExecutionService for multi-step plan execution.
        /// </summary>
        public static readonly CommandResult AgenticPlanRequired = new CommandResult(
            false, "AGENTIC_PLAN_REQUIRED", "AGENTIC_PLAN_REQUIRED");

        /// <summary>
        /// PHASE 5: Fallback to Python FastAPI backend for AI-powered intent extraction.
        /// For complex multi-step commands (4+ words with conjunctions), returns a special
        /// marker so the caller routes to AgenticExecutionService (/api/plan).
        /// For simple commands, uses /api/intent for single-action extraction.
        /// </summary>
        private CommandResult? FallbackToAIBackend(string input)
        {
            // Route complex commands to agentic planner
            if (IsComplexCommand(input))
            {
                _logger?.LogInformation("[AI-FALLBACK] Complex command detected, routing to /api/plan: {Input}", input);
                return AgenticPlanRequired;
            }

            try
            {
                // Phase 6: Include window context so the LLM knows what app is active
                var windowTitle = GetTargetWindowTitle();
                var windowProcess = !string.IsNullOrEmpty(_lastTextEditorProcess) ? _lastTextEditorProcess : "(unknown)";

                var payload = new
                {
                    text = input,
                    window_title = windowTitle,
                    window_process = windowProcess
                };
                var json = System.Text.Json.JsonSerializer.Serialize(payload);
                var content = new System.Net.Http.StringContent(
                    json, System.Text.Encoding.UTF8, "application/json");

                _logger?.LogInformation("[AI-FALLBACK] Sending to /api/intent: '{Input}' (window: {Win})", input, windowTitle);

                var response = _sharedHttpClient.PostAsync(
                    "http://localhost:8000/api/intent", content).Result;

                if (!response.IsSuccessStatusCode)
                {
                    _logger?.LogWarning("[AI-FALLBACK] Server returned {Status}", response.StatusCode);
                    return null;
                }

                var responseJson = response.Content.ReadAsStringAsync().Result;
                using var doc = System.Text.Json.JsonDocument.Parse(responseJson);
                var root = doc.RootElement;

                var action = root.GetProperty("action").GetString() ?? "unknown";
                var target = root.TryGetProperty("target", out var t) ? t.GetString() : null;

                _logger?.LogInformation("[AI-FALLBACK] Granite returned: action={Action}, target={Target}",
                    action, target);

                // PHASE 6: Only open_app/search/minimize handled locally.
                // All document-editing actions route to AgenticExecutionService.
                return action switch
                {
                    "open_app" when target != null => ExecuteOpenApp(target),
                    "search_web" when target != null => ExecuteSearch(target),
                    "minimize" => ExecuteMinimize(),
                    "open_settings" => ExecuteHelp(),
                    // Everything else (type, format, save, copy, paste, etc.)
                    // goes through the agentic pipeline for reliable Win32 execution
                    _ when action != "unknown" => AgenticPlanRequired,
                    _ => null
                };
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                _logger?.LogDebug("[AI-FALLBACK] Python server not available: {Error}", ex.Message);
                return null; // Silent fail — local functionality continues
            }
            catch (TaskCanceledException)
            {
                _logger?.LogDebug("[AI-FALLBACK] Request timed out");
                return null;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[AI-FALLBACK] Error: {Error}", ex.Message);
                return null;
            }
        }

        private string StripPrefixPhrases(string text)
        {
            var result = text.Trim();
            foreach (var prefix in _prefixPhrases)
            {
                if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    result = result.Substring(prefix.Length).TrimStart();
                    break;
                }
            }
            return result;
        }

        private CommandResult ExecuteOpenApp(string appName)
        {
            string processName = appName;
            var cleanedAppName = appName.Trim().ToLowerInvariant();

            // Remove common articles and words that might be transcribed
            cleanedAppName = Regex.Replace(cleanedAppName, @"^(the|a|an)\s+", "");

            // First try exact match
            if (_appMappings.TryGetValue(cleanedAppName, out var mapped))
            {
                processName = mapped;
            }
            else
            {
                // Try fuzzy matching if exact match fails
                var bestMatch = FindBestAppMatch(cleanedAppName);
                if (bestMatch != null)
                {
                    processName = bestMatch;
                    _logger?.LogInformation("Fuzzy matched '{Input}' to '{Match}'", appName, processName);
                }
                else
                {
                    // Use original input as process name if no match found
                    processName = appName;
                }
            }

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = processName,
                    UseShellExecute = true
                };

                System.Diagnostics.Process.Start(psi);

                var friendlyName = GetFriendlyName(processName);
                return new CommandResult(true, $"Opened {friendlyName}",
                    $"Opening {friendlyName} for you!");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Failed to open {App}", appName);
                return new CommandResult(false, $"Could not open {appName}",
                    $"Sorry, I couldn't find or open {appName}.");
            }
        }

        private string? FindBestAppMatch(string input)
        {
            // Use FuzzySharp.Process.ExtractOne for best fuzzy match
            // Score threshold: 80 (out of 100) allows small typos and speech variations
            var choices = _appMappings.Keys.ToList();
            var bestMatch = FuzzyProcess.ExtractOne(input, choices);

            if (bestMatch != null && bestMatch.Score >= 80)
            {
                _logger?.LogInformation("[FUZZY-APP] '{Input}' matched to '{Match}' (score: {Score})",
                    input, bestMatch.Value, bestMatch.Score);

                if (_appMappings.TryGetValue(bestMatch.Value, out var processName))
                {
                    return processName;
                }
            }

            return null;
        }

        private CommandResult ExecuteCloseWindow()
        {
            // PHASE 6: Route close to AgenticExecutionService for reliable Win32 execution
            return AgenticPlanRequired;
        }

        private CommandResult ExecuteGetTime()
        {
            var time = DateTime.Now.ToString("HH:mm");
            return new CommandResult(true, $"Current time: {time}",
                $"The current time is {time}.");
        }

        private CommandResult ExecuteGetDate()
        {
            var date = DateTime.Now.ToString("dddd, MMMM d, yyyy");
            return new CommandResult(true, $"Today's date: {date}",
                $"Today is {date}.");
        }

        private CommandResult ExecuteMinimize()
        {
            try
            {
                // Minimize all windows
                var shell = new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(shell);
                return new CommandResult(true, "Windows minimized", "Done!");
            }
            catch
            {
                return new CommandResult(false, "Could not minimize", "I couldn't minimize windows.");
            }
        }

        private CommandResult ExecuteSearch(string query)
        {
            try
            {
                var searchUrl = $"https://www.google.com/search?q={Uri.EscapeDataString(query)}";
                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = searchUrl,
                    UseShellExecute = true
                });
                return new CommandResult(true, $"Searching for: {query}",
                    $"Searching for '{query}' in your browser.");
            }
            catch
            {
                return new CommandResult(false, "Search failed",
                    "I couldn't open the search.");
            }
        }

        private CommandResult ExecuteHelp()
        {
            var helpText = @"Available commands:

📂 APPLICATION CONTROL:
• 'Open [app]' - Open notepad, calculator, browser, Word, Excel
• 'Close' - Close current window
• 'Minimize' - Minimize all windows

📝 DOCUMENT EDITING (works in Notepad, Word):
• 'Type: [text]' - Type text into active application
• 'Select all' - Select all text
• 'Copy' / 'Paste' / 'Cut' - Clipboard operations
• 'Undo' / 'Redo' - Undo or redo last action
• 'Bold' / 'Italic' / 'Underline' - Format selected text
• 'Save' - Save the document
• 'Delete' - Delete selected text

🎯 WINDOW-AWARE COMMANDS (NEW!):
• 'In Word write [text]' - Type directly into Word
• 'In Notepad write [text]' - Type directly into Notepad
• 'In browser open [url]' - Open URL in browser
• 'В ворде напиши [текст]' - Напишет в Word
• 'В блокноте напиши [текст]' - Напишет в Notepad

🔍 UTILITIES:
• 'Search [query]' - Search the web
• 'What time is it?' - Get current time
• 'What date?' - Get current date

❓ LEARNING:
• 'How do I [task]?' - Get help with a specific task
• 'What commands' - Show all available commands

Все команды работают и на русском языке!";

            return new CommandResult(true, "Help displayed", helpText);
        }

        private CommandResult ExecuteHowDoI(string task)
        {
            // Teaching mode: explain how to perform common tasks
            var taskLower = task.ToLowerInvariant();
            
            var explanations = new Dictionary<string, string>
            {
                { "open", "To open an application, say 'Open' followed by the app name. For example: 'Open Notepad', 'Open Calculator', or 'Open Word'." },
                { "type", "To type text, say 'Type:' followed by what you want to write. For example: 'Type: Hello world'. The text will appear in the active window." },
                { "copy", "To copy text, first select it by saying 'Select all', then say 'Copy'. The text is now in your clipboard." },
                { "paste", "To paste text, position your cursor where you want it and say 'Paste'. This inserts text from your clipboard." },
                { "format", "To format text, first select it, then say 'Bold', 'Italic', or 'Underline'. This works in Word and other rich text editors." },
                { "bold", "To make text bold, select the text first, then say 'Bold' or 'Make this bold'." },
                { "save", "To save a document, say 'Save'. This is the same as pressing Ctrl+S." },
                { "undo", "To undo your last action, say 'Undo'. You can undo multiple times." },
                { "search", "To search the web, say 'Search' followed by your query. For example: 'Search weather today'." },
                { "close", "To close the current window, say 'Close' or 'Close this window'." },
            };

            foreach (var kvp in explanations)
            {
                if (taskLower.Contains(kvp.Key))
                {
                    return new CommandResult(true, $"Help: {task}", kvp.Value);
                }
            }

            return new CommandResult(true, $"Help: {task}",
                $"I can help you open applications, type text, format documents, search the web, and more!");
        }

        private CommandResult ExecuteShowCommands()
        {
            // Show a categorized list of all commands
            var commandList = @"Here are all the commands I understand:

APPLICATIONS: open, start, launch, close, minimize
TYPING: type, write (followed by your text)
EDITING: select all, copy, paste, cut, delete, undo, redo
FORMATTING: bold, italic, underline, save
NAVIGATION: go to start, go to end, new line
UTILITIES: search, time, date, help

Try saying 'How do I copy text?' for detailed instructions on any command.";

            return new CommandResult(true, "Command list", commandList);
        }

        private string GetFriendlyName(string processName)
        {
            var friendlyNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "notepad", "Notepad" },
                { "calc", "Calculator" },
                { "msedge", "Microsoft Edge" },
                { "chrome", "Google Chrome" },
                { "explorer", "File Explorer" },
                { "WINWORD", "Microsoft Word" },
                { "EXCEL", "Microsoft Excel" },
                { "mspaint", "Paint" },
                { "cmd", "Command Prompt" },
                { "wt", "Windows Terminal" },
            };

            return friendlyNames.TryGetValue(processName, out var friendly) ? friendly : processName;
        }

    }

    public class CommandResult
    {
        public bool Success { get; }
        public string Description { get; }
        public string SpeechResponse { get; }

        public CommandResult(bool success, string description, string speechResponse)
        {
            Success = success;
            Description = description;
            SpeechResponse = speechResponse;
        }
    }
}
