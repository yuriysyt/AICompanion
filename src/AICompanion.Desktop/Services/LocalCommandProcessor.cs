using System;
using System.Collections.Generic;
using System.Diagnostics;
using SysProcess = System.Diagnostics.Process;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using AICompanion.Desktop.Models;
using FuzzySharp;
using FuzzyProcess = FuzzySharp.Process;
using Microsoft.Extensions.Logging;
using System.Windows.Automation;

namespace AICompanion.Desktop.Services
{
    public class LocalCommandProcessor
    {
        // ── Win32 ────────────────────────────────────────────────────────────────

        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsDelegate lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")] private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const byte VK_CONTROL      = 0x11;
        private const byte VK_N            = 0x4E;
        private const uint KEYEVENTF_KEYUP = 0x0002;

        private delegate bool EnumWindowsDelegate(IntPtr hWnd, IntPtr lParam);

        // ── HTTP clients ─────────────────────────────────────────────────────────

        private static readonly System.Net.Http.HttpClient _intentHttpClient = new() { Timeout = TimeSpan.FromSeconds(8) };
        private static readonly System.Net.Http.HttpClient _agentHttpClient  = new() { Timeout = TimeSpan.FromSeconds(30) };

        // ── Session & window state ───────────────────────────────────────────────

        private string _sessionId = "default";
        public void SetSession(string userId) => _sessionId = userId;

        private IntPtr _lastTargetWindow     = IntPtr.Zero;
        private string _lastTargetWindowTitle = "";
        private IntPtr _lastTextEditorWindow  = IntPtr.Zero;
        private string _lastTextEditorTitle   = "";
        private string _lastTextEditorProcess = "";
        private string _lastOpenedApp         = "";
        private string _lastTypedText         = "";

        // ── Context tracking ─────────────────────────────────────────────────────
        private enum AppContext { None, Word, Notepad, Calculator, Browser }
        private AppContext _currentContext = AppContext.None;
        private string     _currentDocPath = "";   // path of the most recently created/opened doc
        private string _lastNavigationTarget  = "";

        // ── OpenClaw-style routing table ─────────────────────────────────────────
        // Each entry maps a compiled Regex to a short action tag.
        // ProcessCommand resolves these tags to actual handler calls.

        private static readonly (Regex Pattern, string Action)[] _localRoutes =
        [
            (new Regex(@"(?:открой|открыть|запусти|open|start|launch)\s+(.+)",         RegexOptions.IgnoreCase), "open_app"),
            (new Regex(@"(?:закрой|закрыть|close)\s*(?:это|this|окно|window)?",        RegexOptions.IgnoreCase), "close"),
            (new Regex(@"(?:который час|сколько времени|what time|current time)",       RegexOptions.IgnoreCase), "time"),
            (new Regex(@"(?:какой сегодня день|какая дата|what day|what date|today)",   RegexOptions.IgnoreCase), "date"),
            (new Regex(@"(?:сверни|свернуть|minimize)",                                 RegexOptions.IgnoreCase), "minimize"),
            (new Regex(@"(?:помощь|справка|help|что ты умеешь|what can you do)",        RegexOptions.IgnoreCase), "help"),
            (new Regex(@"(?:найди|поиск|search|find|google|гугли)\s+(.+)",              RegexOptions.IgnoreCase), "search"),
            (new Regex(@"^(?:привет|здравствуй|hello|hi|hey)$",                         RegexOptions.IgnoreCase), "greet"),
            (new Regex(@"(?:how do i|how can i|как мне|как я могу)\s+(.+)",             RegexOptions.IgnoreCase), "how_do_i"),
            (new Regex(@"(?:what commands|list commands|show commands|какие команды|покажи команды)", RegexOptions.IgnoreCase), "list_commands"),
            (new Regex(@"(?:start tutorial|begin tutorial|teach me|начать обучение)",   RegexOptions.IgnoreCase), "tutorial_start"),
            (new Regex(@"(?:stop tutorial|end tutorial|exit tutorial|закончить обучение)", RegexOptions.IgnoreCase), "tutorial_stop"),
            (new Regex(@"(?:skip|skip this|next step|пропустить)",                      RegexOptions.IgnoreCase), "tutorial_skip"),
            (new Regex(@"(?:give me a hint|hint|подсказка)",                            RegexOptions.IgnoreCase), "tutorial_hint"),
            // Write AI content into Word (explicit "in word" target)
            (new Regex(@"^write\s+(?:(?:an?\s+)?essay\s+)?(?:about\s+)?(.+?)\s+(?:in(?:to)?|to)\s+(?:word|ворд)\s*$", RegexOptions.IgnoreCase), "write_word"),
            // Write AI-generated content to a new notebook file
            // Matches: "write essay about Baku in notebook", "write about AI in notepad", etc.
            (new Regex(@"^write\s+(?:(?:an?\s+)?essay\s+)?(?:about\s+)?(.+?)\s+(?:in(?:to)?|to)\s+(?:notebook|notepad|note\s*(?:file|book)|блокнот)\s*$", RegexOptions.IgnoreCase), "write_notebook"),
            // "write essay about X" with no explicit target → context-aware dispatch
            (new Regex(@"^write\s+(?:(?:an?\s+)?essay\s+)?about\s+(.+?)\s*$", RegexOptions.IgnoreCase), "write_ctx"),
            // Explicit Word document creation
            (new Regex(@"^(?:create\s+(?:a\s+)?new\s+|new\s+)(?:word|ворд)\s+(?:document|file|doc|page|файл|документ)", RegexOptions.IgnoreCase), "new_doc_word"),
            // Explicit Notepad/notebook creation
            (new Regex(@"^(?:create\s+(?:a\s+)?new\s+|new\s+)(?:notebook|note\s*book|блокнот|note\s+file|notepad\s+file)", RegexOptions.IgnoreCase), "new_doc_notepad"),
            // Explicit Excel creation
            (new Regex(@"^(?:create\s+(?:a\s+)?new\s+|new\s+)(?:excel|эксель)\s*(?:document|file|doc|sheet|spreadsheet|таблицу)?", RegexOptions.IgnoreCase), "new_doc_excel"),
            // Context-aware new document: "create document", "create a document", "create new document", "new document"
            (new Regex(@"^(?:create\s+(?:a\s+)?(?:new\s+)?|new\s+)(?:document|file|doc|page|файл|документ)", RegexOptions.IgnoreCase), "new_doc"),
            // ── Calculator app automation (System.Windows.Automation) ─────────────
            // "type X in calculator" — opens Calculator and types without returning result
            (new Regex(@"^type\s+(.+?)\s+in(?:to)?\s+calculator\s*$", RegexOptions.IgnoreCase), "calc_type"),
            // "calculate X", "compute X", "посчитай X" → open Calculator, run, read result
            (new Regex(@"^(?:calculate|compute|eval|посчитай|посчитайте|вычисли|please\s+calculate|can\s+you\s+calculate)\s+(.+)$", RegexOptions.IgnoreCase), "calc_app"),
            // Natural word-form: "123 plus 456", "999 times 7", "1000 minus 337", "100 divided by 4"
            (new Regex(@"^(\d[\d\s]*(?:plus|minus|times|divided\s+by|multiplied\s+by)\s*[\d\s]+)[\?\.]*$", RegexOptions.IgnoreCase), "calc_app"),
            // "what is 123 + 456" / "what is 999 times 7" (must start content with a digit)
            (new Regex(@"^what\s+is\s+(\d.+?)\s*[\?\.]*$", RegexOptions.IgnoreCase), "calc_app"),
            // ── Inline symbolic math (local DataTable.Compute — no Calculator window) ──
            (new Regex(@"^([\d\s\+\-\*\/\.\(\)]+(?:\s*[\+\-\*\/]\s*[\d\s\.\(\)]+)+)$", RegexOptions.IgnoreCase), "math_expr"),
            (new Regex(@"^([\d\s\+\-\*\/\.\(\)]+)=\s*\??\s*$", RegexOptions.IgnoreCase), "math_expr"),
        ];

        // ── App name dictionary ──────────────────────────────────────────────────

        private static readonly Dictionary<string, string> _appMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            // Russian
            { "блокнот",        "notepad"   }, { "калькулятор",    "calc"      },
            { "браузер",        "msedge"    }, { "хром",           "chrome"    },
            { "проводник",      "explorer"  }, { "ворд",           "WINWORD"   },
            { "эксель",         "EXCEL"     }, { "пэинт",          "mspaint"   },
            { "пейнт",          "mspaint"   }, { "командную строку","cmd"      },
            { "терминал",       "wt"        }, { "настройки",      "ms-settings:" },
            { "опера",          "opera"     },
            // English — single words
            { "notepad",        "notepad"   }, { "notebook",       "notepad"   },
            { "calculator",     "calc"      }, { "calc",           "calc"      },
            { "browser",        "msedge"    }, { "chrome",         "chrome"    },
            { "explorer",       "explorer"  }, { "word",           "WINWORD"   },
            { "excel",          "EXCEL"     }, { "paint",          "mspaint"   },
            { "terminal",       "wt"        }, { "settings",       "ms-settings:" },
            { "edge",           "msedge"    }, { "firefox",        "firefox"   },
            { "opera",          "opera"     }, { "cmd",            "cmd"       },
            { "powershell",     "powershell"}, { "document",       "WINWORD"   },
            { "internet",       "msedge"    }, { "folder",         "explorer"  },
            { "folders",        "explorer"  },
            // English — multi-word
            { "note pad",           "notepad"   }, { "note book",      "notepad"   },
            { "text editor",        "notepad"   }, { "the notepad",    "notepad"   },
            { "a notepad",          "notepad"   }, { "my notepad",     "notepad"   },
            { "open notepad",       "notepad"   }, { "the calculator", "calc"      },
            { "a calculator",       "calc"      }, { "the browser",    "msedge"    },
            { "web browser",        "msedge"    }, { "the internet",   "msedge"    },
            { "microsoft edge",     "msedge"    }, { "google chrome",  "chrome"    },
            { "chrome browser",     "chrome"    }, { "file explorer",  "explorer"  },
            { "files",              "explorer"  }, { "my files",       "explorer"  },
            { "my documents",       "explorer"  }, { "microsoft word", "WINWORD"   },
            { "ms word",            "WINWORD"   }, { "the word",       "WINWORD"   },
            { "word document",      "WINWORD"   }, { "word doc",       "WINWORD"   },
            { "word file",          "WINWORD"   }, { "a word document","WINWORD"   },
            { "new word document",  "WINWORD"   }, { "новый документ", "WINWORD"   },
            { "ворд документ",      "WINWORD"   }, { "ворд документа", "WINWORD"   },
            { "ворд файл",          "WINWORD"   }, { "документ ворд",  "WINWORD"   },
            { "microsoft excel",    "EXCEL"     },
            { "ms excel",           "EXCEL"     }, { "the excel",      "EXCEL"     },
            { "spreadsheet",        "EXCEL"     }, { "spreadsheets",   "EXCEL"     },
            { "ms paint",           "mspaint"   }, { "command prompt", "cmd"       },
            { "windows terminal",   "wt"        }, { "opera browser",  "opera"     },
            { "opera gx",           "opera"     },
        };

        private static readonly Dictionary<string, string> _friendlyNames = new(StringComparer.OrdinalIgnoreCase)
        {
            { "notepad",    "Notepad"           }, { "calc",       "Calculator"    },
            { "msedge",     "Microsoft Edge"    }, { "chrome",     "Google Chrome" },
            { "explorer",   "File Explorer"     }, { "WINWORD",    "Microsoft Word"},
            { "EXCEL",      "Microsoft Excel"   }, { "mspaint",    "Paint"         },
            { "cmd",        "Command Prompt"    }, { "wt",         "Windows Terminal"},
        };

        // ── Prefix stripping ─────────────────────────────────────────────────────

        private static readonly string[] _prefixPhrases =
        [
            "please", "can you", "could you", "would you", "i want to", "i'd like to",
            "i would like to", "kindly", "go ahead and", "just", "now",
            "пожалуйста", "можешь", "можешь ли ты", "хочу",
        ];

        // ── Fuzzy verb map ───────────────────────────────────────────────────────

        private static readonly Dictionary<string, string> _fuzzyVerbMap = new(StringComparer.OrdinalIgnoreCase)
        {
            { "open",     "open"     }, { "launch",   "open"     }, { "start",  "open"   },
            { "close",    "close"    }, { "minimize",  "minimize" },
            { "help",     "help"     }, { "search",    "search"   }, { "find",   "search" },
            { "открой",   "open"     }, { "запусти",   "open"     },
            { "закрой",   "close"    }, { "сверни",    "minimize" },
            { "помощь",   "help"     }, { "найди",     "search"   },
        };

        // ── Constructor & chat history ───────────────────────────────────────────

        private readonly ILogger<LocalCommandProcessor>? _logger;
        private readonly List<Dictionary<string, string>> _chatHistory = new();
        private const int MaxChatHistoryTurns = 20;

        public LocalCommandProcessor(ILogger<LocalCommandProcessor>? logger = null)
        {
            _logger = logger;
        }

        // ── Public sentinel ──────────────────────────────────────────────────────

        public static readonly CommandResult AgenticPlanRequired =
            new(false, "AGENTIC_PLAN_REQUIRED", "AGENTIC_PLAN_REQUIRED");

        // ── Window capture ───────────────────────────────────────────────────────

        public void CaptureTargetWindow()
        {
            var hwnd  = GetForegroundWindow();
            var title = GetWindowTitle(hwnd);
            if (title.Contains("AI Companion", StringComparison.OrdinalIgnoreCase)) return;

            _lastTargetWindow      = hwnd;
            _lastTargetWindowTitle = title;
            _logger?.LogInformation("[CAPTURE] Target window: '{Title}'", title);

            if (IsTextEditorWindow(title))
            {
                _lastTextEditorWindow  = hwnd;
                _lastTextEditorTitle   = title;
                _lastTextEditorProcess = DetectProcessType(title);
                _logger?.LogInformation("[SMART] Text editor: '{Title}' ({Process})", title, _lastTextEditorProcess);
            }
        }

        public IntPtr GetTargetWindow()
        {
            if (_lastTargetWindow != IntPtr.Zero)
            {
                var title = GetWindowTitle(_lastTargetWindow);
                if (!string.IsNullOrEmpty(title) && title != "(none)")
                    return _lastTargetWindow;
            }
            return GetForegroundWindow();
        }

        public string GetTargetWindowTitle()    => !string.IsNullOrEmpty(_lastTargetWindowTitle) ? _lastTargetWindowTitle : "(none)";
        public string GetCurrentActiveWindowInfo()
        {
            var h = GetForegroundWindow();
            return $"Window: '{GetWindowTitle(h)}' (Handle: {h})";
        }

        public string GetActiveTextEditorTitle()
        {
            if (_lastTextEditorWindow != IntPtr.Zero)
            {
                var t = GetWindowTitle(_lastTextEditorWindow);
                if (!string.IsNullOrEmpty(t) && t != "(none)") return t;
            }
            return "(no text editor)";
        }

        public string GetSessionContextSnapshot()
        {
            var snapshot = new
            {
                last_opened_app        = _lastOpenedApp,
                last_typed_text        = _lastTypedText,
                last_navigation_target = _lastNavigationTarget,
                last_text_editor       = _lastTextEditorTitle,
                last_text_editor_app   = _lastTextEditorProcess,
                target_window          = SanitizeWindowTitle(_lastTargetWindowTitle),
            };
            return System.Text.Json.JsonSerializer.Serialize(snapshot);
        }

        public void ClearTextEditorMemory()
        {
            _lastTextEditorWindow  = IntPtr.Zero;
            _lastTextEditorTitle   = "";
            _lastTextEditorProcess = "";
            _logger?.LogInformation("[SMART] Text editor memory cleared");
        }

        // ── IsComplexCommand ─────────────────────────────────────────────────────

        public bool IsComplexCommand(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;

            var lower = input.ToLowerInvariant();
            var words = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);

            // Conversational prefix: simple sub-command → handle locally
            var conversationalPrefixes = new[]
            {
                "can you", "could you", "would you", "please",
                "i want", "i need", "i'd like", "go ahead",
                "можешь", "пожалуйста", "мне нужно", "я хочу",
            };
            foreach (var prefix in conversationalPrefixes)
            {
                if (!lower.StartsWith(prefix)) continue;

                var stripped      = lower[prefix.Length..].TrimStart();
                var strippedWords = stripped.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                bool isSimple     = strippedWords.Length <= 3
                                    && !stripped.Contains(" and ")
                                    && !stripped.Contains(" then ");

                if (isSimple)
                {
                    _logger?.LogInformation("[LOCAL] Prefix '{P}' stripped — simple command", prefix);
                    return false;
                }

                _logger?.LogInformation("[AGENTIC] Conversational prefix: '{P}'", prefix);
                return true;
            }

            // Conjunctions / sequencing
            if (lower.Contains(" and ")  || lower.Contains(" then ") ||
                lower.Contains(" after ") || lower.Contains(", ")    ||
                lower.Contains(" и ")    || lower.Contains(" потом ") || lower.Contains(" затем "))
                return true;

            // Two or more action verbs
            var actionVerbs = new[]
            {
                "open", "type", "write", "save", "close", "bold", "italic",
                "copy", "paste", "search", "delete", "format", "select",
                "create", "screenshot", "refresh",
                "открой", "напиши", "сохрани", "закрой", "найди", "выдели",
                "создай", "скриншот",
            };
            // Note: "new" and "tab" removed — they are adjectives/nouns, not action verbs.
            // "create new document" → verbCount=1 ("create") → local handler, not complex.
            int verbCount = actionVerbs.Count(v =>
                Regex.IsMatch(lower, @"\b" + Regex.Escape(v) + @"\b"));
            if (verbCount >= 2) return true;

            // Short search commands (≤4 words) are handled by the local search route.
            // e.g. "search for Python tutorials" → local; "google latest news" → local
            if (words.Length <= 4 &&
                Regex.IsMatch(lower, @"^(?:найди|поиск|search|find|google|гугли)\s+"))
                return false;

            // Write-content commands are single-intent regardless of word count.
            if (Regex.IsMatch(lower,
                @"^write\s+(?:(?:an?\s+)?essay\s+)?(?:about\s+)?.+\s+(?:in(?:to)?|to)\s+(?:notebook|notepad|note\s*(?:file|book)|блокнот)\s*$",
                RegexOptions.IgnoreCase))
                return false;
            if (Regex.IsMatch(lower,
                @"^write\s+(?:(?:an?\s+)?essay\s+)?(?:about\s+)?.+\s+(?:in(?:to)?|to)\s+(?:word|ворд)\s*$",
                RegexOptions.IgnoreCase))
                return false;
            if (Regex.IsMatch(lower,
                @"^write\s+(?:(?:an?\s+)?essay\s+)?about\s+.+$",
                RegexOptions.IgnoreCase))
                return false;

            // New document/file creation is always a single-intent command regardless of
            // word count — "create a new word document" must reach MatchRoutes locally.
            if (Regex.IsMatch(lower,
                @"^(?:create\s+(?:a\s+)?new|new)\s+" +
                @"(?:word|ворд|excel|эксель|notepad|notebook|note\s*book|блокнот)?\s*" +
                @"(?:document|file|doc|page|sheet|spreadsheet|таблицу|файл|документ|notebook|блокнот|note)?\s*$",
                RegexOptions.IgnoreCase))
                return false;

            // Calculator/math expressions are single-intent regardless of word count
            if (Regex.IsMatch(lower, @"^(?:calculate|compute|eval|посчитай|вычисли|please\s+calculate)\s+.+$"))
                return false;
            if (Regex.IsMatch(lower, @"^\d.+(?:plus|minus|times|divided\s+by)"))
                return false;
            if (Regex.IsMatch(lower, @"^what\s+is\s+\d.+"))
                return false;
            if (Regex.IsMatch(lower, @"^type\s+.+\s+in(?:to)?\s+calculator\s*$"))
                return false;

            // Four or more words → complex
            if (words.Length >= 4) return true;

            return false;
        }

        // ── ProcessCommand ───────────────────────────────────────────────────────

        public CommandResult ProcessCommand(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return new CommandResult(false, "Empty command", "");

            _logger?.LogInformation("Processing command: {Text}", text);

            if (IsComplexCommand(text))
            {
                _logger?.LogInformation("[AGENTIC] Complex command — bypassing local regex");
                return AgenticPlanRequired;
            }

            var normalized = StripPrefixPhrases(text);

            // Try normalized text, then original text, against routing table
            foreach (var candidate in normalized != text ? new[] { normalized, text } : new[] { normalized })
            {
                var result = MatchRoutes(candidate);
                if (result != null) return result;
            }

            // Fuzzy verb matching
            var fuzzy = FuzzyMatchCommand(normalized);
            if (fuzzy != null)
            {
                _logger?.LogInformation("[FUZZY] Matched '{Input}'", text);
                return fuzzy;
            }

            // AI backend fallback
            var ai = FallbackToAIBackend(text);
            if (ai != null)
            {
                _logger?.LogInformation("[AI-FALLBACK] Matched '{Input}'", text);
                return ai;
            }

            return new CommandResult(false, $"Unknown command: {text}", "");
        }

        // ── FallbackToAIBackend ──────────────────────────────────────────────────

        public CommandResult? FallbackToAIBackend(string input)
        {
            if (IsComplexCommand(input))
                return AgenticPlanRequired;

            try
            {
                var sanitized     = SanitizeWindowTitle(GetTargetWindowTitle());
                var windowProcess = !string.IsNullOrEmpty(_lastTextEditorProcess) ? _lastTextEditorProcess : "(unknown)";

                var payload = new { text = input, session_id = _sessionId, window_title = sanitized, window_process = windowProcess };
                var content = new System.Net.Http.StringContent(
                    System.Text.Json.JsonSerializer.Serialize(payload),
                    System.Text.Encoding.UTF8, "application/json");

                _logger?.LogInformation("[AI-FALLBACK] → /api/intent '{Input}' (window: {Win})", input, sanitized);

                var response = _intentHttpClient.PostAsync("http://localhost:8000/api/intent", content).Result;
                if (!response.IsSuccessStatusCode) return null;

                var responseJson = response.Content.ReadAsStringAsync().Result;
                using var doc = System.Text.Json.JsonDocument.Parse(responseJson);
                var root   = doc.RootElement;
                var action  = root.GetProperty("action").GetString() ?? "unknown";
                var target  = root.TryGetProperty("target", out var t) ? t.GetString() : null;
                var @params = root.TryGetProperty("params", out var p) ? p.GetString() : null;
                var query   = target ?? @params;  // prefer target; fall back to params

                _logger?.LogInformation("[AI-FALLBACK] action={A}, target={T}, params={P}", action, target, @params);

                return action switch
                {
                    "open_app"      when target != null => ExecuteOpenApp(target),
                    "new_document"  when target != null => ExecuteNewDocument(target),
                    "new_document"                      => ExecuteNewDocument(),
                    "search_web"    when query  != null => ExecuteSearch(query),
                    "minimize"                          => ExecuteMinimize(),
                    "open_settings"                     => ExecuteHelp(),
                    _ when action != "unknown"          => AgenticPlanRequired,
                    _                                   => null,
                };
            }
            catch (System.Net.Http.HttpRequestException ex)
            {
                _logger?.LogDebug("[AI-FALLBACK] Server unavailable: {E}", ex.Message);
                return null;
            }
            catch (TaskCanceledException)
            {
                _logger?.LogDebug("[AI-FALLBACK] Request timed out");
                return null;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[AI-FALLBACK] Error: {E}", ex.Message);
                return null;
            }
        }

        // ── ChatAsync ────────────────────────────────────────────────────────────

        public async Task<string> ChatAsync(string message)
        {
            try
            {
                var history = _chatHistory
                    .TakeLast(MaxChatHistoryTurns * 2)
                    .Select(h => new { role = h["role"], content = h["content"] })
                    .ToList<object>();

                var payload = new { message, session_id = _sessionId, history };
                var content = new System.Net.Http.StringContent(
                    System.Text.Json.JsonSerializer.Serialize(payload),
                    System.Text.Encoding.UTF8, "application/json");

                using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(28));
                var response = await _agentHttpClient.PostAsync("http://localhost:8000/api/chat", content, cts.Token);
                if (!response.IsSuccessStatusCode) return "";

                var responseJson = await response.Content.ReadAsStringAsync();
                using var doc    = System.Text.Json.JsonDocument.Parse(responseJson);
                var reply        = doc.RootElement.GetProperty("reply").GetString() ?? "";

                _chatHistory.Add(new() { ["role"] = "user",      ["content"] = message });
                _chatHistory.Add(new() { ["role"] = "assistant",  ["content"] = reply   });
                if (_chatHistory.Count > MaxChatHistoryTurns * 2)
                    _chatHistory.RemoveRange(0, _chatHistory.Count - MaxChatHistoryTurns * 2);

                return reply;
            }
            catch (Exception ex)
            {
                _logger?.LogDebug("[CHAT] Error: {E}", ex.Message);
                return "";
            }
        }

        public void ClearChatHistory() => _chatHistory.Clear();

        // ── Static helpers ───────────────────────────────────────────────────────

        internal static string SanitizeWindowTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return "(unknown)";

            if (title.Contains(" - "))
                return title.Split(new[] { " - " }, StringSplitOptions.RemoveEmptyEntries)[^1].Trim();

            if (Regex.IsMatch(title, @"(?:[A-Za-z]:\\|\\\\)"))
                return "Document Editor";

            var lower = title.ToLowerInvariant();
            if (lower.Contains(".docx") || lower.Contains(".doc") || lower.Contains(".txt") ||
                lower.Contains(".pdf")  || lower.Contains(".xlsx") || lower.Contains(".csv"))
                return "Document Editor";

            return title;
        }

        // ── Private routing helpers ──────────────────────────────────────────────

        private CommandResult? MatchRoutes(string text)
        {
            foreach (var (pattern, action) in _localRoutes)
            {
                var m = pattern.Match(text);
                if (!m.Success) continue;

                try
                {
                    return action switch
                    {
                        "open_app"       => ExecuteOpenApp(m.Groups[1].Value.Trim()),
                        "close"          => ExecuteCloseWindow(),
                        "time"           => ExecuteGetTime(),
                        "date"           => ExecuteGetDate(),
                        "minimize"       => ExecuteMinimize(),
                        "help"           => ExecuteHelp(),
                        "search"         => ExecuteSearch(m.Groups[1].Value.Trim()),
                        "greet"          => new CommandResult(true, "Hello!", "Hello! I'm ready to assist you."),
                        "how_do_i"       => ExecuteHowDoI(m.Groups[1].Value.Trim()),
                        "list_commands"  => ExecuteShowCommands(),
                        "tutorial_start" => new CommandResult(true, "Starting tutorial",  "TUTORIAL_START"),
                        "tutorial_stop"  => new CommandResult(true, "Stopping tutorial",  "TUTORIAL_STOP"),
                        "tutorial_skip"  => new CommandResult(true, "Skipping step",      "TUTORIAL_SKIP"),
                        "tutorial_hint"  => new CommandResult(true, "Hint requested",     "TUTORIAL_HINT"),
                        "write_word"      => WriteContentToWordDocument(m.Groups[1].Value.Trim()),
                        "write_notebook"  => WriteContentToNotebook(m.Groups[1].Value.Trim()),
                        "write_ctx"       => WriteContextAware(m.Groups[1].Value.Trim()),
                        "new_doc"         => ExecuteNewDocument(),
                        "new_doc_word"    => ExecuteNewDocument("WINWORD"),
                        "new_doc_notepad" => ExecuteNewDocument("notepad"),
                        "new_doc_excel"   => ExecuteNewDocument("EXCEL"),
                        "calc_app"       => ExecuteCalculation(m.Groups[1].Value.Trim()),
                        "calc_type"      => ExecuteCalcType(m.Groups[1].Value.Trim()),
                        "math_expr"      => EvalMath(m.Groups[1].Value.Trim()),
                        _                => null,
                    };
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "Error executing action '{Action}'", action);
                    return new CommandResult(false, $"Error: {ex.Message}", "");
                }
            }
            return null;
        }

        private CommandResult? FuzzyMatchCommand(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return null;

            var words = input.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (words.Length == 0) return null;

            string? matchedVerb = null;
            int bestScore = 0;
            foreach (var verb in _fuzzyVerbMap.Keys)
            {
                int score = Fuzz.Ratio(words[0].ToLowerInvariant(), verb.ToLowerInvariant());
                if (score > bestScore && score >= 80)
                {
                    bestScore    = score;
                    matchedVerb  = _fuzzyVerbMap[verb];
                }
            }

            if (matchedVerb == null) return null;
            _logger?.LogInformation("[FUZZY] '{W}' → '{V}' (score {S})", words[0], matchedVerb, bestScore);

            return matchedVerb switch
            {
                "open"     when words.Length > 1 => ExecuteOpenApp(string.Join(" ", words.Skip(1))),
                "close"                          => ExecuteCloseWindow(),
                "minimize"                       => ExecuteMinimize(),
                "help"                           => ExecuteHelp(),
                "search"   when words.Length > 1 => ExecuteSearch(string.Join(" ", words.Skip(1))),
                _                                => null,
            };
        }

        private string StripPrefixPhrases(string text)
        {
            var result = text.Trim();
            foreach (var prefix in _prefixPhrases)
            {
                if (result.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    result = result[prefix.Length..].TrimStart();
                    break;
                }
            }
            return result;
        }

        // ── Command executors ────────────────────────────────────────────────────

        private CommandResult ExecuteOpenApp(string appName)
        {
            var cleaned = Regex.Replace(appName.Trim().ToLowerInvariant(), @"^(the|a|an)\s+", "");
            // Also strip Russian articles/possessives
            cleaned = Regex.Replace(cleaned, @"^(мой|моё|мою|новый|новую|новое)\s+", "");

            string? processName = null;
            if (_appMappings.TryGetValue(cleaned, out var mapped))
            {
                processName = mapped;
            }
            else
            {
                // Try single-word lookup — handle "ворд документ" → try first word "ворд"
                var firstWord = cleaned.Split(' ')[0];
                if (firstWord != cleaned && _appMappings.TryGetValue(firstWord, out var firstWordMapped))
                {
                    processName = firstWordMapped;
                    _logger?.LogInformation("[APP] First-word match: '{W}' → '{P}'", firstWord, processName);
                }
            }

            if (processName == null)
            {
                var bestMatch = FuzzyProcess.ExtractOne(cleaned, _appMappings.Keys.ToList());
                if (bestMatch?.Score >= 80)
                {
                    processName = _appMappings[bestMatch.Value];
                    _logger?.LogInformation("[FUZZY-APP] '{In}' → '{M}' (score {S})", appName, bestMatch.Value, bestMatch.Score);
                }
            }

            // Can't resolve app name → let IBM Granite handle it
            if (processName == null)
            {
                _logger?.LogInformation("[APP] Cannot resolve '{App}' — routing to IBM Granite", appName);
                return AgenticPlanRequired;
            }

            var friendly     = _friendlyNames.TryGetValue(processName, out var fn) ? fn : processName;
            var existingHwnd = FindWindowByProcessOrTitle(processName);

            if (existingHwnd != IntPtr.Zero)
            {
                ShowWindow(existingHwnd, 9 /* SW_RESTORE */);
                SetForegroundWindow(existingHwnd);
                _lastOpenedApp  = friendly;
                _currentContext = ResolveContext(processName);
                return new CommandResult(true, $"Switched to {friendly}", $"Switched to the open {friendly} window.");
            }

            try
            {
                System.Diagnostics.Process.Start(new ProcessStartInfo { FileName = processName, UseShellExecute = true });
                _lastOpenedApp  = friendly;
                _currentContext = ResolveContext(processName);
                return new CommandResult(true, $"Opened {friendly}", $"Opening {friendly} for you!");
            }
            catch (Exception)
            {
                // Launch failed — delegate to IBM Granite which can try alternative methods
                _logger?.LogInformation("[APP] Process.Start failed for '{App}' — routing to IBM Granite", processName);
                return AgenticPlanRequired;
            }
        }

        private IntPtr FindWindowByProcessOrTitle(string processName)
        {
            var stem  = System.IO.Path.GetFileNameWithoutExtension(processName).ToLowerInvariant();
            IntPtr found = IntPtr.Zero;

            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                var sb = new StringBuilder(256);
                GetWindowText(h, sb, 256);
                var title = sb.ToString();
                if (string.IsNullOrWhiteSpace(title)) return true;

                bool titleMatch = title.Contains(stem, StringComparison.OrdinalIgnoreCase)
                    || (title.Contains("Блокнот", StringComparison.OrdinalIgnoreCase) && stem == "notepad");

                if (!titleMatch) return true;

                GetWindowThreadProcessId(h, out uint pid);
                try
                {
                    var proc = System.Diagnostics.Process.GetProcessById((int)pid);
                    if (proc.ProcessName.Equals(stem,        StringComparison.OrdinalIgnoreCase) ||
                        proc.ProcessName.Equals(processName, StringComparison.OrdinalIgnoreCase))
                    {
                        found = h;
                        return false;
                    }
                }
                catch { /* process may have exited */ }
                return true;
            }, IntPtr.Zero);

            return found;
        }

        public string LastOpenedApp => _lastOpenedApp;

        private CommandResult ExecuteNewDocument(string? explicitTarget = null)
        {
            string processName;
            string friendlyName;

            if (explicitTarget != null)
            {
                processName  = explicitTarget;
                friendlyName = _friendlyNames.TryGetValue(explicitTarget, out var fn) ? fn : explicitTarget;
            }
            else if (_currentContext == AppContext.Notepad ||
                     _lastOpenedApp.Contains("Notepad", StringComparison.OrdinalIgnoreCase) ||
                     _lastTextEditorProcess.Equals("notepad", StringComparison.OrdinalIgnoreCase))
            { processName = "notepad"; friendlyName = "Notepad"; }
            else if (_lastOpenedApp.Contains("Excel", StringComparison.OrdinalIgnoreCase) ||
                     _lastTextEditorProcess.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
            { processName = "EXCEL"; friendlyName = "Microsoft Excel"; }
            else
            { processName = "WINWORD"; friendlyName = "Microsoft Word"; }

            var hwnd = FindWindowByProcessOrTitle(processName);
            if (hwnd != IntPtr.Zero)
            {
                ShowWindow(hwnd, 9 /* SW_RESTORE */);
                SetForegroundWindow(hwnd);
                System.Threading.Thread.Sleep(300);
                // Use keybd_event for reliable Ctrl+N — SendKeys can lose the Ctrl modifier on
                // Word's legacy _WwG canvas and on high-DPI displays.
                keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
                keybd_event(VK_N,       0, 0, UIntPtr.Zero);
                keybd_event(VK_N,       0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                System.Threading.Thread.Sleep(500);
                _logger?.LogInformation("[CMD] new_doc: Ctrl+N in '{App}'", friendlyName);
                return new CommandResult(true, $"New document in {friendlyName}", $"Creating a new document in {friendlyName}!");
            }

            // Notepad: create a temp file to bypass Windows 11 session restore
            // (which would otherwise reopen the last closed file instead of a blank document)
            if (processName.Equals("notepad", StringComparison.OrdinalIgnoreCase))
            {
                var tempPath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    $"note_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                System.IO.File.WriteAllText(tempPath, "");
                try
                {
                    SysProcess.Start(new ProcessStartInfo
                    {
                        FileName        = "notepad.exe",
                        Arguments       = $"\"{tempPath}\"",
                        UseShellExecute = true
                    });
                    _lastOpenedApp  = "Notepad";
                    _currentContext = AppContext.Notepad;
                    _currentDocPath = tempPath;
                    _logger?.LogInformation("[CMD] new_doc: Notepad + temp file '{Path}'", tempPath);
                    return new CommandResult(true, "New Notepad file", "Opening a new Notepad file for you!");
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "new_doc: Failed to launch Notepad");
                    return new CommandResult(false, $"Could not create file: {ex.Message}", "Sorry, I couldn't create a new file.");
                }
            }

            // Word: create a blank .docx on disk so the file exists immediately, then open it.
            // This mirrors the Notepad approach — a real file is created before the app launches.
            if (processName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    var docxPath = CreateBlankDocxFile();
                    SysProcess.Start(new ProcessStartInfo
                    {
                        FileName        = "WINWORD.EXE",
                        Arguments       = $"\"{docxPath}\"",
                        UseShellExecute = true
                    });
                    _lastOpenedApp  = "Microsoft Word";
                    _currentContext = AppContext.Word;
                    _currentDocPath = docxPath;
                    _logger?.LogInformation("[CMD] new_doc: Word + file '{Path}'", docxPath);
                    return new CommandResult(true, $"New Word document: {System.IO.Path.GetFileName(docxPath)}",
                        "Opening a new Word document for you!");
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "new_doc: Failed to create .docx");
                    return new CommandResult(false, $"Could not create document: {ex.Message}",
                        "Sorry, I couldn't create a new Word document.");
                }
            }

            // Excel and other apps: launch directly — opens with a blank workbook by default
            try
            {
                SysProcess.Start(new ProcessStartInfo { FileName = processName, UseShellExecute = true });
                _lastOpenedApp = friendlyName;
                _logger?.LogInformation("[CMD] new_doc: Launched '{App}'", friendlyName);
                return new CommandResult(true, $"Opening {friendlyName}", $"Opening {friendlyName} with a new document!");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "new_doc: Failed to launch {App}", processName);
                return new CommandResult(false, $"Could not create document: {ex.Message}", "Sorry, I couldn't create a new document.");
            }
        }

        /// <summary>
        /// Creates a minimal but valid blank .docx in the user's Documents folder.
        /// Uses System.IO.Packaging (WindowsBase) — no extra NuGet dependencies.
        /// </summary>
        private static string CreateBlankDocxFile()
        {
            var dir  = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var path = System.IO.Path.Combine(dir, $"doc_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            const string docRelType      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string docContentType  = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
            const string settingsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
            const string settingsCT      = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

            using var pkg = Package.Open(path, FileMode.Create, FileAccess.ReadWrite);

            // ── word/document.xml ────────────────────────────────────────────────
            var docUri  = PackUriHelper.CreatePartUri(new Uri("word/document.xml", UriKind.Relative));
            var docPart = pkg.CreatePart(docUri, docContentType, CompressionOption.Normal);
            using (var sw = new StreamWriter(docPart.GetStream(FileMode.Create, FileAccess.Write)))
                sw.Write(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr></w:p>" +
                    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");

            // ── word/settings.xml (required by Word to suppress repair prompt) ──
            var setUri  = PackUriHelper.CreatePartUri(new Uri("word/settings.xml", UriKind.Relative));
            var setPart = pkg.CreatePart(setUri, settingsCT, CompressionOption.Normal);
            using (var sw = new StreamWriter(setPart.GetStream(FileMode.Create, FileAccess.Write)))
                sw.Write(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");

            // ── word/_rels/document.xml.rels ────────────────────────────────────
            docPart.CreateRelationship(
                PackUriHelper.CreatePartUri(new Uri("settings.xml", UriKind.Relative)),
                TargetMode.Internal, settingsRelType, "rId1");

            // ── package relationship (_rels/.rels) ──────────────────────────────
            pkg.CreateRelationship(docUri, TargetMode.Internal, docRelType, "rId1");

            return path;
        }

        /// <summary>
        /// Calls IBM Granite to generate an essay on <paramref name="topic"/>,
        /// saves the reply to a timestamped note_*.txt in %TEMP%, then opens Notepad.
        /// The HTTP call runs on a background thread so the UI stays responsive.
        /// </summary>
        private CommandResult WriteContentToNotebook(string topic)
        {
            if (string.IsNullOrWhiteSpace(topic))
                return new CommandResult(false, "No topic specified", "Please tell me what to write about.");

            var prompt = $"Write a detailed, well-structured 3-paragraph essay about {topic}. " +
                         "Be informative, engaging, and specific.";

            // Fire-and-forget: generate text then write file and open Notepad
            _ = Task.Run(async () =>
            {
                try
                {
                    _logger?.LogInformation("[WRITE] Generating essay about '{Topic}'", topic);
                    var essay = await ChatAsync(prompt);

                    if (string.IsNullOrWhiteSpace(essay))
                    {
                        _logger?.LogWarning("[WRITE] Empty reply from backend for topic '{Topic}'", topic);
                        return;
                    }

                    var tempPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"note_{DateTime.Now:yyyyMMdd_HHmmss}.txt");

                    await System.IO.File.WriteAllTextAsync(tempPath, essay, System.Text.Encoding.UTF8);
                    _lastOpenedApp = "Notepad";
                    _logger?.LogInformation("[WRITE] Essay saved to '{Path}' ({Bytes} bytes)",
                        tempPath, new System.IO.FileInfo(tempPath).Length);

                    SysProcess.Start(new ProcessStartInfo
                    {
                        FileName        = "notepad.exe",
                        Arguments       = $"\"{tempPath}\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "[WRITE] Failed to write essay about '{Topic}'", topic);
                }
            });

            return new CommandResult(true,
                $"Generating essay about \"{topic}\"...",
                $"I'm writing an essay about {topic}. Notepad will open with the result in a few seconds!");
        }

        // ── Context-aware write dispatch ─────────────────────────────────────────

        /// <summary>
        /// Dispatches "write essay about X" to the correct handler based on the
        /// current active context (Word / Notepad). If no context is set, returns
        /// a clarification prompt instead of silently doing nothing.
        /// </summary>
        private CommandResult WriteContextAware(string topic)
        {
            _logger?.LogInformation("[WRITE-CTX] Context={Ctx}  Topic='{Topic}'", _currentContext, topic);
            return _currentContext switch
            {
                AppContext.Word    => WriteContentToWordDocument(topic),
                AppContext.Notepad => WriteContentToNotebook(topic),
                _ => new CommandResult(false,
                    "No active document context",
                    "Where should I write? Say 'write essay about [topic] in word' or 'write essay about [topic] in notebook'.")
            };
        }

        /// <summary>
        /// Calls IBM Granite to write an essay about <paramref name="topic"/>, creates a new
        /// .docx in ~/Documents with the essay pre-populated, and opens it in Word.
        /// Uses System.IO.Packaging so the file is immediately valid OpenXML.
        /// Runs on a background thread so the UI stays responsive.
        /// </summary>
        private CommandResult WriteContentToWordDocument(string topic)
        {
            if (string.IsNullOrWhiteSpace(topic))
                return new CommandResult(false, "No topic", "Please tell me what to write about.");

            var prompt = $"Write a detailed, well-structured 3-paragraph essay about {topic}. " +
                         "Be informative, engaging, and specific.";

            _ = Task.Run(async () =>
            {
                try
                {
                    _logger?.LogInformation("[WORD-WRITE] Generating essay about '{Topic}'", topic);
                    var essay = await ChatAsync(prompt);
                    if (string.IsNullOrWhiteSpace(essay)) return;

                    // Create a new .docx with the essay content already inside.
                    // This avoids the file-lock issue that would occur if Word already
                    // has the document open (Word holds an exclusive lock on .docx files).
                    var path = CreateEssayDocxFile(topic, essay);
                    _currentDocPath = path;
                    _currentContext = AppContext.Word;
                    _lastOpenedApp  = "Microsoft Word";

                    _logger?.LogInformation("[WORD-WRITE] Essay saved to '{Path}' ({Bytes} bytes)",
                        path, new System.IO.FileInfo(path).Length);

                    SysProcess.Start(new ProcessStartInfo
                    {
                        FileName        = "WINWORD.EXE",
                        Arguments       = $"\"{path}\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "[WORD-WRITE] Failed for topic '{Topic}'", topic);
                }
            });

            return new CommandResult(true,
                $"Writing essay about \"{topic}\" in Word...",
                $"I'm writing an essay about {topic}. Word will open with the full essay in a few seconds!");
        }

        /// <summary>
        /// Creates a new .docx file in ~/Documents pre-populated with <paramref name="essay"/>
        /// as a sequence of paragraphs. Uses System.IO.Packaging — no extra NuGet required.
        /// </summary>
        private static string CreateEssayDocxFile(string topic, string essay)
        {
            var dir  = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var safe = Regex.Replace(topic.Trim().ToLowerInvariant(), @"[^\w\s-]", "");
            safe = Regex.Replace(safe, @"\s+", "_").TrimEnd('_');
            if (safe.Length > 30) safe = safe[..30];
            var path = System.IO.Path.Combine(dir, $"essay_{safe}_{DateTime.Now:yyyyMMdd_HHmmss}.docx");

            // Build <w:p> elements — one per non-empty line
            var paragraphs = new StringBuilder();
            foreach (var line in essay.Split('\n'))
            {
                var text = line.Trim();
                if (string.IsNullOrEmpty(text)) continue;
                paragraphs.Append(
                    "<w:p><w:r><w:rPr/><w:t xml:space=\"preserve\">" +
                    XmlEscape(text) +
                    "</w:t></w:r></w:p>");
            }
            // Ensure at least one paragraph so Word doesn't complain
            if (paragraphs.Length == 0)
                paragraphs.Append("<w:p/>");

            const string docRelType     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string docContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
            const string settingsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
            const string settingsCT     = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

            using var pkg = Package.Open(path, FileMode.Create, FileAccess.ReadWrite);

            var docUri  = PackUriHelper.CreatePartUri(new Uri("word/document.xml", UriKind.Relative));
            var docPart = pkg.CreatePart(docUri, docContentType, CompressionOption.Normal);
            using (var sw = new StreamWriter(docPart.GetStream(FileMode.Create, FileAccess.Write)))
                sw.Write(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body>" + paragraphs +
                    "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                    "</w:body></w:document>");

            var setUri  = PackUriHelper.CreatePartUri(new Uri("word/settings.xml", UriKind.Relative));
            var setPart = pkg.CreatePart(setUri, settingsCT, CompressionOption.Normal);
            using (var sw = new StreamWriter(setPart.GetStream(FileMode.Create, FileAccess.Write)))
                sw.Write(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");

            docPart.CreateRelationship(
                PackUriHelper.CreatePartUri(new Uri("settings.xml", UriKind.Relative)),
                TargetMode.Internal, settingsRelType, "rId1");
            pkg.CreateRelationship(docUri, TargetMode.Internal, docRelType, "rId1");

            return path;
        }

        private static string XmlEscape(string text) =>
            text.Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");

        /// <summary>Maps a process name to the corresponding AppContext.</summary>
        private static AppContext ResolveContext(string processName) =>
            processName.ToUpperInvariant() switch
            {
                "WINWORD" => AppContext.Word,
                "NOTEPAD" => AppContext.Notepad,
                "CALC"    => AppContext.Calculator,
                "MSEDGE"  => AppContext.Browser,
                "CHROME"  => AppContext.Browser,
                _         => AppContext.None
            };

        private static CommandResult ExecuteCloseWindow()      => AgenticPlanRequired;

        private static CommandResult ExecuteGetTime()
        {
            var time = DateTime.Now.ToString("HH:mm");
            return new CommandResult(true, $"Current time: {time}", $"The current time is {time}.");
        }

        private static CommandResult ExecuteGetDate()
        {
            var date = DateTime.Now.ToString("dddd, MMMM d, yyyy");
            return new CommandResult(true, $"Today's date: {date}", $"Today is {date}.");
        }

        private static CommandResult ExecuteMinimize()
        {
            try
            {
                SysProcess.Start(new ProcessStartInfo
                {
                    FileName      = "explorer.exe",
                    Arguments     = "shell:::{3080F90D-D7AD-11D9-BD98-0000947B0257}",
                    UseShellExecute = true,
                });
                return new CommandResult(true, "Windows minimized", "Done!");
            }
            catch
            {
                return new CommandResult(false, "Could not minimize", "I couldn't minimize windows.");
            }
        }

        private static CommandResult ExecuteSearch(string query)
        {
            try
            {
                SysProcess.Start(new ProcessStartInfo
                {
                    FileName        = $"https://www.google.com/search?q={Uri.EscapeDataString(query)}",
                    UseShellExecute = true,
                });
                return new CommandResult(true, $"Searching for: {query}", $"Searching for '{query}' in your browser.");
            }
            catch
            {
                return new CommandResult(false, "Search failed", "I couldn't open the search.");
            }
        }

        private static CommandResult ExecuteHelp()
        {
            const string helpText = @"Available commands:

APPLICATION CONTROL:
  'Open [app]'   - notepad, calculator, browser, Word, Excel
  'Close'        - close current window
  'Minimize'     - minimize all windows

DOCUMENT EDITING (Notepad, Word):
  'Type: [text]' - type text into active application
  'Select all' / 'Copy' / 'Paste' / 'Cut'
  'Undo' / 'Redo'
  'Bold' / 'Italic' / 'Underline'
  'Save' / 'Delete'

WINDOW-AWARE COMMANDS:
  'In Word write [text]'       - type into Word
  'In Notepad write [text]'    - type into Notepad
  'В ворде напиши [текст]'
  'В блокноте напиши [текст]'

UTILITIES:
  'Search [query]'  - search the web
  'What time is it?' / 'What date?'

LEARNING:
  'How do I [task]?'  - help on a specific task
  'What commands'     - list all commands

All commands work in Russian too!";

            return new CommandResult(true, "Help displayed", helpText);
        }

        private static CommandResult ExecuteHowDoI(string task)
        {
            var lower = task.ToLowerInvariant();
            var explanations = new Dictionary<string, string>
            {
                { "open",   "Say 'Open' followed by the app name, e.g. 'Open Notepad'." },
                { "type",   "Say 'Type:' followed by your text. It will appear in the active window." },
                { "copy",   "Select text first ('Select all'), then say 'Copy'." },
                { "paste",  "Position your cursor, then say 'Paste'." },
                { "format", "Select text, then say 'Bold', 'Italic', or 'Underline'." },
                { "bold",   "Select text, then say 'Bold'." },
                { "save",   "Say 'Save' — same as Ctrl+S." },
                { "undo",   "Say 'Undo'. Repeat to undo multiple actions." },
                { "search", "Say 'Search' followed by your query, e.g. 'Search weather today'." },
                { "close",  "Say 'Close' or 'Close this window'." },
            };

            foreach (var kvp in explanations)
                if (lower.Contains(kvp.Key))
                    return new CommandResult(true, $"Help: {task}", kvp.Value);

            return new CommandResult(true, $"Help: {task}",
                "I can help you open apps, type text, format documents, search the web, and more!");
        }

        private static CommandResult ExecuteShowCommands()
        {
            const string list = @"Commands I understand:

APPLICATIONS : open / start / launch / close / minimize
TYPING       : type / write [text]
EDITING      : select all, copy, paste, cut, delete, undo, redo
FORMATTING   : bold, italic, underline, save
NAVIGATION   : go to start, go to end, new line
UTILITIES    : search, time, date, help

Tip: 'How do I copy text?' gives detailed instructions.";

            return new CommandResult(true, "Command list", list);
        }

        // ── Private window helpers ───────────────────────────────────────────────

        private string GetWindowTitle(IntPtr hWnd)
        {
            if (hWnd == IntPtr.Zero) return "(none)";
            var sb = new StringBuilder(256);
            GetWindowText(hWnd, sb, 256);
            return sb.ToString();
        }

        private static bool IsTextEditorWindow(string title)
        {
            var lower = title.ToLowerInvariant();
            return lower.Contains("word")     || lower.Contains(".docx") ||
                   lower.Contains(".doc")     || lower.Contains("notepad") ||
                   lower.Contains("блокнот")  || lower.Contains(".txt") ||
                   lower.Contains("document") || lower.Contains("документ");
        }

        private CommandResult EvalMath(string expr)
        {
            try
            {
                // Normalise natural-language operators
                var e = expr
                    .Replace("times",        "*").Replace("multiplied by", "*")
                    .Replace("divided by",   "/").Replace("divided",       "/")
                    .Replace("plus",         "+").Replace("minus",         "-")
                    .Replace("умножить на",  "*").Replace("разделить на",  "/")
                    .Replace("плюс",         "+").Replace("минус",         "-")
                    .Replace("×",            "*").Replace("÷",             "/");

                var dt     = new System.Data.DataTable();
                var result = dt.Compute(e, "");
                var value  = Convert.ToDouble(result);
                var answer = value == Math.Floor(value) ? ((long)value).ToString() : value.ToString("G10");

                _logger?.LogInformation("[MATH] {Expr} = {Answer}", expr, answer);
                return new CommandResult(true,
                    $"🧮 {expr} = {answer}",
                    $"The answer is {answer}");
            }
            catch
            {
                return new CommandResult(false,
                    $"Could not evaluate: {expr}",
                    "Sorry, I couldn't calculate that. Try a simpler expression.");
            }
        }

        // ── Calculator UI Automation ─────────────────────────────────────────────

        /// <summary>
        /// Opens the Windows Calculator, types <paramref name="expression"/> via
        /// System.Windows.Automation InvokePattern, reads the display result, and
        /// returns it in the command response.
        /// Falls back to local DataTable.Compute if UI Automation fails.
        /// </summary>
        private CommandResult ExecuteCalculation(string expression)
        {
            var normalized = NormalizeCalcExpression(expression);
            _logger?.LogInformation("[CALC] Expression='{Expr}'  Normalized='{Norm}'",
                expression, normalized);

            if (string.IsNullOrWhiteSpace(normalized))
                return EvalMath(expression);

            string? calcResult = null;
            Exception? uiaEx   = null;

            // System.Windows.Automation uses COM — must run on STA thread.
            var sta = new System.Threading.Thread(() =>
            {
                try   { calcResult = RunCalcAutomation(normalized); }
                catch (Exception ex) { uiaEx = ex; }
            });
            sta.SetApartmentState(System.Threading.ApartmentState.STA);
            sta.IsBackground = true;
            sta.Start();
            sta.Join(TimeSpan.FromSeconds(12));

            if (uiaEx != null)
                _logger?.LogWarning(uiaEx, "[CALC] UIA failed — falling back to local eval");

            if (string.IsNullOrWhiteSpace(calcResult))
                return EvalMath(expression);

            _logger?.LogInformation("[CALC] Result='{R}'", calcResult);
            return new CommandResult(true,
                $"Calculator: {expression} = {calcResult}",
                $"The answer is {calcResult}.");
        }

        /// <summary>
        /// Opens Calculator and types <paramref name="expression"/> without reading the result.
        /// Useful for setting up complex expressions the user wants to inspect.
        /// </summary>
        private CommandResult ExecuteCalcType(string expression)
        {
            var normalized = NormalizeCalcExpression(expression);
            if (string.IsNullOrWhiteSpace(normalized))
                return new CommandResult(false, "Nothing to type",
                    "Please specify what to type in Calculator.");

            var sta = new System.Threading.Thread(() =>
            {
                try
                {
                    if (SysProcess.GetProcessesByName("CalculatorApp").Length == 0)
                    {
                        SysProcess.Start(new ProcessStartInfo
                        { FileName = "ms-calculator:", UseShellExecute = true });
                        System.Threading.Thread.Sleep(2500);
                    }
                    var calc = FindCalcWindow(5000);
                    if (calc == null) return;
                    System.Threading.Thread.Sleep(400);
                    ClickCalcButton(calc, "clearButton");
                    System.Threading.Thread.Sleep(100);
                    TypeInCalc(calc, normalized);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, "[CALC-TYPE] Failed for '{Expr}'", expression);
                }
            });
            sta.SetApartmentState(System.Threading.ApartmentState.STA);
            sta.IsBackground = true;
            sta.Start();

            return new CommandResult(true,
                $"Typing in Calculator: {expression}",
                $"I've typed {expression} into Calculator.");
        }

        /// <summary>
        /// Full Calculator automation sequence (must run on STA thread).
        /// Launches Calculator if needed, clears, types the expression, hits =, reads display.
        /// </summary>
        private static string RunCalcAutomation(string normalized)
        {
            if (SysProcess.GetProcessesByName("CalculatorApp").Length == 0)
            {
                SysProcess.Start(new ProcessStartInfo
                { FileName = "ms-calculator:", UseShellExecute = true });
                System.Threading.Thread.Sleep(2500);
            }
            else
            {
                System.Threading.Thread.Sleep(300);
            }

            var calc = FindCalcWindow(5000)
                ?? throw new InvalidOperationException("Calculator window not found via UIA");

            System.Threading.Thread.Sleep(500);   // allow buttons to fully render

            ClickCalcButton(calc, "clearButton");
            System.Threading.Thread.Sleep(120);

            TypeInCalc(calc, normalized);

            ClickCalcButton(calc, "equalButton");
            System.Threading.Thread.Sleep(400);

            return ReadCalcDisplay(calc);
        }

        /// <summary>Types each character of <paramref name="expr"/> into Calculator via InvokePattern.</summary>
        private static void TypeInCalc(AutomationElement calc, string expr)
        {
            var map = new Dictionary<char, string>
            {
                ['0'] = "num0Button",  ['1'] = "num1Button",  ['2'] = "num2Button",
                ['3'] = "num3Button",  ['4'] = "num4Button",  ['5'] = "num5Button",
                ['6'] = "num6Button",  ['7'] = "num7Button",  ['8'] = "num8Button",
                ['9'] = "num9Button",
                ['+'] = "plusButton",  ['-'] = "minusButton",
                ['*'] = "multiplyButton",  ['/'] = "divideButton",
                ['.'] = "decimalSeparatorButton",
            };
            foreach (var ch in expr)
            {
                if (map.TryGetValue(ch, out var id))
                {
                    ClickCalcButton(calc, id);
                    System.Threading.Thread.Sleep(80);
                }
            }
        }

        /// <summary>
        /// Converts natural-language math to a compact symbolic string.
        /// "123 plus 456" → "123+456",  "999 times 7" → "999*7"
        /// </summary>
        private static string NormalizeCalcExpression(string expr)
        {
            if (string.IsNullOrWhiteSpace(expr)) return "";

            var e = expr.Trim()
                .Replace("multiplied by", "*", StringComparison.OrdinalIgnoreCase)
                .Replace("times",         "*", StringComparison.OrdinalIgnoreCase)
                .Replace("divided by",    "/", StringComparison.OrdinalIgnoreCase)
                .Replace("divided",       "/", StringComparison.OrdinalIgnoreCase)
                .Replace("plus",          "+", StringComparison.OrdinalIgnoreCase)
                .Replace("minus",         "-", StringComparison.OrdinalIgnoreCase)
                .Replace("умножить на",   "*", StringComparison.OrdinalIgnoreCase)
                .Replace("разделить на",  "/", StringComparison.OrdinalIgnoreCase)
                .Replace("плюс",          "+", StringComparison.OrdinalIgnoreCase)
                .Replace("минус",         "-", StringComparison.OrdinalIgnoreCase)
                .Replace("×", "*")
                .Replace("÷", "/");

            // Keep only digits, operators, and decimal point
            var sb = new StringBuilder();
            foreach (var ch in e)
            {
                if (char.IsDigit(ch) || ch == '+' || ch == '-' ||
                    ch == '*' || ch == '/' || ch == '.')
                    sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Polls UIA for a top-level window named "Calculator" up to <paramref name="timeoutMs"/> ms.
        /// Returns the first matching AutomationElement, or null on timeout.
        /// </summary>
        private static AutomationElement? FindCalcWindow(int timeoutMs)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            var cond     = new PropertyCondition(AutomationElement.NameProperty, "Calculator");

            while (DateTime.UtcNow < deadline)
            {
                var found = AutomationElement.RootElement.FindAll(TreeScope.Children, cond);
                if (found.Count > 0)
                    return found[0];
                System.Threading.Thread.Sleep(250);
            }
            return null;
        }

        /// <summary>Clicks a Calculator button by AutomationId via InvokePattern.</summary>
        private static void ClickCalcButton(AutomationElement root, string autoId)
        {
            var cond = new PropertyCondition(AutomationElement.AutomationIdProperty, autoId);
            var btn  = root.FindFirst(TreeScope.Descendants, cond)
                       ?? throw new InvalidOperationException($"Calc button '{autoId}' not found");

            if (btn.GetCurrentPattern(InvokePattern.Pattern) is InvokePattern ip)
                ip.Invoke();
        }

        /// <summary>
        /// Reads the Calculator display element (AutomationId="CalculatorResults").
        /// The accessible name is "Display is 579" — strips prefix and comma separators.
        /// </summary>
        private static string ReadCalcDisplay(AutomationElement root)
        {
            var cond = new PropertyCondition(
                AutomationElement.AutomationIdProperty, "CalculatorResults");
            var display = root.FindFirst(TreeScope.Descendants, cond)
                          ?? throw new InvalidOperationException("CalculatorResults not found");

            var name = display.Current.Name;
            if (name.StartsWith("Display is ", StringComparison.OrdinalIgnoreCase))
                name = name["Display is ".Length..];

            return name.Trim().Replace(",", "");
        }

        private static string DetectProcessType(string title)
        {
            var lower = title.ToLowerInvariant();
            if (lower.Contains("word") || lower.Contains(".docx") || lower.Contains(".doc")) return "WINWORD";
            if (lower.Contains("notepad") || lower.Contains("блокнот") || lower.Contains(".txt")) return "notepad";
            return "unknown";
        }
    }

    // ── CommandResult ────────────────────────────────────────────────────────────

    public class CommandResult
    {
        public bool   Success        { get; }
        public string Description    { get; }
        public string SpeechResponse { get; }

        public CommandResult(bool success, string description, string speechResponse)
        {
            Success        = success;
            Description    = description;
            SpeechResponse = speechResponse;
        }
    }
}
