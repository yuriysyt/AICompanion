using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using FluentAssertions;
using Xunit;
using Xunit.Abstractions;

// ═══════════════════════════════════════════════════════════════════════════
//  RealUserWorkflowTests  –  Tестирование как настоящий пользователь
//
//  RU-1  Запуск → авторизация → главное окно
//  RU-2  Открытие Notepad через кнопку → Notepad появляется
//  RU-3  AI генерирует текст → текст печатается в Notepad → содержимое верифицируется
//  RU-4  Скриншот каждого ключевого шага для наглядности
//
//  Поведение теста:
//    • Патчит output appsettings.json → RequireLogin: false
//    • Если всё равно появляется Login-окно → авто-логин через UI как пользователь
//    • После теста восстанавливает настройки
//    • Делает скриншоты в C:\Users\yyurc\Desktop\TestScreenshots\
// ═══════════════════════════════════════════════════════════════════════════

namespace AICompanion.IntegrationTests.RealAppTests
{
    [Collection("LiveModel")]
    public class RealUserWorkflowTests : IAsyncLifetime
    {
        // ── Win32 ─────────────────────────────────────────────────────────────
        [DllImport("user32.dll")] static extern bool  IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")] static extern int   GetWindowText(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern bool  EnumWindows(EnumWindowsProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern uint  GetWindowThreadProcessId(IntPtr h, out uint pid);
        [DllImport("user32.dll")] static extern bool  SetForegroundWindow(IntPtr h);
        [DllImport("user32.dll")] static extern bool  ShowWindow(IntPtr h, int n);
        [DllImport("user32.dll")] static extern bool  GetWindowRect(IntPtr h, out RECT r);
        [DllImport("user32.dll")] static extern bool  SetCursorPos(int x, int y);
        [DllImport("user32.dll")] static extern void  mouse_event(uint f, int x, int y, uint d, UIntPtr e);
        [DllImport("user32.dll")] static extern void  keybd_event(byte vk, byte scan, uint flags, UIntPtr e);
        [DllImport("user32.dll")] static extern bool  EnumChildWindows(IntPtr p, EnumChildProc cb, IntPtr lp);
        [DllImport("user32.dll")] static extern int   GetClassName(IntPtr h, StringBuilder sb, int n);
        [DllImport("user32.dll")] static extern uint  SendMessage(IntPtr h, uint msg, IntPtr w, IntPtr l);

        [StructLayout(LayoutKind.Sequential)] struct RECT { public int L, T, R, B; }
        delegate bool EnumWindowsProc(IntPtr h, IntPtr lp);
        delegate bool EnumChildProc(IntPtr h, IntPtr lp);

        const uint MOUSEEVENTF_LEFT   = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint KEYUP   = 0x0002;
        const byte VK_TAB  = 0x09;
        const byte VK_RETURN  = 0x0D;
        const byte VK_CONTROL = 0x11;
        const byte VK_V    = 0x56;
        const byte VK_A    = 0x41;
        const byte VK_C    = 0x43;
        const byte VK_ALT  = 0x12;
        const byte VK_F4   = 0x73;
        const byte VK_N    = 0x4E;  // for Don't Save (N)

        // ── State ─────────────────────────────────────────────────────────────
        private readonly ITestOutputHelper _out;
        private readonly HttpClient        _http = BackendFixture.Http;
        private Process? _app;
        private IntPtr   _mainHwnd;
        private string   _exePath     = "";
        private string   _outputSettingsPath = "";
        private string?  _savedOutputSettings;
        private string   _screenshotDir = @"C:\Users\yyurc\Desktop\TestScreenshots";

        public RealUserWorkflowTests(ITestOutputHelper output)
        {
            _out = output;
        }

        // ── Lifecycle ─────────────────────────────────────────────────────────

        public async Task InitializeAsync()
        {
            Directory.CreateDirectory(_screenshotDir);

            var solutionRoot = FindSolutionRoot();
            var exeDir = Path.Combine(solutionRoot, "src", "AICompanion.Desktop",
                                      "bin", "Debug", "net8.0-windows");
            _outputSettingsPath = Path.Combine(exeDir, "appsettings.json");

            // Patch output appsettings: RequireLogin false
            if (File.Exists(_outputSettingsPath))
            {
                _savedOutputSettings = File.ReadAllText(_outputSettingsPath);
                File.WriteAllText(_outputSettingsPath, DisableAuth(_savedOutputSettings));
                _out.WriteLine("[INIT] Patched output appsettings: RequireLogin=false");
            }

            // Build
            _out.WriteLine("[INIT] Building project…");
            var projPath = Path.Combine(solutionRoot, "src", "AICompanion.Desktop",
                                        "AICompanion.Desktop.csproj");
            var build = Process.Start(new ProcessStartInfo
            {
                FileName = "dotnet",
                Arguments = $"build \"{projPath}\" -c Debug --no-restore -v quiet",
                UseShellExecute = false,
                RedirectStandardOutput = true, RedirectStandardError = true,
                CreateNoWindow = true
            })!;
            await build.WaitForExitAsync();
            if (build.ExitCode != 0)
                throw new Exception($"Build failed (exit {build.ExitCode})");
            _out.WriteLine("[INIT] Build OK ✓");

            // Re-patch if build overwrote
            if (File.Exists(_outputSettingsPath))
            {
                var after = File.ReadAllText(_outputSettingsPath);
                if (after.Contains("\"RequireLogin\": true"))
                {
                    File.WriteAllText(_outputSettingsPath, DisableAuth(after));
                    _out.WriteLine("[INIT] Re-patched output appsettings after build ✓");
                }
            }

            // Launch
            _exePath = Path.Combine(exeDir, "AICompanion.exe");
            if (!File.Exists(_exePath)) _exePath = Path.Combine(exeDir, "AICompanion.Desktop.exe");
            _out.WriteLine($"[INIT] Launching: {_exePath}");
            _app = Process.Start(new ProcessStartInfo
            {
                FileName = _exePath, UseShellExecute = false, CreateNoWindow = false
            })!;

            // Wait for window and handle login if needed
            _mainHwnd = await WaitForMainWindowAsync((uint)_app.Id);
            _out.WriteLine($"[INIT] Main window: 0x{_mainHwnd:X}");
        }

        public Task DisposeAsync()
        {
            if (_savedOutputSettings != null && File.Exists(_outputSettingsPath))
            {
                File.WriteAllText(_outputSettingsPath, _savedOutputSettings);
                _out.WriteLine("[CLEANUP] Restored output appsettings ✓");
            }
            try { _app?.Kill(entireProcessTree: true); } catch { }
            _app?.Dispose();
            return Task.CompletedTask;
        }

        // ════════════════════════════════════════════════════════════════════
        //  RU-1  Приложение запускается и показывает главное окно
        // ════════════════════════════════════════════════════════════════════
        [Fact]
        public void RU1_App_LaunchesAndShowsMainWindow()
        {
            Banner("RU-1: Приложение запускается — главное окно видно");

            _mainHwnd.Should().NotBe(IntPtr.Zero, "главное окно должно появиться после запуска");

            var title = GetTitle(_mainHwnd);
            _out.WriteLine($"  Заголовок окна: \"{title}\"");

            GetWindowRect(_mainHwnd, out var r);
            var w = r.R - r.L;
            var h = r.B - r.T;
            _out.WriteLine($"  Размер: {w} × {h}");

            Screenshot("RU1_MainWindow");

            title.Should().ContainEquivalentOf("AI Companion",
                because: "заголовок должен идентифицировать приложение");
            w.Should().BeGreaterThan(200, because: "окно должно иметь разумную ширину");

            _out.WriteLine("✅ RU-1 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  RU-2  Кнопка Notepad → Notepad открывается
        // ════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task RU2_QuickButton_Notepad_Opens()
        {
            Banner("RU-2: Кнопка Notepad → Notepad открывается");

            _mainHwnd.Should().NotBe(IntPtr.Zero);
            Screenshot("RU2_Before");

            var existing = SnapshotWindows("Notepad");
            existing.UnionWith(SnapshotWindows("Блокнот"));
            _out.WriteLine($"  Открытых Notepad до клика: {existing.Count}");

            ClickQuickNotepad(_mainHwnd);
            _out.WriteLine("  Нажата кнопка Notepad");

            var notepad = await WaitForNewWindowAsync("Notepad", existing, 12000);
            if (notepad == IntPtr.Zero)
                notepad = await WaitForNewWindowAsync("Блокнот", existing, 3000);

            Screenshot("RU2_NotepadOpened");

            notepad.Should().NotBe(IntPtr.Zero,
                "клик по кнопке Notepad должен открыть новое окно Notepad");
            _out.WriteLine($"  Notepad hwnd: 0x{notepad:X}, заголовок: \"{GetTitle(notepad)}\"");

            CloseWindowSafe(notepad);
            DismissSaveDialog();

            _out.WriteLine("✅ RU-2 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  RU-3  AI генерирует текст → вставляет в Notepad → проверяем содержимое
        // ════════════════════════════════════════════════════════════════════
        [Fact]
        public async Task RU3_AI_GeneratesText_AndTypesIntoNotepad()
        {
            Banner("RU-3: AI генерирует текст → Notepad → содержимое верифицировано");

            _mainHwnd.Should().NotBe(IntPtr.Zero);

            // Шаг 1: Открыть Notepad
            var existing = SnapshotWindows("Notepad");
            existing.UnionWith(SnapshotWindows("Блокнот"));
            ClickQuickNotepad(_mainHwnd);
            _out.WriteLine("  Нажата кнопка Notepad…");

            var notepad = await WaitForNewWindowAsync("Notepad", existing, 12000);
            if (notepad == IntPtr.Zero)
                notepad = await WaitForNewWindowAsync("Блокнот", existing, 3000);

            notepad.Should().NotBe(IntPtr.Zero, "Notepad должен открыться");
            _out.WriteLine($"  Notepad: 0x{notepad:X}");
            Screenshot("RU3_NotepadOpen");

            try
            {
                // Шаг 2: Спросить AI (Granite) написать текст
                const string session = "ru3_workflow";
                _out.WriteLine("  Запрашиваем у Granite текст для Notepad…");

                var resp = await _http.PostAsJsonAsync(
                    $"{BackendFixture.BackendUrl}/api/chat",
                    new ChatWithHistoryRequest
                    {
                        Message   = "Write exactly 3 short paragraphs about the benefits of using AI assistants on Windows. Plain text only, no headers, no bullet points.",
                        SessionId = session,
                    });
                resp.EnsureSuccessStatusCode();
                var chat = (await resp.Content.ReadFromJsonAsync<ChatResponse>())!;

                _out.WriteLine($"  AI ответил за {chat.LatencyMs} мс ({chat.Reply.Length} символов):");
                foreach (var line in chat.Reply.Split('\n'))
                    _out.WriteLine($"    {line}");

                chat.Reply.Should().NotBeNullOrWhiteSpace("AI должен вернуть непустой текст");
                chat.Reply.Length.Should().BeGreaterThan(60, "3 абзаца должны быть длиннее 60 символов");

                // Шаг 3: Вставить текст в Notepad
                _out.WriteLine("  Вставляем текст в Notepad…");
                TypeViaClipboard(notepad, chat.Reply);
                Thread.Sleep(500);
                Screenshot("RU3_TextTyped");

                // Шаг 4: Прочитать содержимое обратно
                _out.WriteLine("  Читаем содержимое Notepad обратно…");
                var content = ReadWindowContent(notepad);
                _out.WriteLine($"  Прочитано символов: {content.Length}");

                content.Should().NotBeNullOrWhiteSpace(
                    "Notepad должен содержать вставленный текст");

                var expected40 = chat.Reply.Substring(0, Math.Min(40, chat.Reply.Length)).Trim();
                content.Should().Contain(expected40,
                    because: "AI-текст должен быть точно вставлен в Notepad");

                _out.WriteLine("  Содержимое совпадает ✓");

                // Шаг 5: Проверяем что AI помнит сессию (контекст)
                _out.WriteLine("  Проверяем контекстную память AI…");
                var r2 = await _http.PostAsJsonAsync(
                    $"{BackendFixture.BackendUrl}/api/chat",
                    new ChatWithHistoryRequest
                    {
                        Message   = "What topic did I just ask you to write about?",
                        SessionId = session,
                    });
                r2.EnsureSuccessStatusCode();
                var chat2 = (await r2.Content.ReadFromJsonAsync<ChatResponse>())!;
                _out.WriteLine($"  AI помнит: \"{chat2.Reply}\"");
                chat2.Reply.Should().ContainEquivalentOf("AI",
                    because: "модель должна помнить тему предыдущего запроса");

                Screenshot("RU3_Verified");

                // Очистка сессии
                await _http.PostAsJsonAsync($"{BackendFixture.BackendUrl}/api/chat/clear",
                    new ClearHistoryRequest { SessionId = session });
            }
            finally
            {
                CloseWindowSafe(notepad);
                DismissSaveDialog();
            }

            _out.WriteLine("✅ RU-3 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  RU-4  Полный скриншот-отчёт состояния приложения
        // ════════════════════════════════════════════════════════════════════
        [Fact]
        public void RU4_Screenshot_FullAppState()
        {
            Banner("RU-4: Скриншот полного состояния приложения");

            _mainHwnd.Should().NotBe(IntPtr.Zero);
            ShowWindow(_mainHwnd, 9 /*SW_RESTORE*/);
            SetForegroundWindow(_mainHwnd);
            Thread.Sleep(500);

            var path = Screenshot("RU4_FullState");
            _out.WriteLine($"  Скриншот сохранён: {path}");

            GetWindowRect(_mainHwnd, out var r);
            _out.WriteLine($"  Позиция приложения: ({r.L},{r.T}) размер {r.R-r.L}×{r.B-r.T}");

            // Перечислить все дочерние элементы
            var children = new List<string>();
            EnumChildWindows(_mainHwnd, (h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                var cls = new StringBuilder(64); GetClassName(h, cls, 64);
                var t = sb.ToString();
                if (!string.IsNullOrWhiteSpace(t))
                    children.Add($"{cls}: '{t}'");
                return true;
            }, IntPtr.Zero);

            _out.WriteLine($"  Дочерних элементов с текстом: {children.Count}");
            foreach (var c in children)
                _out.WriteLine($"    {c}");

            _out.WriteLine("✅ RU-4 PASSED");
        }

        // ════════════════════════════════════════════════════════════════════
        //  Helpers
        // ════════════════════════════════════════════════════════════════════

        private void Banner(string t)
        {
            _out.WriteLine(""); _out.WriteLine(new string('═', 66));
            _out.WriteLine($"  {t}"); _out.WriteLine(new string('═', 66));
        }

        private static string DisableAuth(string json) =>
            json.Replace("\"RequireLogin\": true",        "\"RequireLogin\": false")
                .Replace("\"RequireSecurityCode\": true", "\"RequireSecurityCode\": false");

        private static string FindSolutionRoot()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                if (Directory.GetFiles(dir.FullName, "*.sln").Length > 0) return dir.FullName;
                dir = dir.Parent;
            }
            throw new DirectoryNotFoundException("Solution root (.sln) not found.");
        }

        private async Task<IntPtr> WaitForMainWindowAsync(uint appPid)
        {
            var deadline = DateTime.UtcNow.AddSeconds(30);
            IntPtr loginHwnd = IntPtr.Zero;
            IntPtr mainHwnd  = IntPtr.Zero;

            while (DateTime.UtcNow < deadline && mainHwnd == IntPtr.Zero)
            {
                await Task.Delay(400);
                loginHwnd = IntPtr.Zero;
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h)) return true;
                    GetWindowThreadProcessId(h, out var wPid);
                    if (wPid != appPid) return true;
                    var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                    var title = sb.ToString();
                    if (string.IsNullOrWhiteSpace(title)) return true;
                    _out.WriteLine($"  [window] '{title}'");
                    if (title.Contains("Login", StringComparison.OrdinalIgnoreCase))
                        loginHwnd = h;
                    else
                        mainHwnd = h;
                    return true;
                }, IntPtr.Zero);

                if (loginHwnd != IntPtr.Zero && mainHwnd == IntPtr.Zero)
                {
                    _out.WriteLine("  Login-окно обнаружено → авто-вход…");
                    Screenshot("AutoLogin_Before");
                    await AutoLoginAsync(loginHwnd);
                    await Task.Delay(2000);
                }
            }

            return mainHwnd != IntPtr.Zero ? mainHwnd : loginHwnd;
        }

        private async Task AutoLoginAsync(IntPtr loginHwnd)
        {
            SetForegroundWindow(loginHwnd);
            await Task.Delay(600);

            // Заполнить поля логина/пароля
            var fields = new List<IntPtr>();
            EnumChildWindows(loginHwnd, (h, _) =>
            {
                var cls = new StringBuilder(64); GetClassName(h, cls, 64);
                var c = cls.ToString();
                if (c.Contains("Edit", StringComparison.OrdinalIgnoreCase) ||
                    c.Contains("TextBox", StringComparison.OrdinalIgnoreCase))
                    fields.Add(h);
                return true;
            }, IntPtr.Zero);

            _out.WriteLine($"    Полей ввода: {fields.Count}");

            if (fields.Count >= 2)
            {
                FocusAndPaste(fields[0], AppExeFixture.TestUser);
                await Task.Delay(200);
                FocusAndPaste(fields[1], AppExeFixture.TestPassword);
                await Task.Delay(200);
            }
            else
            {
                // Через Tab
                SetForegroundWindow(loginHwnd); await Task.Delay(200);
                PasteText(AppExeFixture.TestUser);
                keybd_event(VK_TAB, 0, 0, UIntPtr.Zero); Thread.Sleep(50);
                keybd_event(VK_TAB, 0, KEYUP, UIntPtr.Zero);
                await Task.Delay(200);
                PasteText(AppExeFixture.TestPassword);
                await Task.Delay(200);
            }

            // Нажать Sign In
            IntPtr btn = IntPtr.Zero;
            EnumChildWindows(loginHwnd, (h, _) =>
            {
                var cls = new StringBuilder(64); GetClassName(h, cls, 64);
                if (!cls.ToString().Equals("Button", StringComparison.OrdinalIgnoreCase)) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                var t = sb.ToString();
                if (t.Contains("Sign", StringComparison.OrdinalIgnoreCase) ||
                    t.Contains("Войти", StringComparison.OrdinalIgnoreCase))
                { btn = h; return false; }
                return true;
            }, IntPtr.Zero);

            if (btn != IntPtr.Zero)
            {
                GetWindowRect(btn, out var br);
                Click((br.L + br.R) / 2, (br.T + br.B) / 2);
            }
            else
            {
                keybd_event(VK_RETURN, 0, 0, UIntPtr.Zero);
                keybd_event(VK_RETURN, 0, KEYUP, UIntPtr.Zero);
            }
            await Task.Delay(400);
        }

        private void ClickQuickNotepad(IntPtr appHwnd)
        {
            SetForegroundWindow(appHwnd); Thread.Sleep(400);

            // 1. UI Automation — reliable for WPF buttons
            try
            {
                var root = System.Windows.Automation.AutomationElement.FromHandle(appHwnd);
                var allBtns = root.FindAll(
                    System.Windows.Automation.TreeScope.Descendants,
                    new System.Windows.Automation.PropertyCondition(
                        System.Windows.Automation.AutomationElement.ControlTypeProperty,
                        System.Windows.Automation.ControlType.Button));

                foreach (System.Windows.Automation.AutomationElement btn in allBtns)
                {
                    var name = btn.Current.Name ?? "";
                    if (name.Contains("Notepad", StringComparison.OrdinalIgnoreCase))
                    {
                        var rect = btn.Current.BoundingRectangle;
                        int bx = (int)(rect.Left + rect.Width / 2);
                        int by = (int)(rect.Top  + rect.Height / 2);
                        _out.WriteLine($"  [UIA] Кнопка Notepad: ({bx},{by}) name='{name}'");
                        Click(bx, by);
                        return;
                    }
                }
                _out.WriteLine("  [UIA] Кнопка Notepad не найдена — переключаемся на InvokePattern");

                // 2. Try Invoke without clicking (works even if button not visible)
                foreach (System.Windows.Automation.AutomationElement btn in allBtns)
                {
                    var name = btn.Current.Name ?? "";
                    if (name.Contains("Notepad", StringComparison.OrdinalIgnoreCase))
                    {
                        if (btn.TryGetCurrentPattern(
                                System.Windows.Automation.InvokePattern.Pattern,
                                out var pattern))
                        {
                            ((System.Windows.Automation.InvokePattern)pattern).Invoke();
                            _out.WriteLine("  [UIA] InvokePattern.Invoke() выполнен");
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _out.WriteLine($"  [UIA] Ошибка: {ex.Message}");
            }

            // 3. Scroll down and retry UIA (button might be outside scroll view)
            _out.WriteLine("  [UIA] Прокрутка вниз и повторная попытка…");
            SetForegroundWindow(appHwnd);
            GetWindowRect(appHwnd, out var wr);
            // Scroll to bottom in the window
            int scrollX = (wr.L + wr.R) / 2;
            int scrollY = wr.T + (wr.B - wr.T) * 3 / 4;
            SetCursorPos(scrollX, scrollY);
            Thread.Sleep(100);
            // Simulate scroll down with keybd PageDown
            keybd_event(0x22 /*VK_NEXT/PageDown*/, 0, 0, UIntPtr.Zero);
            keybd_event(0x22, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(300);

            // 4. Final heuristic: 88% down the window height = Quick Commands row
            int hx = (wr.L + wr.R) / 2 - 60; // left of center where Notepad btn is
            int hy = wr.T + (int)((wr.B - wr.T) * 0.88);
            _out.WriteLine($"  [Heuristic] Notepad: ({hx},{hy})");
            Click(hx, hy);
        }

        private void TypeViaClipboard(IntPtr hwnd, string text)
        {
            SetForegroundWindow(hwnd); Thread.Sleep(300);
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2;
            int cy = r.T + (int)((r.B - r.T) * 0.65);
            Click(cx, cy); Thread.Sleep(200);
            PasteText(text);
            Thread.Sleep(400);
        }

        private string ReadWindowContent(IntPtr hwnd)
        {
            SetForegroundWindow(hwnd); Thread.Sleep(200);
            GetWindowRect(hwnd, out var r);
            int cx = (r.L + r.R) / 2;
            int cy = r.T + (int)((r.B - r.T) * 0.65);
            Click(cx, cy); Thread.Sleep(200);

            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_A, 0, 0, UIntPtr.Zero); Thread.Sleep(50);
            keybd_event(VK_A, 0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(200);

            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_C, 0, 0, UIntPtr.Zero); Thread.Sleep(50);
            keybd_event(VK_C, 0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero); Thread.Sleep(600);

            string text = "";
            var t = new Thread(() => { try { text = System.Windows.Forms.Clipboard.GetText(); } catch { } });
            t.SetApartmentState(ApartmentState.STA); t.Start(); t.Join(2000);
            return text;
        }

        private void CloseWindowSafe(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return;
            SetForegroundWindow(hwnd); Thread.Sleep(200);
            keybd_event(VK_ALT, 0, 0, UIntPtr.Zero);
            keybd_event(VK_F4, 0, 0, UIntPtr.Zero); Thread.Sleep(30);
            keybd_event(VK_F4, 0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_ALT, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(400);
        }

        private void DismissSaveDialog()
        {
            // Win11 Notepad "Save changes?" → press N (Don't Save)
            Thread.Sleep(400);
            IntPtr dlg = IntPtr.Zero;
            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                if (sb.ToString().Contains("Notepad", StringComparison.OrdinalIgnoreCase) ||
                    sb.ToString().Contains("Save", StringComparison.OrdinalIgnoreCase))
                { dlg = h; return false; }
                return true;
            }, IntPtr.Zero);
            if (dlg != IntPtr.Zero)
            {
                SetForegroundWindow(dlg); Thread.Sleep(200);
                // Click "Don't Save" button
                IntPtr dontSave = IntPtr.Zero;
                EnumChildWindows(dlg, (h, _) =>
                {
                    var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                    var t = sb.ToString();
                    if (t.Contains("Don't", StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("Не сохранять", StringComparison.OrdinalIgnoreCase) ||
                        t.Contains("Discard", StringComparison.OrdinalIgnoreCase))
                    { dontSave = h; return false; }
                    return true;
                }, IntPtr.Zero);

                if (dontSave != IntPtr.Zero)
                {
                    GetWindowRect(dontSave, out var br);
                    Click((br.L + br.R) / 2, (br.T + br.B) / 2);
                }
                else
                {
                    keybd_event(VK_N, 0, 0, UIntPtr.Zero);
                    keybd_event(VK_N, 0, KEYUP, UIntPtr.Zero);
                }
                Thread.Sleep(300);
            }
        }

        private HashSet<IntPtr> SnapshotWindows(string fragment)
        {
            var set = new HashSet<IntPtr>();
            EnumWindows((h, _) =>
            {
                if (!IsWindowVisible(h)) return true;
                var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                if (sb.ToString().Contains(fragment, StringComparison.OrdinalIgnoreCase))
                    set.Add(h);
                return true;
            }, IntPtr.Zero);
            return set;
        }

        private async Task<IntPtr> WaitForNewWindowAsync(string fragment, HashSet<IntPtr> existing, int timeoutMs)
        {
            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
            IntPtr found = IntPtr.Zero;
            while (DateTime.UtcNow < deadline && found == IntPtr.Zero)
            {
                await Task.Delay(400);
                EnumWindows((h, _) =>
                {
                    if (!IsWindowVisible(h) || existing.Contains(h)) return true;
                    var sb = new StringBuilder(256); GetWindowText(h, sb, 256);
                    if (sb.ToString().Contains(fragment, StringComparison.OrdinalIgnoreCase))
                    { found = h; return false; }
                    return true;
                }, IntPtr.Zero);
            }
            return found;
        }

        private string GetTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, 256);
            return sb.ToString();
        }

        private void FocusAndPaste(IntPtr hwnd, string text)
        {
            SetForegroundWindow(hwnd); Thread.Sleep(100);
            SendMessage(hwnd, 0x0007 /*WM_SETFOCUS*/, IntPtr.Zero, IntPtr.Zero);
            Thread.Sleep(80);
            PasteText(text);
        }

        private void PasteText(string text)
        {
            var sta = new Thread(() => System.Windows.Forms.Clipboard.SetText(text));
            sta.SetApartmentState(ApartmentState.STA); sta.Start(); sta.Join();
            Thread.Sleep(80);
            keybd_event(VK_CONTROL, 0, 0, UIntPtr.Zero);
            keybd_event(VK_V, 0, 0, UIntPtr.Zero); Thread.Sleep(40);
            keybd_event(VK_V, 0, KEYUP, UIntPtr.Zero);
            keybd_event(VK_CONTROL, 0, KEYUP, UIntPtr.Zero);
            Thread.Sleep(100);
        }

        private void Click(int x, int y)
        {
            SetCursorPos(x, y); Thread.Sleep(60);
            mouse_event(MOUSEEVENTF_LEFT,   x, y, 0, UIntPtr.Zero); Thread.Sleep(40);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, UIntPtr.Zero); Thread.Sleep(120);
        }

        private string Screenshot(string name)
        {
            try
            {
                var path = Path.Combine(_screenshotDir, $"{name}_{DateTime.Now:HHmmss}.png");
                System.Drawing.Rectangle bounds;
                using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    bounds = System.Drawing.Rectangle.Round(g.VisibleClipBounds);
                }
                // Use screen size
                bounds = new System.Drawing.Rectangle(0, 0,
                    System.Windows.Forms.Screen.PrimaryScreen!.Bounds.Width,
                    System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height);

                using var bmp = new System.Drawing.Bitmap(bounds.Width, bounds.Height);
                using var gfx = System.Drawing.Graphics.FromImage(bmp);
                gfx.CopyFromScreen(System.Drawing.Point.Empty, System.Drawing.Point.Empty, bounds.Size);
                bmp.Save(path);
                _out.WriteLine($"  📸 {path}");
                return path;
            }
            catch (Exception ex)
            {
                _out.WriteLine($"  [Screenshot failed: {ex.Message}]");
                return "";
            }
        }
    }
}
