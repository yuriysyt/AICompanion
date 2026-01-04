using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Input;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services.Accessibility
{
    /*
        KeyboardShortcutService provides global hotkey support for accessibility.
        
        Users who cannot easily reach the microphone button or speak the wake
        word can use keyboard shortcuts to activate the assistant. This service
        registers system-wide hotkeys that work even when the application
        window is not focused.
        
        Default shortcuts:
        - Ctrl+Shift+A: Activate listening mode
        - Ctrl+Shift+S: Stop listening / cancel current operation
        - Ctrl+Shift+H: Show/hide the companion window
        - Escape: Cancel current operation
        
        Reference: https://docs.microsoft.com/en-us/windows/win32/inputdev/hot-keys
    */
    public class KeyboardShortcutService : IDisposable
    {
        private readonly ILogger<KeyboardShortcutService> _logger;
        private readonly Dictionary<int, Action> _registeredHotkeys;
        private IntPtr _windowHandle;
        private bool _isDisposed;
        private int _nextHotkeyId = 1;

        /*
            Windows API constants for modifier keys.
        */
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;
        private const uint MOD_NOREPEAT = 0x4000;

        /*
            Event raised when the activate shortcut is pressed.
        */
        public event EventHandler? ActivateRequested;

        /*
            Event raised when the stop shortcut is pressed.
        */
        public event EventHandler? StopRequested;

        /*
            Event raised when the toggle visibility shortcut is pressed.
        */
        public event EventHandler? ToggleVisibilityRequested;

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        public KeyboardShortcutService(ILogger<KeyboardShortcutService> logger)
        {
            _logger = logger;
            _registeredHotkeys = new Dictionary<int, Action>();
        }

        /*
            Initializes the service and registers default hotkeys.
            Must be called after the main window is created.
        */
        public void Initialize(IntPtr windowHandle)
        {
            _windowHandle = windowHandle;

            /*
                Register Ctrl+Shift+A for activation.
            */
            RegisterHotkey(
                ModifierKeys.Control | ModifierKeys.Shift,
                Key.A,
                () => ActivateRequested?.Invoke(this, EventArgs.Empty));

            /*
                Register Ctrl+Shift+S for stop.
            */
            RegisterHotkey(
                ModifierKeys.Control | ModifierKeys.Shift,
                Key.S,
                () => StopRequested?.Invoke(this, EventArgs.Empty));

            /*
                Register Ctrl+Shift+H for toggle visibility.
            */
            RegisterHotkey(
                ModifierKeys.Control | ModifierKeys.Shift,
                Key.H,
                () => ToggleVisibilityRequested?.Invoke(this, EventArgs.Empty));

            _logger.LogInformation("Keyboard shortcuts registered successfully");
        }

        /*
            Registers a global hotkey with the specified modifiers and key.
        */
        public bool RegisterHotkey(ModifierKeys modifiers, Key key, Action callback)
        {
            var id = _nextHotkeyId++;
            var fsModifiers = ConvertModifiers(modifiers);
            var virtualKey = (uint)KeyInterop.VirtualKeyFromKey(key);

            if (RegisterHotKey(_windowHandle, id, fsModifiers | MOD_NOREPEAT, virtualKey))
            {
                _registeredHotkeys[id] = callback;
                _logger.LogDebug("Registered hotkey {Id}: {Modifiers}+{Key}", id, modifiers, key);
                return true;
            }
            else
            {
                _logger.LogWarning("Failed to register hotkey: {Modifiers}+{Key}", modifiers, key);
                return false;
            }
        }

        /*
            Processes a WM_HOTKEY message from the window procedure.
        */
        public void ProcessHotkey(int hotkeyId)
        {
            if (_registeredHotkeys.TryGetValue(hotkeyId, out var callback))
            {
                _logger.LogDebug("Hotkey {Id} triggered", hotkeyId);
                callback.Invoke();
            }
        }

        /*
            Converts WPF ModifierKeys to Windows API modifier flags.
        */
        private uint ConvertModifiers(ModifierKeys modifiers)
        {
            uint result = 0;

            if ((modifiers & ModifierKeys.Alt) != 0)
                result |= MOD_ALT;
            if ((modifiers & ModifierKeys.Control) != 0)
                result |= MOD_CONTROL;
            if ((modifiers & ModifierKeys.Shift) != 0)
                result |= MOD_SHIFT;
            if ((modifiers & ModifierKeys.Windows) != 0)
                result |= MOD_WIN;

            return result;
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;

            foreach (var id in _registeredHotkeys.Keys)
            {
                UnregisterHotKey(_windowHandle, id);
            }

            _registeredHotkeys.Clear();
            _isDisposed = true;

            _logger.LogInformation("Keyboard shortcuts unregistered");
        }
    }
}
