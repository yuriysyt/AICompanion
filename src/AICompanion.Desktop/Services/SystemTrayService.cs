using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Extensions.Logging;

namespace AICompanion.Desktop.Services
{
    /*
        SystemTrayService manages the application's presence in the Windows
        notification area (system tray).
        
        When the user minimizes the application, it can optionally hide to
        the system tray instead of appearing in the taskbar. This keeps the
        companion accessible without cluttering the taskbar. A context menu
        provides quick access to common actions.
        
        The tray icon also displays status information through icon changes
        and balloon notifications for important events.
    */
    public class SystemTrayService : IDisposable
    {
        private readonly ILogger<SystemTrayService> _logger;
        private NotifyIcon? _notifyIcon;
        private ContextMenuStrip? _contextMenu;
        private bool _isDisposed;

        /*
            Event raised when the user clicks "Show" in the context menu
            or double-clicks the tray icon.
        */
        public event EventHandler? ShowWindowRequested;

        /*
            Event raised when the user clicks "Exit" in the context menu.
        */
        public event EventHandler? ExitRequested;

        /*
            Event raised when the user clicks "Start Listening" in the menu.
        */
        public event EventHandler? StartListeningRequested;

        public SystemTrayService(ILogger<SystemTrayService> logger)
        {
            _logger = logger;
        }

        /*
            Initializes the system tray icon and context menu.
        */
        public void Initialize()
        {
            _contextMenu = CreateContextMenu();

            _notifyIcon = new NotifyIcon
            {
                Text = "AI Companion",
                Visible = false,
                ContextMenuStrip = _contextMenu
            };

            /*
                Load the application icon from embedded resources.
                Falls back to a default icon if not found.
            */
            try
            {
                _notifyIcon.Icon = SystemIcons.Application;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Could not load tray icon, using default");
                _notifyIcon.Icon = SystemIcons.Application;
            }

            _notifyIcon.DoubleClick += OnTrayIconDoubleClick;

            _logger.LogInformation("System tray service initialized");
        }

        /*
            Creates the right-click context menu for the tray icon.
        */
        private ContextMenuStrip CreateContextMenu()
        {
            var menu = new ContextMenuStrip();

            var showItem = new ToolStripMenuItem("Show AI Companion");
            showItem.Click += (s, e) => ShowWindowRequested?.Invoke(this, EventArgs.Empty);
            showItem.Font = new Font(showItem.Font, FontStyle.Bold);
            menu.Items.Add(showItem);

            menu.Items.Add(new ToolStripSeparator());

            var listenItem = new ToolStripMenuItem("Start Listening");
            listenItem.Click += (s, e) => StartListeningRequested?.Invoke(this, EventArgs.Empty);
            menu.Items.Add(listenItem);

            menu.Items.Add(new ToolStripSeparator());

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ExitRequested?.Invoke(this, EventArgs.Empty);
            menu.Items.Add(exitItem);

            return menu;
        }

        /*
            Shows the tray icon when the window is minimized.
        */
        public void Show()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = true;
            }
        }

        /*
            Hides the tray icon when the window is restored.
        */
        public void Hide()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
            }
        }

        /*
            Displays a balloon notification in the system tray.
        */
        public void ShowNotification(string title, string message, ToolTipIcon icon = ToolTipIcon.Info)
        {
            if (_notifyIcon != null && _notifyIcon.Visible)
            {
                _notifyIcon.ShowBalloonTip(3000, title, message, icon);
            }
        }

        /*
            Updates the tray icon tooltip text.
        */
        public void UpdateStatus(string status)
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Text = $"AI Companion - {status}";
            }
        }

        private void OnTrayIconDoubleClick(object? sender, EventArgs e)
        {
            ShowWindowRequested?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            if (_isDisposed)
                return;

            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
            }

            _contextMenu?.Dispose();
            _isDisposed = true;

            _logger.LogInformation("System tray service disposed");
        }
    }
}
