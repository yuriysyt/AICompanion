using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using WpfColor = System.Windows.Media.Color;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Configuration;
using AICompanion.Desktop.Services.Voice;
using AICompanion.Desktop.Services.Screen;
using AICompanion.Desktop.Services.Automation;
using AICompanion.Desktop.Services.Communication;
using AICompanion.Desktop.Services.Database;
using AICompanion.Desktop.Services;
using AICompanion.Desktop.Services.Dictation;
using AICompanion.Desktop.Services.Security;
using AICompanion.Desktop.ViewModels;
using AICompanion.Desktop.Views;
using Serilog;

namespace AICompanion.Desktop
{
    public partial class App : System.Windows.Application
    {
        public static ServiceProvider? ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    path: System.IO.Path.Combine(
                        System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                        "AICompanion", "Logs", "app-.log"),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7)
                .CreateLogger();

            Log.Information("AI Companion starting up");

            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();

            Log.Information("Services configured successfully");

            try
            {
                var dbService = ServiceProvider.GetRequiredService<DatabaseService>();
                Task.Run(() => dbService.InitializeAsync()).GetAwaiter().GetResult();
                Log.Information("Database initialized successfully");
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "Failed to initialize local database");
                System.Windows.MessageBox.Show(
                    "Fatal Error: Could not initialize application database.\n\n" + ex.Message,
                    "AI Companion Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                Shutdown(-1);
                return;
            }

            ApplyThemeFromConfig();

            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Closed += (s, args) => Shutdown();
            mainWindow.Show();
        }

        private void ApplyThemeFromConfig()
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                if (!File.Exists(configPath)) return;

                var json = File.ReadAllText(configPath);
                using var doc = JsonDocument.Parse(json);

                if (doc.RootElement.TryGetProperty("Appearance", out var appearance) &&
                    appearance.TryGetProperty("Theme", out var themeProp))
                {
                    var theme = themeProp.GetString() ?? "Light";
                    ApplyTheme(theme);
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Error reading theme settings");
            }
        }

        public static void ApplyTheme(string theme)
        {
            var resources = Current.Resources;

            if (theme == "System")
            {
                try
                {
                    using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                    var useLightTheme = key?.GetValue("AppsUseLightTheme");
                    theme = (useLightTheme is int v && v == 0) ? "Dark" : "Light";
                }
                catch
                {
                    theme = "Light";
                }
            }

            if (theme == "Dark")
            {
                resources["PrimaryColor"] = ColorFromHex("#14B8A6");
                resources["SecondaryColor"] = ColorFromHex("#2DD4BF");
                resources["BackgroundColor"] = ColorFromHex("#1A1A2E");
                resources["SurfaceColor"] = ColorFromHex("#242442");
                resources["TextPrimaryColor"] = ColorFromHex("#E8E8F0");
                resources["TextSecondaryColor"] = ColorFromHex("#A0A0B8");
                resources["SuccessColor"] = ColorFromHex("#22C55E");
                resources["ErrorColor"] = ColorFromHex("#E5484D");
                resources["WarningColor"] = ColorFromHex("#F59E0B");
            }
            else
            {
                resources["PrimaryColor"] = ColorFromHex("#0F766E");
                resources["SecondaryColor"] = ColorFromHex("#14B8A6");
                resources["BackgroundColor"] = ColorFromHex("#F3F7F6");
                resources["SurfaceColor"] = ColorFromHex("#FFFFFF");
                resources["TextPrimaryColor"] = ColorFromHex("#102A26");
                resources["TextSecondaryColor"] = ColorFromHex("#4B5B58");
                resources["SuccessColor"] = ColorFromHex("#22C55E");
                resources["ErrorColor"] = ColorFromHex("#E5484D");
                resources["WarningColor"] = ColorFromHex("#F59E0B");
            }

            resources["PrimaryBrush"] = new SolidColorBrush((WpfColor)resources["PrimaryColor"]);
            resources["SecondaryBrush"] = new SolidColorBrush((WpfColor)resources["SecondaryColor"]);
            resources["BackgroundBrush"] = new SolidColorBrush((WpfColor)resources["BackgroundColor"]);
            resources["SurfaceBrush"] = new SolidColorBrush((WpfColor)resources["SurfaceColor"]);
            resources["TextPrimaryBrush"] = new SolidColorBrush((WpfColor)resources["TextPrimaryColor"]);
            resources["TextSecondaryBrush"] = new SolidColorBrush((WpfColor)resources["TextSecondaryColor"]);
            resources["SuccessBrush"] = new SolidColorBrush((WpfColor)resources["SuccessColor"]);
            resources["ErrorBrush"] = new SolidColorBrush((WpfColor)resources["ErrorColor"]);

            if (theme == "Dark")
            {
                resources["AvatarPanelBrush"] = new SolidColorBrush(ColorFromHex("#1E2A3A"));
                resources["CardPanelBrush"] = new SolidColorBrush(ColorFromHex("#2A2A48"));
                resources["InputPanelBrush"] = new SolidColorBrush(ColorFromHex("#262644"));
                resources["LogPanelBrush"] = new SolidColorBrush(ColorFromHex("#262644"));
                resources["LogHeaderBrush"] = new SolidColorBrush(ColorFromHex("#2A2A48"));
                resources["StatusPanelBrush"] = new SolidColorBrush(ColorFromHex("#2A2A48"));
                resources["StatusBorderBrush"] = new SolidColorBrush(ColorFromHex("#3A3A58"));
                resources["InputBorderBrush"] = new SolidColorBrush(ColorFromHex("#3A3A58"));
                resources["LogBorderBrush"] = new SolidColorBrush(ColorFromHex("#3A3A58"));
                resources["CardBorderBrush"] = new SolidColorBrush(ColorFromHex("#3A3A58"));
            }
            else
            {
                resources["AvatarPanelBrush"] = new SolidColorBrush(ColorFromHex("#E7F5F3"));
                resources["CardPanelBrush"] = new SolidColorBrush(ColorFromHex("#ECF8F6"));
                resources["InputPanelBrush"] = new SolidColorBrush(ColorFromHex("#F5FBFA"));
                resources["LogPanelBrush"] = new SolidColorBrush(ColorFromHex("#F7FBFA"));
                resources["LogHeaderBrush"] = new SolidColorBrush(ColorFromHex("#ECF8F6"));
                resources["StatusPanelBrush"] = new SolidColorBrush(ColorFromHex("#EEF2F1"));
                resources["StatusBorderBrush"] = new SolidColorBrush(ColorFromHex("#D1E0DC"));
                resources["InputBorderBrush"] = new SolidColorBrush(ColorFromHex("#D7E6E3"));
                resources["LogBorderBrush"] = new SolidColorBrush(ColorFromHex("#DDEAE7"));
                resources["CardBorderBrush"] = new SolidColorBrush(ColorFromHex("#D1E8E2"));
            }

            var gradient = new LinearGradientBrush
            {
                StartPoint = new System.Windows.Point(0, 0),
                EndPoint = new System.Windows.Point(1, 1)
            };
            gradient.GradientStops.Add(new GradientStop((WpfColor)resources["PrimaryColor"], 0));
            gradient.GradientStops.Add(new GradientStop((WpfColor)resources["SecondaryColor"], 1));
            resources["PrimaryGradientBrush"] = gradient;

            Log.Information("Theme applied: {Theme}", theme);
        }

        private static WpfColor ColorFromHex(string hex)
        {
            return (WpfColor)System.Windows.Media.ColorConverter.ConvertFromString(hex);
        }

        private void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(builder =>
            {
                builder.AddSerilog(dispose: true);
            });

            var configPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            var configuration = new ConfigurationBuilder()
                .AddJsonFile(configPath, optional: true)
                .Build();
            services.AddSingleton<IConfiguration>(configuration);

            services.AddSingleton<DatabaseService>();

            // SecureApiKeyManager: DPAPI-encrypted key storage (%LOCALAPPDATA%\AICompanion\keys.dat)
            services.AddSingleton<SecureApiKeyManager>();
            services.AddSingleton<SecurityService>();

            services.AddSingleton<ElevenLabsService>();
            services.AddSingleton<ElevenLabsSpeechService>();

            services.AddSingleton<UnifiedVoiceManager>(sp =>
            {
                var logger = sp.GetService<ILogger<UnifiedVoiceManager>>();
                var elevenLabsTTS = sp.GetRequiredService<ElevenLabsService>();
                var elevenLabsSTT = sp.GetRequiredService<ElevenLabsSpeechService>();
                var keyManager = sp.GetRequiredService<SecureApiKeyManager>();

                var manager = new UnifiedVoiceManager(logger, elevenLabsTTS, elevenLabsSTT);

                // 1. Load from DPAPI (secure, encrypted)
                var apiKey = keyManager.LoadApiKey(SecureApiKeyManager.ElevenLabsKeyName) ?? "";

                // 2. If not in DPAPI yet, fall back to appsettings.json and migrate
                if (string.IsNullOrEmpty(apiKey))
                {
                    var config = sp.GetService<IConfiguration>();
                    apiKey = config?["ElevenLabs:ApiKey"] ?? "";
                    if (!string.IsNullOrEmpty(apiKey))
                    {
                        keyManager.SaveApiKey(SecureApiKeyManager.ElevenLabsKeyName, apiKey);
                        Log.Information("[Voice] API key migrated from appsettings.json → DPAPI secure storage");
                    }
                }

                var voiceId = sp.GetService<IConfiguration>()?["ElevenLabs:VoiceId"] ?? "JBFqnCBsd6RMkjVDRZzb";

                Log.Information("[Voice] ElevenLabs API key present: {HasKey}", !string.IsNullOrEmpty(apiKey));

                if (!string.IsNullOrEmpty(apiKey))
                    manager.Configure(new VoiceSettings { ElevenLabsApiKey = apiKey, ElevenLabsVoiceId = voiceId });

                return manager;
            });

            services.AddSingleton<DictationService>();

            services.AddSingleton<ScreenCaptureService>();

            services.AddSingleton<UIAutomationService>();

            services.AddSingleton<WindowAutomationHelper>();

            services.AddSingleton<AgenticExecutionService>(sp =>
            {
                var automation = sp.GetRequiredService<WindowAutomationHelper>();
                var logger = sp.GetService<ILogger<AgenticExecutionService>>();
                return new AgenticExecutionService(automation, logger);
            });

            services.AddSingleton<AIEngineClient>();

            services.AddSingleton<LocalCommandProcessor>();

            services.AddTransient<MainViewModel>();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("AI Companion shutting down");
            ServiceProvider?.Dispose();
            Log.CloseAndFlush();
            base.OnExit(e);
        }
    }
}
