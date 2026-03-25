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
using AICompanion.Desktop.Services.Security;
using AICompanion.Desktop.Services;
using AICompanion.Desktop.Services.Dictation;
using AICompanion.Desktop.ViewModels;
using AICompanion.Desktop.Views;
using Serilog;

namespace AICompanion.Desktop
{
    /*
        Application entry point and dependency injection configuration.
        
        This class sets up the service container with all required dependencies
        for the AI Companion application. The dependency injection pattern allows
        for easier testing and loose coupling between components.
        
        Services are registered as singletons where appropriate (voice recognition,
        AI client) to maintain state across the application lifetime.
        
        Reference: https://docs.microsoft.com/en-us/dotnet/core/extensions/dependency-injection
    */
    public partial class App : System.Windows.Application
    {
        /*
            The service provider manages dependency injection throughout the application.
            Components request their dependencies through constructor injection.
        */
        public static ServiceProvider? ServiceProvider { get; private set; }

        /*
            Application startup handler that configures services and logging.
        */
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            /*
                Configure Serilog for structured logging to files.
                Logs are stored in the user's local application data folder
                and rotated daily to prevent excessive disk usage.
            */
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

            /*
                Configure the dependency injection container.
                All services and view models are registered here.
            */
            var services = new ServiceCollection();
            ConfigureServices(services);
            ServiceProvider = services.BuildServiceProvider();

            Log.Information("Services configured successfully");

            // Initialize database before any views load
            try
            {
                var dbService = ServiceProvider.GetRequiredService<DatabaseService>();
                dbService.InitializeAsync().GetAwaiter().GetResult();
                Log.Information("Database initialized successfully");
            }
            catch (Exception ex)
            {
                Log.Fatal(ex, "Failed to initialize local database");
                System.Windows.MessageBox.Show("Fatal Error: Could not initialize application database.\n\n" + ex.Message, 
                                "AI Companion Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                Shutdown(-1);
                return;
            }

            // Apply saved theme
            ApplyThemeFromConfig();

            // Check if login is required at startup
            if (ShouldShowLoginAtStartup())
            {
                var loginWindow = new LoginWindow();
                if (loginWindow.ShowDialog() != true)
                {
                    // User cancelled login or failed to authenticate
                    Log.Information("User cancelled login, shutting down");
                    Shutdown();
                    return;
                }
                Log.Information("User authenticated: {User}", loginWindow.AuthenticatedUser);
            }

            // Manually create and show MainWindow (StartupUri removed to avoid
            // shutdown when LoginWindow closes before MainWindow opens)
            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            mainWindow.Closed += (s, args) => Shutdown();
            mainWindow.Show();
        }

        private bool ShouldShowLoginAtStartup()
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
                if (!File.Exists(configPath)) return false;

                var json = File.ReadAllText(configPath);
                using var doc = JsonDocument.Parse(json);

                if (doc.RootElement.TryGetProperty("Security", out var security))
                {
                    if (security.TryGetProperty("RequireLogin", out var requireLogin))
                    {
                        return requireLogin.GetBoolean();
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Error reading login settings, defaulting to no login");
            }

            return false;
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

        /// <summary>
        /// Apply a theme ("Light", "Dark", or "System") to the application.
        /// Updates all color resources in the application-level resource dictionary.
        /// </summary>
        public static void ApplyTheme(string theme)
        {
            var resources = Current.Resources;

            if (theme == "System")
            {
                // Detect Windows system theme
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

            // Update brushes to reference the new colors
            resources["PrimaryBrush"] = new SolidColorBrush((WpfColor)resources["PrimaryColor"]);
            resources["SecondaryBrush"] = new SolidColorBrush((WpfColor)resources["SecondaryColor"]);
            resources["BackgroundBrush"] = new SolidColorBrush((WpfColor)resources["BackgroundColor"]);
            resources["SurfaceBrush"] = new SolidColorBrush((WpfColor)resources["SurfaceColor"]);
            resources["TextPrimaryBrush"] = new SolidColorBrush((WpfColor)resources["TextPrimaryColor"]);
            resources["TextSecondaryBrush"] = new SolidColorBrush((WpfColor)resources["TextSecondaryColor"]);
            resources["SuccessBrush"] = new SolidColorBrush((WpfColor)resources["SuccessColor"]);
            resources["ErrorBrush"] = new SolidColorBrush((WpfColor)resources["ErrorColor"]);

            // Update panel background brushes
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

            // Update gradient brush
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

        /*
            Registers all application services with the DI container.
            
            Services are organized by their responsibility:
            - Logging infrastructure
            - Voice input/output services
            - Screen capture and analysis
            - UI automation for computer control
            - AI engine communication
            - View models for MVVM pattern
        */
        private void ConfigureServices(IServiceCollection services)
        {
            /*
                Add logging services using Serilog as the provider.
            */
            services.AddLogging(builder =>
            {
                builder.AddSerilog(dispose: true);
            });

            /*
                Load configuration from appsettings.json
            */
            var configPath = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            var configuration = new ConfigurationBuilder()
                .AddJsonFile(configPath, optional: true)
                .Build();
            services.AddSingleton<IConfiguration>(configuration);

            /*
                Database service for context, history, and security data.
            */
            services.AddSingleton<DatabaseService>();

            /*
                Security service for authentication and authorization.
            */
            services.AddSingleton<SecurityService>();

            /*
                Voice services - ElevenLabs only (Windows and Watson removed).
            */
            services.AddSingleton<ElevenLabsService>();
            services.AddSingleton<ElevenLabsSpeechService>();

            /*
                Unified Voice Manager - ElevenLabs only mode.
            */
            services.AddSingleton<UnifiedVoiceManager>(sp =>
            {
                var logger = sp.GetService<ILogger<UnifiedVoiceManager>>();
                var elevenLabsTTS = sp.GetRequiredService<ElevenLabsService>();
                var elevenLabsSTT = sp.GetRequiredService<ElevenLabsSpeechService>();

                var manager = new UnifiedVoiceManager(logger, elevenLabsTTS, elevenLabsSTT);

                // Configure from settings
                var config = sp.GetService<IConfiguration>();
                if (config != null)
                {
                    var settings = new VoiceSettings
                    {
                        ElevenLabsApiKey = config["ElevenLabs:ApiKey"] ?? "",
                        ElevenLabsVoiceId = config["ElevenLabs:VoiceId"] ?? "21m00Tcm4TlvDq8ikWAM"
                    };

                    Log.Information("Voice config: ElevenLabs API key present: {HasKey}",
                        !string.IsNullOrEmpty(settings.ElevenLabsApiKey));

                    manager.Configure(settings);
                }

                return manager;
            });

            /*
                Dictation service for typing into Word/Notepad.
            */
            services.AddSingleton<DictationService>();

            /*
                Screen services capture and analyze the display.
            */
            services.AddSingleton<ScreenCaptureService>();

            /*
                Automation service controls Windows applications.
            */
            services.AddSingleton<UIAutomationService>();

            /*
                Window automation helper for reliable Win32 focus and text input.
            */
            services.AddSingleton<WindowAutomationHelper>();

            /*
                Agentic execution service for multi-step plan execution from /api/plan.
            */
            services.AddSingleton<AgenticExecutionService>(sp =>
            {
                var automation = sp.GetRequiredService<WindowAutomationHelper>();
                var logger = sp.GetService<ILogger<AgenticExecutionService>>();
                return new AgenticExecutionService(automation, logger);
            });

            /*
                AI engine client communicates with the Python backend.
            */
            services.AddSingleton<AIEngineClient>();

            /*
                Local command processor for voice command recognition and execution.
                Registered as Singleton to maintain window state across commands.
            */
            services.AddSingleton<LocalCommandProcessor>();

            /*
                View models implement the MVVM pattern for data binding.
            */
            services.AddTransient<MainViewModel>();
        }

        /*
            Application shutdown handler for cleanup.
        */
        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("AI Companion shutting down");
            
            /*
                Dispose all services to release resources.
            */
            ServiceProvider?.Dispose();
            
            /*
                Flush and close the log file.
            */
            Log.CloseAndFlush();
            
            base.OnExit(e);
        }
    }
}
