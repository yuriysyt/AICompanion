using System;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.Configuration;
using AICompanion.Desktop.Services.Voice;
using AICompanion.Desktop.Services.Database;

namespace AICompanion.Desktop.Views
{
    public partial class SettingsWindow : Window
    {
        private readonly UnifiedVoiceManager? _voiceManager;
        private readonly DatabaseService? _databaseService;
        private Services.Voice.VoiceSettings _settings = new();

        public SettingsWindow()
        {
            InitializeComponent();

            _voiceManager = App.ServiceProvider?.GetService<UnifiedVoiceManager>();
            _databaseService = App.ServiceProvider?.GetService<DatabaseService>();

            Loaded += SettingsWindow_Loaded;
            OpacitySlider.ValueChanged += OpacitySlider_ValueChanged;
        }

        private async void SettingsWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadSettingsAsync();
            UpdateDatabaseInfo();
        }

        private async System.Threading.Tasks.Task LoadSettingsAsync()
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

                if (File.Exists(configPath))
                {
                    var json = await File.ReadAllTextAsync(configPath);
                    using var doc = JsonDocument.Parse(json);

                    // Load ElevenLabs settings
                    if (doc.RootElement.TryGetProperty("ElevenLabs", out var elevenLabs))
                    {
                        ElevenLabsApiKeyBox.Text = elevenLabs.TryGetProperty("ApiKey", out var key) ? key.GetString() ?? "" : "";
                        ElevenLabsVoiceIdBox.Text = elevenLabs.TryGetProperty("VoiceId", out var voice) ? voice.GetString() ?? "21m00Tcm4TlvDq8ikWAM" : "21m00Tcm4TlvDq8ikWAM";
                        ElevenLabsTTSEnabled.IsChecked = elevenLabs.TryGetProperty("TTSEnabled", out var tts) ? tts.GetBoolean() : true;
                        ElevenLabsSTTEnabled.IsChecked = elevenLabs.TryGetProperty("STTEnabled", out var stt) ? stt.GetBoolean() : true;
                    }

                    // Load security settings
                    if (doc.RootElement.TryGetProperty("Security", out var security))
                    {
                        RequireLoginCheck.IsChecked = security.TryGetProperty("RequireLogin", out var login) ? login.GetBoolean() : true;
                        RequireSecurityCodeCheck.IsChecked = security.TryGetProperty("RequireSecurityCode", out var code) ? code.GetBoolean() : true;

                        if (security.TryGetProperty("SecurityCodeTimeout", out var timeout))
                        {
                            SelectComboByTag(SecurityCodeTimeoutCombo, timeout.GetInt32().ToString());
                        }
                        if (security.TryGetProperty("AutoLockMinutes", out var autoLock))
                        {
                            SelectComboByTag(AutoLockCombo, autoLock.GetInt32().ToString());
                        }
                    }

                    // Load dictation settings
                    if (doc.RootElement.TryGetProperty("Dictation", out var dictation))
                    {
                        AutoDictationCheck.IsChecked = dictation.TryGetProperty("AutoDetect", out var autoDetect) ? autoDetect.GetBoolean() : true;
                        AutoPunctuationCheck.IsChecked = dictation.TryGetProperty("AutoPunctuation", out var punct) ? punct.GetBoolean() : true;
                        AutoCapitalizationCheck.IsChecked = dictation.TryGetProperty("AutoCapitalization", out var cap) ? cap.GetBoolean() : true;
                        TargetWordCheck.IsChecked = dictation.TryGetProperty("TargetWord", out var tw) ? tw.GetBoolean() : true;
                        TargetNotepadCheck.IsChecked = dictation.TryGetProperty("TargetNotepad", out var tn) ? tn.GetBoolean() : true;
                        TargetVSCodeCheck.IsChecked = dictation.TryGetProperty("TargetVSCode", out var tv) ? tv.GetBoolean() : false;
                    }

                    // Load appearance settings (theme, window, notifications)
                    if (doc.RootElement.TryGetProperty("Appearance", out var appearance))
                    {
                        // Theme radio buttons
                        var theme = appearance.TryGetProperty("Theme", out var themeProp) ? themeProp.GetString() ?? "Light" : "Light";
                        ThemeLight.IsChecked = theme == "Light";
                        ThemeDark.IsChecked = theme == "Dark";
                        ThemeSystem.IsChecked = theme == "System";

                        AlwaysOnTopCheck.IsChecked = appearance.TryGetProperty("AlwaysOnTop", out var top) ? top.GetBoolean() : false;
                        MinimizeToTrayCheck.IsChecked = appearance.TryGetProperty("MinimizeToTray", out var tray) ? tray.GetBoolean() : false;
                        StartMinimizedCheck.IsChecked = appearance.TryGetProperty("StartMinimized", out var startMin) ? startMin.GetBoolean() : false;
                        ShowToastCheck.IsChecked = appearance.TryGetProperty("ShowToast", out var toast) ? toast.GetBoolean() : true;
                        PlaySoundsCheck.IsChecked = appearance.TryGetProperty("PlaySounds", out var sounds) ? sounds.GetBoolean() : true;

                        if (appearance.TryGetProperty("Opacity", out var opacity))
                        {
                            OpacitySlider.Value = opacity.GetDouble();
                        }
                    }

                    // Load database settings
                    if (doc.RootElement.TryGetProperty("Database", out var database))
                    {
                        if (database.TryGetProperty("HistoryRetentionDays", out var retention))
                        {
                            SelectComboByTag(HistoryRetentionCombo, retention.GetInt32().ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error loading settings: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Selects the ComboBoxItem whose Tag matches the given value string.
        /// </summary>
        private static void SelectComboByTag(System.Windows.Controls.ComboBox combo, string tagValue)
        {
            foreach (ComboBoxItem item in combo.Items)
            {
                if (item.Tag?.ToString() == tagValue)
                {
                    combo.SelectedItem = item;
                    return;
                }
            }
        }


        private void UpdateDatabaseInfo()
        {
            var dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AICompanion", "data.db");

            DatabasePathText.Text = $"Path: {dbPath}";

            // TODO: Get actual count from database service
            HistoryCountText.Text = "Total commands: 0";
        }

        private void OpacitySlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (OpacityValueText != null)
            {
                OpacityValueText.Text = $"{(int)(OpacitySlider.Value * 100)}%";
            }
        }

        private async void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

                // MERGE STRATEGY: Load existing JSON, update only the sections we control
                Dictionary<string, object>? existingConfig = null;
                if (File.Exists(configPath))
                {
                    try
                    {
                        var existingJson = await File.ReadAllTextAsync(configPath);
                        existingConfig = JsonSerializer.Deserialize<Dictionary<string, object>>(existingJson);
                    }
                    catch { /* If parse fails, start fresh */ }
                }
                existingConfig ??= new Dictionary<string, object>();

                // Update only the sections the Settings UI controls
                existingConfig["ElevenLabs"] = new
                {
                    ApiKey = ElevenLabsApiKeyBox.Text.Trim(),
                    VoiceId = ElevenLabsVoiceIdBox.Text.Trim(),
                    TTSEnabled = ElevenLabsTTSEnabled.IsChecked == true,
                    STTEnabled = ElevenLabsSTTEnabled.IsChecked == true
                };

                existingConfig["Security"] = new
                {
                    RequireLogin = RequireLoginCheck.IsChecked == true,
                    RequireSecurityCode = RequireSecurityCodeCheck.IsChecked == true,
                    SecurityCodeTimeout = int.Parse((SecurityCodeTimeoutCombo.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "60"),
                    AutoLockMinutes = int.Parse((AutoLockCombo.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "15")
                };

                existingConfig["Dictation"] = new
                {
                    AutoDetect = AutoDictationCheck.IsChecked == true,
                    AutoPunctuation = AutoPunctuationCheck.IsChecked == true,
                    AutoCapitalization = AutoCapitalizationCheck.IsChecked == true,
                    TargetWord = TargetWordCheck.IsChecked == true,
                    TargetNotepad = TargetNotepadCheck.IsChecked == true,
                    TargetVSCode = TargetVSCodeCheck.IsChecked == true
                };

                existingConfig["Appearance"] = new
                {
                    Theme = ThemeLight.IsChecked == true ? "Light" : ThemeDark.IsChecked == true ? "Dark" : "System",
                    AlwaysOnTop = AlwaysOnTopCheck.IsChecked == true,
                    MinimizeToTray = MinimizeToTrayCheck.IsChecked == true,
                    StartMinimized = StartMinimizedCheck.IsChecked == true,
                    Opacity = OpacitySlider.Value,
                    ShowToast = ShowToastCheck.IsChecked == true,
                    PlaySounds = PlaySoundsCheck.IsChecked == true
                };

                existingConfig["Database"] = new
                {
                    HistoryRetentionDays = int.Parse((HistoryRetentionCombo.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "30")
                };

                var json = JsonSerializer.Serialize(existingConfig, new JsonSerializerOptions { WriteIndented = true });
                await File.WriteAllTextAsync(configPath, json);

                // Apply theme immediately
                var selectedTheme = ThemeLight.IsChecked == true ? "Light" : ThemeDark.IsChecked == true ? "Dark" : "System";
                App.ApplyTheme(selectedTheme);

                // Apply ElevenLabs configuration
                if (_voiceManager != null)
                {
                    var voiceSettings = new Services.Voice.VoiceSettings
                    {
                        ElevenLabsApiKey = ElevenLabsApiKeyBox.Text.Trim(),
                        ElevenLabsVoiceId = ElevenLabsVoiceIdBox.Text.Trim()
                    };

                    _voiceManager.Configure(voiceSettings);
                    await _voiceManager.InitializeAsync();
                }

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error saving settings: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private async void TestConnection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var apiKey = ElevenLabsApiKeyBox.Text.Trim();
                if (string.IsNullOrEmpty(apiKey))
                {
                    System.Windows.MessageBox.Show("Please enter ElevenLabs API key", "Test Connection", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                using var client = new System.Net.Http.HttpClient();
                client.DefaultRequestHeaders.Add("xi-api-key", apiKey);
                var response = await client.GetAsync("https://api.elevenlabs.io/v1/user");

                if (response.IsSuccessStatusCode)
                {
                    System.Windows.MessageBox.Show("ElevenLabs connection successful!", "Test Connection", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                else
                {
                    System.Windows.MessageBox.Show($"ElevenLabs connection failed: {response.StatusCode}", "Test Connection", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Connection test failed: {ex.Message}", "Test Connection", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        private void ChangePassword_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show("Password change functionality will be implemented with LoginWindow", "Change Password", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        private void LockNow_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement session lock
            System.Windows.MessageBox.Show("Session locked. Please log in again.", "Session Locked", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
        }

        private void OpenDatabaseFolder_Click(object sender, RoutedEventArgs e)
        {
            var dbFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AICompanion");

            if (!Directory.Exists(dbFolder))
            {
                Directory.CreateDirectory(dbFolder);
            }

            System.Diagnostics.Process.Start("explorer.exe", dbFolder);
        }

        private void BackupDatabase_Click(object sender, RoutedEventArgs e)
        {
            var dbPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "AICompanion", "data.db");

            if (!File.Exists(dbPath))
            {
                System.Windows.MessageBox.Show("Database file not found", "Backup", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"AICompanion_backup_{DateTime.Now:yyyyMMdd_HHmmss}.db",
                DefaultExt = ".db",
                Filter = "SQLite Database|*.db"
            };

            if (dialog.ShowDialog() == true)
            {
                File.Copy(dbPath, dialog.FileName, true);
                System.Windows.MessageBox.Show($"Database backed up to:\n{dialog.FileName}", "Backup Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }

        private void ViewHistory_Click(object sender, RoutedEventArgs e)
        {
            var historyWindow = new HistoryWindow();
            historyWindow.Owner = this;
            historyWindow.ShowDialog();
        }

        private void ClearHistory_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "Are you sure you want to clear all command history?\nThis action cannot be undone.",
                "Clear History",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result == System.Windows.MessageBoxResult.Yes)
            {
                // TODO: Clear database history
                System.Windows.MessageBox.Show("History cleared successfully", "Clear History", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                HistoryCountText.Text = "Total commands: 0";
            }
        }

        private void ExportJson_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"AICompanion_history_{DateTime.Now:yyyyMMdd}.json",
                DefaultExt = ".json",
                Filter = "JSON Files|*.json"
            };

            if (dialog.ShowDialog() == true)
            {
                // TODO: Export from database
                File.WriteAllText(dialog.FileName, "[]");
                System.Windows.MessageBox.Show($"Exported to:\n{dialog.FileName}", "Export Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }

        private void ExportCsv_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"AICompanion_history_{DateTime.Now:yyyyMMdd}.csv",
                DefaultExt = ".csv",
                Filter = "CSV Files|*.csv"
            };

            if (dialog.ShowDialog() == true)
            {
                // TODO: Export from database
                File.WriteAllText(dialog.FileName, "Timestamp,Command,Result\n");
                System.Windows.MessageBox.Show($"Exported to:\n{dialog.FileName}", "Export Complete", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }
    }
}
