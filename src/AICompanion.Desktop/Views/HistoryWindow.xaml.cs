using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using WpfBrush = System.Windows.Media.Brush;
using WpfSolidColorBrush = System.Windows.Media.SolidColorBrush;
using WpfColor = System.Windows.Media.Color;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.Services.Database;

namespace AICompanion.Desktop.Views
{
    public partial class HistoryWindow : Window
    {
        private readonly DatabaseService? _databaseService;
        private ObservableCollection<HistoryItem> _allItems = new();
        private ObservableCollection<HistoryItem> _filteredItems = new();

        public HistoryWindow()
        {
            InitializeComponent();

            _databaseService = App.ServiceProvider?.GetService<DatabaseService>();
            HistoryListView.ItemsSource = _filteredItems;

            Loaded += HistoryWindow_Loaded;
        }

        private async void HistoryWindow_Loaded(object sender, RoutedEventArgs e)
        {
            await LoadHistoryAsync();
        }

        private async System.Threading.Tasks.Task LoadHistoryAsync()
        {
            _allItems.Clear();

            try
            {
                // TODO: Load from database service
                // For now, add sample data
                var sampleItems = new List<HistoryItem>
                {
                    new HistoryItem
                    {
                        Timestamp = DateTime.Now.AddMinutes(-5),
                        Command = "Open Microsoft Word",
                        Type = "Voice",
                        Success = true,
                        Details = "Launched WINWORD.EXE"
                    },
                    new HistoryItem
                    {
                        Timestamp = DateTime.Now.AddMinutes(-10),
                        Command = "Type: Hello world",
                        Type = "Dictation",
                        Success = true,
                        Details = "Typed 11 characters"
                    },
                    new HistoryItem
                    {
                        Timestamp = DateTime.Now.AddMinutes(-15),
                        Command = "Save file",
                        Type = "Voice",
                        Success = true,
                        Details = "Sent Ctrl+S"
                    },
                    new HistoryItem
                    {
                        Timestamp = DateTime.Now.AddHours(-1),
                        Command = "Open calculator",
                        Type = "Quick",
                        Success = true,
                        Details = "Launched calc.exe"
                    },
                    new HistoryItem
                    {
                        Timestamp = DateTime.Now.AddHours(-2),
                        Command = "Search weather forecast",
                        Type = "Voice",
                        Success = true,
                        Details = "Opened browser search"
                    }
                };

                foreach (var item in sampleItems)
                {
                    _allItems.Add(item);
                }

                ApplyFilters();
                UpdateCountText();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error loading history: {ex.Message}", "Error",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }

            await System.Threading.Tasks.Task.CompletedTask;
        }

        private void ApplyFilters()
        {
            var searchText = SearchBox?.Text?.ToLowerInvariant() ?? "";
            var dateFilter = (DateFilterCombo?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All Time";
            var typeFilter = (TypeFilterCombo?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All Types";

            var filtered = _allItems.AsEnumerable();

            // Text search
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                filtered = filtered.Where(item =>
                    item.Command.ToLowerInvariant().Contains(searchText) ||
                    item.Details.ToLowerInvariant().Contains(searchText));
            }

            // Date filter
            filtered = dateFilter switch
            {
                "Today" => filtered.Where(item => item.Timestamp.Date == DateTime.Today),
                "Last 7 Days" => filtered.Where(item => item.Timestamp >= DateTime.Today.AddDays(-7)),
                "Last 30 Days" => filtered.Where(item => item.Timestamp >= DateTime.Today.AddDays(-30)),
                _ => filtered
            };

            // Type filter
            filtered = typeFilter switch
            {
                "Voice Commands" => filtered.Where(item => item.Type == "Voice"),
                "Quick Actions" => filtered.Where(item => item.Type == "Quick"),
                "Dictation" => filtered.Where(item => item.Type == "Dictation"),
                "System" => filtered.Where(item => item.Type == "System"),
                _ => filtered
            };

            _filteredItems.Clear();
            foreach (var item in filtered.OrderByDescending(i => i.Timestamp))
            {
                _filteredItems.Add(item);
            }

            UpdateCountText();
        }

        private void UpdateCountText()
        {
            var total = _allItems.Count;
            var filtered = _filteredItems.Count;

            if (total == filtered)
            {
                HistoryCountText.Text = $"{total} commands";
            }
            else
            {
                HistoryCountText.Text = $"{filtered} of {total} commands";
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void DateFilter_Changed(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private void TypeFilter_Changed(object sender, SelectionChangedEventArgs e)
        {
            ApplyFilters();
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            await LoadHistoryAsync();
        }

        private void ExportSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = HistoryListView.SelectedItems.Cast<HistoryItem>().ToList();

            if (selected.Count == 0)
            {
                System.Windows.MessageBox.Show("Please select items to export", "Export",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            ExportItems(selected, $"AICompanion_selected_{DateTime.Now:yyyyMMdd_HHmmss}");
        }

        private void ExportAll_Click(object sender, RoutedEventArgs e)
        {
            if (_filteredItems.Count == 0)
            {
                System.Windows.MessageBox.Show("No items to export", "Export",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            ExportItems(_filteredItems.ToList(), $"AICompanion_history_{DateTime.Now:yyyyMMdd}");
        }

        private void ExportItems(List<HistoryItem> items, string defaultName)
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = defaultName,
                DefaultExt = ".json",
                Filter = "JSON Files|*.json|CSV Files|*.csv"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    if (dialog.FileName.EndsWith(".csv"))
                    {
                        var csv = "Timestamp,Command,Type,Success,Details\n";
                        foreach (var item in items)
                        {
                            csv += $"\"{item.Timestamp:yyyy-MM-dd HH:mm:ss}\",\"{item.Command.Replace("\"", "\"\"")}\",\"{item.Type}\",\"{item.Success}\",\"{item.Details.Replace("\"", "\"\"")}\"\n";
                        }
                        File.WriteAllText(dialog.FileName, csv);
                    }
                    else
                    {
                        var json = JsonSerializer.Serialize(items, new JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(dialog.FileName, json);
                    }

                    System.Windows.MessageBox.Show($"Exported {items.Count} items to:\n{dialog.FileName}", "Export Complete",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Export failed: {ex.Message}", "Error",
                        System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                }
            }
        }

        private void DeleteSelected_Click(object sender, RoutedEventArgs e)
        {
            var selected = HistoryListView.SelectedItems.Cast<HistoryItem>().ToList();

            if (selected.Count == 0)
            {
                System.Windows.MessageBox.Show("Please select items to delete", "Delete",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            var result = System.Windows.MessageBox.Show(
                $"Delete {selected.Count} selected item(s)?\nThis cannot be undone.",
                "Confirm Delete",
                System.Windows.MessageBoxButton.YesNo,
                System.Windows.MessageBoxImage.Warning);

            if (result == System.Windows.MessageBoxResult.Yes)
            {
                foreach (var item in selected)
                {
                    _allItems.Remove(item);
                    _filteredItems.Remove(item);
                }

                // TODO: Delete from database
                UpdateCountText();
            }
        }

        private void ReplayCommand_Click(object sender, RoutedEventArgs e)
        {
            if (HistoryListView.SelectedItem is HistoryItem item)
            {
                var result = System.Windows.MessageBox.Show(
                    $"Replay command:\n\"{item.Command}\"",
                    "Replay Command",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Question);

                if (result == System.Windows.MessageBoxResult.Yes)
                {
                    // TODO: Send command to LocalCommandProcessor
                    System.Windows.MessageBox.Show("Command replay functionality will be implemented with MainWindow integration",
                        "Replay", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("Please select a command to replay", "Replay",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }

    public class HistoryItem
    {
        public DateTime Timestamp { get; set; }
        public string Command { get; set; } = "";
        public string Type { get; set; } = "Voice";
        public bool Success { get; set; } = true;
        public string Details { get; set; } = "";

        public string ResultIcon => Success ? "✓" : "✗";

        public WpfBrush TypeColor => Type switch
        {
            "Voice" => new WpfSolidColorBrush(WpfColor.FromRgb(0x0F, 0x76, 0x6E)),
            "Dictation" => new WpfSolidColorBrush(WpfColor.FromRgb(0x14, 0xB8, 0xA6)),
            "Quick" => new WpfSolidColorBrush(WpfColor.FromRgb(0x6B, 0x7B, 0x8D)),
            "System" => new WpfSolidColorBrush(WpfColor.FromRgb(0xF5, 0x9E, 0x0B)),
            _ => new WpfSolidColorBrush(WpfColor.FromRgb(0x9C, 0xA3, 0xAF))
        };
    }
}
