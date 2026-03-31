using System;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using AICompanion.Desktop.Services.Database;

namespace AICompanion.Desktop.Views
{
    public partial class PrivacyPolicyWindow : Window
    {
        private readonly DatabaseService? _database;

        public bool Accepted { get; private set; }

        public PrivacyPolicyWindow()
        {
            InitializeComponent();
            _database = App.ServiceProvider?.GetService<DatabaseService>();
        }

        private async void Accept_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_database != null)
                    await _database.AcceptPrivacyPolicyAsync();
            }
            catch
            {
                // Non-fatal — user can still proceed
            }

            Accepted = true;
            DialogResult = true;
            Close();
        }

        private void Decline_Click(object sender, RoutedEventArgs e)
        {
            Accepted = false;
            DialogResult = false;
            Close();
        }
    }
}
