using System;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Npgsql;


namespace LibraryAssistant
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DatabaseConnectionConnectButton_Click(object sender, RoutedEventArgs e)
        {
            TextBox host = (TextBox)this.FindName("DatabaseConnectionHostInput");
            TextBox port = (TextBox)this.FindName("DatabaseConnectionPortInput");
            TextBox user = (TextBox)this.FindName("DatabaseConnectionUserInput");
            TextBox password = (TextBox)this.FindName("DatabaseConnectionPasswordInput");
            TextBox databaseName = (TextBox)this.FindName("DatabaseConnectionDatabaseNameInput");

            if (host.Text == "" || port.Text == "" || user.Text == "" || password.Text == "" || databaseName.Text == "") {
                TextBox errorMessage = (TextBox)this.FindName("DatabaseConnectionErrorMessage");
                errorMessage.Foreground = Brushes.Red;
                errorMessage.FontWeight = FontWeights.Bold;
                errorMessage.Text = "Все поля формы должны быть заполнены!";
                return;
            }

            string connectionString = String.Format("Host={0};Port={1};Database={2};Username={3};Password={4};", host.Text, port.Text, databaseName.Text, user.Text, password.Text);
            NpgsqlConnection connection = new NpgsqlConnection(connectionString);

            try {
                connection.Open();
                connection.Close();
            }
            catch {
                connection.Close();
                TextBox errorMessage = (TextBox)this.FindName("DatabaseConnectionErrorMessage");
                errorMessage.Foreground = Brushes.Red;
                errorMessage.FontWeight = FontWeights.Bold;
                errorMessage.Text = "Не удалось подключиться к БД!";
                return;
            }

            ApplicationWindow application = new ApplicationWindow(connection);
            application.Show();
            this.Close();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e) {
            Regex template = new Regex("[^0-9]+");
            e.Handled = template.IsMatch(e.Text);
        }
    }
}
