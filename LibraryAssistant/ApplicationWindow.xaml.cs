using LibraryAssistant.Models;
using Npgsql;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;

namespace LibraryAssistant
{
    /// <summary>
    /// Логика взаимодействия для ApplicationWindow.xaml
    /// </summary>
    /// 
    public partial class ApplicationWindow : Window
    {
        public string ReportsPath { get; set; }
        public NpgsqlConnection connection { get; set; }
        
        public ApplicationWindow(NpgsqlConnection nc)
        {
            InitializeComponent();

            this.connection = nc;
            this.DataContext = DateTime.Now;

            createTables();
        }
        private void restartConnection()
        {
            this.connection.Close();
            this.connection.Open();
        }
        private void createTables()
        {
            this.connection.Open();
            NpgsqlCommand command = new NpgsqlCommand("SELECT table_name FROM information_schema.tables WHERE table_schema='public';", this.connection);
            NpgsqlDataReader reader = command.ExecuteReader();

            if (reader.HasRows) {
                this.connection.Close();
                return;
            }

            restartConnection();
            using (var transaction = connection.BeginTransaction())
            {
                command = new NpgsqlCommand("CREATE TABLE genres (id SERIAL PRIMARY KEY, name VARCHAR(128) UNIQUE NOT NULL);", this.connection);
                command.ExecuteNonQuery();

                command = new NpgsqlCommand("CREATE TABLE authors (id SERIAL PRIMARY KEY, name VARCHAR(128) NOT NULL, surname VARCHAR(128) NOT NULL, patronymic VARCHAR(128) NOT NULL);", this.connection);
                command.ExecuteNonQuery();

                command = new NpgsqlCommand("CREATE TABLE users (id SERIAL PRIMARY KEY, name VARCHAR(128) NOT NULL, surname VARCHAR(128) NOT NULL, patronymic VARCHAR(128) NOT NULL, email VARCHAR(128) UNIQUE NOT NULL, register_date DATE NOT NULL);", this.connection);
                command.ExecuteNonQuery();
                command = new NpgsqlCommand("INSERT INTO users VALUES (0, 'default', 'default', 'default', 'default', '11.11.1111'::date);", connection);
                command.ExecuteNonQuery();

                command = new NpgsqlCommand("CREATE TABLE books (id SERIAL PRIMARY KEY, name VARCHAR(256) NOT NULL, genres INTEGER[] NOT NULL, authors INTEGER[] NOT NULL, status INTEGER NOT NULL DEFAULT 0, register_date DATE NOT NULL, FOREIGN KEY (status) REFERENCES users (id) ON DELETE SET DEFAULT);", this.connection);
                command.ExecuteNonQuery();

                command = new NpgsqlCommand("CREATE TABLE facts (user_id INTEGER NOT NULL, type VARCHAR(6) NOT NULL, taking_date DATE NOT NULL, return_date DATE NOT NULL, real_return_date VARCHAR(10), books INTEGER[] NOT NULL, CHECK (taking_date < return_date), FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE);", this.connection);
                command.ExecuteNonQuery();

                transaction.Commit();
            }
        }

        // Users TAB
        private void SearchUserSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchUserSectionNameInput.Text = "";
            SearchUserSectionSurnameInput.Text = "";
            SearchUserSectionPatronymicInput.Text = "";
            SearchUserSectionResultsList.ItemsSource = new List<User> { };

            CreateEditUserSectionNameInput.Text = "";
            CreateEditUserSectionSurnameInput.Text = "";
            CreateEditUserSectionPatronymicInput.Text = "";
            CreateEditUserSectionEmailInput.Text = "";
            CreateEditUserSectionRegisterDateInput.Text = "";
            CreateEditUserSectionBooksHistoryList.ItemsSource = new List<string> { };
        }
        private void SearchUserSectionSearchButton_Click(object sender, RoutedEventArgs e)
        {
            string name = SearchUserSectionNameInput.Text.Trim();
            string surname = SearchUserSectionSurnameInput.Text.Trim();
            string patronymic = SearchUserSectionPatronymicInput.Text.Trim();

            if (name == "" && surname == "" && patronymic == "") {
                MessageBox.Show("Заполните данными для поиска хотя бы одно поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }
            
            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, surname, patronymic, email, register_date FROM users WHERE name LIKE '{name}%' AND surname LIKE '{surname}%' AND patronymic LIKE '{patronymic}%' AND id > 0;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var users = new List<User> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        var ssurname = reader.GetString(reader.GetOrdinal("surname"));
                        var ppatronymic = reader.GetString(reader.GetOrdinal("patronymic"));
                        var email = reader.GetString(reader.GetOrdinal("email"));
                        var registerDate = reader.GetDateTime(reader.GetOrdinal("register_date")).ToString("dd.MM.yyyy");

                        users.Add(new User(id, nname, ssurname, ppatronymic, email, registerDate));
                    }

                    SearchUserSectionResultsList.ItemsSource = users;
                }
                else
                {
                    SearchUserSectionResultsList.ItemsSource = new List<User> { };
                }
            }
            connection.Close();
        }
        private void UserSectionSaveUserButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditUserSectionNameInput.Text.Trim();
            string surname = CreateEditUserSectionSurnameInput.Text.Trim();
            string patronymic = CreateEditUserSectionPatronymicInput.Text.Trim();
            string email = CreateEditUserSectionEmailInput.Text.Trim();
            string register_date = CreateEditUserSectionRegisterDateInput.Text;
            Regex template = new Regex(".+@(gmail|yandex|mail|yahoo|outlook)\\.(ru|com|net)");

            if (name == "" ||  surname == "" || patronymic == "" || email == "") {
                MessageBox.Show("Все поля формы должны быть заполнены.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            if (!template.IsMatch(email)) {
                MessageBox.Show("Введённая почта не соответствует стандарту.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            // Проверка на наличие пользователя в БД 
            if (register_date == "") // Пользователь отсутствует в БД
            {
                using (var command = new NpgsqlCommand($"SELECT email FROM users WHERE email = '{email}';", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        MessageBox.Show("Пользователь с таким адресом электронной почты уже существует.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        connection.Close();
                        return;
                    }
                }
                restartConnection();

                register_date = DateTime.Now.Date.ToString();
                using (var command = new NpgsqlCommand("INSERT INTO users (name, surname, patronymic, email, register_date) VALUES (@p1, @p2, @p3, @p4, @p5::date);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.Parameters.AddWithValue("p4", email);
                    command.Parameters.AddWithValue("p5", register_date);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый пользователь был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);
            }
            else // Пользователь уже существует
            {
                int id = 1;
                if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
                {
                    id = selectedUser.Id;
                }

                using (var command = new NpgsqlCommand($"UPDATE users SET name = @p1, surname = @p2, patronymic = @p3, email = @p4 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.Parameters.AddWithValue("p4", email);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Пользователь был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);
            }
            connection.Close();
        }
        private void UserSectionDeleteUserButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
            {
                int id = selectedUser.Id;

                connection.Open();
                var command = new NpgsqlCommand($"DELETE FROM users WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Пользователь был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Удаляемый пользователь не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void SearchUserSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Получение данных о выбранном пользователе и истории его взятия/возврата книг.
            if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
            {
                CreateEditUserSectionNameInput.Text = selectedUser.Name;
                CreateEditUserSectionSurnameInput.Text = selectedUser.Surname;
                CreateEditUserSectionPatronymicInput.Text = selectedUser.Patronymic;
                CreateEditUserSectionEmailInput.Text = selectedUser.Email;
                CreateEditUserSectionRegisterDateInput.Text = selectedUser.RegisterDate;

                // Формирование истории взятия/возврата пользователем книг библиотеки.
                connection.Open();
                using (var command = new NpgsqlCommand($"SELECT f.type, f.taking_date, f.return_date, f.real_return_date, b.name FROM facts AS f CROSS JOIN unnest(f.books) AS book JOIN books AS b ON b.id = book WHERE f.user_id = {selectedUser.Id};", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        var facts = new List<Fact>();

                        while (reader.Read())
                        {
                            var type = reader.GetString(reader.GetOrdinal("type"));
                            if (type == "Taking")
                            {
                                type = "Взята";
                            }
                            else if (type == "Return")
                            {
                                type = "Возвращена";
                            }
                            else
                            {
                                type = "Возвращена с опозданием";
                            }
                            var takingDate = reader.GetDateTime(reader.GetOrdinal("taking_date")).ToString("dd.MM.yyyy");
                            var returnDate = reader.GetDateTime(reader.GetOrdinal("return_date")).ToString("dd.MM.yyyy");
                            var realReturnDate = reader.GetString(reader.GetOrdinal("real_return_date"));
                            var book = reader.GetString(reader.GetOrdinal("name"));

                            facts.Add(new Fact(type, takingDate, returnDate, realReturnDate, book));
                        }

                        CreateEditUserSectionBooksHistoryList.ItemsSource = facts;
                    }
                    else
                    {
                        CreateEditUserSectionBooksHistoryList.ItemsSource = new List<User> { };
                    }
                }
                connection.Close();
            }
        }
        
        // Authors TAB
        private void SearchAuthorSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchAuthorSectionNameInput.Text = "";
            SearchAuthorSectionSurnameInput.Text = "";
            SearchAuthorSectionPatronymicInput.Text = "";
            SearchAuthorSectionResultsList.ItemsSource = new List<User> { };

            CreateEditAuthorSectionNameInput.Text = "";
            CreateEditAuthorSectionSurnameInput.Text = "";
            CreateEditAuthorSectionPatronymicInput.Text = "";
        }
        private void SearchAuthorSectionSearchButton_Click(Object sender, RoutedEventArgs e)
        {
            string name = SearchAuthorSectionNameInput.Text.Trim();
            string surname = SearchAuthorSectionSurnameInput.Text.Trim();
            string patronymic = SearchAuthorSectionPatronymicInput.Text.Trim();

            if (name == "" && surname == "" && patronymic == "") {
                MessageBox.Show("Заполните данными для поиска хотя бы одно поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, surname, patronymic FROM authors WHERE name LIKE '{name}%' AND surname LIKE '{surname}%' AND patronymic LIKE '{patronymic}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        var ssurname = reader.GetString(reader.GetOrdinal("surname"));
                        var ppatronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        authors.Add(new Author(id, nname, ssurname, ppatronymic));
                    }

                    SearchAuthorSectionResultsList.ItemsSource = authors;
                }
                else
                {
                    SearchAuthorSectionResultsList.ItemsSource = new List<Author> { };
                }
            }
            connection.Close();
        }
        private void SearchAuthorSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Заполнение полей формы для создания/редактирования автора данными о выбранном авторе.
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor)
            {
                CreateEditAuthorSectionNameInput.Text = selectedAuthor.Name;
                CreateEditAuthorSectionSurnameInput.Text = selectedAuthor.Surname;
                CreateEditAuthorSectionPatronymicInput.Text = selectedAuthor.Patronymic;
            }
        }
        private void AuthorSectionSaveAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditAuthorSectionNameInput.Text.Trim();
            string surname = CreateEditAuthorSectionSurnameInput.Text.Trim();
            string patronymic = CreateEditAuthorSectionPatronymicInput.Text.Trim();

            if (name == "" || surname == "" || patronymic == "") {
                MessageBox.Show("Все поля формы должны быть заполнены.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            int id = -1;
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor) // Сработает только при выборе автора из списка результатов поиска
            {
                id = selectedAuthor.Id;
            }

            connection.Open();
            // Проверка на наличие автора в таблице БД
            if (id == -1)
            {
                using (var command = new NpgsqlCommand($"SELECT * FROM authors WHERE name = '{name}' AND surname = '{surname}' AND patronymic = '{patronymic}';", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        MessageBox.Show("Данный автор уже существует в базе данных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        connection.Close();
                        return;
                    }
                }
                restartConnection();

                using (var command = new NpgsqlCommand("INSERT INTO authors (name, surname, patronymic) VALUES (@p1, @p2, @p3);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый автор был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                using (var command = new NpgsqlCommand($"UPDATE authors SET name = @p1, surname = @p2, patronymic = @p3 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Автор был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);
            }
            connection.Close();
        }
        private void AuthorSectionDeleteAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor)
            {
                int id = selectedAuthor.Id;

                connection.Open();
                var command = new NpgsqlCommand($"DELETE FROM authors WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Автор был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Удаляемый автор не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Genres TAB
        private void SearchGenresSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchGenresSectionNameInput.Text = "";
            SearchGenresSectionResultsList.ItemsSource = new List<Genre> { };

            CreateEditGenreSectionNameInput.Text = "";
        }
        private void SearchGenresSectionSearchButton_Click(Object sender, RoutedEventArgs e)
        {
            string name = SearchGenresSectionNameInput.Text.Trim();

            if (name == "") {
                MessageBox.Show("Заполните данными для поиска поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name FROM genres WHERE name LIKE '{name}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));

                        genres.Add(new Genre(id, nname));
                    }

                    SearchGenresSectionResultsList.ItemsSource = genres;
                }
                else
                {
                    SearchGenresSectionResultsList.ItemsSource = new List<Genre> { };
                }
            }
            connection.Close();
        }
        private void SearchGenresSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Заполнение полей формы для создания/редактирования жанра данными о выбранном жанре.
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre)
            {
                CreateEditGenreSectionNameInput.Text = selectedGenre.Name;
            }
        }
        private void GenresSectionSaveGenreButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditGenreSectionNameInput.Text.Trim();

            if (name == "") {
                MessageBox.Show("Все поля формы должны быть заполнены.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            int id = -1;
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre) // Сработает только при выборе жанра из списка результатов поиска
            {
                id = selectedGenre.Id;
            }

            connection.Open();
            // Проверка на наличие жанра в таблице БД
            if (id == -1)
            {
                using (var command = new NpgsqlCommand($"SELECT * FROM genres WHERE name = '{name}';", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        MessageBox.Show("Данный жанр уже существует в базе данных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        connection.Close();
                        return;
                    }
                }
                restartConnection();

                using (var command = new NpgsqlCommand("INSERT INTO genres (name) VALUES (@p1);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый жанр был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                using (var command = new NpgsqlCommand($"UPDATE genres SET name = @p1 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Жанр был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);
            }
            connection.Close();
        }
        private void GenresSectionDeleteGenreButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre)
            {
                int id = selectedGenre.Id;

                connection.Open();
                var command = new NpgsqlCommand($"DELETE FROM genres WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                connection.Close();

                MessageBox.Show("Жанр был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                MessageBox.Show("Удаляемый жанр не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Books TAB
        private void SearchBookSectionClearFieldsButton_Click(Object sender, RoutedEventArgs e)
        {
            SearchBookSectionNameInput.Text = "";
            SearchBookSectionResultsList.ItemsSource = new List<Book> { };

            CreateEditBookSectionNameInput.Text = "";
            CreateEditBookSectionStatus.Text = "";

            ChoiceGenresSectionGenreInput.Text = "";
            ChoiceGenresSectionResultsList.ItemsSource = new List<Genre> { };

            ChoiceAuthorsSectionAuthorInput.Text = "";
            ChoiceAuthorsSectionResultsList.ItemsSource = new List<Author> { };

            BookGenresSectionGenresList.ItemsSource = new List<Genre> { };
            BookAuthorsSectionAuthorsList.ItemsSource = new List<Author> { };
        }
        private void SearchBookSectionSearchButton_Click(object sender, RoutedEventArgs e)
        {
            string name = SearchBookSectionNameInput.Text.Trim();

            if (name == "") {
                MessageBox.Show("Заполните данными для поиска поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, genres, array_length(genres, 1) as genres_count, authors, array_length(authors, 1) as authors_count, status, CASE WHEN status = 0 THEN 'Свободна' ELSE 'Занята' END AS text_status FROM books WHERE name LIKE '{name}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var books = new List<Book> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        var genres = (int[])reader.GetValue(reader.GetOrdinal("genres"));
                        var genresCount = reader.GetInt16(reader.GetOrdinal("genres_count"));
                        var authors = (int[])reader.GetValue(reader.GetOrdinal("authors"));
                        var authorsCount = reader.GetInt16(reader.GetOrdinal("authors_count"));
                        var status = reader.GetInt16(reader.GetOrdinal("status"));
                        var textStatus = reader.GetString(reader.GetOrdinal("text_status"));

                        books.Add(new Book(id, nname, genres, genresCount, authors, authorsCount, status, textStatus));
                    }

                    SearchBookSectionResultsList.ItemsSource = books;
                }
                else
                {
                    SearchBookSectionResultsList.ItemsSource = new List<Book> { };
                }
            }
            connection.Close();
        }
        private void SearchBookSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Заполнение полей формы для создания/редактирования книги данными о выбранной книге.
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook)
            {
                CreateEditBookSectionNameInput.Text = selectedBook.Name;

                // Формирование статуса книги
                connection.Open();
                if (selectedBook.Status == 0) {
                    CreateEditBookSectionStatus.Text = selectedBook.TextStatus;
                }
                else
                {
                    using (var command = new NpgsqlCommand($"SELECT name, surname, patronymic FROM users WHERE id = {selectedBook.Status}", connection))
                    {
                        NpgsqlDataReader reader = command.ExecuteReader();
                        reader.Read();

                        string name = reader.GetString(reader.GetOrdinal("name"));
                        string surname = reader.GetString(reader.GetOrdinal("surname"));
                        string patronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        CreateEditBookSectionStatus.Text = $"На руках у {surname} {name} {patronymic}.";
                    }
                    restartConnection();
                }
                
                // Формирование списка жанров книги
                string g = "{" + String.Join(", ", selectedBook.Genres) + "}";
                using (var command = new NpgsqlCommand($"SELECT g.id, g.name FROM unnest('{g}'::integer[]) as genre INNER JOIN genres AS g ON genre = g.id;", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var name = reader.GetString(reader.GetOrdinal("name"));

                        genres.Add(new Genre(id, name));
                    }

                    BookGenresSectionGenresList.ItemsSource = genres;
                }
                restartConnection();

                // Формирование списка аворов книги
                string a = "{" + String.Join(", ", selectedBook.Authors) + "}";
                using (var command = new NpgsqlCommand($"SELECT a.id, a.name, a.surname, a.patronymic FROM unnest('{a}'::integer[]) as author INNER JOIN authors AS a ON author = a.id;", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var name = reader.GetString(reader.GetOrdinal("name"));
                        var surname = reader.GetString(reader.GetOrdinal("surname"));
                        var patronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        authors.Add(new Author(id, name, surname, patronymic));
                    }

                    BookAuthorsSectionAuthorsList.ItemsSource = authors;
                }
                connection.Close();
            }
        }
        private void ChoiceGenresSectionSearchGenreButton_Click(object sender, RoutedEventArgs e)
        {
            string name = ChoiceGenresSectionGenreInput.Text.Trim();
            
            if (name == "") {
                MessageBox.Show("Заполните данными для поиска поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name FROM genres WHERE name LIKE '{name}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));

                        genres.Add(new Genre(id, nname));
                    }

                    ChoiceGenresSectionResultsList.ItemsSource = genres;
                }
                else
                {
                    ChoiceGenresSectionResultsList.ItemsSource = new List<Genre> { };
                }
            }
            connection.Close();
        } 
        private void ChoiceAuthorsSectionSearchAuthorButton_Click(Object sender, RoutedEventArgs e)
        {
            string input = ChoiceAuthorsSectionAuthorInput.Text.Trim();

            if (input == "") {
                MessageBox.Show("Заполните данными для поиска поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            string name = input.Split(' ').ElementAtOrDefault(0);
            string surname = input.Split(' ').ElementAtOrDefault(1);
            string patronymic = input.Split(' ').ElementAtOrDefault(2);

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, surname, patronymic FROM authors WHERE name LIKE '{name}%' AND surname LIKE '{surname}%' AND patronymic LIKE '{patronymic}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        var ssurname = reader.GetString(reader.GetOrdinal("surname"));
                        var ppatronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        authors.Add(new Author(id, nname, ssurname, ppatronymic));
                    }

                    ChoiceAuthorsSectionResultsList.ItemsSource = authors;
                }
                else
                {
                    ChoiceAuthorsSectionResultsList.ItemsSource = new List<Author> { };
                }
            }
            connection.Close();
        }
        private void ChoiceGenresSectionAddGenreButton_Click(object sender, RoutedEventArgs e)
        {
            // Добавление найденного жанра в список жанров книги
            if (ChoiceGenresSectionResultsList.SelectedItem is Genre selectedGenre) {
                var foundGenres = ChoiceGenresSectionResultsList.ItemsSource.Cast<Genre>().ToList();
                foundGenres.Remove(selectedGenre);
                ChoiceGenresSectionResultsList.ItemsSource = foundGenres;
                ChoiceGenresSectionGenreInput.Clear();

                var items = BookGenresSectionGenresList.ItemsSource;
                var bookGenres = (items == null) ? new List<Genre> { } : items.Cast<Genre>().ToList();
                if (bookGenres.Find(g => g.Id == selectedGenre.Id) == null)
                {
                    bookGenres.Add(selectedGenre);
                    BookGenresSectionGenresList.ItemsSource = bookGenres;
                }
            } else
            {
                MessageBox.Show("Выберите жанр, который хотите добавить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void ChoiceAuthorsSectionAddAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            // Добавление найденного автора в список авторов книги
            if (ChoiceAuthorsSectionResultsList.SelectedItem is Author selectedAuthor) {
                var foundAuthors = ChoiceAuthorsSectionResultsList.ItemsSource.Cast<Author>().ToList();
                foundAuthors.Remove(selectedAuthor);
                ChoiceAuthorsSectionResultsList.ItemsSource = foundAuthors;
                ChoiceAuthorsSectionAuthorInput.Clear();

                var items = BookAuthorsSectionAuthorsList.ItemsSource;
                var bookAuthors = (items == null) ? new List<Author> { } : items.Cast<Author>().ToList();
                if (bookAuthors.Find(a => a.Id == selectedAuthor.Id) == null) {
                    bookAuthors.Add(selectedAuthor);
                    BookAuthorsSectionAuthorsList.ItemsSource = bookAuthors;
                }
            } else
            {
                MessageBox.Show("Выберите автора, которого хотите добавить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void BookGenresSectionDeleteGenreButton_Click(Object sender, RoutedEventArgs e) {
            // Удаление жанра из списка жанров книги
            Genre selectedGenre = (Genre)BookGenresSectionGenresList.SelectedItem;

            if (selectedGenre == null) {
                MessageBox.Show("Выберите жанр, который хотели бы удалить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            var bookGenres = BookGenresSectionGenresList.ItemsSource.Cast<Genre>().ToList();
            bookGenres.Remove(selectedGenre);
            BookGenresSectionGenresList.ItemsSource = bookGenres;
        }
        private void BookAuthorsSectionDeleteAuthorButton_Click(object sender, RoutedEventArgs e) {
            // Удаление автора из списка авторов книги
            Author selectedAuthor = (Author)BookAuthorsSectionAuthorsList.SelectedItem;

            if (selectedAuthor == null) {
                MessageBox.Show("Выберите автора, которого хотели бы удалить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            var bookAuthors = BookAuthorsSectionAuthorsList.ItemsSource.Cast<Author>().ToList();
            bookAuthors.Remove(selectedAuthor);
            BookAuthorsSectionAuthorsList.ItemsSource = bookAuthors;
        }
        private void SearchBookSectionSaveBookButton_Click(object sender, RoutedEventArgs e) {
            string name = CreateEditBookSectionNameInput.Text.Trim();

            var items = BookGenresSectionGenresList.ItemsSource;
            var bookGenres = (items == null) ? new List<Genre> { } : items.Cast<Genre>().ToList();
            var genresIds = new List<int> { };
            bookGenres.ForEach(g => genresIds.Add(g.Id));
            var genres = "{" + String.Join(",", genresIds) + "}";

            var iitems = BookAuthorsSectionAuthorsList.ItemsSource;
            var bookAuthors = (iitems == null) ? new List<Author> { } : iitems.Cast<Author>().ToList();
            var authorsIds = new List<int> { };
            bookAuthors.ForEach(a => authorsIds.Add(a.Id));
            var authors = "{" + String.Join(",", authorsIds) + "}";
            
            if (name == "" || bookGenres.Count == 0 || bookAuthors.Count == 0) {
                MessageBox.Show("Все поля формы должны быть заполнены.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            int id = -1;
            int status = -1;
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook) {
                id = selectedBook.Id;
                status = selectedBook.Status;
            }

            connection.Open();
            // Проверка на наличие книги в таблице БД
            if (id == -1) {
                var command = new NpgsqlCommand($"INSERT INTO books (name, genres, authors, register_date) VALUES ('{name}', '{genres}'::integer[], '{authors}'::integer[], 'today'::date);", connection);
                command.ExecuteNonQuery();

                MessageBox.Show("Новая книга была успешно создана!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchBookSectionClearFieldsButton_Click(sender, e);
            } else {
                // Проверка на занятость книги кем-либо из читателей
                if (status == 0) {
                    var command = new NpgsqlCommand($"UPDATE books SET name = '{name}', genres = '{genres}'::integer[], authors = '{authors}'::integer[] WHERE id = {id};", connection);
                    command.ExecuteNonQuery();
                    
                    MessageBox.Show("Книга была успешно отредактирована!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    SearchBookSectionClearFieldsButton_Click(sender, e);
                } else {
                    MessageBox.Show("Невозможно отредактировать книгу, находящуюся на руках у читателя!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                }
            }
            connection.Close();
        }
        private void SearchBookSectionDeleteBookButton_Click(object sender, RoutedEventArgs e) {
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook) {
                if (selectedBook.Status != 0) {
                    MessageBox.Show("Невозможно удалить книгу, находящуюся на руках у читателя!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }

                int id = selectedBook.Id;

                connection.Open();
                var command = new NpgsqlCommand($"DELETE FROM books WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                connection.Close();
                
                MessageBox.Show("Книга была успешно удалёна!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchBookSectionClearFieldsButton_Click(sender, e);
            }
            else {
                MessageBox.Show("Удаляемая книга не выбрана из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Facts TAB
        private void FactSectionClearFieldsButton_Click(object sender, RoutedEventArgs e) {
            FactSectionFactUserInput.Text = "";
            FactSectionFactTypeTaking.IsChecked = false;
            FactSectionFactTypeReturn.IsChecked = false;
            FactSectionBooksList.ItemsSource = new List<Book> { };

            AddUserSectionNameInput.Clear();
            AddUserSectionSurnameInput.Clear();
            AddUserSectionPatronymicInput.Clear();
            AddUserSectionUsersList.ItemsSource = new List<User> { };

            AddBookSectionNameInput.Clear();
            AddBookSectionBooksList.ItemsSource = new List<Book> { };
        }
        private void AddUserSectionSearchUserButton_Click(object sender, RoutedEventArgs e)
        {
            string name = AddUserSectionNameInput.Text.Trim();
            string surname = AddUserSectionSurnameInput.Text.Trim();
            string patronymic = AddUserSectionPatronymicInput.Text.Trim();

            if (name == "" && surname == "" && patronymic == "") {
                MessageBox.Show("Заполните данными для поиска хотя бы одно поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, surname, patronymic, email FROM users WHERE name LIKE '{name}%' AND surname LIKE '{surname}%' AND patronymic LIKE '{patronymic}%';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var users = new List<User> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        var ssurname = reader.GetString(reader.GetOrdinal("surname"));
                        var ppatronymic = reader.GetString(reader.GetOrdinal("patronymic"));
                        var email = reader.GetString(reader.GetOrdinal("email"));

                        users.Add(new User(id, nname, ssurname, ppatronymic, email));
                    }

                    AddUserSectionUsersList.ItemsSource = users;
                }
                else
                {
                    AddUserSectionUsersList.ItemsSource = new List<User> { };
                }
            }
            connection.Close();
        }
        private void AddBookSectionSearchBookButton_Click(object sender, RoutedEventArgs e)
        {
            string name = AddBookSectionNameInput.Text.Trim();

            if (name == "") {
                MessageBox.Show("Заполните данными для поиска поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, name, array_length(genres, 1) as genres_count, array_length(authors, 1) as authors_count, status, (CASE WHEN status = 0 THEN 'Свободна' ELSE 'Занята' END) AS text_status FROM books WHERE name LIKE '{name}%';", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows) {
                    var books = new List<Book> { };

                    while (reader.Read()) {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var nname = reader.GetString(reader.GetOrdinal("name"));
                        int[] genres = { };
                        var genresCount = reader.GetInt16(reader.GetOrdinal("genres_count"));
                        int[] authors = { };
                        var authorsCount = reader.GetInt16(reader.GetOrdinal("authors_count"));
                        var status = reader.GetInt16(reader.GetOrdinal("status"));
                        var textStatus = reader.GetString(reader.GetOrdinal("text_status"));

                        books.Add(new Book(id, nname, genres, genresCount, authors, authorsCount, status, textStatus));
                    }

                    AddBookSectionBooksList.ItemsSource = books;
                }
                else
                {
                    AddBookSectionBooksList.ItemsSource = new List<Book> { };
                }
            }
            connection.Close();
        }
        private void AddUserSectionAddSelectedUserButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddUserSectionUsersList.SelectedItem is User selectedUser) {
                FactSectionFactUserInput.Text = $"{selectedUser.Name} {selectedUser.Surname} {selectedUser.Patronymic}";
            } else
            {
                MessageBox.Show("Выберите читателя, которого хотите добавить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void AddBookSectionAddSelectedBookButton_Click(object sender, RoutedEventArgs e)
        {
            // Добавление найденной книги в список книг факта
            if (AddBookSectionBooksList.SelectedItem is Book selectedBook) {
                var foundBooks = AddBookSectionBooksList.ItemsSource.Cast<Book>().ToList();
                foundBooks.Remove(selectedBook);
                AddBookSectionBooksList.ItemsSource = foundBooks;
                AddBookSectionNameInput.Clear();

                var items = FactSectionBooksList.ItemsSource;
                var factBooks = (items == null) ? new List<Book>() : items.Cast<Book>().ToList();
                if (factBooks.Find(b => b.Id == selectedBook.Id) == null) {
                    factBooks.Add(selectedBook);
                    FactSectionBooksList.ItemsSource = factBooks;
                }
            } else
            {
                MessageBox.Show("Выберите книгу, которую хотите добавить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void FactSectionDeleteSelectedBookButton_Click(object sender, RoutedEventArgs e)
        {
            // Удаление книги из списка книг факта
            Book selectedBook = (Book)FactSectionBooksList.SelectedItem;

            if (selectedBook == null) {
                MessageBox.Show("Выберите книгу, которую хотели бы удалить.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            var factBooks = FactSectionBooksList.ItemsSource.Cast<Book>().ToList();
            factBooks.Remove(selectedBook);
            FactSectionBooksList.ItemsSource = factBooks;
        }
        private void FactSectionCreateFactButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddUserSectionUsersList.SelectedItem is User selectedUser)
            {
                if (FactSectionFactTypeReturn.IsChecked == false && FactSectionFactTypeTaking.IsChecked == false)
                {
                    MessageBox.Show("Тип факта не был определён!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }

                var items = FactSectionBooksList.ItemsSource;
                var factBooks = (items == null) ? new List<Book> { } : FactSectionBooksList.ItemsSource.Cast<Book>().ToList();
                if (factBooks.Count == 0) {
                    MessageBox.Show("Факт должен содержать как минимум 1 книгу!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }
                var booksIds = new List<int>();
                foreach (Book b in factBooks) {
                    if (b.Status != 0 && FactSectionFactTypeReturn.IsChecked == false) {
                        MessageBox.Show("Факт не может содержать книги, находящиеся на руках у других читателей!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        return;
                    }
                    booksIds.Add(b.Id);
                }
                booksIds.Sort();
                string books = "{" + String.Join(",", booksIds) + "}";

                if (FactSectionFactTypeTaking.IsChecked == true) {
                    var taking_date = DateTime.Now;
                    var return_date = FactSectionReturnDateInput.SelectedDate.Value;

                    if (taking_date >= return_date) {
                        MessageBox.Show("Дата возврата не может предшествовать или быть такой же, как текущая!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        return;
                    }

                    connection.Open();
                    var command = new NpgsqlCommand($"UPDATE books SET status = {selectedUser.Id} WHERE id IN ({books.Substring(1, books.Length - 2)});", connection);
                    command.ExecuteNonQuery();
                    restartConnection();
                    var ccommand = new NpgsqlCommand($"INSERT INTO facts VALUES ({selectedUser.Id}, 'Taking', '{taking_date.ToString("dd.MM.yyyy")}'::date, '{return_date.ToString("dd.MM.yyyy")}'::date, ' ', '{books}'::integer[]);", connection);
                    ccommand.ExecuteNonQuery();
                    connection.Close();
                    
                    MessageBox.Show("Факт взятия был успешно зарегистрирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    FactSectionClearFieldsButton_Click(sender, e);
                } 
                else if (FactSectionFactTypeReturn.IsChecked == true) {
                    connection.Open();
                    using (var command = new NpgsqlCommand($"UPDATE facts SET type = CASE WHEN return_date >= '{DateTime.Now.ToString("dd.MM.yyyy")}'::date THEN 'Return' ELSE 'Expire' END, real_return_date = '{DateTime.Now.ToString("dd.MM.yyyy")}' WHERE user_id = {selectedUser.Id} AND type = 'Taking' AND books = '{books}'::integer[];", connection)) {
                        int rows = command.ExecuteNonQuery();
                        
                        if (rows == 1) {
                            MessageBox.Show("Факт возврата был успешно зарегистрирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                            FactSectionClearFieldsButton_Click(sender, e);
                        } else
                        {
                            MessageBox.Show("Ошибка при регистрации факта возврата!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                            connection.Close();
                            return;
                        }
                    }
                    restartConnection();
                    var ccommand = new NpgsqlCommand($"UPDATE books SET status = 0 WHERE id IN ({books.Substring(1, books.Length - 2)});", connection);
                    ccommand.ExecuteNonQuery();
                    connection.Close();
                }
            } 
            else {
                MessageBox.Show("Пользователь не был выбран!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Reports TAB
        private void ReportsSectionSaveReportsPathButton_Click(object sender, RoutedEventArgs e) {
            var path = ReportsSectionReportsPathInput.Text.Trim();

            if (path == "") {
                MessageBox.Show("Заполните данными поле формы.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            if (Directory.Exists(path)) {
                this.ReportsPath = path;
                MessageBox.Show("Путь для сохранения отчётов был успешно установлен!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                ReportsSectionReportsPathInput.Clear();
            } else {
                MessageBox.Show("При установке пути для сохранения отчётов возникла ошибка!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void GetBooksRecievedListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null) {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH book_genres AS (SELECT b.id AS book_id, array_agg(g.name) AS genre_name FROM books b INNER JOIN genres g ON b.genres && ARRAY[g.id] GROUP BY b.id), book_authors AS (SELECT b.id AS book_id, array_agg(a.surname || ' ' || a.name || ' ' || a.patronymic) AS author_fullnames FROM books b INNER JOIN authors a ON b.authors && ARRAY[a.id] GROUP BY b.id) SELECT b.id AS book_id, b.name AS book_name, CASE WHEN b.status = 0 THEN 'Свободна' ELSE 'На руках у ' || COALESCE(u.surname || ' ' || u.name || ' ' || u.patronymic, '') END AS book_status, array_to_string(array_agg(DISTINCT bg.genre_name), ', ') AS genre_names, array_to_string(array_agg(DISTINCT ba.author_fullnames), ', ') AS author_fullnames FROM books b LEFT JOIN book_genres bg ON b.id = bg.book_id LEFT JOIN book_authors ba ON b.id = ba.book_id LEFT JOIN users u ON b.status = u.id WHERE b.register_date = CURRENT_DATE GROUP BY b.id, b.name, b.status, u.surname, u.name, u.patronymic;", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень поступивших книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Название");
                row.CreateCell(1).SetCellValue("Статус");
                row.CreateCell(2).SetCellValue("Жанры");
                row.CreateCell(3).SetCellValue("Авторы");

                while (reader.Read()) {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(1));
                    row.CreateCell(1).SetCellValue(reader.GetString(2));
                    row.CreateCell(2).SetCellValue(reader.GetString(3));
                    row.CreateCell(3).SetCellValue(reader.GetString(4));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень поступивших книг за {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write)) {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetNewUsersListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null) {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT surname, name, patronymic, email FROM users WHERE register_date = CURRENT_DATE;", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень новых пользователей");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Фамилия");
                row.CreateCell(1).SetCellValue("Имя");
                row.CreateCell(2).SetCellValue("Отчество");
                row.CreateCell(3).SetCellValue("Почта");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень новых читателей за {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetBooksTakenListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null)
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH taken_books AS (SELECT DISTINCT ON (b.id) b.name AS book_name, array_agg(a.name || ' ' || a.surname || ' ' || a.patronymic) AS author_names, u.name || ' ' || u.surname || ' ' || u.patronymic AS user_fullname, u.email, f.return_date FROM facts f JOIN books b ON f.books @> ARRAY[b.id] JOIN users u ON f.user_id = u.id JOIN authors a ON b.authors @> ARRAY[a.id] WHERE f.taking_date = CURRENT_DATE GROUP BY b.id, u.id, f.return_date) SELECT book_name, array_to_string(author_names, ', ') AS authors, user_fullname, email, return_date FROM taken_books;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень взятых книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Книга");
                row.CreateCell(1).SetCellValue("Авторы");
                row.CreateCell(2).SetCellValue("Кем взята");
                row.CreateCell(3).SetCellValue("Почта взявшего");
                row.CreateCell(4).SetCellValue("Ожидаемая дата возврата");

                while (reader.Read()) {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetDateTime(4).ToString("dd.MM.yyyy"));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень взятых книг за {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetBooksReturnedListButton_Click(object sender, RoutedEventArgs e) {
            if (ReportsPath == null) {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"WITH return_books AS (SELECT DISTINCT ON (b.id) b.name AS book_name, array_agg(a.name || ' ' || a.surname || ' ' || a.patronymic) AS author_names, u.name || ' ' || u.surname || ' ' || u.patronymic AS user_fullname, u.email, (CASE WHEN f.return_date = f.real_return_date::date THEN 'Возвращена вовремя' ELSE 'Возвращена с опозданием' END) AS text_status FROM facts f JOIN books b ON f.books @> ARRAY[b.id] JOIN users u ON f.user_id = u.id JOIN authors a ON b.authors @> ARRAY[a.id] WHERE f.real_return_date = '{DateTime.Now.ToString("dd.MM.yyyy")}' GROUP BY b.id, u.id, f.return_date, f.real_return_date) SELECT book_name, array_to_string(author_names, ', ') AS authors, user_fullname, email, text_status FROM return_books;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень возвращённых книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Книга");
                row.CreateCell(1).SetCellValue("Авторы");
                row.CreateCell(2).SetCellValue("Кем возвращена");
                row.CreateCell(3).SetCellValue("Почта вернувшего");
                row.CreateCell(4).SetCellValue("Статус возврата");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень возвращённых книг за {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetBooksShouldReturnedListButton_Click(object sender, RoutedEventArgs e) {
            if (ReportsPath == null) {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"WITH should_return_books AS (SELECT DISTINCT ON (b.id) b.name AS book_name, array_agg(a.name || ' ' || a.surname || ' ' || a.patronymic) AS author_names, u.name || ' ' || u.surname || ' ' || u.patronymic AS user_fullname, u.email, (CASE WHEN f.real_return_date = ' ' THEN 'Ещё не возвращена' ELSE 'Возвращена' END) AS text_status FROM facts f JOIN books b ON f.books @> ARRAY[b.id] JOIN users u ON f.user_id = u.id JOIN authors a ON b.authors @> ARRAY[a.id] WHERE f.return_date = CURRENT_DATE GROUP BY b.id, u.id, f.real_return_date) SELECT book_name, array_to_string(author_names, ', ') AS authors, user_fullname, email, text_status FROM should_return_books;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Книга");
                row.CreateCell(1).SetCellValue("Авторы");
                row.CreateCell(2).SetCellValue("Кем должна быть возвращена");
                row.CreateCell(3).SetCellValue("Почта возвращающего");
                row.CreateCell(4).SetCellValue("Статус возврата");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень книг, которые должны быть возвращены {DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetMostReadGenresListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null)
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH current_month AS (SELECT date_trunc('month', CURRENT_DATE) AS start_of_month, date_trunc('month', CURRENT_DATE + INTERVAL '1 month') AS end_of_month), book_takings AS (SELECT b.id, unnest(b.genres) AS genre_id FROM facts f JOIN books b ON f.books @> ARRAY[b.id] WHERE f.taking_date >= (SELECT start_of_month FROM current_month) AND f.return_date <= (SELECT end_of_month FROM current_month) ) SELECT ROW_NUMBER() OVER (ORDER BY count(*) DESC) AS place, g.name AS genre_name, COUNT(*) AS book_count FROM book_takings bt JOIN genres g ON bt.genre_id = g.id GROUP BY g.name ORDER BY book_count DESC LIMIT 5;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень самых читаемых жанров");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Место");
                row.CreateCell(1).SetCellValue("Жанр");
                row.CreateCell(2).SetCellValue("Число книг, взятых за месяц");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetInt32(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень самых читаемых жанров за {DateTime.Now.ToString("MMMM")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetMostReadAuthorsListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "") {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH current_month AS (SELECT generate_series(date_trunc('month', CURRENT_DATE), date_trunc('month', CURRENT_DATE + INTERVAL '1 month') - INTERVAL '1 day', '1 day')::DATE AS month_day) SELECT ROW_NUMBER() OVER (ORDER BY count(DISTINCT b.id) DESC) AS place, a.name || ' ' || a.surname || ' ' || a.patronymic AS author, COUNT(DISTINCT b.id) AS book_count FROM facts f JOIN current_month cm ON f.taking_date <= cm.month_day AND f.return_date >= cm.month_day JOIN books b ON f.books @> ARRAY[b.id] JOIN UNNEST(b.authors) ua(author_id) ON TRUE JOIN authors a ON ua.author_id = a.id GROUP BY a.id, a.name, a.surname, a.patronymic ORDER BY book_count DESC LIMIT 5;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень самых читаемых авторов");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Место");
                row.CreateCell(1).SetCellValue("Автор");
                row.CreateCell(2).SetCellValue("Число книг, взятых за месяц");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetInt32(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень самых читаемых авторов за {DateTime.Now.ToString("MMMM")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetMostReadUsersListButton_Click(object sender, RoutedEventArgs e) {
            if (ReportsPath == null)
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"WITH user_books AS (SELECT user_id, ARRAY_AGG(DISTINCT book_id) AS unique_books FROM ( SELECT user_id, unnest(books) AS book_id FROM facts WHERE return_date BETWEEN '01.{DateTime.Now.ToString("MM.yy")}' AND '31.{DateTime.Now.ToString("MM.yy")}') AS unnested_books GROUP BY user_id), ranked_users AS (SELECT ROW_NUMBER() OVER (ORDER BY cardinality(unique_books) DESC) AS place, u.name || ' ' || u.surname || ' ' || u.patronymic AS full_name, u.email, cardinality(unique_books) AS total_books FROM user_books ub JOIN users u ON ub.user_id = u.id ) SELECT place, full_name, email, total_books FROM ranked_users LIMIT 5;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень пользователей");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Место");
                row.CreateCell(1).SetCellValue("Читатель");
                row.CreateCell(2).SetCellValue("Почта читателя");
                row.CreateCell(3).SetCellValue("Число книг, прочитанных за месяц");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetInt32(3));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень самых читающих пользователей за {DateTime.Now.ToString("MMMM")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullBooksListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH book_genres AS (SELECT b.id AS book_id, array_agg(g.name) AS genre_name FROM books b INNER JOIN genres g ON b.genres && ARRAY[g.id] GROUP BY b.id), book_authors AS (SELECT b.id AS book_id, array_agg(a.surname || ' ' || a.name || ' ' || a.patronymic) AS author_fullnames FROM books b INNER JOIN authors a ON b.authors && ARRAY[a.id] GROUP BY b.id) SELECT b.id AS book_id, b.name AS book_name, CASE WHEN b.status = 0 THEN 'Свободна' ELSE 'На руках у ' || COALESCE(u.surname || ' ' || u.name || ' ' || u.patronymic, '') END AS book_status, CASE WHEN b.status != 0 THEN f.return_date::varchar(10) ELSE ' ' END AS return_date, b.register_date AS book_register_date, array_to_string(array_agg(DISTINCT bg.genre_name), ', ') AS genre_names, array_to_string(array_agg(DISTINCT ba.author_fullnames), ', ') AS author_fullnames FROM books b LEFT JOIN book_genres bg ON b.id = bg.book_id LEFT JOIN book_authors ba ON b.id = ba.book_id LEFT JOIN users u ON b.status = u.id LEFT JOIN facts f ON u.id = f.user_id GROUP BY b.id, b.name, b.status, u.surname, u.name, u.patronymic, f.return_date;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Название");
                row.CreateCell(1).SetCellValue("Статус");
                row.CreateCell(2).SetCellValue("Ожидаемая дата возврата");
                row.CreateCell(3).SetCellValue("Дата поступления");
                row.CreateCell(4).SetCellValue("Жанры");
                row.CreateCell(5).SetCellValue("Авторы");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(1));
                    row.CreateCell(1).SetCellValue(reader.GetString(2));
                    row.CreateCell(2).SetCellValue(reader.GetString(3));
                    row.CreateCell(3).SetCellValue(reader.GetDateTime(4).ToString("dd.MM.yyyy"));
                    row.CreateCell(4).SetCellValue(reader.GetString(5));
                    row.CreateCell(5).SetCellValue(reader.GetString(6));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень книг.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullUsersListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH fact_books AS (SELECT user_id, unnest(books) AS book_id FROM facts) SELECT u.surname, u.name, u.patronymic, u.email, u.register_date, COUNT(fb.book_id) AS books_count FROM users u LEFT JOIN fact_books fb ON u.id = fb.user_id WHERE u.id > 1 GROUP BY u.id, u.name, u.surname, u.patronymic, u.email, u.register_date ORDER BY u.id;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень пользователей");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Фамилия");
                row.CreateCell(1).SetCellValue("Имя");
                row.CreateCell(2).SetCellValue("Отчество");
                row.CreateCell(3).SetCellValue("Почта");
                row.CreateCell(4).SetCellValue("Дата регистрации");
                row.CreateCell(5).SetCellValue("Число взятых книг");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetDateTime(4).ToString("dd.MM.yyyy"));
                    row.CreateCell(5).SetCellValue(reader.GetInt16(5));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень читателей.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullAuthorsListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT a.surname, a.name, a.patronymic, COUNT(b.id) AS books_count FROM authors a JOIN books b ON a.id = ANY(b.authors) GROUP BY a.id, a.name, a.surname, a.patronymic ORDER BY a.id;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень авторов");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Фамилия");
                row.CreateCell(1).SetCellValue("Имя");
                row.CreateCell(2).SetCellValue("Отчество");
                row.CreateCell(3).SetCellValue("Число книг");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetInt16(3));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень авторов.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullGenresListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT g.name, COUNT(b.id) AS books_count FROM genres g JOIN books b ON g.id = ANY(b.genres) GROUP BY g.id, g.name ORDER BY g.id;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень жанров");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Название жанра");
                row.CreateCell(1).SetCellValue("Число книг");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetInt16(1));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень жанров.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
    }
}
