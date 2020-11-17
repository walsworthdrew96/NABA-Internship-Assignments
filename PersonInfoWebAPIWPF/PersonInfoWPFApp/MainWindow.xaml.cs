using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.OleDb;
using System.Data;
using PersonInfoWebAPIWPF.Models;
using System.Net.Http;
using System.Text.Json;
using System.Collections.ObjectModel;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http.Headers;


namespace PersonInfoWPFApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private PersonAccess pa;
        //Set the scope for API call to user.read
        string[] scopes = new string[] { "user.read" };

        private string txt_file_name = "full_name.txt";
        private string xlsx_file_name = "full_name.xlsx";
        private string txt_file_path;
        private string xlsx_file_path;
        private string access_db_file_name;
        private string access_db_path;
        private string msAccessConnectionString;
        private string msSQLServerConnectionString;
        private string azureConnectionString;
        private List<Person> people;
        private CheckBox[] checkboxes;
        public HttpClient client;

        public static string MakeFilePath(string file_name)
        {
            return System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\..\files", file_name);
        }

        public MainWindow()
        {
            InitializeComponent();

            txt_file_path = MakeFilePath(txt_file_name);
            xlsx_file_path = xlsx_file_path = MakeFilePath(xlsx_file_name);

            // Create a PersonAccess object to work with files containing Person objects.
            pa = new PersonAccess();
            checkboxes = new CheckBox[]
            {
                AccessDbCheckBox, SqlDbCheckBox, AzureDbCheckBox
            };

            client = new HttpClient();
        }

        public static void CreateFileIfNotExists(string file_name)
        {
            string file_path = MakeFilePath(file_name);
            if (!System.IO.File.Exists(file_path))
            {
                File.Create(file_path);
            }
        }

        async void GetAll(string url)
        {
            try
            {
                MessageTextBlock.Text = "Processing Request...";
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    using (HttpContent content = response.Content)
                    {
                        if (response.StatusCode.ToString() != "OK")
                        {
                            MessageTextBlock.Text = $"{response.StatusCode}";
                        }
                        //MessageTextBlock.Text = $"response.StatusCode = {response.StatusCode}\n";
                        //MessageTextBlock.Text += $"response.Headers = {response.Headers}\n";
                        //MessageTextBlock.Text += $"response.Content = {response.Content}\n";
                        //MessageTextBlock.Text += $"response.RequestMessage = {response.RequestMessage}\n";
                        //MessageTextBlock.Text += $"response.TrailingHeaders = {response.TrailingHeaders}\n\n";
                        string mycontent = await content.ReadAsStringAsync();
                        List<Person> people = JsonConvert.DeserializeObject<List<Person>>(mycontent);
                        MessageTextBlock.Text = mycontent;
                        MessageTextBlock.Text += "\n";
                        foreach (Person p in people)
                        {
                            MessageTextBlock.Text += $"{p}\n";
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageTextBlock.Text = e.Message;
                return;
            }
            MessageTextBlock.Text += "GET Request Completed Successfully.";
        }

        async void GetById(string url)
        {
            try
            {
                MessageTextBlock.Text = "Processing Request...";
                using (HttpResponseMessage response = await client.GetAsync(url))
                {
                    using (HttpContent content = response.Content)
                    {
                        string mycontent = await content.ReadAsStringAsync();
                        Person p = JsonConvert.DeserializeObject<Person>(mycontent);
                        MessageTextBlock.Text = $"{mycontent}\n";
                        MessageTextBlock.Text += $"{p}\n";
                    }
                }
            }
            catch (Exception e)
            {
                MessageTextBlock.Text = e.Message;
                return;
            }
            MessageTextBlock.Text += "GET Request Completed Successfully.";
        }

        async void PostRequest(string url)
        {
            try
            {
                MessageTextBlock.Text = "Processing Request...\n";
                Person p = new Person
                {
                    FirstName = FirstNameTextBox.Text,
                    LastName = LastNameTextBox.Text 
                };

                string jsonString;
                jsonString = JsonConvert.SerializeObject(p);
                MessageTextBlock.Text += $"jsonString = {jsonString}\n";

                using (HttpResponseMessage response = await client.PostAsync(url, new StringContent(jsonString, Encoding.UTF8, "application/json")))
                {
                    using (HttpContent content = response.Content)
                    {
                        string mycontent = await content.ReadAsStringAsync();
                        HttpContentHeaders headers = content.Headers;
                        MessageTextBlock.Text += $"{mycontent}\n";
                    }
                }
            }
            catch (Exception e)
            {
                MessageTextBlock.Text = e.Message;
                return;
            }
            MessageTextBlock.Text += "POST Request Completed Successfully.";
        }

        async void PutRequest(string url)
        {
            try
            {
                //MessageTextBlock.Text = "Processing Request...\n";
                MessageTextBlock.Text += "Processing Request...\n";
                Person p = new Person
                {
                    Id = Int32.Parse(IdTextBox.Text),
                    FirstName = FirstNameTextBox.Text,
                    LastName = LastNameTextBox.Text
                };

                string jsonString;
                jsonString = JsonConvert.SerializeObject(p);
                MessageTextBlock.Text += $"jsonString = {jsonString}\n";
                StringContent stringContent = new StringContent(jsonString, Encoding.UTF8, "application/json");

                using (HttpResponseMessage response = await client.PutAsync(url, stringContent))
                {
                    using (HttpContent content = response.Content)
                    {
                        string mycontent = await content.ReadAsStringAsync();
                        HttpContentHeaders headers = content.Headers;
                        MessageTextBlock.Text += $"{mycontent}\n";
                    }
                }
            }
            catch (Exception e)
            {
                MessageTextBlock.Text = e.Message;
                return;
            }
            MessageTextBlock.Text += "PUT Request Completed Successfully.";
        }

        async void DeleteRequest(string url)
        {
            try
            {
                //MessageTextBlock.Text = "Processing Request...\n";
                MessageTextBlock.Text += "Processing Request...\n";

                using (HttpResponseMessage response = await client.DeleteAsync(url))
                {
                    using (HttpContent content = response.Content)
                    {
                        string mycontent = await content.ReadAsStringAsync();
                        HttpContentHeaders headers = content.Headers;
                        MessageTextBlock.Text += $"{mycontent}\n";
                    }
                }
            }
            catch (Exception e)
            {
                MessageTextBlock.Text = e.Message;
                return;
            }
            MessageTextBlock.Text += "DELETE Request Completed Successfully.";
        }

        //private async Task<string> PostHTTPRequestAsync(string url, Dictionary<string, string> data)
        //{
        //    using (HttpContent formContent = new FormUrlEncodedContent(data))
        //    {
        //        using (HttpResponseMessage response = await client.PostAsync(url, formContent).ConfigureAwait(false))
        //        {
        //            response.EnsureSuccessStatusCode();
        //            return await response.Content.ReadAsStringAsync().ConfigureAwait(false);
        //        }
        //    }
        //}

        public bool ValidatePersonForm()
        {
            if ((FirstNameTextBox.Text != null && FirstNameTextBox.Text != "")
                && (LastNameTextBox.Text != null && LastNameTextBox.Text != ""))
            {
                return true;
            }
            MessageBox.Show("Please enter a Person's First Name and Last Name.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        public bool ValidatePersonFormWithId()
        {
            if ((IdTextBox.Text != null && IdTextBox.Text != null)
                && (FirstNameTextBox.Text != null && FirstNameTextBox.Text != "")
                && (LastNameTextBox.Text != null && LastNameTextBox.Text != ""))
            {
                return true;
            }
            MessageBox.Show("Please enter a Person's Id, First Name, and Last Name.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        public bool ValidatePersonId()
        {
            if (IdTextBox.Text != null && IdTextBox.Text != "")
            {
                return true;
            }
            MessageBox.Show("Please enter a Person's Id.", "Alert", MessageBoxButton.OK, MessageBoxImage.Information);
            return false;
        }

        // READ FILE BUTTON CLICKS
        private void Read_Button_Click(object sender, RoutedEventArgs e)
        {
            if (TextFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(txt_file_name);
                people = pa.ReadFromTextFile(txt_file_path);

                MessageTextBlock.Text = $"Current \"{txt_file_name}\" Contents:\n";
                foreach (var person in people)
                {
                    MessageTextBlock.Text += $"{person}\n";
                }
            }

            if (ExcelFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(xlsx_file_name);
                people = pa.ReadFromExcelFile(xlsx_file_path);

                if (TextFileCheckBox.IsChecked == true)
                {
                    MessageTextBlock.Text += $"Current \"{xlsx_file_name}\" Contents:\n";
                    foreach (var person in people)
                    {
                        MessageTextBlock.Text += $"{person}\n";
                    }
                }
                else
                {
                    MessageTextBlock.Text = $"Current \"{xlsx_file_name}\" Contents:\n";
                    foreach (var person in people)
                    {
                        MessageTextBlock.Text += $"{person}\n";
                    }
                }
            }
        }

        // WRITE FILE BUTTON CLICKS
        private void Write_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonForm())
            {
                return;
            }

            if (TextFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(txt_file_name);

                List<Person> people = new List<Person>
                {
                    new Person
                    {
                        FirstName = FirstNameTextBox.Text,
                        LastName = LastNameTextBox.Text
                    }
                };

                pa.WriteToTextFile(people, txt_file_path);

                MessageTextBlock.Text = $"Overwrote {txt_file_name}\".\n";
            }

            if (ExcelFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(txt_file_name);

                List<Person> people = new List<Person>
                {
                    new Person
                    {
                        FirstName = FirstNameTextBox.Text,
                        LastName = LastNameTextBox.Text
                    }
                };

                pa.WriteToExcelFile(people, xlsx_file_path);

                if (TextFileCheckBox.IsChecked == true)
                {
                    MessageTextBlock.Text += $"Overwrote \"{xlsx_file_name}\".\n";
                }
                else
                {
                    MessageTextBlock.Text = $"Overwrote \"{xlsx_file_name}\".\n";
                }
            }
            Read_Button_Click(null, null);
        }

        // APPEND FILE BUTTON CLICKS
        private void Append_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonForm())
            {
                return;
            }

            if (TextFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(txt_file_name);

                List<Person> people = new List<Person>
                {
                    new Person
                    {
                        FirstName = FirstNameTextBox.Text,
                        LastName = LastNameTextBox.Text
                    }
                };

                pa.AppendToTextFile(people, txt_file_path);

                MessageTextBlock.Text = $"Appended to \"{txt_file_name}\".\n";
            }

            if (ExcelFileCheckBox.IsChecked == true)
            {
                CreateFileIfNotExists(txt_file_name);

                List<Person> people = new List<Person>
                {
                    new Person
                    {
                        FirstName = FirstNameTextBox.Text,
                        LastName = LastNameTextBox.Text
                    }
                };

                pa.AppendToExcelFile(people, xlsx_file_path);

                if (TextFileCheckBox.IsChecked == true)
                {
                    MessageTextBlock.Text += $"Appended to \"{xlsx_file_name}\".\n";
                }
                else
                {
                    MessageTextBlock.Text = $"Appended to \"{xlsx_file_name}\".\n";
                }
            }
            Read_Button_Click(null, null);
        }
        private void Select_All_Button_Click(object sender, RoutedEventArgs e)
        {
            if (AccessDbCheckBox.IsChecked == true)
            {
                GetAll($"http://localhost:5000/api/person?dbSelection=msAccessConnection");
            }
            if (SqlDbCheckBox.IsChecked == true)
            {
                GetAll($"http://localhost:5000/api/person?dbSelection=msSqlConnection");
            }
            if (AzureDbCheckBox.IsChecked == true)
            {
                GetAll($"http://localhost:5000/api/person?dbSelection=azureSqlConnection");
            }
        }

        private void Select_By_Id_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonId())
            {
                return;
            }
            if (AccessDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                GetById($"{url}{IdTextBox.Text}?dbSelection=msAccessConnection");
            }
            if (SqlDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                GetById($"{url}{IdTextBox.Text}?dbSelection=msSqlConnection");
            }
            if (AzureDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                GetById($"{url}{IdTextBox.Text}?dbSelection=azureSqlConnection");
            }
        }

        private void Insert_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonFormWithId())
            {
                return;
            }
            if (AccessDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person";
                PostRequest($"{url}?dbSelection=msAccessConnection");
            }
            if (SqlDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person";
                PostRequest($"{url}?dbSelection=msSqlConnection");
            }
            if (AzureDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person";
                PostRequest($"{url}?dbSelection=azureSqlConnection");
            }
        }
        private void Update_By_Id_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonFormWithId())
            {
                return;
            }
            if (AccessDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                PutRequest($"{url}{IdTextBox.Text}?dbSelection=msAccessConnection");
            }
            if (SqlDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                PutRequest($"{url}{IdTextBox.Text}?dbSelection=msSqlConnection");
            }
            if (AzureDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                PutRequest($"{url}{IdTextBox.Text}?dbSelection=azureSqlConnection");
            }
        }
        private void Delete_By_Id_Button_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidatePersonFormWithId())
            {
                return;
            }
            if (AccessDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                DeleteRequest($"{url}{IdTextBox.Text}?dbSelection=msAccessConnection");
            }
            if (SqlDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                DeleteRequest($"{url}{IdTextBox.Text}?dbSelection=msSqlConnection");
            }
            if (AzureDbCheckBox.IsChecked == true)
            {
                MessageTextBlock.Text = "\n";
                string url = "http://localhost:5000/api/person/";
                MessageTextBlock.Text += $"url:{url}";
                DeleteRequest($"{url}{IdTextBox.Text}?dbSelection=azureSqlConnection");
            }
        }

        private void Delete_All_Button_Click(object sender, RoutedEventArgs e)
        {
            //string message = "";
            //if (AccessDbCheckBox.IsChecked == true)
            //{
            //    message = "";
            //}
            //if (SqlDbCheckBox.IsChecked == true)
            //{
            //    message = "";
            //}
            //if (AzureDbCheckBox.IsChecked == true)
            //{
            //    message = "";
            //}
        }

        private void FirstName_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void LastName_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextFile_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void ExcelFile_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void AccessDB_Checked(object sender, RoutedEventArgs e)
        {
            int true_index = Array.IndexOf(checkboxes, sender as CheckBox);
            for (int i = 0; i < checkboxes.Length; i++)
            {
                if (i != true_index)
                {
                    checkboxes[i].IsChecked = false;
                }
            }
        }

        private void SQLServerDB_Checked(object sender, RoutedEventArgs e)
        {
            int true_index = Array.IndexOf(checkboxes, sender as CheckBox);
            for (int i = 0; i < checkboxes.Length; i++)
            {
                if (i != true_index)
                {
                    checkboxes[i].IsChecked = false;
                }
            }
        }

        private void AzureSQLDB_Checked(object sender, RoutedEventArgs e)
        {
            int true_index = Array.IndexOf(checkboxes, sender as CheckBox);
            for (int i = 0; i < checkboxes.Length; i++)
            {
                if (i != true_index)
                {
                    checkboxes[i].IsChecked = false;
                }
            }
        }

        private void MessageText_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void SQLDb_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void AzureDb_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void IdTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void FirstNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void LastNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TempDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void SavedDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
