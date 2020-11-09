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
using Azure.Identity;
using Microsoft.Identity.Client;

namespace PersonInfoWPFApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }

        //Set the scope for API call to user.read
        string[] scopes = new string[] { "user.read" };

        private string access_db_file_name;
        private string access_db_path;
        private string msAccessConnectionString;
        private string msSQLServerConnectionString;
        private string azureConnectionString;

        public static void CreateFileIfNotExists(string file_path, string file_name)
        {
            if (!System.IO.File.Exists(file_path))
            {
                using (System.IO.FileStream fs = System.IO.File.Create(file_path))
                {
                }
            }
        }

        public MainWindow()
        {
            //file paths
            access_db_file_name = "naba_db.accdb";
            access_db_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", access_db_file_name);
            //connection strings
            msAccessConnectionString = $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={access_db_path};Persist Security Info=False;";
            msSQLServerConnectionString = @"Data Source=localhost;Initial Catalog=NABA;Integrated Security=True";
            azureConnectionString = "Server=tcp:naba-server.database.windows.net,1433;Initial Catalog=naba-db;Persist Security Info=False;User ID=naba-server-admin;Password=rz0f-396v-xr54;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

            InitializeComponent();
        }

        //private void MSAccessQuery(string sql)
        //{
        //    messageTextBox.Text = "";

        //    OleDbConnection connection = new OleDbConnection();
        //    connection.ConnectionString = msAccessConnectionString;
        //    connection.Open();
        //    messageTextBox.Text += "MS Access DB Connection Open!\n";

        //    OleDbCommand command = new OleDbCommand();
        //    command.Connection = connection;
        //    command.CommandText = sql;
        //    command.ExecuteNonQuery();

        //    OleDbDataAdapter da = new OleDbDataAdapter(command);
        //    DataTable dt = new DataTable();
        //    da.Fill(dt);
        //    DataRowCollection rows = dt.Rows;
        //    messageTextBox.Text += $"Data from {access_db_file_name}:\n";
        //    foreach (DataRow row in rows)
        //    {
        //        int count = 0;
        //        foreach (var item in row.ItemArray)
        //        {
        //            messageTextBox.Text += item.ToString();
        //            if (count != row.ItemArray.Length)
        //            {
        //                messageTextBox.Text += " ";
        //            }
        //            count += 1;
        //            messageTextBox.Text += "\n";
        //        }
        //    }

        //    connection.Close();
        //    messageTextBox.Text += "MS Access DB Connection Closed!\n";
        //}

        // READ FILE BUTTON CLICKS
        private void Read_File_Button_Click(object sender, RoutedEventArgs e)
        {
            // Clear Text Box Contents
            messageTextBox.Text = "";

            if (textFileCheckBox.IsChecked == true)
            {
                PersonInfo pi = new PersonInfo();
                string txt_file_name = "full_name.txt";
                string txt_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", txt_file_name);
                CreateFileIfNotExists(txt_file_path, txt_file_name);
                pi.DisplayNameFromTextFile(messageTextBox, txt_file_path);
            }

            if (excelFileCheckBox.IsChecked == true)
            {
                if (textFileCheckBox.IsChecked == true)
                {
                    messageTextBox.Text += "\n";
                }

                PersonInfo pi = new PersonInfo();
                string xlsx_file_name = "full_name.xlsx";
                string xlsx_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", xlsx_file_name);
                Console.WriteLine("xlsx_file_path: " + xlsx_file_path);
                CreateFileIfNotExists(xlsx_file_path, xlsx_file_name);
                pi.DisplayNameFromExcelFile(messageTextBox, xlsx_file_path);
            }
        }

        // WRITE FILE BUTTON CLICKS
        private void Write_File_Button_Click(object sender, RoutedEventArgs e)
        {
            // Clear Text Box Contents
            messageTextBox.Text = "";

            if (textFileCheckBox.IsChecked == true)
            {
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                string txt_file_name = "full_name.txt";
                string txt_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", txt_file_name);
                CreateFileIfNotExists(txt_file_path, txt_file_path);
                pi.WriteToTextFile(txt_file_path);
                messageTextBox.Text += $"Overwrote \"{pi}\" to \"{txt_file_name}\".";
            }

            if (excelFileCheckBox.IsChecked == true)
            {
                if (textFileCheckBox.IsChecked == true)
                {
                    messageTextBox.Text += "\n";
                }
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                string xlsx_file_name = "full_name.xlsx";
                string xlsx_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", xlsx_file_name);
                CreateFileIfNotExists(xlsx_file_path, xlsx_file_name);
                pi.WriteToExcelFile(xlsx_file_path);
                messageTextBox.Text += $"Overwrote \"{pi}\" to \"{xlsx_file_name}\".";
            }
        }

        // APPEND FILE BUTTON CLICKS
        private void Append_File_Button_Click(object sender, RoutedEventArgs e)
        {
            // Clear Text Box Contents
            messageTextBox.Text = "";

            if (textFileCheckBox.IsChecked == true)
            {
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                string txt_file_name = "full_name.txt";
                string txt_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", txt_file_name);
                CreateFileIfNotExists(txt_file_path, txt_file_path);
                pi.AppendToTextFile(txt_file_path);
                messageTextBox.Text += $"Appended \"{pi}\" to \"{txt_file_name}\".";
            }

            if (excelFileCheckBox.IsChecked == true)
            {
                if (textFileCheckBox.IsChecked == true)
                {
                    messageTextBox.Text += "\n";
                }
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                string xlsx_file_name = "full_name.xlsx";
                string xlsx_file_path = System.IO.Path.Combine(Environment.CurrentDirectory, @"..\..\", xlsx_file_name);
                CreateFileIfNotExists(xlsx_file_path, xlsx_file_name);
                pi.AppendToExcelFile(xlsx_file_path);
                messageTextBox.Text += $"Appended \"{pi}\" to \"{xlsx_file_name}\".";
            }
        }

        // DB BUTTON CLICKS

        private void SelectAll_Button_Click(object sender, RoutedEventArgs e)
        {
            if (msAccessDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = msAccessConnectionString;
                connection.Open();
                messageTextBox.Text += "MS Access DB Connection Open!\n";

                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                string sql = $"SELECT * FROM Person;";
                command.CommandText = sql;
                command.ExecuteNonQuery();

                OleDbDataAdapter da = new OleDbDataAdapter(command);
                DataTable dt = new DataTable();
                da.Fill(dt);

                //display data
                messageTextBox.Text += $"Data from {access_db_file_name}:\n";
                DataRowCollection rows = dt.Rows;
                foreach (DataRow row in rows)
                {
                    messageTextBox.Text += $"{row["FirstName"]} {row["LastName"]}\n";
                }

                connection.Close();
                messageTextBox.Text += "MS Access DB Connection Closed!\n";
            }
            if (sqlServerDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(msSQLServerConnectionString);
                cnn.Open();

                messageTextBox.Text += "MS SQL Server Connection Open!\n";
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                messageTextBox.Text += $"PEOPLE IN DATABASE:\n{pi.SelectAll(cnn)}";

                cnn.Close();
                messageTextBox.Text += "MS SQL Server Connection Closed!\n";
            }
            if (azureSQLDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(azureConnectionString);
                cnn.Open();

                messageTextBox.Text += "Azure SQL Server Connection Open!\n";
                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                
                messageTextBox.Text += $"PEOPLE IN DATABASE:\n{pi.SelectAll(cnn)}";

                cnn.Close();
                messageTextBox.Text += "Azure SQL Server Connection Closed!\n";
            }
        }

        private void Insert_Button_Click(object sender, RoutedEventArgs e)
        {
            if (msAccessDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = msAccessConnectionString;
                connection.Open();
                messageTextBox.Text += "MS Access DB Connection Open!\n";

                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"INSERT INTO Person (FirstName, LastName) VALUES ('{firstNameTextBox.Text}', '{lastNameTextBox.Text}');";
                command.ExecuteNonQuery();
                messageTextBox.Text += $"Inserted {pi} into {access_db_file_name}.\n";

                connection.Close();
                messageTextBox.Text += "MS Access DB Connection Closed!\n";
            }
            if (sqlServerDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(msSQLServerConnectionString);
                cnn.Open();
                messageTextBox.Text += "MS SQL Server Connection Open!\n";

                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                pi.Insert(cnn);
                messageTextBox.Text += $"INSERTED PERSON: {pi.FullName}\n";

                cnn.Close();
                messageTextBox.Text += "MS SQL Server Connection Closed!\n";
            }
            if (azureSQLDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(azureConnectionString);
                cnn.Open();
                messageTextBox.Text += "Azure SQL Server Connection Open!\n";

                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                pi.Insert(cnn);
                messageTextBox.Text += $"INSERTED PERSON: {pi.FullName}\n";

                cnn.Close();
                messageTextBox.Text += "Azure SQL Server Connection Closed!\n";
            }
        }

        private void Delete_All_Button_Click(object sender, RoutedEventArgs e)
        {
            if (msAccessDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = msAccessConnectionString;
                connection.Open();
                messageTextBox.Text += "MS Access DB Connection Open!\n";

                OleDbCommand command = new OleDbCommand();
                command.Connection = connection;
                command.CommandText = $"DELETE FROM Person;";
                command.ExecuteNonQuery();
                messageTextBox.Text += "All Persons deleted from Person table.\n";

                connection.Close();
                messageTextBox.Text += "MS Access DB Connection Closed!\n";
            }
            if (sqlServerDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(msSQLServerConnectionString);
                cnn.Open();
                messageTextBox.Text += "MS SQL Server Connection Open!\n";

                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                pi.Delete_All(cnn);
                messageTextBox.Text += $"DELETED ALL ROWS FROM THE PERSON TABLE.\n";

                cnn.Close();
                messageTextBox.Text += "MS SQL Server Connection Closed!\n";
            }
            if (azureSQLDBCheckBox.IsChecked == true)
            {
                messageTextBox.Text = "";

                SqlConnection cnn = new SqlConnection(azureConnectionString);
                cnn.Open();
                messageTextBox.Text += "Azure SQL Server Connection Open!\n";

                PersonInfo pi = new PersonInfo(firstNameTextBox.Text, lastNameTextBox.Text);
                pi.Delete_All(cnn);
                messageTextBox.Text += $"DELETED ALL ROWS FROM THE PERSON TABLE.\n";

                cnn.Close();
                messageTextBox.Text += "Azure SQL Server Connection Closed!\n";
            }
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

        }

        private void SQLServerDB_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void AzureSQLDB_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void MessageText_TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
