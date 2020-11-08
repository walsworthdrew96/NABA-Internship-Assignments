using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Common;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;

namespace PersonInfoConsoleApp
{
    class Program
    {
        // Check if file exists, if not create it.
        static void CreateFileIfNotExists(string file_path, string file_name)
        {
            if (!System.IO.File.Exists(file_path))
            {
                using (System.IO.FileStream fs = System.IO.File.Create(file_path))
                {
                }
            }
            else
            {
                //Console.WriteLine($"File {file_name} already exists.");
                return;
            }
        }
        static void Main(string[] args)
        {
            // Write to and Display from a Text File.
            string txt_file_name = "full_name.txt";
            string txt_file_path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\", txt_file_name);
            PersonInfo pi = new PersonInfo();
            pi.DisplayName();
            pi.WriteToTextFile(txt_file_path);
            pi.DisplayNameFromTextFile(txt_file_path);


            // Write to and Display from an Excel File.
            string xlsx_file_name = "full_name.xlsx";
            string xlsx_file_path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\", xlsx_file_name);
            //Console.WriteLine("xlsx_file_path: " + xlsx_file_path);
            CreateFileIfNotExists(xlsx_file_path, xlsx_file_name);
            pi.WriteToExcelFile(xlsx_file_path);
            pi.DisplayNameFromExcelFile(xlsx_file_path);


            // Create a SqlConnection object to open a connection with the database.
            string connectionString;
            SqlConnection cnn;
            connectionString = @"Data Source=localhost;Initial Catalog=NABA;Integrated Security=True";
            cnn = new SqlConnection(connectionString);
            cnn.Open();
            Console.WriteLine("Connection Open!");

            // INSERT INTO PERSON (this Person object)
            pi.Insert(cnn);
            Console.WriteLine($"INSERTED PERSON: {pi.FullName}");

            // SELECT * FROM PERSON;
            Console.WriteLine($"PEOPLE IN DATABASE:\n{pi.SelectAll(cnn)}");

            cnn.Close();
            Console.WriteLine("Connection Closed!");
        }
    }
}
