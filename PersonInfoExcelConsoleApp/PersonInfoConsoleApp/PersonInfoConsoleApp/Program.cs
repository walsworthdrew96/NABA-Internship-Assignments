using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
            string txt_file_name = "full_name.txt";
            string txt_file_path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\", txt_file_name);
            PersonInfo pi = new PersonInfo();
            pi.DisplayName();
            pi.WriteToTextFile(txt_file_path);
            pi.DisplayNameFromTextFile(txt_file_path);


            string xlsx_file_name = "full_name.xlsx";
            string xlsx_file_path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\", xlsx_file_name);
            //Console.WriteLine("xlsx_file_path: " + xlsx_file_path);
            CreateFileIfNotExists(xlsx_file_path, xlsx_file_name);
            pi.WriteToExcelFile(xlsx_file_path);
            pi.DisplayNameFromExcelFile(xlsx_file_path);
        }
    }
}
