using OfficeOpenXml;
using PersonInfoWebAPIWPF.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net.Mime;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

namespace PersonInfoWPFApp
{
    class PersonAccess
    {
        public PersonAccess()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public List<int?> GetIds(IEnumerable<Person> people)
        {
            List<int?> ids = new List<int?>();
            foreach (Person p in people)
            {
                ids.Add(p.Id);
            }
            return ids;
        }

        public List<string> GetFirstNames(IEnumerable<Person> people)
        {
            List<string> firstNames = new List<string>();
            foreach (Person p in people)
            {
                firstNames.Add(p.FirstName);
            }
            return firstNames;
        }

        public List<string> GetLastNames(IEnumerable<Person> people)
        {
            List<string> lastNames = new List<string>();
            foreach (Person p in people)
            {
                lastNames.Add(p.LastName);
            }
            return lastNames;
        }

        public List<Person> ReadFromTextFile(string filePath)
        {
            List<string> lines = File.ReadAllLines(filePath).ToList();
            List<Person> peopleFromFile = new List<Person>();
            foreach (string line in lines)
            {
                List<string> pieces = line.Trim().Split(' ').ToList();
                peopleFromFile.Add(new Person
                {
                    FirstName = pieces[0],
                    LastName = pieces[1]
                });
            }
            return peopleFromFile;
        }

        public void WriteToTextFile(List<Person> peopleToWrite, string filePath)
        {
            List<string> linesToWrite = new List<string>();
            foreach (Person p in peopleToWrite)
            {
                linesToWrite.Add(p.ToString());
            }
            File.WriteAllLines(filePath, linesToWrite);
        }

        public void AppendToTextFile(List<Person> peopleToAppend, string filePath)
        {
            List<string> linesToAppend = new List<string>();
            foreach (Person p in peopleToAppend)
            {
                linesToAppend.Add(p.ToString());
            }
            File.AppendAllLines(filePath, linesToAppend);
        }

        public List<Person> ReadFromExcelFile(string filePath)
        {
            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

                // Create worksheet if it doesn't exist.
                if (worksheet == null)
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                    worksheet = package.Workbook.Worksheets["Sheet1"];
                }

                // Exit method if there is no content in the worksheet.
                var start = worksheet?.Dimension?.Start;
                var end = worksheet?.Dimension?.End;
                if (end == null)
                {
                    return null;
                }

                // Create a dictionary with column count matching the excel file.
                Dictionary<string, List<string>> excelDict = new Dictionary<string, List<string>>();
                
                // Assign each column header's value the column list as a key value pair.
                for (int col = start.Column; col <= end.Column; col++)
                {
                    string column_header = worksheet.Cells[1, col].Value.ToString();
                    excelDict[column_header] = new List<string>();
                }

                // Add each worksheet cell as data in the dictionary.
                for (int row = start.Row+1; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        string current_header_key = worksheet.Cells[1, col].Value.ToString();
                        string current_data_value = worksheet.Cells[row, col].Value.ToString();
                        excelDict[current_header_key].Add(current_data_value);
                    }
                }

                List<Person> people = new List<Person>();

                // Initialize list to the number of entries in the dictionary.
                for (int i = 0; i < excelDict["First Name"].Count; i++)
                {
                    people.Add(new Person());
                }

                // For each column/column list entry:
                foreach (KeyValuePair<string, List<string>> entry in excelDict)
                {
                    string header = entry.Key;
                    List<string> dataList = entry.Value;

                    // For each value in the column data list assign the field of a person based on the header.
                    int i = 0;
                    foreach (string value in dataList)
                    {
                        if (header == "First Name")
                        {
                            people[i].FirstName = value;
                        }
                        if (header == "Last Name")
                        {
                            people[i].LastName = value;
                        }
                        i += 1;
                    }
                }

                return people;
            }
        }

        public void WriteToExcelFile(List<Person> people, string filePath)
        {

            FileInfo excel_file = new FileInfo(filePath);

            using var package = new ExcelPackage(excel_file);

            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

            // Create worksheet if it doesn't exist.
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("Sheet1");
            }

            // Write Header values.
            worksheet.Cells["A1"].Value = "First Name";
            worksheet.Cells["B1"].Value = "Last Name";

            // Write Data values.
            int r = 1;
            foreach (Person p in people)
            {
                r += 1;
                worksheet.Cells[$"A{r}"].Value = p.FirstName;
                worksheet.Cells[$"B{r}"].Value = p.LastName;
            }

            package.SaveAs(excel_file);
        }

        public void AppendToExcelFile(List<Person> people, string filePath)
        { 
            List<Person> peopleFromFile = ReadFromExcelFile(filePath);
            for (int i = 0; i < people.Count; i++)
            {
                peopleFromFile.Add(people[i]);
            }
            WriteToExcelFile(peopleFromFile, filePath);
        }
    }
}
