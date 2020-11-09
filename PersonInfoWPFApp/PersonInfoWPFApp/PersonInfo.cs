using OfficeOpenXml;
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
    class PersonInfo
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        protected int ID { get; set; }

        public string FullName
        {
            get
            {
                return $"{FirstName} {LastName}";
            }
        }

        // 1. Take first and last name information:
        public PersonInfo()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public PersonInfo(string firstName, string lastName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.FirstName = firstName;
            this.LastName = lastName;
        }

        public override string ToString()
        {
            return FullName;
        }

        public void DisplayName(TextBox messageTextBox)
        {
            messageTextBox.Text = $"Name from Class: {FullName}";
        }

        public List<string> ReadFromTextFile(string filePath)
        {
            return File.ReadAllLines(filePath).ToList();
        }

        public void WriteToTextFile(string filePath)
        {
            List<string> lines = new List<string>
            {
                $"{FullName}"
            };
            File.WriteAllLines(filePath, lines);
        }

        public void AppendToTextFile(string filePath)
        {
            List<string> fileLines = ReadFromTextFile(filePath);
            fileLines.Add(FullName);
            File.WriteAllLines(filePath, fileLines);
        }

        public void DisplayNameFromTextFile(TextBox messageTextBox, string filePath)
        {
            List<string> lines = ReadFromTextFile(filePath);
            if (lines == null || lines.Count == 0)
            {
                messageTextBox.Text += "Text file is empty.";
                return;
            }
            messageTextBox.Text += $"Names from \"{filePath.Split('\\').Last()}\":\n";
            foreach (string line in lines)
            {
                messageTextBox.Text += $"{line}\n";
            }
        }

        public Dictionary<string, List<string>> ReadFromExcelFile(string filePath)
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
                    Console.WriteLine("Worksheet is empty.\n");
                    return null;
                }

                // create a dictionary with column count matching the excel file
                Dictionary<string, List<string>> excelDict = new Dictionary<string, List<string>>();
                
                // assign each column header key a data list as a value
                for (int col = start.Column; col <= end.Column; col++)
                {
                    string current_header_key = worksheet.Cells[1, col].Value.ToString();
                    Console.WriteLine($"current_header_key: {current_header_key}");
                    excelDict[current_header_key] = new List<string>();
                }

                // add each worksheet cell as data in the dict.
                for (int row = start.Row+1; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        string current_header_key = worksheet.Cells[1, col].Value.ToString();
                        string current_data_value = worksheet.Cells[row, col].Value.ToString();
                        Console.WriteLine($"current_data_value: {current_data_value}");
                        excelDict[current_header_key].Add(current_data_value);
                    }
                }

                return excelDict;
            }
        }

        public void WriteToExcelFile(string filePath)
        {
            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                // Create worksheet if it doesn't exist.
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");
                }

                worksheet.Cells["A1"].Value = "First Name";
                worksheet.Cells["B1"].Value = "Last Name";
                worksheet.Cells["A2"].Value = FirstName;
                worksheet.Cells["B2"].Value = LastName;

                package.SaveAs(excel_file);
            }
        }

        public void AppendToExcelFile(string filePath)
        {
            var excelDict = ReadFromExcelFile(filePath);
            if (excelDict == null)
            {
                return;
            }
            // Display information from dict
            foreach (var kvp in excelDict)
            {
                Console.WriteLine($"{kvp.Key}: {kvp.Value}");
            }

            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");
                }

                int col_index = 0;
                int row_index = 0;
                // write header information
                col_index = 0;
                foreach (KeyValuePair<string, List<string>> col_kvp in excelDict)
                {
                    col_index += 1;
                    Console.WriteLine($"col_kvp.Key: {col_kvp.Key}");
                    worksheet.Cells[1, col_index].Value = col_kvp.Key;
                }

                // write data information
                col_index = 0;
                // foreach col
                foreach (KeyValuePair<string, List<string>> col in excelDict)
                {
                    col_index += 1;
                    // foreach row in each col
                    // start from row 2 (first data row, after header row)
                    row_index = 1;
                    List<string> col_data = col.Value;
                    foreach (string cell in col_data)
                    {
                        row_index += 1;
                        string col_item = col_data[row_index - 1 - 1];
                        Console.WriteLine($"col_item: {col_item}");
                        // col_data row index is 0, so -1-1 because row index is 1-based worksheet is 2 (2nd row)
                        worksheet.Cells[row_index, col_index].Value = col_item;
                    }
                }

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                // write contents back to excel file.
                worksheet.Cells[end.Row + 1, 1].Value = FirstName;
                worksheet.Cells[end.Row + 1, 2].Value = LastName;

                package.SaveAs(excel_file);
            }
        }

        public void DisplayNameFromExcelFile(TextBox messageTextBox, string filePath)
        {
            var excelDict = ReadFromExcelFile(filePath);

            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                // Create worksheet if it doesn't exist.
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");
                }
                
                var start = worksheet?.Dimension?.Start;
                var end = worksheet?.Dimension?.End;
                if (end == null)
                {
                    messageTextBox.Text += "Worksheet is empty.\n";
                    return;
                }

                //Determine max width for each column
                int[] maxWidthPerCol = new int[end.Column];

                //Note: The worksheet array is 1-based, while the maxWidthPerCol array is 0-based.
                //Console.WriteLine($"{start.Column} to {end.Column}");
                //Console.WriteLine($"{start.Row} to {end.Row}");
                for (int col = start.Column; col <= end.Column; col++)
                {
                    for (int row = start.Row; row <= end.Row; row++)
                    {
                        object cellValue = worksheet.Cells[row, col].Value;
                        if (cellValue.ToString().Length > maxWidthPerCol[col - 1])
                        {
                            maxWidthPerCol[col - 1] = cellValue.ToString().Length;
                        }
                    }
                }

                // Calculate all columns character width
                int all_cols_width = 0;
                foreach (int col_width in maxWidthPerCol)
                {
                    all_cols_width += col_width;
                }

                messageTextBox.Text += "Names from Excel File:\n";
                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        object cellValue = worksheet.Cells[row, col].Value;
                        messageTextBox.Text += String.Format($"{cellValue.ToString().PadRight(maxWidthPerCol[col - 1])}");
                        //messageTextBox.Text += String.Format("{0,"+(-maxWidthPerCol[col - 1]).ToString()+"}", cellValue);
                        if (col != end.Column)
                        {
                            messageTextBox.Text += " | ";
                        }
                        else
                        {
                            messageTextBox.Text += "\n";
                            if (row == start.Row)
                            {
                                messageTextBox.Text += $"{"".PadRight(all_cols_width, '-')}\n";
                            }
                        }
                    }
                }
            }
        }

        public string SelectAll(SqlConnection cnn)
        {
            SqlCommand command;
            SqlDataReader dataReader;
            string sql = "SELECT FirstName, LastName FROM Person;";
            string Output = "";
            command = new SqlCommand(sql, cnn);

            dataReader = command.ExecuteReader();
            while (dataReader.Read())
            {
                Console.WriteLine($"dataReader: {dataReader}");
                Output += $"{dataReader.GetValue(0)} | {dataReader.GetValue(1)}\n";
            }

            dataReader.Close();
            command.Dispose();

            return Output;
        }

        public void Insert(SqlConnection cnn)
        {
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            string sql = $"INSERT INTO Person(FirstName, LastName) VALUES('{FirstName}', '{LastName}')";
            command = new SqlCommand(sql, cnn);

            adapter.InsertCommand = new SqlCommand(sql, cnn);
            adapter.InsertCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public void Update(SqlConnection cnn)
        {
            //UPDATE QUERY
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            string sql = $"UPDATE Person SET FirstName='{FirstName}', LastName='{LastName}', ID={ID}";
            command = new SqlCommand(sql, cnn);

            adapter.UpdateCommand = new SqlCommand(sql, cnn);
            adapter.UpdateCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public void Delete(SqlConnection cnn)
        {
            //DELETE QUERY
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            string sql = $"DELETE FROM Person WHERE ID='{ID}'";
            command = new SqlCommand(sql, cnn);

            adapter.UpdateCommand = new SqlCommand(sql, cnn);
            adapter.UpdateCommand.ExecuteNonQuery();

            command.Dispose();
        }

        public void Delete_All(SqlConnection cnn)
        {
            //DELETE QUERY
            SqlCommand command;
            SqlDataAdapter adapter = new SqlDataAdapter();
            string sql = $"DELETE FROM Person";
            command = new SqlCommand(sql, cnn);

            adapter.UpdateCommand = new SqlCommand(sql, cnn);
            adapter.UpdateCommand.ExecuteNonQuery();

            command.Dispose();
        }
    }
}
