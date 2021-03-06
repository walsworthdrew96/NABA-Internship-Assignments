﻿using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mime;

namespace PersonInfoConsoleApp
{
    class PersonInfo
    {
        private readonly string firstName;
        private readonly string lastName;

        // 1. Take first and last name information:
        public PersonInfo()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Console.Write("Enter your First Name: ");
            firstName = Console.ReadLine();
            Console.Write("Enter your Last Name: ");
            lastName = Console.ReadLine();
        }

        // 2. Present firstName and lastName to the screen:
        public void DisplayName()
        {
            Console.WriteLine($"Name from Class: {firstName} {lastName}");
        }

        // 3. Write the information to a text file:
        public void WriteToTextFile(string filePath)
        {
            List<string> lines = new List<string>
            {
                $"{firstName} {lastName}"
            };
            File.WriteAllLines(filePath, lines);
        }

        // 4. Read the data from the text file and present it back to the screen
        public void DisplayNameFromTextFile(string filePath)
        {
            List<string> lines = File.ReadAllLines(filePath).ToList();

            Console.Write("Name from File: ");
            foreach (string line in lines)
            {
                Console.WriteLine(line);
            }
        }

        public void WriteToExcelFile(string filePath)
        {
            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Sheet1");
                }

                worksheet.Cells["A1"].Value = "First Name";
                worksheet.Cells["B1"].Value = "Last Name";
                worksheet.Cells["A2"].Value = firstName;
                worksheet.Cells["B2"].Value = lastName;

                package.SaveAs(excel_file);
            }
        }

        

        public void DisplayNameFromExcelFile(string filePath)
        {
            FileInfo excel_file = new FileInfo(filePath);
            using (var package = new ExcelPackage(excel_file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;


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
                        if (cellValue.ToString().Length > maxWidthPerCol[col-1])
                        {
                            maxWidthPerCol[col-1] = cellValue.ToString().Length;
                        }
                    }
                }

                Console.WriteLine("Name from Excel File:");
                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        object cellValue = worksheet.Cells[row, col].Value;
                        Console.Write(String.Format("{0,"+(-maxWidthPerCol[col - 1]).ToString()+"}", cellValue));
                        if (col != end.Column)
                        {
                            Console.Write(" ");
                        }
                        else
                        {
                            Console.Write("\n");
                        }
                    }
                }
            }
        }
    }
}
