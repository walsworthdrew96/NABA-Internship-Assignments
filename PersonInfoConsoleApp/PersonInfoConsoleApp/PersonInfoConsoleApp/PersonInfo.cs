using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace PersonInfoConsoleApp
{
    class PersonInfo
    {
        private readonly string firstName;
        private readonly string lastName;

        // 1. Take first and last name information:
        public PersonInfo()
        {
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
        public void DisplayNameFromFile(string filePath)
        {
            List<string> lines = File.ReadAllLines(filePath).ToList();

            Console.Write("Name from File: ");
            foreach (string line in lines)
            {
                Console.WriteLine(line);
            }
        }
    }
}
