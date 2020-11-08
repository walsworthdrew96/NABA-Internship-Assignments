using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PersonInfoConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "full_name.txt";
            string path = Path.Combine(Environment.CurrentDirectory, @"..\..\..\", fileName);
            PersonInfo pi = new PersonInfo();
            pi.DisplayName();
            pi.WriteToTextFile(path);
            pi.DisplayNameFromFile(path);
        }
    }
}
