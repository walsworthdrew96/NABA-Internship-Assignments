using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PersonInfoWebAPIWPF.Models
{
    public class Person
    {
        public int? Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }

        public override string ToString()
        {
            if (Id != null)
            {
                return $"{Id} {FirstName} {LastName}";
            }
            return $"{FirstName} {LastName}";
        }
    }
}
