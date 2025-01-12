using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class Author
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Patronymic { get; set; }

        public Author(int id, string name, string surname, string patronymic) {
            Id = id;
            Name = name;
            Surname = surname;
            Patronymic = patronymic;
        }
    }
}
