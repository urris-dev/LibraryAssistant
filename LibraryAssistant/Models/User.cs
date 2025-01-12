using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public string Patronymic { get; set; }
        public string Email { get; set; }
        public string RegisterDate {  get; set; }

        public User(int  id, string name, string surname, string patronymic, string email, string registerDate = "11.11.1111")
        {
            Id = id;
            Name = name;
            Surname = surname;
            Patronymic = patronymic;
            Email = email;
            RegisterDate = registerDate;
        }
    }
}
