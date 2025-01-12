using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class Genre
    {
        public int Id { get; set; }
        public string Name { get; set; }
        
        public Genre(int id, string name) {
            Id = id;
            Name = name;
        }
    }
}
