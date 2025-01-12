using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class Book
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int[] Genres { get; set; }
        public int GenresCount { get; set; }
        public int[] Authors { get; set; }
        public int AuthorsCount { get; set; }
        public int Status { get; set; }
        public string TextStatus { get; set; }

        public Book(int id, string name, int[] genres, int genresCount, int[] authors, int authorsCount, int status, string textStatus) { 
            Id = id;
            Name = name;
            Genres = genres;
            GenresCount = genresCount;
            Authors = authors;
            AuthorsCount = authorsCount;
            Status = status;
            TextStatus = textStatus;
        }
    }
}
