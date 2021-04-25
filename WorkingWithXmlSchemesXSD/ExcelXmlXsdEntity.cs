using System;
using System.Collections.Generic;
using System.Text;

namespace WorkingWithXmlSchemesXSD
{
    public class Author
    {
        public byte Yas { get; set; }
        public string Name { get; set; }
        public char Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string PlaceOfBirth { get; set; }
        public string Address { get; set; }
    }

    public class Book 
    {
        public string Title { get; set; }
        public DateTime ReleaseDate { get; set; }
        public Decimal Price { get; set; }
        public Author author { get; set; }
    }
}
