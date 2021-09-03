using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication1.Models
{
    public class HomeViewModel
    {
        public string json { get; set; }

        public class Book
        {
            public int ID { get; set; }
            public string Name { get; set; }
            public string Category { get; set; }
            public DateTime Date { get; set; }
            public bool Published { get; set; }
            public List<Book> SpinOffs { get; set; }
            public string Author { get; set; }
            public List<string> ISBN { get; set; }
            public Decimal Price { get; set; }
        }
    }
}