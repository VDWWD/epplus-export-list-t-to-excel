using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using System.Web;
using Newtonsoft.Json;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            var model = new HomeViewModel()
            {
                json = JsonConvert.SerializeObject(getBooks(), Formatting.Indented)
            };

            return View(model);
        }


        [HttpPost]
        [ValidateAntiForgeryToken]
        public FileStreamResult Index(HomeViewModel model)
        {
            var books = new List<HomeViewModel.Book>();

            try
            {
                books = JsonConvert.DeserializeObject<List<HomeViewModel.Book>>(model.json);
            }
            catch
            {
                books = getBooks();
            }

            byte[] excel = EPPlus.createExcel(books, "VDWWD", "My Book List");

            return new FileStreamResult(new MemoryStream(excel), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") { 
                FileDownloadName = "VDWWD-list-to-excel-demo.xlsx" 
            };
        }


        public List<HomeViewModel.Book> getBooks()
        {
            var books = new List<HomeViewModel.Book>();

            //create some book data
            for (int i = 0; i < 10; i++)
            {
                books.Add(new HomeViewModel.Book()
                {
                    ID = i,
                    Name = "Name " + i,
                    Category = "Category " + i,
                    Date = DateTime.Now.AddDays(i).AddYears(i - i * 2),
                    Author = "Author " + i,
                    Price = i * i,
                    Published = i % 2 == 0,
                    SpinOffs = i % 5 == 0 ? new List<HomeViewModel.Book>() : null,
                    ISBN = new List<string>()
                    {
                        "978-3-16-148410-0",
                        "978-3-16-148410-1"
                    }
                });
            }

            return books;
        }
    }
}