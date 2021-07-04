using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OnlineCinemaAdminApplication.Models;

namespace OnlineCinemaAdminApplication.Controllers
{
    public class UserController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult ImportUsers(IFormFile file)
        {

            
            string path = $"{Directory.GetCurrentDirectory()}\\files\\{file.FileName}";

            using (FileStream fileStream = System.IO.File.Create(path))
            {
                file.CopyTo(fileStream);

                fileStream.Flush();
            }

            
            List<TicketUser> users = getAllUsersFromFile(file.FileName);


            HttpClient client = new HttpClient();


            string URI = "https://localhost:44389/api/Admin/ImportUsers";


            HttpContent content = new StringContent(JsonConvert.SerializeObject(users), Encoding.UTF8, "application/json");

            HttpResponseMessage response = client.PostAsync(URI, content).Result;


            var result = response.Content.ReadAsAsync<bool>().Result;


            return RedirectToAction("Index", "Order");
        }

        private List<TicketUser> getAllUsersFromFile(string fileName)
        {

            List<TicketUser> users = new List<TicketUser>();

            string filePath = $"{Directory.GetCurrentDirectory()}\\files\\{fileName}";

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);


            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        users.Add(new Models.TicketUser
                        {
                            Email = reader.GetValue(0).ToString(),
                            Password = reader.GetValue(1).ToString(),
                            ConfirmPassword = reader.GetValue(2).ToString(),
                            Role = reader.GetValue(3).ToString()
                        });
                    }
                }
            }


            return users;
        }
    }
}

