using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using GemBox.Document;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OnlineCinemaAdminApplication.Models;

namespace OnlineCinemaAdminApplication.Controllers
{
    public class OrderController : Controller
    {
        public OrderController()
        {
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        }
        public IActionResult Index()
        {
            HttpClient client = new HttpClient();

            string URI = "https://localhost:44389/api/Admin/GetOrders";

            HttpResponseMessage response = client.GetAsync(URI).Result;

            var result = response.Content.ReadAsAsync<List<Order>>().Result;

            return View(result);
        }

        public IActionResult Details(Guid id)
        {
            HttpClient client = new HttpClient();

            string URI = "https://localhost:44389/api/Admin/GetDetailsForTicket";

            var model = new
            {
                Id = id
            };

            HttpContent content = new StringContent(JsonConvert.SerializeObject(model), Encoding.UTF8, "application/json");

            HttpResponseMessage response = client.PostAsync(URI, content).Result;

            var result = response.Content.ReadAsAsync<Order>().Result;

            return View(result);
        }
        public FileContentResult Invoice(Guid id)
        {
            HttpClient client = new HttpClient();

            string URI = "https://localhost:44389/api/Admin/GetDetailsForTicket";

            var model = new
            {
                Id = id
            };

            HttpContent content = new StringContent(JsonConvert.SerializeObject(model),
                Encoding.UTF8, "application/json");

            HttpResponseMessage response = client.PostAsync(URI, content).Result;

            var result = response.Content.ReadAsAsync<Order>().Result;

            var path = Path.Combine(Directory.GetCurrentDirectory(), "Invoice.docx");
            var document = DocumentModel.Load(path);

            document.Content.Replace("{{OrderNumber}}", result.Id.ToString());
            document.Content.Replace("{{UserName}}", result.User.UserName);

            StringBuilder stringBuilder = new StringBuilder();
            var total = 0.0;

            foreach(var ticket in result.TicketInOrders)
            {
                total += ticket.Quantity * ticket.ChosenTicket.TicketPrice;
                stringBuilder.AppendLine(ticket.ChosenTicket.TicketName + "ordered with quantity of: " + ticket.Quantity + " and with  price of: " + ticket.ChosenTicket.TicketPrice + "$");
            }

            document.Content.Replace("{{TicketsList}}", stringBuilder.ToString());
            document.Content.Replace("{{TotalPrice}}", total.ToString() + "$");

            var stream = new MemoryStream();

            document.Save(stream, new PdfSaveOptions());

            return File(stream.ToArray(), new PdfSaveOptions().ContentType, "ExportInvoice.pdf");
        }

        [HttpGet]
        public FileContentResult ExportOrders(string genre)
        {
            string fileName = "Orders.xlsx";
            string type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            using (var workbook = new XLWorkbook())
            {

                IXLWorksheet worksheet = workbook.Worksheets.Add("Orders");

                worksheet.Cell(1, 1).Value = "Order Id";
                worksheet.Cell(1, 2).Value = "Costumer Email";
                worksheet.Cell(1, 3).Value = "Movie Genre";

                HttpClient client = new HttpClient();


                string URI = "https://localhost:44389/api/Admin/GetOrders";

                HttpResponseMessage responseMessage = client.GetAsync(URI).Result;

                var result = responseMessage.Content.ReadAsAsync<List<Order>>().Result;

                for (int i = 1; i <= result.Count(); i++)
                {
                    var item = result[i - 1];

                    worksheet.Cell(i + 1, 1).Value = item.Id.ToString();
                    worksheet.Cell(i + 1, 2).Value = item.User.Email;

                    for (int t = 0; t < item.TicketInOrders.Count(); t++)
                    { 
                        worksheet.Cell(1, t + 4).Value = "Ticket-" + (t + 1);
                        worksheet.Cell(i + 1, t + 4).Value = item.TicketInOrders.ElementAt(t).ChosenTicket.TicketName;
                        worksheet.Cell(i + 1, t + 3).Value = item.TicketInOrders.ElementAt(t).ChosenTicket.MovieGenre;

                        worksheet.RangeUsed().SetAutoFilter().Column(3).EqualTo(genre);
                    }
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(content, type, fileName);
                }
            }
        }

        [HttpPost]
        public string ExportOrders()
        {
            return "Action Done";
        }
    }
}
