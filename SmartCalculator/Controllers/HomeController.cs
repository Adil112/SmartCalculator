using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SmartCalculator.Models;
using System;
using System.Diagnostics;
using System.Net;
using OfficeOpenXml;
using System.IO;
using System.Web;
using System.Collections.Generic;
using System.Globalization;

namespace SmartCalculator.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }
        [HttpGet]
        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult Index(Calc calc)
        {
            double a, b;
            a = calc.num1;
            b = calc.num2;
            string sign = null;
            switch(calc.oper)
            {
                case "Сложение" : 
                    calc.result = a + b;
                    sign = "+";
                    break;
                case "Вычитание":
                    calc.result = a - b;
                    sign = "-";
                    break;
                case "Умножение":
                    calc.result = a * b;
                    sign = "*";
                    break;
                case "Деление":
                    calc.result = a / b;
                    sign = "/";
                    break;
            }
            calc.showres = calc.result.ToString();
            ViewData["showresult"] = calc.showres;

            History h = new History();
            h.dateTime = DateTime.Now.ToString();
            h.operation = a + " " + sign + " " + b + " = " + calc.result;
            h.ip = GetIP();
            SetData(h);

            return View();
        }
        [HttpGet]
        public IActionResult GetData()
        {
            return View();
        }
        [HttpPost]
        public IActionResult GetData(DateTime d1, DateTime d2)
        {
            List<History> data1 = new List<History>();
            List<History> data = new List<History>();
            string path = @"D:\Adil\ElesyTest\Data.xlsx";
            FileInfo existingFile = new FileInfo(path);
            using (var package = new ExcelPackage(existingFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                
                double val = (double)worksheet.Cells[1, 4].Value;
                int value = (int)val;
                for(int i = 1; i <= value; i++)
                {
                    History hh = new History();
                    hh.dateTime = worksheet.Cells[i, 1].Value.ToString();
                    hh.operation = worksheet.Cells[i, 2].Value.ToString();
                    hh.ip = worksheet.Cells[i, 3].Value.ToString();
                    data.Add(hh);
                }
                
                foreach(var t in data)
                {
                    DateTime time = DateTime.ParseExact(t.dateTime, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture); //System.Globalization.CultureInfo.InvariantCulture
                    if (time > d1 && time < d2) data1.Add(t);
                }
            }
            return View(data1);
            /*<div>
    <hr />
    <dl class="dl-horizontal">
        @foreach (var t in Model)
        {
            <dt>
                @Html.DisplayName(t)
            </dt>

            <dd>
                @Html.ActionLink("Информация", "Index2", new { id = t })
            </dd>
        }
    </dl>
</div> */
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
        void SetData(History h)
        {
            string path = @"D:\Adil\ElesyTest\Data.xlsx";
            FileInfo existingFile = new FileInfo(path);
            using (var package = new ExcelPackage(existingFile))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                var value = worksheet.Cells[1, 4].Value;
                double count = (double)value + 1;
                worksheet.Cells[(int)count, 1].Value = h.dateTime;
                worksheet.Cells[(int)count, 2].Value = h.operation;
                worksheet.Cells[(int)count, 3].Value = h.ip;
                worksheet.Cells[1, 4].Value = count++;

                package.Save();
            }
        }
        private string GetIP()
        {
            string strHostName = "";
            strHostName = System.Net.Dns.GetHostName();

            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);

            IPAddress[] addr = ipEntry.AddressList;

            return addr[addr.Length - 1].ToString();

        }

    }
}
