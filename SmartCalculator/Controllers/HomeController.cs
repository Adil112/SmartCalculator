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
            DateTime dCount1 = new DateTime(2021, 1, 1, 0, 0, 0); 
            DateTime dCount2 = new DateTime(2021, 1, 1, 0, 0, 0);
            if (d1.Minute != 0 || d1.Second != 0) dCount1 = d1.AddMinutes(-d1.Minute).AddSeconds(-d1.Second);
            else dCount1 = d1;
            if (d2.Minute != 0 || d2.Second != 0) dCount2 = d2.AddMinutes(-d1.Minute).AddSeconds(-d1.Second);
            else dCount2 = d2;
            int dayC = (dCount2 - dCount1).Days;
            int hourC = dCount2.Hour - dCount1.Hour;
            int maxCounter = dayC * 24 + hourC;
            DataCount[] dataCounts = new DataCount[++maxCounter];
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
                
                for(int i =0; i < maxCounter; i++)
                {
                    DataCount dt = new DataCount();
                    dt.DateTime = dCount1.AddHours(i).ToString();
                    dt.Count = 0;
                    dataCounts[i] = dt;
                }
                foreach(var t in data)
                {
                    DateTime time = DateTime.ParseExact(t.dateTime, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture); 
                    if (time > d1 && time < d2)
                    { 
                        data1.Add(t);
                        for(int i =0; i < maxCounter; i++)
                        {
                            int j = i + 1;
                            bool oio = (dCount1.Hour + i) == time.Hour && time.Hour < (dCount1.Hour + j);
                            if ((dCount1.Hour + i) == time.Hour && time.Hour < (dCount1.Hour + j))
                            {
                                dataCounts[i].Count++;
                            }
                        }
                        
                    }
                }
            }
            ViewBag.dataCounts = dataCounts;
            ViewBag.datas = data1;
            return View(); 
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
