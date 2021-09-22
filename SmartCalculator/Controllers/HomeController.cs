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
            Calculator calculator = new Calculator();
            ViewData["showresult"] = calculator.Calculate(calc).showres;
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
            Calculator calculator = new Calculator();
            Lists lists = new Lists();
            if (d1 > d2)
            {
                ViewBag.Error = "Неправильно ввденный интервал времени!";
                return View();
            }
                
            //lists = calculator.GetLists(d1, d2);
            lists = calculator.GetListsHistory(d1, d2);
            ViewBag.dataCounts = lists.list1; //dataCounts1;
            ViewBag.datas = lists.list2;  //data1;
            ViewBag.Error = null;
            ViewBag.d1 = d1;
            ViewBag.d2 = d2;
            //calculator.GetDoc(lists, d1, d2);
            return View(); 
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


    }
}
