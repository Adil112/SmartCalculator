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
            lists = calculator.GetLists(d1, d2);
            ViewBag.dataCounts = lists.list1; //dataCounts1;
            ViewBag.datas = lists.list2;  //data1;

            calculator.GetDoc(lists);
            return View(); 
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


    }
}
