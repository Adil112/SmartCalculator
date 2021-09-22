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
        ICalculator _calculator;
        public HomeController(ICalculator calculator)
        {
            _calculator = calculator;
        }
        [HttpGet]
        public IActionResult Index()
        {
            CalculatorInfo calc = new CalculatorInfo();
            return View(calc);
        }
        [HttpPost]
        public IActionResult Index(CalculatorInfo calc)
        {
            calc.showres = _calculator.Calculate(calc).showres;
            return View(calc);
        }
        [HttpGet]
        public IActionResult GetData()
        {
            return View();
        }
        [HttpPost]
        public IActionResult GetData(DateTime d1, DateTime d2)
        {
            ListsOfReport lists = new ListsOfReport();
            if (d1 > d2)
            {
                ModelState.AddModelError("", "Неккоректный интервал времени!");
                return View();
            }
            if (ModelState.IsValid)
            {
                lists = _calculator.GetListsHistory(d1, d2);
                ViewBag.dataCounts = lists.list1;
                ViewBag.datas = lists.list2;  
                ViewBag.d1 = d1;
                ViewBag.d2 = d2;
                _calculator.GetDoc(lists, d1, d2);
                return View();
            }
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }


    }
}
