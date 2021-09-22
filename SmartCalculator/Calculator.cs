using OfficeOpenXml;
using SmartCalculator.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Net;


namespace SmartCalculator
{
    public class Calculator
    {
        public CalculatorInfo Calculate(CalculatorInfo calc)// сам калькулятор
        {
            double a, b;
            a = calc.num1;
            b = calc.num2;
            string sign = null;
            switch (calc.oper) 
            {
                case "Сложение":
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

            History h = new History();
            h.dateTime = DateTime.Now.ToString();
            h.operation = a + " " + sign + " " + b + " = " + calc.result;


            h.ip = GetIP();
            SaveHistory(h);
            return calc;
        }
        public string GetIP() // получаем IP адрес
        {
            string strHostName = "";
            strHostName = System.Net.Dns.GetHostName();

            IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(strHostName);

            IPAddress[] addr = ipEntry.AddressList;

            return addr[addr.Length - 1].ToString();
        }
        public void GetDoc(ListsOfReport lists, DateTime d1, DateTime d2) //сохранение отчета в документе excel (исправить!!!)
        {
            Random rnd = new Random();
            int z = rnd.Next(1, 100);
            string name = "Report" + z.ToString() + ".xlsx";
            FileInfo newFile = new FileInfo(name);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pck = new ExcelPackage(newFile);
            
            var worksheet = pck.Workbook.Worksheets.Add("History");

            worksheet.Cells[1, 1].Value = "Время и дата";
            worksheet.Cells[1, 2].Value = "Операция";
            worksheet.Cells[1, 3].Value = "IP адрес";
            worksheet.Cells[1, 5].Value = "Дата и время";
            worksheet.Cells[1, 6].Value = "Кол-во операции";
            worksheet.Cells[1, 8].Value = d1.ToString();
            worksheet.Cells[1, 9].Value = d2.ToString();
            var list1 = lists.list2;
            var list2 = lists.list1;
            int list1Num = list1.Count;
            int list2Num = list2.Count;
            int list1Counter = 2;
            int list2Counter = 2;
            foreach (var t in list1)
            {
                if (list1Counter > (list1Num + 2)) break;
                list1Counter++;
                worksheet.Cells[list1Counter, 1].Value = t.dateTime;
                worksheet.Cells[list1Counter, 2].Value = t.operation;
                worksheet.Cells[list1Counter, 3].Value = t.ip;
            }
            foreach (var t in list2)
            {
                if (list2Counter > (list2Num + 2)) break;
                list2Counter++;
                worksheet.Cells[list2Counter, 5].Value = t.dateTime;
                worksheet.Cells[list2Counter, 6].Value = t.count;
            }


            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            pck.Save();
        }

        public void SaveHistory(History h)// записываем данные в csv
        {
            string path = @"Data.csv";
            var dateTime = h.dateTime;
            var operation = h.operation;
            var ip = h.ip;
            File.AppendAllLines(path, new[] { $"{dateTime};{operation};{ip}" });
        }

        public ListsOfReport GetListsHistory(DateTime d1, DateTime d2) //получение записей истории в виде списков
        {
            DateTime dCount1 = new DateTime(2021, 1, 1, 0, 0, 0);
            DateTime dCount2 = new DateTime(2021, 1, 1, 0, 0, 0);
            if (d1.Minute != 0 || d1.Second != 0) dCount1 = d1.AddMinutes(-d1.Minute).AddSeconds(-d1.Second);
            else dCount1 = d1;
            if (d2.Minute != 0 || d2.Second != 0) dCount2 = d2.AddMinutes(-d1.Minute).AddSeconds(-d1.Second);
            else dCount2 = d2;
            var max = (dCount2 - dCount1).TotalHours;
            int maxCounter = (int)max;
            DataCount[] dataCounts = new DataCount[++maxCounter];
            List<History> data1 = new List<History>();
            List<History> data = new List<History>();
            string path = @"Data.csv";


            int value = 0;
            string line;
            TextReader reader = new StreamReader(path);
            while ((line = reader.ReadLine()) != null) // определение количества строк
            {
                value++;
            }
            reader.Close();

            using (StreamReader sr = new StreamReader(path)) //чтение записей из файла
            {
                string[] row = new string[3];
                while ((line = sr.ReadLine()) != null)
                {
                    row = line.Split(';');
                    History hh = new History();
                    hh.dateTime = row[0];
                    hh.operation = row[1];
                    hh.ip = row[2];
                    data.Add(hh);
                }
            }
         

            for (int i = 0; i < maxCounter; i++) // создание почасового списка для каждого часа
            {
                DataCount dt = new DataCount();
                dt.dateTime = dCount1.AddHours(i).ToString();
                dt.count = 0;
                dataCounts[i] = dt;
            }
            foreach (var t in data) // перебор записей
            {
                DateTime time = DateTime.ParseExact(t.dateTime, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                if (time > d1 && time < d2) // фильтрация по времени
                {
                    data1.Add(t); // обычный список истории
                    for (int i = 0; i < maxCounter; i++)
                    {
                        int j = i + 1;
                        DateTime dCount3 = dCount1.AddHours(i);
                        DateTime dCount4 = dCount1.AddHours(j);
                        if ((dCount3 < time) && (time < dCount4))
                        {
                            dataCounts[i].count++; // почасовой список истории
                        }
                    }
                }
            }

            List<DataCount> dataCounts1 = new List<DataCount>();
            for (int i = 0; i < dataCounts.Length; i++) // убираем лишние записи
            {
                if (dataCounts[i].count != 0) dataCounts1.Add(dataCounts[i]);
            }
            ListsOfReport lists = new ListsOfReport();
            lists.list1 = dataCounts1;
            lists.list2 = data1;
            return lists;
        }
    }
}
