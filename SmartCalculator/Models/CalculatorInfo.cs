using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SmartCalculator.Models
{
    public enum Operation
    {
        Сложение,
        Вычитание,
        Умножение,
        Деление
    }
    public class CalculatorInfo
    {
        public double num1 { get; set; }
        public double num2 { get; set; }
        public double result { get; set; }
        public string showres { get; set; }
        public Operation oper { get; set; }
    }
}
