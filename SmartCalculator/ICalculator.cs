using SmartCalculator.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SmartCalculator
{
    public interface ICalculator
    {
        CalculatorInfo Calculate(CalculatorInfo calc);

        string GetIP();
        void GetDoc(ListsOfReport lists, DateTime d1, DateTime d2);
        void SaveHistory(History h);
        ListsOfReport GetListsHistory(DateTime d1, DateTime d2);

    }
}
