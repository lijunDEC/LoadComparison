using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoadComparison
{
    class Program
    {
        
        static void Main(string[] args)
        {
            ExcelOperation excelOperation = new ExcelOperation();
//            try
            {
                excelOperation.CreateMainEequivalentFatigueLoadsSheet();
                excelOperation.CreateMainUltimateLoadsSheet();
                excelOperation.CreateMainUltimateLoadsThermodynamicChartSheet();
            }
//            catch
            {
  //              excelOperation.QuitExcel();
            }
        }

    }
}
