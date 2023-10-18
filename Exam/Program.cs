using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Exam
{
    class Program
    {
        static void Main(string[] args)
        {
            double A;
            string Path = @"D:\FPT Poytechnic\DuAnC#\ConsoleApp1VsExcel\Excel\dataset (1).xlsx";
            FileStream fileStream = new FileStream(Path, FileMode.Open, FileAccess.Read);
            XSSFWorkbook wb = new XSSFWorkbook(fileStream);
            ISheet sheet1 = wb.GetSheetAt(0);

            if (sheet1 != null)
            {
                for (int i = 1; i < 10; i++)
                {
                    IRow row = sheet1.GetRow(i);
                    A = row.GetCell(3).NumericCellValue;
                    Console.WriteLine(":AA A" + A + "\n");
                }
            }
            Console.ReadKey();
        }
    }
}
