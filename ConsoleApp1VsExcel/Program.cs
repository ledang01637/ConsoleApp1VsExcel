using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApp1VsExcel
{
    class Program
    {
        class Match
        {
            public int PlayerA { get; set; }
            public int PlayerB { get; set; }
        }
        class Round
        {
            public Match[] Matches { get; set; }
        }
        static void Main(string[] args)
        {
            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet worksheet = null;
            int lastUsedRow = 0;
            int lastUsedColumn = 0;
            string srcFile = @"D:\FPT Poytechnic\DuAnC#\ConsoleApp1VsExcel\Excel\dataset (1).xlsx";

            xlApp = new Excel.Application();
            xlApp.Visible = false;
            wb = xlApp.Workbooks.Open(srcFile,
                                           0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                                           true, false, 0, true, false, false);

            worksheet = (Excel.Worksheet)wb.Worksheets[1];
            //Excel.Range range = currentWS.Range[cell.Cells[1, 1], cell.Cells[nRowCount, nColumnCount]];
            // Find the last real row
            lastUsedRow = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            // Find the last real column
            lastUsedColumn = worksheet.Cells.Find("*", System.Reflection.Missing.Value,
                                           System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                           Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                           false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;


            var rounds = Generate(lastUsedRow - 1);
            foreach (var round in rounds)
            {
                foreach (var match in round.Matches)
                {
                    Console.WriteLine("{0} vs {1}", match.PlayerA, match.PlayerB);
                }
                Console.WriteLine();
            }
            worksheet.Cells[11,10].Borders[Excel.XlBordersIndex.xlEdgeTop].Weight = 3;
            worksheet.Cells[11,11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
            worksheet.Cells[12,11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
            worksheet.Cells[12,11].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3;
            worksheet.Cells[13,11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
            worksheet.Cells[14,11].Borders[Excel.XlBordersIndex.xlEdgeLeft].Weight = 3;
            worksheet.Cells[14,10].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3;
            xlApp.Workbooks.Close();
            xlApp.Quit();
            //Marshal.ReleaseComObject(worksheet);
            //Marshal.ReleaseComObject(wb);
            //Marshal.ReleaseComObject(xlApp);
            Console.ReadKey();
        }
        static Round[] Generate(int playersNumber)
        {
            var roundsNumber = (int)Math.Log(playersNumber, 2);
            var rounds = new Round[roundsNumber];
            for (int i = 0; i < roundsNumber; i++)
            {
                var round = new Round();
                var prevRound = i > 0 ? rounds[i - 1] : null;
                if (prevRound == null)
                {
                    round.Matches = new[] {
                    new Match() {
                        PlayerA = 1,
                        PlayerB = 2
                    }
                };
                }
                else
                {
                    round.Matches = new Match[prevRound.Matches.Length * 2];
                    var median = (round.Matches.Length * 2 + 1) / 2f;
                    var next = 0;
                    foreach (var match in prevRound.Matches)
                    {
                        round.Matches[next] = new Match()
                        {
                            PlayerA = match.PlayerA,
                            PlayerB = (int)(median + Math.Abs(match.PlayerA - median))
                        };
                        next++;
                        round.Matches[next] = new Match()
                        {
                            PlayerA = match.PlayerB,
                            PlayerB = (int)(median + Math.Abs(match.PlayerB - median))
                        };
                        next++;
                    }
                }
                rounds[i] = round;
            }
            return rounds.Reverse().ToArray();
        }
    }
}