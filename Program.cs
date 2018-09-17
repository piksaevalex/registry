using System;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;


namespace registry
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Stopwatch swTotal = new Stopwatch();
                swTotal.Start();
                DataTable dt = new DataTable();
                Model.NewDT(ref dt);
                string directory = AppDomain.CurrentDomain.BaseDirectory;
                // Обработка реестров excel 
                MyExcel.DirSearchEx(directory, ref dt);
                // Обработка реестров word 
                MyWord.DirSearchWord(directory, ref dt);

                swTotal.Stop();
                Console.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                Logger.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                swTotal.Reset();
                Console.WriteLine("Будет вставленно строк : " + dt.Rows.Count);
                Logger.WriteLine("Будет вставленно строк : " + dt.Rows.Count);
                swTotal.Start();
                // Заполнение таблицы
                ExportData.ExportDT(dt);
                swTotal.Stop();
                Console.WriteLine("Writing (new): " + swTotal.ElapsedMilliseconds + " ms");
                Logger.WriteLine("Writing (new): " + swTotal.ElapsedMilliseconds + " ms");
                Console.WriteLine("----------ВСЁ!-------------");
                Console.ReadKey();
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }
    }
}
