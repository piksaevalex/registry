using System;
using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
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
                int countrow = 5;
                string directory = AppDomain.CurrentDomain.BaseDirectory;
                // Создаём ссылку на Excel приложение
                Excel.Workbook excelappworkbook;                                    // Создаём ссылку на рабочую книгу Excel-приложения
                Excel.Sheets excelsheets;                                           // Создаём ссылку для работы со страницами Excel-приложения
                Excel.Worksheet excelworksheet;                                    // Создаём ссылку на рабочую страницу Excel-приложения
                Excel.Application excelapp = new Excel.Application();

                excelappworkbook = excelapp.Workbooks.Open(directory + "Ш-01.07.03.03-38.xls",           // Устанавливаем ссылку рабочей книги на книгу по пути взятого из TextBox. Параметры(FileName(Имя открываемого файла файла), 
                        Type.Missing, Type.Missing, Type.Missing,                       // UpdateLinks(Способ обновления ссылок в файле), ReadOnly(При значении true открытие только для чтения), Format(Определение формата символа разделителя)
                        "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,     // Password(Пароль доступа к файлу до 15 символов), WriteResPassword(Пароль на сохранение файла), IgnoreReadOnlyRecommended(При значении true отключается вывода запроса на работу без внесения изменений), Origin(Тип текстового файла)
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,         // Delimiter(Разделитель при Format = 6), Editable(Используется только для надстроек Excel 4.0), Notify(При значении true имя файла добавляется в список нотификации файлов), 
                        Type.Missing, Type.Missing);                                    // Converter(Используется для передачи индекса конвертера файла используемого для открытия файла), AddToMRU(При true имя файла добавляется в список открытых файлов)
                excelsheets = excelappworkbook.Worksheets;                      // Устанавливаем ссылку Страниц на страницы новой книги
                excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);
                //Search_list(directory, "", ref excelworksheet, ref countrow);
                DirSearchEx(directory, ref excelworksheet, ref countrow);
                excelappworkbook.Save();
                excelapp.Quit();



                Excel.Application excelapp2;                                         // Создаём ссылку на Excel приложение
                Excel.Workbook excelappworkbook2;                                    // Создаём ссылку на рабочую книгу Excel-приложения
                Excel.Sheets excelsheets2;                                           // Создаём ссылку для работы со страницами Excel-приложения
                Excel.Worksheet excelworksheet2;                                    // Создаём ссылку на рабочую страницу Excel-приложения

                excelapp2 = new Excel.Application();
                excelappworkbook2 = excelapp2.Workbooks.Open(directory + "Ш-01.07.03.03-38.xls",           // Устанавливаем ссылку рабочей книги на книгу по пути взятого из TextBox. Параметры(FileName(Имя открываемого файла файла), 
                        Type.Missing, Type.Missing, Type.Missing,                       // UpdateLinks(Способ обновления ссылок в файле), ReadOnly(При значении true открытие только для чтения), Format(Определение формата символа разделителя)
                        "WWWWW", "WWWWW", Type.Missing, Type.Missing, Type.Missing,     // Password(Пароль доступа к файлу до 15 символов), WriteResPassword(Пароль на сохранение файла), IgnoreReadOnlyRecommended(При значении true отключается вывода запроса на работу без внесения изменений), Origin(Тип текстового файла)
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing,         // Delimiter(Разделитель при Format = 6), Editable(Используется только для надстроек Excel 4.0), Notify(При значении true имя файла добавляется в список нотификации файлов), 
                        Type.Missing, Type.Missing);                                    // Converter(Используется для передачи индекса конвертера файла используемого для открытия файла), AddToMRU(При true имя файла добавляется в список открытых файлов)
                excelsheets2 = excelappworkbook2.Worksheets;                      // Устанавливаем ссылку Страниц на страницы новой книги
                excelworksheet2 = (Excel.Worksheet)excelsheets2.get_Item(1);

                DirSearchWord(directory, ref excelworksheet2, ref countrow);

                excelappworkbook2.Save();
                excelapp2.Quit();
                Console.WriteLine("----------ВСЁ!-------------");
                Console.ReadKey();
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        static void DirSearchEx(string sDir, ref Excel.Worksheet excelworksheet, ref int countrow)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d, "Реестр РД*.xls*"))
                    {
                        if (new DirectoryInfo(d).Name == "5-РД5")
                        {
                            Console.WriteLine(f);
                            Logger.WriteLine(f);
                            FileInfo fi1 = new FileInfo(f);
                            File_Selected_New(fi1, ref excelworksheet, ref countrow);
                        }
                    }
                    DirSearchEx(d, ref excelworksheet, ref countrow);
                }  
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
            
        }

        static void DirSearchWord(string sDir, ref Excel.Worksheet excelworksheet, ref int countrow)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {              
                    foreach (string f in Directory.GetFiles(d, "Реестр РД*.doc*"))
                    {
                        if (new DirectoryInfo(d).Name == "5-РД")
                        {
                            Console.WriteLine(f);
                            Logger.WriteLine(f);
                            FileInfo fi1 = new FileInfo(f);
                            Word(fi1, ref excelworksheet, ref countrow);
                        }
                    }
                    DirSearchWord(d, ref excelworksheet, ref countrow);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        static string GetConnectionString(FileInfo _file)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();


            // XLSX - Excel 2007, 2010, 2012, 2013
            if (_file.Extension == ".xlsx")
            {
                props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
                props["Extended Properties"] = "Excel 12.0 XML";
                props["Data Source"] = _file.FullName;
            }
            else if (_file.Extension == ".xls")
            {
                props["Provider"] = "Microsoft.Jet.OLEDB.4.0";
                props["Extended Properties"] = "Excel 8.0";
                props["Data Source"] = _file.FullName;
            }
            else throw new Exception("Неизвестное расширение файла!");

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        private static void File_Selected_New(FileInfo _file, ref Excel.Worksheet excelworksheet, ref int countrow)
        {
            Stopwatch sw_total = new Stopwatch();
            sw_total.Start();
            DataSet ds = new DataSet();
            string connectionString = GetConnectionString(_file);
            using (OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connectionString))
            {
                conn.Open();
                System.Data.OleDb.OleDbCommand cmd = new System.Data.OleDb.OleDbCommand();
                cmd.Connection = conn;
                // Get all Sheets in Excel File
                System.Data.DataTable dtSheet = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                // Loop through all Sheets to get data
                foreach (DataRow dr in dtSheet.Rows)
                {
                    string sheetName = dr["TABLE_NAME"].ToString();
                    // Get all rows from the Sheet
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.TableName = sheetName;
                    System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(cmd);
                    da.Fill(dt);
                    ds.Tables.Add(dt);
                }
                string[,] table = new string[ds.Tables[0].Rows.Count, ds.Tables[0].Columns.Count];
                int shifrdoc = 1; // индекс столбца шифрдокумента
                int sooruz = 2; // индекс столбца сооружение
                int template = 0;
                Regex regExOboz = new Regex(@"^Шифр.*$"); // находим индекс столбца 
                Regex regExNaim = new Regex(@"^Сооружение.*$"); // находим индекс столбца 
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        table[j, i] = ds.Tables[0].Rows[j].ItemArray[i].ToString();
                        if (regExOboz.Match(table[j, i]).Success) shifrdoc = i; // находим индекс столбца обозначение 
                        if (regExNaim.Match(table[j, i]).Success) sooruz = i; // находим индекс столбца наименование
                    }                      
                }

                bool state = false;
                Regex regExProject = new Regex("^Шифр..*$"); // находим шифр проекта по этой маске
                Regex regExStage = new Regex("^..*[Э,э][Т,т][А,а][П,п]$"); // находим этап по этой маске
                Regex regExStage2 = new Regex("^[Э,э][Т,т][А,а][П,п]..*$"); // находим этап по этой маске
                Regex regExStage_number = new Regex("^/d*$"); // находим цифру по этой маске
                Row row = new Row();
                row.STAGE = "1";
                row.NAIMPROJE = table[0, 0];
                row.Directory = _file.FullName;
                row.DATEOFLASTWRITE = _file.LastWriteTime.ToShortDateString();
                //string directory = AppDomain.CurrentDomain.BaseDirectory;
                string directory = _file.DirectoryName;
                for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                {
                    
                    switch (template)
                    {
                        case 0:
                            row.SHFRDOC = table[j, shifrdoc]; row.NAIMOBJ = table[j, sooruz];
                            for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                            {
                                if (regExProject.Match(table[j, i]).Success) { row.SHFR = table[j, i]; row.SHFR = row.SHFR.Replace("Шифр ", ""); }
                                if (regExStage.Match(table[j, i]).Success)
                                {
                                    foreach (var item in table[j, 1].Split(' ')) 
                                    {
                                        if (regExStage_number.Match(item).Success) row.STAGE = item;
                                    }
                                    //row.STAGE = table[j, i].Remove(table[j, i].IndexOf(' '));
                                }
                                if (regExStage2.Match(table[j, i]).Success)
                                {
                                    foreach (var item in table[j, 1].Split(' '))
                                    {
                                        if (regExStage_number.Match(item).Success) row.STAGE = item;
                                    }
                                    //row.STAGE = table[j, i].Remove(0, table[j, i].IndexOf(' '));
                                }
                                if (i > 1 && table[j, i] != "")
                                {
                                    row.MARKA = table[j, i];
                                    if (row.SHFRDOC != "" && row.SHFRDOC != "Шифр" && row.SHFRDOC != "Договор №")
                                    {
                                        if (row.MARKA != "")
                                        {
                                            row.OBOSDOC = row.SHFRDOC + "-" + row.MARKA;
                                            bool done = true;
                                            Search_list(directory, row.SHFRDOC + "*" + row.MARKA, ref excelworksheet, ref countrow, ref row, ref done);
                                            if (done)
                                            {
                                                row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                                Export(row, ref countrow, ref excelworksheet);

                                            }
                                            //Export(row, ref countrow, ref excelworksheet);
                                        }
                                        
                                    }

                                }
                            }
                            break;
                        case 1:
                            if (table[j, 0] != "") row.SHFRDOC = table[j, 0];
                            if (table[j, 1] != "") row.NAIMOBJ = table[j, 1];
                            if (table[j, 2] != "")
                            {
                                //row.STAGE = table[j, 2];
                                foreach (var item in table[j, 2].Split(' '))
                                {
                                    if (regExStage_number.Match(item).Success) row.STAGE = item;
                                }
                            }
                            if (row.SHFRDOC != "Шифр" && table[j,3] != "")
                            {
                                //row.MARKA = table[j, 3];
                                var Marks = table[j, 3].Split(',');
                                foreach (var item in Marks)
                                {
                                    row.OBOSDOC = row.SHFRDOC + "-" + item.Replace(" ", "");
                                    //row.OBOSDOC = row.OBOSDOC.Replace(@"--",@"-");
                                    bool done = true;
                                    Search_list(directory, row.SHFRDOC + "*" + item.Replace(" ", ""), ref excelworksheet, ref countrow, ref row, ref done);
                                    //Export(row, ref countrow, ref excelworksheet);
                                    if (done)
                                    {
                                        row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                        Export(row, ref countrow, ref excelworksheet);

                                    }
                                }
                                
                            }

                            break;
                    }
                    if (table[j, 0] == "Шифр" && table[j, 1] == "Сооружение (площадка, трасса, система)" && table[j, 3] == "Марка" && table[j, 4] == "Изменения") { template = 1; }

                }
            }
            sw_total.Stop();
            Console.WriteLine("Reading (new): " + sw_total.ElapsedMilliseconds + " ms");
            
        }
        

        private static void Export(Row row, ref int countrow, ref Excel.Worksheet excelworksheet)
        {
            if (row.NAIMPROJE != null) excelworksheet.Cells[countrow, 6].Value = row.NAIMPROJE;
            if (row.SHFR != null) excelworksheet.Cells[countrow, 5].Value = row.SHFR;
            if (row.OBOSDOC != null) excelworksheet.Cells[countrow, 7].Value = row.OBOSDOC;
            if (row.NAIMOBJ != null) excelworksheet.Cells[countrow, 8].Value = row.NAIMOBJ;
            if (row.NAIMIZOBR != null) excelworksheet.Cells[countrow, 9] = row.NAIMIZOBR;
            if (row.STAGE != null) excelworksheet.Cells[countrow, 19] = row.STAGE;
            //if (row.Directory != null) excelworksheet.Cells[countrow, 20] = row.Directory;
            if (row.DATEOFLASTWRITE != null) excelworksheet.Cells[countrow, 3] = row.DATEOFLASTWRITE;
            excelworksheet.Hyperlinks.Add(
                (Excel.Range)excelworksheet.Cells[countrow, 20],
                row.Directory,
                string.Empty,
                "Screen Tip Text",
                row.Directory);
            countrow++;
        }

        private static void Word(FileInfo _file, ref Excel.Worksheet excelworksheet, ref int countrow)
        {
            Stopwatch sw_total = new Stopwatch();
            sw_total.Start();
            Word.Application wdapp = null;
            Word.Document wddoc = null;
            Word.Table wdtbl = null;
            //Word.Section wdcoll = null;
           
            wdapp = new Word.Application();
            wddoc = wdapp.Documents.Open(_file.FullName, ReadOnly: true, AddToRecentFiles: false);
            wdtbl = wddoc.Tables[1];
            //wdcoll = wddoc.Sections[1];
            Row row = new Row();
            row.Directory = _file.FullName;
            row.DATEOFLASTWRITE = _file.LastWriteTime.ToShortDateString();
            string directory = _file.DirectoryName;
            Regex regExProject = new Regex("^Шифр..*$"); // находим шифр проекта по этой маске
            Regex regExStage = new Regex("^.*[Э|э]тап.*$"); // находим этап по этой маске
            Regex regExStage2 = new Regex("^[Э|э]тап.*$"); // вариация для другого шаблона
            Regex regDoc = new Regex(@"^.*\d\d\r\a$"); // проверка на то что в колонке обозночение шифр без марки
            Regex regExStage3 = new Regex(@"^\d+-й этап.*$"); // обработка этапа для ещё однго шаблона
            Regex regExOboz = new Regex(@"^[О|о]бозначение.*$"); // находим индекс столбца 
            Regex regExNaim = new Regex(@"^[Н|н]аименование.*$"); // находим индекс столбца 
            int oboz = 2; // индекс столбца обозначение
            int naim = 3; // индекс столбца наименование
            for (int i = 1; i < wdtbl.Rows.Count; i++)
            {   
                for (int j = 1; j < wdtbl.Columns.Count; j++)
                {
                    if (regExOboz.Match(wdtbl.Cell(i, j).Range.Text).Success) oboz = j; // находим индекс столбца обозначение 
                    if (regExNaim.Match(wdtbl.Cell(i, j).Range.Text).Success) naim = j; // находим индекс столбца наименование
                }
                if (wdtbl.Cell(i, oboz).Range.Text != "\r\a" && wdtbl.Cell(i, oboz).Range.Text != "Обозначение\r\a") 
                {
                    if (regDoc.Match(wdtbl.Cell(i, oboz).Range.Text).Success)
                    {
                        row.NAIMOBJ = wdtbl.Cell(i, naim).Range.Text.Replace("\r\a", "");
                    }
                    else
                    {
                        if (row.NAIMIZOBR == "Наименование") row.NAIMIZOBR = "";
                        row.OBOSDOC = wdtbl.Cell(i, oboz).Range.Text.Replace("\r\a", "");
                        Export(row, ref countrow, ref excelworksheet);
                    }
                        
                }
                else
                {
                    if (regExStage.Match(wdtbl.Cell(i, naim).Range.Text).Success)
                    {
                        if (wdtbl.Cell(i, naim).Range.Text.Length > 10)
                        {
                            row.STAGE = wdtbl.Cell(i, naim).Range.Text.Replace("\r", "");
                            row.STAGE = row.STAGE.Replace("\a", "");
                            int a = row.STAGE.IndexOf(' ');
                            string str = row.STAGE.Remove(0, row.STAGE.IndexOf(' ') + 1);
                            int b = str.IndexOf(' ');
                            string str2 = str.Remove(0, str.IndexOf(' ') + 1); // наименование объекта
                            int c = a + b;
                            string str4 = row.STAGE.Substring(0, c + 1); // значение этапа

                            row.STAGE = str4.Replace(".", "");
                            row.STAGE = row.STAGE.Replace(":", "");
                            if (regExStage3.Match(row.STAGE).Success)
                            {
                                row.STAGE = row.STAGE.Remove(row.STAGE.IndexOf("-"));
                                row.STAGE = row.STAGE.Remove(0, row.STAGE.IndexOf(' ') + 1);
                            }
                            //str = str.Split('.')[1];
                            str = str.Remove(0, str.IndexOf('.')+2);
                            row.NAIMOBJ = str;
                            row.STAGE = row.STAGE.Replace("Этап ", "");
                        }
                        else
                        {
                            if (regExStage2.Match(wdtbl.Cell(i, naim).Range.Text).Success)
                            {
                                row.STAGE = wdtbl.Cell(i, naim).Range.Text.Replace("\r", "");
                                row.STAGE = row.STAGE.Replace("\a", "");
                                string str = row.STAGE.Remove(0, row.STAGE.IndexOf(' ') + 1);
                                row.STAGE = str.Remove(0, str.IndexOf(' ') + 1);
                                row.STAGE = row.STAGE.Replace(":", "");
                            }
                            else
                            {
                                row.STAGE = wdtbl.Cell(i, naim).Range.Text.Replace("\r", "");
                                row.STAGE = row.STAGE.Replace("\a", "");
                                string str = row.STAGE.Remove(row.STAGE.IndexOf(' ') + 1);
                                row.STAGE = str.Remove(str.IndexOf(' '));
                                row.STAGE = row.STAGE.Replace(":", "");
                            }                           
                        }                            
                    }
                    else
                    {
                        row.NAIMIZOBR = wdtbl.Cell(i, naim).Range.Text.Replace("\a", "").Replace("\r", "");
                    }
                }
            }
            if (wdtbl.Cell(0, oboz).Range.Text != "\r\a") // фича ворда, последняя строка табоицы записывается в строку с индексом 0
            {
                row.OBOSDOC = wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", "");
                bool done = true;
                Search_list(directory, wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", ""), ref excelworksheet, ref countrow, ref row, ref done);
                if (done)
                {
                    row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                    Export(row, ref countrow, ref excelworksheet);

                }
                //Export(row, ref countrow, ref excelworksheet);
            }
            wddoc.Close(SaveChanges: false);
            wdapp.Quit(SaveChanges: false);
            sw_total.Stop();
            Console.WriteLine("Reading (new): " + sw_total.ElapsedMilliseconds + " ms");

        }

        private static void Search_list(string sDir, string file_name, ref Excel.Worksheet excelworksheet, ref int countrow, ref Row row, ref bool done)
        {
            try
            {
                
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    file_name = file_name.Replace(@"/", "*").Replace(@".", @"_");
                    foreach (string f in Directory.GetFiles(d, "*" + file_name + "*.doc*"))
                    {
                        //if (new DirectoryInfo(d).Name == "*")
                        //{

                            Console.WriteLine(f);
                            Logger.WriteLine(f);
                            FileInfo fi1 = new FileInfo(f);
                            done = false;
                            GetList(fi1, ref excelworksheet, ref countrow, ref row);
                        //}
                        
                    }
                    Search_list(d, file_name, ref excelworksheet, ref countrow, ref row, ref done);
                }
            }
            catch (System.Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        private static void GetList(FileInfo _file, ref Excel.Worksheet excelworksheet, ref int countrow, ref Row row)
        {
            Word.Application wdapp = null;
            Word.Document wddoc = null;
            Word.Table wdtbl = null;
            //Word.Section wdcoll = null;

            wdapp = new Word.Application();
            wddoc = wdapp.Documents.Open(_file.FullName, ReadOnly: true, AddToRecentFiles: false);
            wdtbl = wddoc.Tables[1];
            //wdcoll = wddoc.Sections[1];
            string obosnachdocleft = row.OBOSDOC;
            string[,] table = new string[wdtbl.Rows.Count, wdtbl.Columns.Count];
            for (int i = 0; i < wdtbl.Rows.Count; i++)
            {
                for (int j = 0; j < wdtbl.Columns.Count; j++)
                {
                    table[i, j] = wdtbl.Cell(i, j).Range.Text;
                }
            }
            for (int i = 1; i < wdtbl.Rows.Count; i++)
            {
                //for (int j = 1; j < wdtbl.Columns.Count; j++)
                //{
                    //Console.WriteLine(i + ", " + j + " : " + wdtbl.Cell(i, j).Range.Text.Replace("\a", "").Replace("\r", ""));
                //}
                if (wdtbl.Cell(i, 1).Range.Text.Replace("\a", "").Replace("\r", "") != "")
                {
                    row.NAIMIZOBR = wdtbl.Cell(i, 2).Range.Text.Replace("\a", "").Replace("\r", "");
                    foreach (var item in wdtbl.Cell(i, 1).Range.Text.Split('\r'))
                    {
                        row.OBOSDOC = obosnachdocleft + "_Л." + item.Replace("\a", "").Replace("\r", "");
                        if (row.NAIMIZOBR != "Наименование" && item.Replace("\a", "").Replace("\r", "") != "") Export(row, ref countrow, ref excelworksheet);
                    }
                    
                    
                }

            }
            // фишка ворда индекс последней строки 0
            if (wdtbl.Cell(0, 1).Range.Text.Replace("\a", "").Replace("\r", "") != "")
            {
                string[] izobr = wdtbl.Cell(0, 2).Range.Text.Split('\r');
                int flag = 0;
                //row.NAIMIZOBR = wdtbl.Cell(0, 2).Range.Text.Replace("\a", "").Replace("\r", "");
                foreach (var item in wdtbl.Cell(0, 1).Range.Text.Split('\r'))
                {
                    row.NAIMIZOBR = izobr[flag].Replace("\a", "").Replace("\r", ""); flag++;
                    row.OBOSDOC = obosnachdocleft + "_Л." + item.Replace("\a", "").Replace("\r", "");
                    if (row.NAIMIZOBR != "Наименование" && item.Replace("\a", "").Replace("\r", "") != "") Export(row, ref countrow, ref excelworksheet);
                }
            }
            wddoc.Close(SaveChanges: false);
            wdapp.Quit(SaveChanges: false);
        }
    }
}
