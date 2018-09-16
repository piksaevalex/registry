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
                dt.Clear();
                dt.Columns.Add("SHFR");
                dt.Columns.Add("SHFRDOC");
                dt.Columns.Add("MARKA");
                dt.Columns.Add("OBOSDOC");
                dt.Columns.Add("NAIMOBJ");
                dt.Columns.Add("STAGE");
                dt.Columns.Add("NAIMIZOBR");
                dt.Columns.Add("NAIMPROJE");
                dt.Columns.Add("DATEOFLASTWRITE");
                dt.Columns.Add("Directory");
                string directory = AppDomain.CurrentDomain.BaseDirectory;
                DirSearchEx(directory, ref dt);

                DirSearchWord(directory, ref dt);

                swTotal.Stop();
                Console.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                Logger.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                swTotal.Reset();
                Console.WriteLine("Будет вставленно строк : " + dt.Rows.Count);
                Logger.WriteLine("Будет вставленно строк : " + dt.Rows.Count.ToString());
                swTotal.Start();
                ExportData.ExportDT(dt);
                swTotal.Stop();
                Console.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                Logger.WriteLine("Reading (new): " + swTotal.ElapsedMilliseconds + " ms");
                Console.WriteLine("----------ВСЁ!-------------");
                Console.ReadKey();
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        static void DtAdd(ref DataTable dt, ref Row row)
        {
            DataRow dr = dt.NewRow();
            dr["SHFR"] = row.SHFR;
            dr["SHFRDOC"] = row.SHFRDOC;
            dr["MARKA"] = row.MARKA;
            dr["OBOSDOC"] = row.OBOSDOC;
            dr["NAIMOBJ"] = row.NAIMOBJ;
            dr["STAGE"] = row.STAGE;
            dr["NAIMIZOBR"] = row.NAIMIZOBR;
            dr["NAIMPROJE"] = row.NAIMPROJE;
            dr["DATEOFLASTWRITE"] = row.DATEOFLASTWRITE;
            dr["Directory"] = row.Directory;
            dt.Rows.Add(dr);
        }

        static void DirSearchEx(string sDir, ref DataTable dt)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d, "Реестр РД*.xls*"))
                    {
                        if (new DirectoryInfo(d).Name == "5-РД")
                        {
                            Console.WriteLine(f);
                            Logger.WriteLine(f);
                            FileInfo fi1 = new FileInfo(f);
                            File_Selected_New(fi1, ref dt);
                        }
                    }
                    DirSearchEx(d, ref dt);
                }  
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
            
        }

        static void DirSearchWord(string sDir, ref DataTable dt)
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
                            Word(fi1, ref dt);
                        }
                    }
                    DirSearchWord(d, ref dt);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        static string GetConnectionString(FileInfo _file)
        {
            if (_file == null)
            {
                throw new ArgumentNullException(nameof(_file));
            }

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

        private static void File_Selected_New(FileInfo _file, ref DataTable dt)
        {
            


                if (_file == null)
                {
                    throw new ArgumentNullException(nameof(_file));
                }


                DataSet ds = new DataSet();
                string connectionString = GetConnectionString(_file);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    try
                    {
                        OleDbCommand cmd = new OleDbCommand();
                        cmd.Connection = conn;
                        // Get all Sheets in Excel File
                        DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        // Loop through all Sheets to get data
                        foreach (DataRow dr in dtSheet.Rows)
                        {
                            string sheetName = dr["TABLE_NAME"].ToString();
                            // Get all rows from the Sheet
                            cmd.CommandText = "SELECT * FROM [" + sheetName + "]";
                            DataTable dts = new DataTable();
                            dts.TableName = sheetName;
                            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                            da.Fill(dts);
                            ds.Tables.Add(dts);
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
                                if (regExOboz.Match(table[j, i]).Success)
                                    shifrdoc = i; // находим индекс столбца обозначение 
                                if (regExNaim.Match(table[j, i]).Success)
                                    sooruz = i; // находим индекс столбца наименование
                            }
                        }

                        bool state = false;
                        Regex regEx_marksLS = new Regex("^.*ЛС.*$"); 
                        Regex regExProject = new Regex("^Шифр..*$"); // находим шифр проекта по этой маске
                        Regex regExStage = new Regex("^..*[Э,э][Т,т][А,а][П,п]$"); // находим этап по этой маске
                        Regex regExStage2 = new Regex("^[Э,э][Т,т][А,а][П,п]..*$"); // находим этап по этой маске
                        Regex regExStage_number = new Regex("^/d*$"); // находим цифру по этой маске
                        Row row = new Row();
                        row.STAGE = "1";
                        row.NAIMPROJE = table[0, 0];
                        row.Directory = _file.FullName;
                        row.DATEOFLASTWRITE = _file.LastWriteTime.ToShortDateString();
                        string directory = _file.DirectoryName;
                        for (int j = 0; j < ds.Tables[0].Rows.Count; j++)
                        {

                            switch (template)
                            {
                                case 0:
                                    row.SHFRDOC = table[j, shifrdoc];
                                    row.NAIMOBJ = table[j, sooruz];
                                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                                    {
                                        if (regExProject.Match(table[j, i]).Success)
                                        {
                                            row.SHFR = table[j, i];
                                            row.SHFR = row.SHFR.Replace("Шифр ", "");
                                        }

                                        if (regExStage.Match(table[j, i]).Success)
                                        {
                                            foreach (var item in table[j, 1].Split(' '))
                                            {
                                                if (regExStage_number.Match(item).Success) row.STAGE = item;
                                            }
                                        }

                                        if (regExStage2.Match(table[j, i]).Success)
                                        {
                                            foreach (var item in table[j, 1].Split(' '))
                                            {
                                                if (regExStage_number.Match(item).Success) row.STAGE = item;
                                            }
                                        }

                                        if (i > 1 && table[j, i] != "")
                                        {
                                            row.MARKA = table[j, i];
                                            if (row.SHFRDOC != "" && row.SHFRDOC != "Шифр" &&
                                                row.SHFRDOC != "Договор №")
                                            {
                                                if (row.MARKA != "")
                                                {
                                                    row.OBOSDOC = row.SHFRDOC + "-" + row.MARKA;
                                                    bool done = true;
                                                    List<Model> FilesNames = new List<Model>();
                                                    Search_list(directory, row.SHFRDOC + "*" + row.MARKA, ref done,
                                                        ref dt,
                                                        ref FilesNames);
                                                    ChooseList(FilesNames, ref row, ref dt);
                                                    if (done)
                                                    {
                                                        if (regEx_marksLS.Match(row.SHFRDOC + "*" + row.MARKA).Success)
                                                        {
                                                            row.NAIMIZOBR = "";
                                                            DtAdd(ref dt, ref row);
                                                        }
                                                        else
                                                        {
                                                            row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                                            DtAdd(ref dt, ref row);
                                                        }
                                                    }
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

                                    if (row.SHFRDOC != "Шифр" && table[j, 3] != "")
                                    {
                                        //row.MARKA = table[j, 3];
                                        foreach (var item in table[j, 3].Split(','))
                                        {
                                            row.OBOSDOC = row.SHFRDOC + "-" + item.Replace(" ", "");
                                            bool done = true;
                                            List<Model> FilesNames = new List<Model>();
                                            Search_list(directory, row.SHFRDOC + "*" + item.Replace(" ", ""), ref done,
                                                ref dt, ref FilesNames);
                                            ChooseList(FilesNames, ref row, ref dt);
                                            if (done)
                                            {
                                                if (regEx_marksLS.Match(row.SHFRDOC + "*" + item.Replace(" ", "")).Success)
                                                {
                                                    row.NAIMIZOBR = "";
                                                    DtAdd(ref dt, ref row);
                                                }
                                                else
                                                {
                                                    row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                                    DtAdd(ref dt, ref row);
                                                }
                                                
                                            }
                                        }

                                    }

                                    break;
                            }

                            if (table[j, 0] == "Шифр" && table[j, 1] == "Сооружение (площадка, трасса, система)" &&
                                table[j, 3] == "Марка" && table[j, 4] == "Изменения")
                            {
                                template = 1;
                            }
                        }
                    }
                    catch (Exception excpt)
                    {
                        Console.WriteLine(excpt.Message);
                        Logger.WriteLine(excpt.Message);
                    }
                    finally
                    {
                        conn.Close();
                    }
                }
            
        }


        private static void Word(FileInfo _file, ref DataTable dt)
        {
            Word.Application wdapp = null;
            Word.Document wddoc = null;
            Word.Table wdtbl = null;
            //Word.Section wdcoll = null;

            wdapp = new Word.Application();
            wddoc = wdapp.Documents.Open(_file.FullName, ReadOnly: true, AddToRecentFiles: false);
            try
            {
  
                wdtbl = wddoc.Tables[1];
                Row row = new Row();
                row.Directory = _file.FullName;
                row.DATEOFLASTWRITE = _file.LastWriteTime.ToShortDateString();
                string directory = _file.DirectoryName;
                Regex regEx_marksLS = new Regex("^.*ЛС.*$"); // находим шифр проекта по этой маске
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
                        if (regExOboz.Match(wdtbl.Cell(i, j).Range.Text).Success)
                            oboz = j; // находим индекс столбца обозначение 
                        if (regExNaim.Match(wdtbl.Cell(i, j).Range.Text).Success)
                            naim = j; // находим индекс столбца наименование
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
                            DtAdd(ref dt, ref row);

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
                                    row.STAGE = row.STAGE.Remove(startIndex: row.STAGE.IndexOf("-"));
                                    row.STAGE = row.STAGE.Remove(0, row.STAGE.IndexOf(' ') + 1);
                                }

                                //str = str.Split('.')[1];
                                str = str.Remove(0, str.IndexOf('.') + 2);
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

                if (wdtbl.Cell(0, oboz).Range.Text != "\r\a"
                ) // фича ворда, последняя строка табоицы записывается в строку с индексом 0
                {
                    row.OBOSDOC = wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", "");
                    bool done = true;
                    List<Model> FilesNames = new List<Model>();
                    Search_list(directory, wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", ""), ref done, ref dt,
                        ref FilesNames);
                    ChooseList(FilesNames, ref row, ref dt);
                    if (done)
                    {
                        if (regEx_marksLS.Match(wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", "")).Success)
                        {
                            row.NAIMIZOBR = "";
                            DtAdd(ref dt, ref row);
                        }
                        else
                        {
                            row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                            DtAdd(ref dt, ref row);
                        }
                            
                    }
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
            finally
            {
                wddoc.Close(SaveChanges: false);
                wdapp.Quit(SaveChanges: false);
            }
        }

        private static void Search_list(string sDir, string file_name, ref bool done, ref DataTable dt, ref List<Model> FilesNames)
        {           
            if (file_name == null)
            {
                throw new ArgumentNullException(nameof(file_name));
            }
            try
            {
                Regex regExOprosList = new Regex("^.*ОЛ.*$");
                foreach (string d in Directory.GetDirectories(sDir))
                {                    
                    file_name = file_name.Replace(@"/", "*").Replace(@".", @"_");
                    foreach (string f in Directory.GetFiles(d, "*" + file_name + "*.doc*"))
                    {
                        if (!regExOprosList.Match(f).Success)
                        {
                            done = false;
                            FilesNames.Add(new Model() { Path = f, index = 0 });
                        }
                                            
                    }
                    Search_list(d, file_name, ref done, ref dt, ref FilesNames);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
            
        }

        // Нахождение самой последней версии инфо о листе
        private static void ChooseList(List<Model> Files, ref Row row, ref DataTable dt)
        {
            try
            {
                Regex regExIZM = new Regex("^.*Изм.*$");
                if (Files.Count > 1)
                {
                    int maxindex = -1;
                    int truefileindex = 0;
                    foreach (var File in Files)
                    {
                        // разобьём путь к файлу по каталогам 
                        string[] item = File.Path.Split('\\');
                        foreach (var catalog in item)
                        {
                            // Нахожим каталог со значением Изменения
                            if (regExIZM.Match(catalog).Success)
                            {
                                int numberofchanche = Convert.ToInt32(Regex.Replace(catalog, @"[^\d]+", ""));
                                File.index = numberofchanche;
                            }
                        }

                        if (File.index > maxindex)
                        {
                            maxindex = File.index;
                            truefileindex = Files.IndexOf(File);
                        }
                    }

                    Console.WriteLine(Files[truefileindex].Path);
                    Logger.WriteLine(Files[truefileindex].Path);
                    FileInfo fi1 = new FileInfo(Files[truefileindex].Path);
                    GetList(fi1, ref row, ref dt);
                }

                if (Files.Count == 1)
                {
                    Console.WriteLine(Files[0].Path);
                    Logger.WriteLine(Files[0].Path);
                    FileInfo fi1 = new FileInfo(Files[0].Path);
                    GetList(fi1, ref row, ref dt);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }

        // Распарсить инфу о листах
        private static void GetList(FileInfo _file, ref Row row, ref DataTable dt)
        {
            Word.Application wdapp = null;
            Word.Document wddoc = null;
            Word.Table wdtbl = null;
            wdapp = new Word.Application();
            wddoc = wdapp.Documents.Open(_file.FullName, ReadOnly: true, AddToRecentFiles: false);
            try
            {
                wdtbl = wddoc.Tables[1];
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
                    string[] izobr = wdtbl.Cell(i, 2).Range.Text.Split('\r');
                    int flag = 0;
                    if (wdtbl.Cell(i, 1).Range.Text.Replace("\a", "").Replace("\r", "") != "")
                    {
                        foreach (var item in wdtbl.Cell(i, 1).Range.Text.Split('\r'))
                        {
                            row.NAIMIZOBR = izobr[flag].Replace("\a", "").Replace("\r", "");
                            row.OBOSDOC = obosnachdocleft + "_Л." + item.Replace("\a", "").Replace("\r", "");
                            if (row.NAIMIZOBR != "Наименование" && item.Replace("\a", "").Replace("\r", "") != "")
                            {
                                DtAdd(ref dt, ref row);
                                flag++;
                            }
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
                        row.NAIMIZOBR = izobr[flag].Replace("\a", "").Replace("\r", "");
                        row.OBOSDOC = obosnachdocleft + "_Л." + item.Replace("\a", "").Replace("\r", "");
                        if (row.NAIMIZOBR != "Наименование" && item.Replace("\a", "").Replace("\r", "") != "")
                        {
                            DtAdd(ref dt, ref row);
                            flag++;
                        }
                    }
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
                row.NAIMIZOBR = "Ошибка в файле описи листов";
                DtAdd(ref dt, ref row);
            }
            finally
            {
                wddoc.Close(SaveChanges: false);
                wdapp.Quit(SaveChanges: false);
            }
            
        }
    }
}
