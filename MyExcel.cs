using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace registry
{
    class MyExcel
    {
        // Здесь находятся все функции для обработки excel таблиц
        //
        //
        //Поиск Реестров excel формата
        public static void DirSearchEx(string sDir, ref DataTable dt)
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
                            ReestrParse(fi1, ref dt);
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
        // Обработка excel (Выгрузка из него данных)
        private static void ReestrParse(FileInfo _file, ref DataTable dt)
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
                                                MyWord.Searchfiles_izobrlist(directory, row.SHFRDOC + "*" + row.MARKA, ref done,
                                                    ref dt,
                                                    ref FilesNames);
                                                MyWord.ChooseList(FilesNames, ref row, ref dt);
                                                if (done)
                                                {
                                                    if (regEx_marksLS.Match(row.SHFRDOC + "*" + row.MARKA).Success)
                                                    {
                                                        row.NAIMIZOBR = "";
                                                        Model.DtAdd(ref dt, ref row);
                                                    }
                                                    else
                                                    {
                                                        row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                                        Model.DtAdd(ref dt, ref row);
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
                                        MyWord.Searchfiles_izobrlist(directory, row.SHFRDOC + "*" + item.Replace(" ", ""), ref done,
                                            ref dt, ref FilesNames);
                                        MyWord.ChooseList(FilesNames, ref row, ref dt);
                                        if (done)
                                        {
                                            if (regEx_marksLS.Match(row.SHFRDOC + "*" + item.Replace(" ", "")).Success)
                                            {
                                                row.NAIMIZOBR = "";
                                                Model.DtAdd(ref dt, ref row);
                                            }
                                            else
                                            {
                                                row.NAIMIZOBR = "Не удалось найти вспомогательный файл";
                                                Model.DtAdd(ref dt, ref row);
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
        // Вспомогательная функция для парса реестра
        private static string GetConnectionString(FileInfo _file)
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
    }
}
