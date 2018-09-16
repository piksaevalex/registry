using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;

namespace registry
{
    class MyWord
    {
        // Здесь находятся все функции для обработки word таблиц
        //
        //
        //Поиск Реестров word формата
        public static void DirSearchWord(string sDir, ref DataTable dt)
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
        // Обработка word реестра (Выгрузка из него данных)
        private static void Word(FileInfo _file, ref DataTable dt)
        {
            Application wdapp = null;
            Document wddoc = null;
            Table wdtbl = null;;
            wdapp = new Application();
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
                            Model.DtAdd(ref dt, ref row);

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
                    Searchfiles_izobrlist(directory, wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", ""), ref done, ref dt,
                        ref FilesNames);
                    ChooseList(FilesNames, ref row, ref dt);
                    if (done)
                    {
                        if (regEx_marksLS.Match(wdtbl.Cell(0, oboz).Range.Text.Replace("\r\a", "")).Success)
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
        // Поиск файлов описывающих информацию о листах в документе
        public static void Searchfiles_izobrlist(string sDir, string file_name, ref bool done, ref DataTable dt, ref List<Model> FilesNames)
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
                    Searchfiles_izobrlist(d, file_name, ref done, ref dt, ref FilesNames);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }

        }

        // Нахождение самой последней версии инфо о листе
        public static void ChooseList(List<Model> Files, ref Row row, ref DataTable dt)
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
                    GetInfFromList(fi1, ref row, ref dt);
                }

                if (Files.Count == 1)
                {
                    Console.WriteLine(Files[0].Path);
                    Logger.WriteLine(Files[0].Path);
                    FileInfo fi1 = new FileInfo(Files[0].Path);
                    GetInfFromList(fi1, ref row, ref dt);
                }
            }
            catch (Exception excpt)
            {
                Console.WriteLine(excpt.Message);
                Logger.WriteLine(excpt.Message);
            }
        }
        // Распарсить инфу о листах
        private static void GetInfFromList(FileInfo _file, ref Row row, ref DataTable dt)
        {
            Microsoft.Office.Interop.Word.Application wdapp = null;
            Microsoft.Office.Interop.Word.Document wddoc = null;
            Microsoft.Office.Interop.Word.Table wdtbl = null;
            wdapp = new Microsoft.Office.Interop.Word.Application();
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
                                Model.DtAdd(ref dt, ref row);
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
                            Model.DtAdd(ref dt, ref row);
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
                Model.DtAdd(ref dt, ref row);
            }
            finally
            {
                wddoc.Close(SaveChanges: false);
                wdapp.Quit(SaveChanges: false);
            }

        }
    }
}
