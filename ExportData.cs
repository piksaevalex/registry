using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DataTable = System.Data.DataTable;

namespace registry
{
    class ExportData
    {
        public static void ExportDT(DataTable dt)
        {
            XSSFWorkbook workbook;

            using (FileStream file = new FileStream("Ш-01.07.03.03-38.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(file);
            }
            ISheet worksheet = workbook.GetSheet("Лист1");
            ICellStyle hlink_style = workbook.CreateCellStyle();
            IFont hlink_font = workbook.CreateFont();
            hlink_font.Underline = FontUnderlineType.Single;
            hlink_font.Color = HSSFColor.Blue.Index;
            hlink_style.WrapText = true;
            hlink_style.SetFont(hlink_font);
            ICellStyle easy_style = workbook.CreateCellStyle();
            easy_style.WrapText = true;
            IFont easy_font = workbook.CreateFont();
            easy_font.FontHeightInPoints = 9;
            easy_font.FontName = "Arial Cyr";
            easy_style.SetFont(easy_font);
            easy_style.Alignment = HorizontalAlignment.Center;
            easy_style.VerticalAlignment = VerticalAlignment.Center;
            for (int rownum = 2; rownum < dt.Rows.Count + 2; rownum++)
            {
                IRow row = worksheet.CreateRow(rownum);
                ICell Cell_5 = row.CreateCell(4); Cell_5.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["SHFR"])); Cell_5.CellStyle = easy_style;
                ICell Cell_19 = row.CreateCell(18); Cell_19.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["STAGE"])); Cell_19.CellStyle = easy_style;
                ICell Cell_7 = row.CreateCell(6); Cell_7.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["OBOSDOC"])); Cell_7.CellStyle = easy_style;
                ICell Cell_6 = row.CreateCell(5); Cell_6.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["NAIMPROJE"])); Cell_6.CellStyle = easy_style;
                ICell Cell_8 = row.CreateCell(7); Cell_8.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["NAIMOBJ"])); Cell_8.CellStyle = easy_style;
                ICell Cell_9 = row.CreateCell(8); Cell_9.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["NAIMIZOBR"])); Cell_9.CellStyle = easy_style;
                ICell Cell_3 = row.CreateCell(2); Cell_3.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["DATEOFLASTWRITE"])); Cell_3.CellStyle = easy_style;
                ICell Cell_20 = row.CreateCell(19); Cell_20.SetCellValue(Convert.ToString(dt.Rows[rownum - 2]["Directory"])); Cell_20.CellStyle = easy_style;
                string link = Convert.ToString(dt.Rows[rownum - 2]["Directory"]);
                var url = new Uri(link);
                XSSFHyperlink FileLink = new XSSFHyperlink(HyperlinkType.File);
                FileLink.Address = Convert.ToString(url);
                Cell_20.Hyperlink = (FileLink);
                Cell_20.CellStyle = (hlink_style);


                for (int celnum = 0; celnum < dt.Columns.Count; celnum++)
                {
                    //ICell Cell = row.CreateCell(celnum);
                    //Cell.SetCellValue(Convert.ToString(dt.Rows[rownum - 2][celnum]));
                    
                }
                
            }

            for (int celnum = 0; celnum < dt.Columns.Count; celnum++)
            {
                //worksheet.AutoSizeColumn(celnum);
            }

            if (!File.Exists("test2.xlsx"))
            {
                File.Delete("test2.xlsx");
            }
            using (FileStream file = new FileStream("test2.xlsx", FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
            {               
                workbook.Write(file);
            }
        }
    }
}
