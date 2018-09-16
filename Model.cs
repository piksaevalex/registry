using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = Microsoft.Office.Interop.Excel.DataTable;

namespace registry
{
    public class Model
    { 
        public string Path { get; set; }
        public int index { get; set; }

        public static void NewDT(ref System.Data.DataTable dt)
        {
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
        }

        public static void DtAdd(ref System.Data.DataTable dt, ref Row row)
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
    }
}
