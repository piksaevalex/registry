using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace registry
{
    public class Row
    {
        public string SHFR { get; set; } // Шифр проекта
        public string SHFRDOC { get; set; } // Шифр документа
        public string MARKA { get; set; } // Марка в шифре документа
        public string OBOSDOC { get; set; } // Обозначение документа
        public string NAIMOBJ { get; set; } // Наименование объекта
        public string STAGE { get; set; } // Этап
        public string NAIMIZOBR { get; set; } // Наименование изображения
        public string NAIMPROJE { get; set; } // Наименование проекта
        public string DATEOFLASTWRITE { get; set; } // Дата факт. поступления (в нашем случае, дата последней записи
        public string Directory { get; set; } // Ссылка на документ
    }
}
