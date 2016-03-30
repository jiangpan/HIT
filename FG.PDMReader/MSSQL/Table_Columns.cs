using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FG.PDMReader.MSSQL
{
    public class Table_Column
    {
        public string TABLE_NAME { set; get; }
        public string COLUMN_NAME { set; get; }
        public string DATA_TYPE { set; get; }
        public string COLUMN_DEFAULT { set; get; }
        public string IS_NULLABLE { set; get; }
        public string IS_PRIMARY_KEY { set; get; }
        public string IS_FOREIGN_KEY { set; get; }
        public string FOREIGN_KEY { set; get; }
        public string FOREIGN_TABLE { set; get; }
        public string COLUMN_DESC { set; get; }
    }
}
