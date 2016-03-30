using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FG.PDMReader.Oracle
{
    public class USER_TAB_COLUMNS
    {
        public string TABLE_NAME { get; set; }
        public string COLUMN_NAME { get; set; }
        public string DATA_TYPE { get; set; }
        public string DATA_LENGTH { get; set; }
        public string DATA_PRECISION { get; set; }
        public string DATA_SCALE { get; set; }
        public string NULLABLE { get; set; }
        public string COLUMN_ID { get; set; }
        public string DEFAULT_LENGTH { get; set; }
        public string DATA_DEFAULT { get; set; }


        public string CHARACTER_SET_NAME { get; set; }

        public string COMMENTS { get; set; }


        

    }
}
