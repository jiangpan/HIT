using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FG.PDMReader.Oracle
{
    class sysdatatypemappings
    {
        public string mapping_id { get; set; }
        public string source_dbms { get; set; }
        public string source_version { get; set; }
        public string source_type { get; set; }
        public string source_length_min { get; set; }
        public string source_length_max { get; set; }
        public string source_precision_min { get; set; }
        public string source_precision_max { get; set; }
        public string source_scale_min { get; set; }
        public string source_scale_max { get; set; }
        public string source_nullable { get; set; }
        public string source_createparams { get; set; }
        public string destination_dbms { get; set; }
        public string destination_version { get; set; }
        public string destination_type { get; set; }
        public string destination_length { get; set; }
        public string destination_precision { get; set; }
        public string destination_scale { get; set; }
        public string destination_nullable { get; set; }
        public string destination_createparams { get; set; }
        public string dataloss { get; set; }
        public string is_default { get; set; }
    }
}
