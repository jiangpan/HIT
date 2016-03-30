using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess;
using System.Data.Common;
using System.Diagnostics;
using System.Configuration;
using Dapper;
using System.Data.SqlClient;
using System.IO;

namespace FG.PDMReader.Oracle
{
    public partial class frmFGMDM : Form
    {
        private List<USER_TABLES> tablist;
        private List<USER_TAB_COLUMNS> tabcolslist;
        private List<sysdatatypemappings> sysdatatypemappingsList;

        public frmFGMDM()
        {
            InitializeComponent();
        }

        private void btnGenScriptAll_Click(object sender, EventArgs e)
        {
            if (tabcolslist == null || tabcolslist.Count == 0)
            {
                return;
            }
            List<USER_TAB_COLUMNS> curTabColsList = null;
            StringBuilder sb = new StringBuilder();

            foreach (USER_TABLES tab in tablist)
            {
                curTabColsList = null;
                curTabColsList = tabcolslist.Where(p => p.TABLE_NAME == tab.TABLE_NAME).ToList<USER_TAB_COLUMNS>();


                sb.AppendLine(GetCreateTableScript(tab, curTabColsList));
            }


            string fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".sql";

            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);


            File.WriteAllText(filePath, sb.ToString());





        }

        private string GetCreateTableScript(USER_TABLES ut, List<USER_TAB_COLUMNS> cols)
        {
            string snip = string.Empty;
            StringBuilder sql = new StringBuilder();
            sql.AppendFormat("CREATE TABLE {0}\r\n(\r\n", ut.TABLE_NAME);
            for (int i = 0; i < cols.Count; i++)
            {
                USER_TAB_COLUMNS utc = cols[i];
                snip = GetColumnSql(utc);
                sql.AppendFormat((i < cols.Count - 1) ? snip : snip.TrimEnd(',', '\r', '\n').Replace(",",""));
            }
            sql.AppendFormat("\r\n)");
            return sql.ToString();
        }

        private string GetColumnSql(USER_TAB_COLUMNS dr)
        {
            StringBuilder sql = new StringBuilder();
            sql.AppendFormat("\t[{0}] {1}{2} {3} {4},  --{5}\r\n",
                dr.COLUMN_NAME,
               OracleDataTypeToMSSQL(dr),
               (HasSize(OracleDataTypeToMSSQL( dr))) ? "(" + dr.DATA_LENGTH + ")" : (HasPrecisionAndScale(dr.DATA_TYPE)) ? "(" + dr.DATA_PRECISION + "," + dr.DATA_SCALE + ")" : "",
                "", //(dr["IsIdentity"].ToString() == "true") ? "IDENTITY" : "",
               ( dr.NULLABLE == "Y") ? "NULL" : "NOT NULL",
               dr.COMMENTS
                );
            return sql.ToString();
        }

        private string OracleDataTypeToMSSQL(USER_TAB_COLUMNS oracleType)
        {
            if (sysdatatypemappingsList == null || sysdatatypemappingsList.Count == 0)
            {
                return string.Empty;
            }
            if (oracleType.CHARACTER_SET_NAME == "CHAR_CS")
            {
                return "varchar";
            }
            
            if (oracleType.CHARACTER_SET_NAME == "NCHAR_CS")
            {
                return "nvarchar";
            }
            if (oracleType.DATA_TYPE == "NUMBER")
            {
                if (string.IsNullOrEmpty( oracleType.DATA_PRECISION))
                {
                    return "bigint";
                }
                if (int.Parse(oracleType.DATA_PRECISION) <= 1 && (string.IsNullOrEmpty(oracleType.DATA_SCALE) || int.Parse(oracleType.DATA_SCALE) == 0))
                {
                    return "bit";
                }
                if (int.Parse(oracleType.DATA_PRECISION) <= 3 && (string.IsNullOrEmpty(oracleType.DATA_SCALE) || int.Parse(oracleType.DATA_SCALE) == 0))
                {
                    return "bit";
                }
                if (int.Parse(oracleType.DATA_PRECISION) <= 5 && (string.IsNullOrEmpty(oracleType.DATA_SCALE) || int.Parse(oracleType.DATA_SCALE) == 0))
                {
                    return "smallint";
                }
                if (  int.Parse(oracleType.DATA_PRECISION) <= 10  && (string.IsNullOrEmpty(oracleType.DATA_SCALE) || int.Parse(oracleType.DATA_SCALE) ==0 ))
                {
                    return "int";
                }
                return "int";
            }
            if (oracleType.DATA_TYPE == "DATE")
            {
                return "datetime";
            }
            if (oracleType.DATA_TYPE == "TIMESTAMP")
            {
                string sqltype = string.Empty;
                switch (oracleType.DATA_LENGTH)
                {
                    case "3":
                        sqltype =  "datetime";
                        break;
                    case "7":
                        sqltype = "datetime2";
                        break;
                    default:
                        break;
                }
                return sqltype;
            }
            //sysdatatypemappings datatypemap = sysdatatypemappingsList.Find(p => (p.destination_type == oracleType.DATA_TYPE ));
            
            return string.Empty;

        }

        private bool HasSize(string dataType)
        {
            Dictionary<string, bool> dataTypes = new Dictionary<string, bool>();

            if (dataTypes.Count <=0)
            {
                dataTypes.Add("bigint", false);
                dataTypes.Add("binary", true);
                dataTypes.Add("bit", false);
                dataTypes.Add("char", true);
                dataTypes.Add("date", false);
                dataTypes.Add("datetime", false);
                dataTypes.Add("datetime2", false);
                dataTypes.Add("datetimeoffset", false);
                dataTypes.Add("decimal", false);
                dataTypes.Add("float", false);
                dataTypes.Add("geography", false);
                dataTypes.Add("geometry", false);
                dataTypes.Add("hierarchyid", false);
                dataTypes.Add("image", true);
                dataTypes.Add("int", false);
                dataTypes.Add("money", false);
                dataTypes.Add("nchar", true);
                dataTypes.Add("ntext", true);
                dataTypes.Add("numeric", false);
                dataTypes.Add("nvarchar", true);
                dataTypes.Add("real", false);
                dataTypes.Add("smalldatetime", false);
                dataTypes.Add("smallint", false);
                dataTypes.Add("smallmoney", false);
                dataTypes.Add("sql_variant", false);
                dataTypes.Add("sysname", false);
                dataTypes.Add("text", true);
                dataTypes.Add("time", false);
                dataTypes.Add("timestamp", false);
                dataTypes.Add("tinyint", false);
                dataTypes.Add("uniqueidentifier", false);
                dataTypes.Add("varbinary", true);
                dataTypes.Add("varchar", true);
                dataTypes.Add("xml", false);
            }
            if (dataTypes.ContainsKey(dataType))
                return dataTypes[dataType];
            return false;
        }

        private static bool HasPrecisionAndScale(string dataType)
        {
            Dictionary<string, bool> dataTypes = new Dictionary<string, bool>();
            dataTypes.Add("bigint", false);
            dataTypes.Add("binary", false);
            dataTypes.Add("bit", false);
            dataTypes.Add("char", false);
            dataTypes.Add("date", false);
            dataTypes.Add("datetime", false);
            dataTypes.Add("datetime2", false);
            dataTypes.Add("datetimeoffset", false);
            dataTypes.Add("decimal", true);
            dataTypes.Add("float", true);
            dataTypes.Add("geography", false);
            dataTypes.Add("geometry", false);
            dataTypes.Add("hierarchyid", false);
            dataTypes.Add("image", false);
            dataTypes.Add("int", false);
            dataTypes.Add("money", false);
            dataTypes.Add("nchar", false);
            dataTypes.Add("ntext", false);
            dataTypes.Add("numeric", false);
            dataTypes.Add("nvarchar", false);
            dataTypes.Add("real", true);
            dataTypes.Add("smalldatetime", false);
            dataTypes.Add("smallint", false);
            dataTypes.Add("smallmoney", false);
            dataTypes.Add("sql_variant", false);
            dataTypes.Add("sysname", false);
            dataTypes.Add("text", false);
            dataTypes.Add("time", false);
            dataTypes.Add("timestamp", false);
            dataTypes.Add("tinyint", false);
            dataTypes.Add("uniqueidentifier", false);
            dataTypes.Add("varbinary", false);
            dataTypes.Add("varchar", false);
            dataTypes.Add("xml", false);
            if (dataTypes.ContainsKey(dataType))
                return dataTypes[dataType];
            return false;
        }

        private void frmFGMDM_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void frmFGMDM_Load(object sender, EventArgs e)
        {
            using (DbConnection conn = OracleClientFactory.Instance.CreateConnection())
            {

                Stopwatch sw = Stopwatch.StartNew();
                conn.ConnectionString = ConfigurationManager.ConnectionStrings["FGMDM"].ConnectionString;
                try
                {
                    conn.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
               

                //查询出所有的表

                tablist = conn.Query<USER_TABLES>("SELECT TABLE_NAME,STATUS,NUM_ROWS FROM USER_TABLES").AsList<USER_TABLES>();
                tabcolslist = conn.Query<USER_TAB_COLUMNS>(@"SELECT c.TABLE_NAME,c.COLUMN_NAME,c.DATA_TYPE,c.DATA_LENGTH,c.DATA_PRECISION,
        c.DATA_SCALE,c.NULLABLE,c.COLUMN_ID,c.DEFAULT_LENGTH,c.DATA_DEFAULT,c.CHARACTER_SET_NAME,co.COMMENTS
        FROM USER_TAB_COLUMNS c INNER JOIN USER_COL_COMMENTS co
        ON c.TABLE_NAME = co.table_name AND c.COLUMN_NAME = co.column_name").AsList<USER_TAB_COLUMNS>();

                sw.Stop();
                this.label1.Text = "";
                this.label1.Text = string.Format("当前时间：{0}，执行耗时:{1}ms", DateTime.Now, sw.ElapsedMilliseconds);

                this.dgvTabs.AutoGenerateColumns = true;

                this.dgvTabs.DataSource = tablist;

                this.dgvTabCols.AutoGenerateColumns = true;
                List<USER_TAB_COLUMNS> colList = tabcolslist.Where(p => p.TABLE_NAME == "APPSYSDOMAIN").ToList<USER_TAB_COLUMNS>();//tablist[0].TABLE_NAME
                this.dgvTabCols.DataSource = colList;

            }
            using (var connsql = SqlClientFactory.Instance.CreateConnection())
            {
                Stopwatch sw = Stopwatch.StartNew();
                connsql.ConnectionString = ConfigurationManager.ConnectionStrings["EMPI"].ConnectionString;
                try
                {
                    connsql.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                //查询出所有的表

                sysdatatypemappingsList = connsql.Query<sysdatatypemappings>("SELECT * FROM msdb.dbo.sysdatatypemappings where destination_dbms = 'ORACLE' AND destination_version = '11'").AsList<sysdatatypemappings>();

                sw.Stop();
                this.label1.Text = "";
                this.label1.Text = string.Format("当前时间：{0}，执行耗时:{1}ms", DateTime.Now, sw.ElapsedMilliseconds);

               



            }




        }
    }
}
