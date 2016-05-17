using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Dapper;
using System.IO;
using System.Reflection;
using FG.PDMReader.Helper;

namespace FG.PDMReader.MSSQL
{
    public partial class frmFGMDM : Form
    {
        /// <summary>
        /// sql语句还需要调整，优化结构
        /// 调整sql20160429 091907
        /// </summary>
        private string sqlgetallcolumns = @"SELECT  /*CASE WHEN c.ORDINAL_POSITION = 1
             THEN c.TABLE_SCHEMA + '.' + c.TABLE_NAME
             ELSE ''
        END AS*/c.TABLE_SCHEMA + '.' + c.TABLE_NAME TABLE_NAME ,
        c.COLUMN_NAME ,
        c.ORDINAL_POSITION,
        CASE WHEN ( ( CHARINDEX('char', c.DATA_TYPE) > 0
                      OR CHARINDEX('binary', c.DATA_TYPE) > 0
                    )
                    AND c.CHARACTER_MAXIMUM_LENGTH <> -1
                  )
             THEN c.DATA_TYPE + '('
                  + CAST(c.CHARACTER_MAXIMUM_LENGTH AS VARCHAR(4)) + ')'
             WHEN ( ( CHARINDEX('CHAR', c.DATA_TYPE) > 0
                      OR CHARINDEX('binary', c.DATA_TYPE) > 0
                    )
                    AND c.CHARACTER_MAXIMUM_LENGTH = -1
                  ) THEN c.DATA_TYPE + '(max)'
             WHEN ( CHARINDEX('numeric', c.DATA_TYPE) > 0 )
             THEN c.DATA_TYPE + '(' + CAST(c.NUMERIC_PRECISION AS VARCHAR(4))
                  + ',' + CAST(c.NUMERIC_SCALE AS VARCHAR(4)) + ')'
             ELSE c.DATA_TYPE
        END AS DATA_TYPE ,
        CASE WHEN COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME),c.COLUMN_NAME,'IsIdentity') = 1 THEN '√' ELSE '' END AS IS_IDENTITY, 
        ISNULL(c.COLUMN_DEFAULT, '') AS COLUMN_DEFAULT ,
        CASE WHEN c.IS_NULLABLE = 'YES' THEN '√'
             ELSE ''
        END IS_NULLABLE ,
        CASE WHEN tc.CONSTRAINT_TYPE = 'PRIMARY KEY' THEN '√'
             ELSE ''
        END AS IS_PRIMARY_KEY ,
        CASE WHEN tc.CONSTRAINT_TYPE = 'FOREIGN KEY' THEN '√'
             ELSE ''
        END AS IS_FOREIGN_KEY ,
        ISNULL(fkcu.COLUMN_NAME, '') AS FOREIGN_KEY ,
        ISNULL(fkcu.TABLE_NAME, '') AS FOREIGN_TABLE,
        ep.value AS COLUMN_DESC
FROM    [INFORMATION_SCHEMA].[COLUMNS] c
        LEFT JOIN [INFORMATION_SCHEMA].[KEY_COLUMN_USAGE] kcu ON kcu.TABLE_SCHEMA = c.TABLE_SCHEMA
                                                              AND kcu.TABLE_NAME = c.TABLE_NAME
                                                              AND kcu.COLUMN_NAME = c.COLUMN_NAME
        LEFT JOIN [INFORMATION_SCHEMA].[TABLE_CONSTRAINTS] tc ON tc.CONSTRAINT_SCHEMA = kcu.CONSTRAINT_SCHEMA
                                                              AND tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME
        LEFT JOIN [INFORMATION_SCHEMA].[REFERENTIAL_CONSTRAINTS] fc ON kcu.CONSTRAINT_SCHEMA = fc.CONSTRAINT_SCHEMA
                                                              AND kcu.CONSTRAINT_NAME = fc.CONSTRAINT_NAME
        LEFT JOIN [INFORMATION_SCHEMA].[KEY_COLUMN_USAGE] fkcu ON fkcu.CONSTRAINT_SCHEMA = fc.UNIQUE_CONSTRAINT_SCHEMA
                                                              AND fkcu.CONSTRAINT_NAME = fc.UNIQUE_CONSTRAINT_NAME
        LEFT JOIN sys.extended_properties ep ON OBJECT_ID(c.TABLE_SCHEMA+ '.' +c.TABLE_NAME) = ep.major_id AND c.ORDINAL_POSITION = ep.minor_id
WHERE NOT EXISTS(SELECT 1 FROM INFORMATION_SCHEMA.VIEWS v WHERE v.TABLE_SCHEMA = c.TABLE_SCHEMA AND v.TABLE_NAME =  c.TABLE_NAME) 
ORDER BY c.TABLE_NAME ASC, c.ORDINAL_POSITION ";


        private string path = string.Empty;

        private List<Table_Column> _curDbTabColsList = null;


        public frmFGMDM()
        {
            InitializeComponent();
        }

        private void frmFGMDM_Load(object sender, EventArgs e)
        {

            using (var connsql = SqlClientFactory.Instance.CreateConnection())
            {
                Stopwatch sw = Stopwatch.StartNew();
                connsql.ConnectionString = ConfigurationManager.ConnectionStrings["FGMDMMSSQL"].ConnectionString;
                try
                {
                    connsql.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


                //查询出所有的表

                _curDbTabColsList = connsql.Query<Table_Column>(sqlgetallcolumns).AsList<Table_Column>();

                sw.Stop();
                //this.label1.Text = "";
                //this.label1.Text = string.Format("当前时间：{0}，执行耗时:{1}ms", DateTime.Now, sw.ElapsedMilliseconds);


            }
            if (_curDbTabColsList != null)
            {
                this.dataGridView1.AutoGenerateColumns = true;
                this.dataGridView1.DataSource = _curDbTabColsList;
            }

        }

        private void frmFGMDM_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void btnExportWord_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            CreateWord2();
            this.Cursor = Cursors.Default;
        }

        private void CreateWord2()
        {
            path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }

            bool wordApp = WordHelper.CreateWord(_curDbTabColsList,path);

            if (wordApp)
            {
                MessageBox.Show("success");

            }
            else
            {
                MessageBox.Show("error");
            }
        }



        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + path.ToString();
            System.Diagnostics.Process.Start(psi);
        }
    }
}
