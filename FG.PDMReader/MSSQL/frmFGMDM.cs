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
using Aspose.Words;
using Aspose.Words.Tables;

namespace FG.PDMReader.MSSQL
{
    public partial class frmFGMDM : Form
    {
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


        private object path = string.Empty;

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
            DocumentBuilder wordApp; //   a   reference   to   Word   application  

            Document wordDoc;//   a   reference   to   the   document  
            path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyyMMddHHmmss") + ".docx");
            wordDoc = new Document();//初始化
            wordApp = new DocumentBuilder(wordDoc);
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }

            wordApp = ThreadWork1(wordApp, _curDbTabColsList);

            if (wordApp != null)
            {
                wordDoc.Save(path.ToString());
                MessageBox.Show("success");
                
            }
            else
            {
                MessageBox.Show("error");
            }
        }

        private DocumentBuilder ThreadWork1(DocumentBuilder wordApp, List<Table_Column> list)
        {
            try
            {
                string text = "数据库名：";
                var gpbDbTabCols = list.GroupBy(p => p.TABLE_NAME);

                int count = gpbDbTabCols.Count();// this.listTable2.Items.Count;

                wordApp.Bold = true;
                wordApp.Font.Size = 12f;
                wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Left;
                // Create the headers.
                wordApp.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
                wordApp.Write("[复高自动生成器www.fugao.com]");
                wordApp.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
                wordApp.Write("[复高自动生成器www.fugao.com]");
                wordApp.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                wordApp.Write("[复高自动生成器www.fugao.com]");

                wordApp.MoveToDocumentStart();
                wordApp.Writeln(text);
                
                //
                int i = 0;
                foreach (var tab in gpbDbTabCols)
                {
                    string tableName = tab.Key;// this.listTable2.Items[i].ToString();
                    object missing = System.Type.Missing;
                    object length = text.Trim().Length;

                    wordApp.InsertBreak(BreakType.LineBreak);
                    wordApp.Writeln("表名：" + tableName);

                    Table tbl = wordApp.StartTable();
                    ParagraphAlignment paragraphAlignmentValue = wordApp.ParagraphFormat.Alignment;
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;

                    wordApp.RowFormat.Height = 25;

                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("序号");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("列名");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("数据类型");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("长度");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("小数位");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("自增列");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("主键");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("允许空");


                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("默认值");

                    wordApp.InsertCell();
                    wordApp.Font.Size = 10.5;
                    wordApp.Font.Name = "宋体";
                    wordApp.CellFormat.VerticalAlignment = Aspose.Words.Tables.CellVerticalAlignment.Center;//垂直居中对齐 
                    wordApp.ParagraphFormat.Alignment = ParagraphAlignment.Center;//水平居中对齐 
                    wordApp.CellFormat.Width = 50.0;
                    wordApp.CellFormat.PreferredWidth = Aspose.Words.Tables.PreferredWidth.FromPoints(50);
                    //设置外框样式    
                    wordApp.CellFormat.Borders.LineStyle = LineStyle.Single;
                    //样式设置结束    
                    wordApp.Write("说明");
                    wordApp.EndRow();
                    //var columnInfoList = list.Where(p => p.TABLE_NAME == "");//this.dbobj.GetColumnInfoList(this.dbname, tableName);
                    int row = 2;
                    foreach (var info in tab)
                    {
                        wordApp.InsertCell();
                        wordApp.Write(info.ORDINAL_POSITION ?? "");// info.ColumnOrder;
                        wordApp.InsertCell();
                        wordApp.Write(info.COLUMN_NAME ?? "");//info.ColumnName;
                        wordApp.InsertCell();
                        wordApp.Write(info.DATA_TYPE ?? "");// info.DATA_TYPE;
                        wordApp.InsertCell();
                        wordApp.Write("");//info.Length;;
                        wordApp.InsertCell();
                        wordApp.Write("");// info.Scale;;
                        wordApp.InsertCell();
                        wordApp.Write(info.IS_IDENTITY ?? "");//(info.IsIdentity.ToString().ToLower() == "true") ? "是" : "";;
                        wordApp.InsertCell();
                        wordApp.Write(info.IS_PRIMARY_KEY ?? "");//(info.IsPrimaryKey.ToString().ToLower() == "true") ? "是" : "";;
                        wordApp.InsertCell();
                        wordApp.Write(info.IS_NULLABLE ?? "");//(info.Nullable.ToString().ToLower() == "true") ? "是" : "否";;
                        wordApp.InsertCell();
                        wordApp.Write(info.COLUMN_DEFAULT ?? "");// info.DefaultVal.ToString();;
                        wordApp.InsertCell();
                        wordApp.Write(info.COLUMN_DESC ?? "");//info.Description.ToString();;

                        wordApp.EndRow();
                    }

                    wordApp.EndTable();

                    wordApp.InsertBreak(BreakType.LineBreak);
                    i++;
                }
                return wordApp;
            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(exception.Message);
                //MessageBox.Show("文档生成失败！(" + exception.Message + ")。\r\n请关闭重试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
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
