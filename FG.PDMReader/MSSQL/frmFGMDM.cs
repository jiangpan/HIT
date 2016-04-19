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
using MSWord = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;

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
            //var tttt = this._curDbTabColsList.GroupBy(p => p.TABLE_NAME);

            //foreach (var item in tttt)
            //{
            //    string sfds = item.Key;
            //    foreach (var tabcol in item)
            //    {
            //        string sssdfe = tabcol.COLUMN_NAME;
            //    }

            //}

            CreateWord1();
            this.Cursor = Cursors.Default;

        }

        private void CreateWord()
        {
            object path;//文件路径
            string strContent;//文件内容
            MSWord.Application wordApp;//Word应用程序变量
            MSWord.Document wordDoc;//Word文档变量
            path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc");//保存为Word2003文档
                                                                                                                         // path = "d:\\myWord.doc";//保存为Word2007文档
            wordApp = new MSWord.ApplicationClass();//初始化
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            //由于使用的是COM 库，因此有许多变量需要用Missing.Value 代替
            Object Nothing = Missing.Value;
            //新建一个word对象
            wordDoc = wordApp.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);

            #region 设置格式
            //页面设置
            wordDoc.PageSetup.PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4;//设置纸张样式
            wordDoc.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientPortrait;//排列方式为垂直方向
            wordDoc.PageSetup.TopMargin = 57.0f;
            wordDoc.PageSetup.BottomMargin = 57.0f;
            wordDoc.PageSetup.LeftMargin = 57.0f;
            wordDoc.PageSetup.RightMargin = 57.0f;
            wordDoc.PageSetup.HeaderDistance = 30.0f;//页眉位置

            //设置页眉
            wordApp.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdOutlineView;//视图样式
            wordApp.ActiveWindow.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekPrimaryHeader;//进入页眉设置，其中页眉边距在页面设置中已完成
            wordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;

            //插入页眉图片(测试结果图片未插入成功)
            wordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.ActiveWindow.ActivePane.Selection.InsertAfter("  文档页眉");
            //去掉页眉的横线
            wordApp.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleNone;
            wordApp.ActiveWindow.ActivePane.Selection.Borders[Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom].Visible = false;
            wordApp.ActiveWindow.ActivePane.View.SeekView = Microsoft.Office.Interop.Word.WdSeekView.wdSeekMainDocument;//退出页眉设置

            //为当前页添加页码
            MSWord.PageNumbers pns = wordApp.Selection.Sections[1].Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers;//获取当前页的号码
            pns.NumberStyle = Microsoft.Office.Interop.Word.WdPageNumberStyle.wdPageNumberStyleNumberInDash;
            pns.HeadingLevelForChapter = 0;
            pns.IncludeChapterNumber = false;
            pns.RestartNumberingAtSection = false;
            pns.StartingNumber = 0;
            object pagenmbetal = Microsoft.Office.Interop.Word.WdPageNumberAlignment.wdAlignPageNumberCenter;//将号码设置在中间
            object first = true;
            wordApp.Selection.Sections[1].Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].PageNumbers.Add(ref pagenmbetal, ref first);

            #endregion


            var gpbDbTabCols = this._curDbTabColsList.GroupBy(p => p.TABLE_NAME);

            var gpbDbtabColsSingle = this._curDbTabColsList.Where(p => p.TABLE_NAME == "FGMDM.AppSysDomain");

            #region 增加数据
            int rows = gpbDbtabColsSingle.Count() + 1;//表格行数加1是为了标题栏
            int cols = 9;//表格列数

            //输出大标题加粗加大字号水平居中
            wordApp.Selection.Font.Bold = 700;
            wordApp.Selection.Font.Size = 16;
            wordApp.Selection.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
            wordApp.Selection.Text = "FGMDM数据字典";
            //换行添加表名
            object lineTableName = MSWord.WdUnits.wdLine;
            wordApp.Selection.MoveDown(ref lineTableName, Nothing, Nothing);
            wordApp.Selection.TypeParagraph();//换行
            MSWord.Range rangeTabName = wordApp.Selection.Range;
            wordApp.Selection.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
            wordApp.Selection.Font.Size = 10;
            wordApp.Selection.Text = "FGMDM.AppSysDomain";



            //换行添加表格
            object line = MSWord.WdUnits.wdLine;
            wordApp.Selection.MoveDown(ref line, Nothing, Nothing);
            wordApp.Selection.TypeParagraph();//换行
            MSWord.Range range = wordApp.Selection.Range;
            MSWord.Table table = wordApp.Selection.Tables.Add(range, rows, cols, ref Nothing, ref Nothing);


            //设置表格的字体大小粗细
            table.Range.Font.Size = 10;
            table.Range.Font.Bold = 0;
            table.Borders.InsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;
            table.Borders.OutsideLineStyle = MSWord.WdLineStyle.wdLineStyleSingle;

            //设置表格标题
            int rowIndex = 1;
            table.Cell(rowIndex, 1).Range.Text = "列名";
            table.Cell(rowIndex, 2).Range.Text = "数据类型";
            table.Cell(rowIndex, 3).Range.Text = "默认值";
            table.Cell(rowIndex, 4).Range.Text = "是否为空";
            table.Cell(rowIndex, 5).Range.Text = "是否主键";
            table.Cell(rowIndex, 6).Range.Text = "是否外键";
            table.Cell(rowIndex, 7).Range.Text = "外键列";
            table.Cell(rowIndex, 8).Range.Text = "外键表";
            table.Cell(rowIndex, 9).Range.Text = "说明";


            //循环数据创建数据行
            rowIndex++;
            foreach (var i in gpbDbtabColsSingle)
            {

                table.Cell(rowIndex, 1).Range.Text = i.COLUMN_NAME;// 列名;
                table.Cell(rowIndex, 2).Range.Text = i.DATA_TYPE;// 数据类型;
                table.Cell(rowIndex, 3).Range.Text = i.COLUMN_DEFAULT;// 默认值;
                table.Cell(rowIndex, 4).Range.Text = i.IS_NULLABLE;// 是否为空;
                table.Cell(rowIndex, 5).Range.Text = i.IS_PRIMARY_KEY;// 是否主键; 
                table.Cell(rowIndex, 6).Range.Text = i.IS_PRIMARY_KEY;// 是否外键; 
                table.Cell(rowIndex, 7).Range.Text = i.FOREIGN_KEY;// 外键列; 
                table.Cell(rowIndex, 8).Range.Text = i.FOREIGN_TABLE;// 外键表; 
                table.Cell(rowIndex, 9).Range.Text = i.COLUMN_DESC;// 说明;

                //table.Cell(rowIndex, 2).Split(2, 1);//分割名字单元格
                //table.Cell(rowIndex, 3).Split(2, 1);//分割成绩单元格

                table.Cell(rowIndex, 1).VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(rowIndex, 4).VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                table.Cell(rowIndex, 5).VerticalAlignment = MSWord.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                //构建姓名，成绩数据
                //foreach (var x in _curDbTabColsList)
                //{
                //    table.Cell(rowIndex, 2).Range.Text = x.TABLE_NAME;
                //    table.Cell(rowIndex, 3).Range.Text = x.COLUMN_NAME;
                //    rowIndex++;
                //}

                rowIndex++;
            }



            #endregion




            //WdSaveDocument为Word2003文档的保存格式(文档后缀.doc)\wdFormatDocumentDefault为Word2007的保存格式(文档后缀.docx)
            object format = MSWord.WdSaveFormat.wdFormatDocument;
            //将wordDoc 文档对象的内容保存为DOC 文档,并保存到path指定的路径
            wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
            //关闭wordDoc文档
            wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        private void CreateWord1()
        {
            string strContent;//文件内容
            MSWord.Application wordApp;//Word应用程序变量
            MSWord.Document wordDoc;//Word文档变量
            path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc");//保存为Word2003文档
                                                                                                                         // path = "d:\\myWord.doc";//保存为Word2007文档
            wordApp = new MSWord.ApplicationClass();//初始化
            if (File.Exists((string)path))
            {
                File.Delete((string)path);
            }
            //由于使用的是COM 库，因此有许多变量需要用Missing.Value 代替
            Object Nothing = Missing.Value;


            wordDoc = ThreadWork1(wordApp, _curDbTabColsList);

            if (wordDoc != null)
            {
                //WdSaveDocument为Word2003文档的保存格式(文档后缀.doc)\wdFormatDocumentDefault为Word2007的保存格式(文档后缀.docx)
                object format = MSWord.WdSaveFormat.wdFormatDocument;
                //将wordDoc 文档对象的内容保存为DOC 文档,并保存到path指定的路径
                wordDoc.SaveAs(ref path, ref format, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                //关闭wordDoc文档
                wordDoc.Close(ref Nothing, ref Nothing, ref Nothing);
                return;
            }
            //关闭wordApp组件对象
            wordApp.Quit(ref Nothing, ref Nothing, ref Nothing);
        }


        private void ThreadWork(List<Table_Column> list)
        {
            try
            {
                //this.SetBtnDisable();
                string str = "数据库名：";// + this.dbname;
                int count = 0;// this.listTable2.Items.Count;
                //this.SetprogressBar1Max(count);
                //this.SetprogressBar1Val(1);
                //this.SetlblStatuText("0");
                object template = Missing.Value;
                object obj3 = @"\endofdoc";
                MSWord.Application application = new MSWord.ApplicationClass { Visible = false };
                MSWord.Document document = application.Documents.Add(ref template, ref template, ref template, ref template);
                application.ActiveWindow.View.Type = MSWord.WdViewType.wdOutlineView;
                application.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekPrimaryHeader;
                application.ActiveWindow.ActivePane.Selection.InsertAfter("动软自动生成器 www.maticsoft.com");
                application.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                application.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                MSWord.Paragraph paragraph = document.Content.Paragraphs.Add(ref template);
                paragraph.Range.Text = str;
                paragraph.Range.Font.Bold = 1;
                paragraph.Range.Font.Name = "宋体";
                paragraph.Range.Font.Size = 12f;
                paragraph.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                paragraph.Format.SpaceAfter = 5f;
                paragraph.Range.InsertParagraphAfter();
                DataTable tablesExProperty = null;// this.dbobj.GetTablesExProperty(this.dbname);
                for (int i = 0; i < count; i++)
                {
                    string tableName = "";// this.listTable2.Items[i].ToString();
                    string str3 = "表名：" + tableName;
                    List<Table_Column> columnInfoList = list;//this.dbobj.GetColumnInfoList(this.dbname, tableName);
                    int num3 = columnInfoList.Count;
                    if ((columnInfoList != null) && (columnInfoList.Count > 0))
                    {
                        object obj4 = document.Bookmarks[ref obj3].Range;
                        MSWord.Paragraph paragraph2 = document.Content.Paragraphs.Add(ref obj4);
                        paragraph2.Range.Text = str3;
                        paragraph2.Range.Font.Bold = 1;
                        paragraph2.Range.Font.Name = "宋体";
                        paragraph2.Range.Font.Size = 10f;
                        paragraph2.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        paragraph2.Format.SpaceBefore = 15f;
                        paragraph2.Format.SpaceAfter = 1f;
                        paragraph2.Range.InsertParagraphAfter();
                        string str4 = "";
                        if (tablesExProperty != null)
                        {
                            try
                            {
                                DataRow[] rowArray = tablesExProperty.Select("objname='" + tableName + "'");
                                if ((rowArray.Length > 0) && (rowArray[0]["value"] != null))
                                {
                                    str4 = rowArray[0]["value"].ToString();
                                }
                            }
                            catch
                            {
                            }
                        }
                        obj4 = document.Bookmarks[ref obj3].Range;
                        MSWord.Paragraph paragraph3 = document.Content.Paragraphs.Add(ref obj4);
                        paragraph3.Range.Text = str4;
                        paragraph3.Range.Font.Bold = 0;
                        paragraph3.Range.Font.Name = "宋体";
                        paragraph3.Range.Font.Size = 9f;
                        paragraph3.Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                        paragraph3.Format.SpaceBefore = 1f;
                        paragraph3.Format.SpaceAfter = 1f;
                        paragraph3.Range.InsertParagraphAfter();
                        MSWord.Range range = document.Bookmarks[ref obj3].Range;
                        MSWord.Table table2 = document.Tables.Add(range, num3 + 1, 11, ref template, ref template);
                        table2.Range.Font.Name = "宋体";
                        table2.Range.Font.Size = 9f;
                        table2.Borders.Enable = 1;
                        table2.Rows.Height = 10f;
                        table2.AllowAutoFit = true;
                        table2.Cell(1, 1).Range.Text = "序号";
                        table2.Cell(1, 2).Range.Text = "列名";
                        table2.Cell(1, 3).Range.Text = "数据类型";
                        table2.Cell(1, 4).Range.Text = "长度";
                        table2.Cell(1, 5).Range.Text = "小数位";
                        table2.Cell(1, 6).Range.Text = "标识";
                        table2.Cell(1, 7).Range.Text = "主键";
                        table2.Cell(1, 8).Range.Text = "外键";
                        table2.Cell(1, 9).Range.Text = "允许空";
                        table2.Cell(1, 10).Range.Text = "默认值";
                        table2.Cell(1, 11).Range.Text = "说明";
                        table2.Columns[1].Width = 33f;
                        table2.Columns[3].Width = 60f;
                        table2.Columns[4].Width = 33f;
                        table2.Columns[5].Width = 43f;
                        table2.Columns[6].Width = 33f;
                        table2.Columns[7].Width = 33f;
                        table2.Columns[8].Width = 43f;
                        table2.Columns[9].Width = 33f;
                        for (int j = 0; j < num3; j++)
                        {
                            Table_Column info = columnInfoList[j];
                            string columnOrder = "";// info.ColumnOrder;
                            string columnName = info.COLUMN_NAME;// info.ColumnName;
                            string typeName = info.DATA_TYPE;//info.TypeName;
                            string length = "";//info.Length;
                            string scale = "";//info.Scale;
                            string str10 = "";//(info.IsIdentity.ToString().ToLower() == "true") ? "是" : "";
                            string str11 = "";//(info.IsPrimaryKey.ToString().ToLower() == "true") ? "是" : "";
                            string str12 = "";// (info.IsForeignKey.ToString().ToLower() == "true") ? "是" : "";
                            string str13 = "";//(info.Nullable.ToString().ToLower() == "true") ? "是" : "否";
                            string str14 = "";//info.DefaultVal.ToString();
                            string str15 = "";//info.Description.ToString();
                            if (length.Trim() == "-1")
                            {
                                length = "MAX";
                            }
                            table2.Cell(j + 2, 1).Range.Text = columnOrder;
                            table2.Cell(j + 2, 2).Range.Text = columnName;
                            table2.Cell(j + 2, 3).Range.Text = typeName;
                            table2.Cell(j + 2, 4).Range.Text = length;
                            table2.Cell(j + 2, 5).Range.Text = scale;
                            table2.Cell(j + 2, 6).Range.Text = str10;
                            table2.Cell(j + 2, 7).Range.Text = str11;
                            table2.Cell(j + 2, 8).Range.Text = str12;
                            table2.Cell(j + 2, 9).Range.Text = str13;
                            table2.Cell(j + 2, 10).Range.Text = str14;
                            table2.Cell(j + 2, 11).Range.Text = str15;
                        }
                        table2.Rows[1].Range.Font.Bold = 1;
                        table2.Rows[1].Range.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                        table2.Rows.First.Shading.Texture = MSWord.WdTextureIndex.wdTexture25Percent;
                    }
                    //this.SetprogressBar1Val(i + 1);
                    //this.SetlblStatuText((i + 1).ToString());
                }
                application.Visible = true;
                document.Activate();
                //this.SetBtnEnable();
                MessageBox.Show("文档生成成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch (Exception exception)
            {
                MessageBox.Show("文档生成失败！(" + exception.Message + ")。\r\n请关闭重试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private MSWord.Document ThreadWork1(MSWord.Application docApp, List<Table_Column> list)
        {
            try
            {
                //this.SetBtnDisable();
                string text = "数据库名：";
                var gpbDbTabCols = list.GroupBy(p => p.TABLE_NAME);

                int count = gpbDbTabCols.Count();// this.listTable2.Items.Count;
                //this.SetprogressBar1Max(count);
                //this.SetprogressBar1Val(1);
                //this.SetlblStatuText("0");
                object template = Missing.Value;
                MSWord.Application WordApp = docApp;//初始化


                MSWord.Document document = WordApp.Documents.Add(ref template, ref template, ref template, ref template);
                WordApp.ActiveWindow.View.Type = MSWord.WdViewType.wdOutlineView;
                WordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekPrimaryHeader;
                WordApp.ActiveWindow.ActivePane.Selection.InsertAfter("[复高自动生成器www.fugao.com]");
                WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphRight;
                WordApp.ActiveWindow.View.SeekView = MSWord.WdSeekView.wdSeekMainDocument;
                WordApp.Selection.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphCenter;
                WordApp.Selection.Font.Bold = 0x98967e;
                WordApp.Selection.Font.Size = 12f;
                WordApp.Selection.TypeText(text);
                //
                int i = 0;
                foreach (var tab in gpbDbTabCols)
                {
                    string tableName = tab.Key;// this.listTable2.Items[i].ToString();
                    object missing = System.Type.Missing;
                    object length = text.Trim().Length;
                    MSWord.Range range = document.Range(ref length, ref length);
                    WordApp.Selection.Tables.Add(range, 2, 10, ref template, ref template);
                    object type = 11;
                    range.InsertBreak(ref type);
                    range.InsertBreak(ref type);
                    range.InsertBefore("表名：" + tableName);
                    MSWord.Table table = document.Tables[1];
                    table.Borders.Enable = 1;
                    table.AllowAutoFit = true;
                    table.Rows.Height = 15f;
                    MSWord.Range range2 = table.Rows[1].Range;
                    range2.Font.Size = 9f;
                    range2.Font.Name = "宋体";
                    range2.Font.Bold = 1;
                    table.Cell(1, 1).Range.Text = "序号";
                    table.Cell(1, 2).Range.Text = "列名";
                    table.Cell(1, 3).Range.Text = "数据类型";
                    table.Cell(1, 4).Range.Text = "长度";
                    table.Cell(1, 5).Range.Text = "小数位";
                    table.Cell(1, 6).Range.Text = "自增列";
                    table.Cell(1, 7).Range.Text = "主键";
                    table.Cell(1, 8).Range.Text = "允许空";
                    table.Cell(1, 9).Range.Text = "默认值";
                    table.Cell(1, 10).Range.Text = "说明";
                    table.Columns[1].Width = 33f;
                    table.Columns[3].Width = 60f;
                    table.Columns[4].Width = 33f;
                    table.Columns[5].Width = 43f;
                    table.Columns[6].Width = 33f;
                    table.Columns[7].Width = 33f;
                    table.Columns[8].Width = 43f;
                    //var columnInfoList = list.Where(p => p.TABLE_NAME == "");//this.dbobj.GetColumnInfoList(this.dbname, tableName);
                    int row = 2;
                    foreach (var info in tab)
                    {
                        object beforeRow = System.Type.Missing;
                        table.Rows.Add(ref beforeRow);
                        table.Rows[row].Range.Select();
                        MSWord.Range range3 = table.Rows[row].Range;
                        range3.Font.Size = 9f;
                        range3.Font.Name = "宋体";
                        range3.Font.Bold = 0;
                        range3.ParagraphFormat.Alignment = MSWord.WdParagraphAlignment.wdAlignParagraphLeft;
                        table.Cell(row, 1).Range.Text = info.ORDINAL_POSITION;// info.ColumnOrder;
                        table.Cell(row, 2).Range.Text = info.COLUMN_NAME;//info.ColumnName;
                        table.Cell(row, 3).Range.Text = info.DATA_TYPE;// info.DATA_TYPE;
                        table.Cell(row, 4).Range.Text = "";//info.Length;;
                        table.Cell(row, 5).Range.Text = "";// info.Scale;;
                        table.Cell(row, 6).Range.Text = info.IS_IDENTITY;//(info.IsIdentity.ToString().ToLower() == "true") ? "是" : "";;
                        table.Cell(row, 7).Range.Text = info.IS_PRIMARY_KEY;//(info.IsPrimaryKey.ToString().ToLower() == "true") ? "是" : "";;
                        table.Cell(row, 8).Range.Text = info.IS_NULLABLE;//(info.Nullable.ToString().ToLower() == "true") ? "是" : "否";;
                        table.Cell(row, 9).Range.Text = info.COLUMN_DEFAULT;// info.DefaultVal.ToString();;
                        table.Cell(row, 10).Range.Text = info.COLUMN_DESC;//info.Description.ToString();;
                        row++;
                    }
                    //this.SetprogressBar1Val(i + 1);
                    //this.SetlblStatuText((i + 1).ToString());
                    table.Rows.First.Shading.Texture = MSWord.WdTextureIndex.wdTexture25Percent;
                    i++;
                }
                WordApp.Visible = true;

                return document;

                //this.SetBtnEnable();
                //MessageBox.Show("文档生成成功！", "完成", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
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
