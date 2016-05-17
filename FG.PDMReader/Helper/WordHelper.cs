using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Tables;
using FG.PDMReader.MSSQL;
using System.IO;

namespace FG.PDMReader.Helper
{
    public class WordHelper
    {

        public static bool CreateWord(List<Table_Column> list, string path)
        {

            DocumentBuilder wordApp; //   a   reference   to   Word   application  

            Document wordDoc;//   a   reference   to   the   document
            wordDoc = new Document();//初始化
            wordApp = new DocumentBuilder(wordDoc);

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
               

                if (wordApp != null)
                {
                    wordDoc.Save(path.ToString());
                    return true;

                }
                else
                {
                    return false;
                }


            }
            catch (Exception exception)
            {
                System.Diagnostics.Debug.WriteLine(exception.Message);
                //MessageBox.Show("文档生成失败！(" + exception.Message + ")。\r\n请关闭重试。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }
    }
}
