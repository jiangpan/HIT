using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace FG.PDMReader
{
    public partial class frmExportExcel : Form
    {
        public frmExportExcel()
        {
            InitializeComponent();
        }

        private PdmReader pr = new PdmReader();

        private string path = string.Empty;

        public frmExportExcel(PdmReader pr, string path)
        {
            InitializeComponent();

            this.pr = pr;
            this.path = path;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            pr.InitData();

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = pr.Tables;
            


            string[] names = null;
            List<string> hms = new List<string>();
            List<string> res = null;
            for (int i = 0; names != null && i < names.Length; i++)
            {
                string[] str = names[i].Split(new char[] { '!' }, StringSplitOptions.RemoveEmptyEntries);
                res = hms.FindAll(M => M.Equals(str[0].Substring(24)));
                if (!(res != null && res.Count > 0))
                {
                    hms.Add(str[0].Substring(24));
                }
            }
            foreach (string s in hms)
            {
                this.comboBox1.Items.Add(s);
            }
        }

        private void btnReturn_Click(object sender, EventArgs e)
        {
            this.Visible = false;
            if (this.Owner != null)
            {
                this.Owner.Show();


            }
            else
            {
                this.Close();
            }
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.SelectedItem != null && !this.comboBox1.SelectedItem.Equals(""))
            {
                FileInfo file = new FileInfo(path);
                if (!Directory.Exists("D:\\work\\Test\\XML\\" + file.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0]))
                {
                    Directory.CreateDirectory("D:\\work\\Test\\XML\\" + file.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0]);
                }
                if (Directory.Exists("D:\\work\\Test\\XML\\" + file.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0] + "\\" + this.comboBox1.SelectedItem.ToString() + ".txt"))
                {
                    File.Delete("D:\\work\\Test\\XML\\" + file.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0] + "\\" + this.comboBox1.SelectedItem.ToString() + ".txt");
                }
                string str = string.Empty;
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    if (dr.Cells[0].Value != null && int.Parse(dr.Cells[0].Value.ToString()) == 1)
                    {
                        str += dr.Cells[1].Value.ToString() + "|";
                    }
                }
                if (str.Length > 0)
                {
                    str = str.Substring(0, str.Length - 1);
                }
                FileStream fs = new FileStream("D:\\work\\Test\\XML\\" + file.Name.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries)[0] + "\\" + this.comboBox1.SelectedItem.ToString() + ".txt", FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(str);
                sw.Flush();
                sw.Close();
                fs.Close();
            }
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            int n = 1;
            FileInfo file = new FileInfo(path);
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();
            app.Visible = false;
            Workbook wb = app.Workbooks.Add(true);
            Worksheet ws = (Worksheet)wb.ActiveSheet;
            ws.Name = "所有表";
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["B", Type.Missing]).ColumnWidth = 15.50;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["A", Type.Missing]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["E", Type.Missing]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["F", Type.Missing]).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["C", Type.Missing]).ColumnWidth = 16.5;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["D", Type.Missing]).ColumnWidth = 8.50;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["E", Type.Missing]).ColumnWidth = 8.5;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["F", Type.Missing]).ColumnWidth = 8.50;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["G", Type.Missing]).ColumnWidth = 8.50;
            ((Microsoft.Office.Interop.Excel.Range)ws.Columns["H", Type.Missing]).ColumnWidth = 33.50;
            Range r = ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 8]);
            r.Interior.ColorIndex = 37;
            r.Font.Size = 12;
            r.Font.Bold = true;
            Borders borders = r.Borders;
            borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            r.Merge(false);
            r.Value = "数据字典";
            //ws.Cells[n, 1] = "序号";
            //ws.Cells[n, 2] = "表名";
            //ws.Cells[n, 3] = "字段名";
            //ws.Cells[n, 4] = "表说明";
            //ws.Cells[n, 5] = "字段类型";
            //ws.Cells[n, 6] = "长度";
            //ws.Cells[n, 7] = "允许空";
            
            n = 2;

            List<string> list = new List<string>();
            for (int i = 0; i < this.dataGridView1.Rows.Count; i++)
            {
                DataGridViewRow dr = this.dataGridView1.Rows[i];
                string str = dr.Cells[1].Value.ToString();
                if (dr.Cells["cbxTable"].Value != null && dr.Cells["cbxTable"].Value.ToString().Equals("1"))
                {
                    list.Add(dr.Cells[1].Value.ToString());
                }
            }

            foreach (TableInfo table in pr.Tables)
            {
                bool print = false;
                if (this.checkBox1.Checked)
                {
                    print = true;
                }
                else
                {
                    foreach (string s in list)
                    {
                        if (s.Equals(table.Code))
                        {
                            print = true;
                        }
                    }
                }
                if (print)
                {
                    Range rt = ws.get_Range(ws.Cells[n, 1], ws.Cells[n, 8]);
                    rt.Interior.ColorIndex = 35;
                    rt.Font.Size = 12;
                    Borders border = rt.Borders;
                    border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    Range tblRange = ws.get_Range(ws.Cells[n, 1], ws.Cells[n, 8]);
                    tblRange.Merge(false);
                    tblRange.Value = "表名：" + table.Code + " " + table.Name;
                    tblRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft; //水平左对齐
                    //ws.Cells[n, 1] = table.Code; //"T";
                    //ws.Cells[n, 2] = table.Code;
                    //ws.Cells[n, 3] = table.Name;
                    //ws.Cells[n, 4] = table.Comment;
                    //ws.Cells[n, 5] = "";
                    //ws.Cells[n, 6] = "";
                    //ws.Cells[n, 7] = "";
                    n = n + 1;

                    //设置每个表的所有列的 表头
                    Range rngColsTitle = ws.get_Range(ws.Cells[n, 1], ws.Cells[n, 8]);
                    rngColsTitle.Interior.ColorIndex = 37;
                    rngColsTitle.Font.Size = 12;
                    rngColsTitle.Font.Bold = true;
                    Borders borders1 = rngColsTitle.Borders;
                    borders1.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    ws.Cells[n, 1] = "序号";
                    ws.Cells[n, 2] = "字段名";
                    ws.Cells[n, 3] = "作用";
                    ws.Cells[n, 4] = "类型";
                    ws.Cells[n, 5] = "长度";
                    ws.Cells[n, 6] = "是否为空";
                    ws.Cells[n, 7] = "默认值";
                    ws.Cells[n, 8] = "说明";

                    n = n + 1;

                    for (int i = 0; i <= table.Columns.Count; i++)
                    {
                        if (i == table.Columns.Count - 1)
                        {
                            Range rtcBlankLine = ws.get_Range(ws.Cells[n, 1], ws.Cells[n, 8]);
                            rtcBlankLine.Merge(false);
                            Borders rtcBlankLineborderc = rtcBlankLine.Borders;
                            rtcBlankLineborderc.get_Item(XlBordersIndex.xlEdgeTop).LineStyle = XlLineStyle.xlContinuous;
                            //rtcBlankLineborderc.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//设置所有边框
                            n = n + 1;
                            break;
                        }

                        Range rtc = ws.get_Range(ws.Cells[n, 1], ws.Cells[n, 8]);
                        rtc.Font.Size = 12;
                        Borders borderc = rtc.Borders;
                        borderc.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        ws.Cells[n, 1] = i + 1;
                        ws.Cells[n, 2] = table.Columns[i].Code;
                        ws.Cells[n, 3] = table.Columns[i].Name;
                        

                        //ws.Cells[n, 4] = table.Columns[i].Comment;
                        ws.Cells[n, 4] = table.Columns[i].DataType.Contains("(") ? table.Columns[i].DataType.Split(new char[] { '(' }, StringSplitOptions.RemoveEmptyEntries)[0] : table.Columns[i].DataType;
                        if (table.Columns[i].DataType.Equals("int"))
                        {
                            ws.Cells[n, 5] = 10;
                        }
                        else if (table.Columns[i].DataType.Equals("datetime"))
                        {
                            ws.Cells[n, 5] = 23;
                        }
                        else
                        {
                            ws.Cells[n, 5] = table.Columns[i].Length;
                        }
                        ws.Cells[n, 6] = table.Columns[i].Mandatory ? "" : "√";
                        //ws.Cells[n, 7] = table.Columns[i].; //默认值
                        ws.Cells[n, 8] = table.Columns[i].Comment;


                        if (table.Primary != null)
                        {
                            foreach (string pk in table.Primary)
                            {
                                if (pk.Equals(table.Columns[i].ColumnId))
                                {
                                    rtc.Interior.ColorIndex = 6;
                                }
                            }
                        }

                        


                        n = n + 1;

                        
                    }
                }
            }

            wb.Saved = true;

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel 2003 文件(*.xls)|*.xls|所有文件(*.*)|*.*";
            sfd.FilterIndex = 1;
            sfd.RestoreDirectory = true;
            DialogResult drSfd = sfd.ShowDialog();
            
            if (drSfd == System.Windows.Forms.DialogResult.OK)
            {
                app.ActiveWorkbook.SaveCopyAs(sfd.FileName);
            }
            MessageBox.Show("保存成功！");
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }


    }
}
