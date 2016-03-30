using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using FG.PDMReader.Oracle;

namespace FG.PDMReader
{
    public partial class frmMain : Form
    {
        XNamespace o = "object";
        XNamespace c = "collection";
        XNamespace a = "attribute";
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnGenScript_Click(object sender, EventArgs e)
        {
            //XDocument.Load()

            FileStream fs = new FileStream(this.textBox1.Text, FileMode.Open);

            XDocument xdoc = XDocument.Load(XmlReader.Create(fs));
            XElement root = xdoc.Root; //获取根元素



            var model = root.Elements().Where(elem => elem.Name.LocalName == "RootObject");//.Element("Model");

            XElement rootobj = root.Element(o + "RootObject").Element(c + "Children").Element(o + "Model").Element(c + "Tables");//.Element(o + "Table");

            //TreeNode treeNode = this.treeView1.Nodes.Add(root.Name.ToString());
            this.richTextBox1.Clear();
            foreach (var item in rootobj.Elements())
            {
                this.richTextBox1.AppendText(GenerateTableScripts(item));
            }

            fs.Close();
            
            
            //LoadNodes(rootobj, treeNode);
        }


        private void LoadNodes(XElement root, TreeNode treeNode)
        {
            foreach (XElement element in root.Elements())
            {
                if (element.Elements().Count() > 0)
                {
                    TreeNode node = treeNode.Nodes.Add(element.Name.ToString());
                    //获取属性
                    foreach (XAttribute attr in element.Attributes())
                    {
                        node.Text += " [" + attr.Name.ToString() + "=" + attr.Value + "]";
                    }
                    LoadNodes(element, node);
                }
                else
                {
                    TreeNode node = treeNode.Nodes.Add(element.Value);
                }
            }
        }


        private string GenerateTableScripts(XElement root)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("CREATE TABLE ");
            sb.Append(root.Element(a + "Code").Value);
            sb.Append(" (" + Environment.NewLine);
            foreach (XElement element in root.Element(c + "Columns").Elements())
            {
                sb.AppendLine(GenerateColumnScripts(element));
            }
            sb.Append(" )" + Environment.NewLine);
            return sb.ToString();
        }

        private string GenerateColumnScripts(XElement root)
        {
            StringBuilder sb = new StringBuilder();

            ColumnInfo sd = new ColumnInfo();

            foreach (XElement element in root.Elements())
            {
                if (element.Name.LocalName == "Name")
                {
                    sd.Name = element.Value;
                }
                if (element.Name.LocalName == "Code")
                {
                    sd.Code = element.Value;
                }
                if (element.Name.LocalName == "DataType")
                {
                    sd.DataType = element.Value;
                }
                if (element.Name.LocalName == "Mandatory")
                {
                    sd.Mandatory = element.Value == "1" ? true:false;
                }
                if (element.Name.LocalName == "Comment")
                {
                    sd.Comment = element.Value;
                }

            }

            return sd.Code + " " + sd.DataType + " " + (sd.Mandatory ? " NOT NULL " : " ") + " , --" + sd.Name + (string.IsNullOrWhiteSpace(sd.Comment) ?"": " ; "+sd.Comment);  
            
        }

        private void btnBrowser_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            DialogResult dr = ofd.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                this.textBox1.Text = ofd.FileName;
            }
        }

        private void btnGenExcel_Click(object sender, EventArgs e)
        {
            try
            {
                PdmReader mTest = new PdmReader(this.textBox1.Text);
                frmExportExcel f2 = new frmExportExcel(mTest, this.textBox1.Text);
                f2.Owner = this;
                f2.Show();
                this.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void tsmiOracleFGMDM_Click(object sender, EventArgs e)
        {
            try
            {
                frmFGMDM fmFgMDM = new frmFGMDM();
                fmFgMDM.Show();
                fmFgMDM.Owner = this;
                this.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void tsmiMSSQLFGMDMDict_Click(object sender, EventArgs e)
        {
            try
            {
                MSSQL.frmFGMDM fmFgMDM = new MSSQL.frmFGMDM();
                fmFgMDM.Show();
                fmFgMDM.Owner = this;
                this.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }
    }
}
