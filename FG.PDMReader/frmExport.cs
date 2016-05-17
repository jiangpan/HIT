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


namespace FG.PDMReader
{
    public partial class frmExport : Form
    {
        public frmExport()
        {
            InitializeComponent();
        }

        private PdmReader pr = new PdmReader();

        private string path = string.Empty;

        public frmExport(PdmReader pr, string path)
        {
            InitializeComponent();

            this.pr = pr;
            this.path = path;
        }

        private void frmExport_Load(object sender, EventArgs e)
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

        private void frmExport_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();

        }

        private void btnExport2Word_Click(object sender, EventArgs e)
        {

        }
    }
}
