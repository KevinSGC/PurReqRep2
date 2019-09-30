using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PurReqRep2
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = openFileDialog2.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //check if the reports has been selected
            if(textBox1.Text=="" || textBox2.Text=="")
            {
                MessageBox.Show("Please select the F/A and MISC PR reports!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //load the FA report
            FileStream fs1 = new FileStream(textBox1.Text,FileMode.Open,FileAccess.Read);
            IWorkbook wb1 = WorkbookFactory.Create(fs1);
            MessageBox.Show(wb1.GetSheetAt(0).SheetName);
            wb1.Close();
            fs1.Close();
        }
    }
}
