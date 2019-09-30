using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
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
            //connect database
            string datasource = "Data Source=" + AppDomain.CurrentDomain.BaseDirectory + "Data\\PurReqRep.db";
            SQLiteConnection conn = new SQLiteConnection(datasource);
            conn.Open();
            SQLiteCommand cmd = conn.CreateCommand();
            string sql = "INSERT INTO MX_PR(DEPT,APPLICANT,WORKFLOW_NO,SAP_PR_NO,SAP_PO_NO,CATEGORY,SPEC,BUYER,CURRENCY,QUANTITY,UNIT,UNIT_PRICE_UN_TAX,TAX_RATE,UNIT_PRICE_TAX,AMOUNT_UN_TAX,AMOUNT_TAX,STATUS,HANDLER,REQUEST_DATE,APPLY_REASON,SITE,CUSTOMER) VALUES (@DEPT,@APPLICANT,@WORKFLOW_NO,@SAP_PR_NO,@SAP_PO_NO,@CATEGORY,@SPEC,@BUYER,@CURRENCY,@QUANTITY,@UNIT,@UNIT_PRICE_UN_TAX,@TAX_RATE,@UNIT_PRICE_TAX,@AMOUNT_UN_TAX,@AMOUNT_TAX,@STATUS,@HANDLER,@REQUEST_DATE,@APPLY_REASON,@SITE,@CUSTOMER);";
            SQLiteTransaction tran = conn.BeginTransaction();
            cmd.Transaction = tran;
            cmd.CommandText = sql;
            cmd.Connection = conn;

            //load the FA report
            toolStripStatusLabel1.Text = "Working on FA report...";
            Application.DoEvents();
            FileStream fs1 = new FileStream(textBox1.Text,FileMode.Open,FileAccess.Read);
            IWorkbook wb1 = WorkbookFactory.Create(fs1);
            ISheet sheet = wb1.GetSheetAt(0);
            int row_cnt = sheet.LastRowNum;
            for(int i=1;i<row_cnt;i++)
            {
                cmd.Parameters.Clear();
                cmd.Parameters.Add("@DEPT", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@APPLICANT", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@WORKFLOW_NO", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@SAP_PR_NO", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(10,MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@SAP_PO_NO", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(10, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CATEGORY", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@SPEC", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(8, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@BUYER", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CURRENCY", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(7, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@QUANTITY", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(20, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                //cmd.Parameters.Add("@UNIT", DbType.String).Value = sheet.GetRow(i).GetCell(10,MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                cmd.Parameters.Add("@UNIT", DbType.String).Value = "";
                cmd.Parameters.Add("@UNIT_PRICE_UN_TAX", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(19, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                //cmd.Parameters.Add("@TAX_RATE", DbType.String).Value = sheet.GetRow(i).GetCell(12).StringCellValue;
                cmd.Parameters.Add("@TAX_RATE", DbType.String).Value = "0";
                cmd.Parameters.Add("@UNIT_PRICE_TAX", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(19, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@AMOUNT_UN_TAX", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(17, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@AMOUNT_TAX", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(17, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@STATUS", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(11, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@HANDLER", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(12, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@REQUEST_DATE", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                //cmd.Parameters.Add("@APPLY_REASON", DbType.String).Value = sheet.GetRow(i).GetCell(19,MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                cmd.Parameters.Add("@APPLY_REASON", DbType.String).Value ="";
                //cmd.Parameters.Add("@SITE", DbType.String).Value = sheet.GetRow(i).GetCell(20,MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                cmd.Parameters.Add("@SITE", DbType.String).Value = "";
                cmd.Parameters.Add("@CUSTOMER", DbType.String).Value = GetCellValue(sheet.GetRow(i).GetCell(15, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.ExecuteNonQuery();
                toolStripStatusLabel1.Text = string.Format("Processing on row {0} of {1}", i, row_cnt);
                Application.DoEvents();
            }
            
            wb1.Close();
            fs1.Close();

            tran.Commit();
            conn.Close();
        }

        private string GetCellValue(ICell cell)
        {
            string result = "";
            if (cell.CellType == CellType.String)
            {
                result = cell.StringCellValue;
            }
            else if(cell.CellType == CellType.Numeric)
            {
                if (DateUtil.IsCellDateFormatted(cell))
                {
                    result = cell.DateCellValue.ToString();
                }
                else
                {
                    result = cell.NumericCellValue.ToString();
                }
            }
            else if(cell.CellType==CellType.Boolean)
            {
                result = cell.BooleanCellValue.ToString();
            }
            else if (cell.CellType == CellType.Blank)
            {
                result = "";
            }
            else if (cell.CellType == CellType.Formula)
            {
                result = cell.StringCellValue;
            }
            else if (cell.CellType == CellType.Error)
            {
                result = "Error";
            }
            else if (cell.CellType == CellType.Unknown)
            {
                result = "Unknown";
            }
            return result;
        }
    }
}
