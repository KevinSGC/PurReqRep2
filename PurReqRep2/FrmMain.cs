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
            string sql = "INSERT INTO MX_PR(WORKFLOW_NO,APPLICANT,CSR,DEPARTMENT,APPLY_DATE,CATEGORY,TOTAL_AMT_EST,CURRENCY,PROD_NAME,PROD_SPEC,FA_NO,SAP_PO_NO,STATUS,HANDLER,BU,PAY_BY_CUST,CUSTOMER,ME_TYPE,TOTAL_AMT,TAX_RATE,VENDOR,UNIT_PRICE,QTY,V_SPEC,INVOICE,CUST_CURRENCY,CUST_AMOUNT,REQUEST_DATE,ETD) VALUES (@WORKFLOW_NO,@APPLICANT,@CSR,@DEPARTMENT,@APPLY_DATE,@CATEGORY,@TOTAL_AMT_EST,@CURRENCY,@PROD_NAME,@PROD_SPEC,@FA_NO,@SAP_PO_NO,@STATUS,@HANDLER,@BU,@PAY_BY_CUST,@CUSTOMER,@ME_TYPE,@TOTAL_AMT,@TAX_RATE,@VENDOR,@UNIT_PRICE,@QTY,@V_SPEC,@INVOICE,@CUST_CURRENCY,@CUST_AMOUNT,@REQUEST_DATE,@ETD);";
            SQLiteTransaction tran = conn.BeginTransaction();
            cmd.Transaction = tran;
            cmd.CommandText = sql;
            cmd.Connection = conn;

            //load the FA report
            toolStripStatusLabel1.Text = "Working on FA report...";
            Application.DoEvents();
            FileStream fs1 = new FileStream(textBox1.Text,FileMode.Open,FileAccess.Read);
            IWorkbook wb1 = WorkbookFactory.Create(fs1);
            ISheet sheet1 = wb1.GetSheetAt(0);
            int row_cnt = sheet1.LastRowNum;
            for(int i=1;i<row_cnt;i++)
            {
                cmd.Parameters.Clear();

                cmd.Parameters.Add("@WORKFLOW_NO", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(0, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@APPLICANT", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CSR", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@DEPARTMENT", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(3, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@APPLY_DATE", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CATEGORY", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(5, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@TOTAL_AMT_EST", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(6, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CURRENCY", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(7, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PROD_NAME", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(8, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PROD_SPEC", DbType.String).Value = "";
                cmd.Parameters.Add("@FA_NO", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(9, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@SAP_PO_NO", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(10, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@STATUS", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(11, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@HANDLER", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(12, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@BU", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(13, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PAY_BY_CUST", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(14, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CUSTOMER", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(15, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@ME_TYPE", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(16, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@TOTAL_AMT", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(17, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@TAX_RATE", DbType.String).Value = "";
                cmd.Parameters.Add("@VENDOR", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(18, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@UNIT_PRICE", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(19, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@QTY", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(20, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@V_SPEC", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(21, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@INVOICE", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(22, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CUST_CURRENCY", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(23, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CUST_AMOUNT", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(24, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@REQUEST_DATE", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(25, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@ETD", DbType.String).Value = GetCellValue(sheet1.GetRow(i).GetCell(26, MissingCellPolicy.CREATE_NULL_AS_BLANK));

                cmd.ExecuteNonQuery();
                toolStripStatusLabel1.Text = string.Format("Processing on row {0} of {1}", i, row_cnt);
                Application.DoEvents();
            }
            
            wb1.Close();
            fs1.Close();

            //LOAD MISC PURCHASE REQUEST
            FileStream fs2 = new FileStream(textBox2.Text, FileMode.Open, FileAccess.Read);
            IWorkbook wb2 = WorkbookFactory.Create(fs2);
            ISheet sheet2 = wb2.GetSheetAt(0);
            row_cnt = sheet2.LastRowNum;
            for (int i = 1; i < row_cnt; i++)
            {
                cmd.Parameters.Clear();

                cmd.Parameters.Add("@WORKFLOW_NO", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(4, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@APPLICANT", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(2, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@CSR", DbType.String).Value = "";
                cmd.Parameters.Add("@DEPARTMENT", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(1, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@APPLY_DATE", DbType.String).Value = "";
                cmd.Parameters.Add("@CATEGORY", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(7, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@TOTAL_AMT_EST", DbType.String).Value = "";
                cmd.Parameters.Add("@CURRENCY", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(11, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PROD_NAME", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(8, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PROD_SPEC", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(9, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@FA_NO", DbType.String).Value = "";
                cmd.Parameters.Add("@SAP_PO_NO", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(6, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@STATUS", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(19, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@HANDLER", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(20, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@BU", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(23, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@PAY_BY_CUST", DbType.String).Value = "";
                cmd.Parameters.Add("@CUSTOMER", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(24, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@ME_TYPE", DbType.String).Value = "";
                cmd.Parameters.Add("@TOTAL_AMT", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(18, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@TAX_RATE", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(15, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@VENDOR", DbType.String).Value = "";
                cmd.Parameters.Add("@UNIT_PRICE", DbType.String).Value = "";
                cmd.Parameters.Add("@QTY", DbType.String).Value = "";
                cmd.Parameters.Add("@V_SPEC", DbType.String).Value = "";
                cmd.Parameters.Add("@INVOICE", DbType.String).Value = "";
                cmd.Parameters.Add("@CUST_CURRENCY", DbType.String).Value = "";
                cmd.Parameters.Add("@CUST_AMOUNT", DbType.String).Value = "";
                cmd.Parameters.Add("@REQUEST_DATE", DbType.String).Value = GetCellValue(sheet2.GetRow(i).GetCell(21, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                cmd.Parameters.Add("@ETD", DbType.String).Value = "";

                cmd.ExecuteNonQuery();
                toolStripStatusLabel1.Text = string.Format("Processing on row {0} of {1}", i, row_cnt);
                Application.DoEvents();
            }

            wb2.Close();
            fs2.Close();

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
