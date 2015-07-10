using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ExcelToSql
{
    public partial class FrmMain : Form
    {
        public FrmMain()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 使用sqlbulkcopy类高效导入
        /// </summary>
        /// <param name="excelFile">excel文件名</param>
        /// <param name="sheetName">sheet表名</param>
        /// <param name="connectionString">连接字符串</param>
         public void TransferData(string excelFile, string sheetName, string connectionString)
        {
            DataSet ds = new DataSet();
            try
            {
                //获取全部数据
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + excelFile + ";" + "Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                strExcel = string.Format("select * from [{0}$]", sheetName);
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                myCommand.Fill(ds, sheetName);

                //如果目标表不存在则创建
                string strSql = string.Format("if object_id('{0}') is null create table {0}(", sheetName);
                foreach (System.Data.DataColumn c in ds.Tables[0].Columns)
                {
                    
                    strSql += string.Format("[{0}] varchar(255),", c.ColumnName);
                }
                strSql = strSql.Trim(',') + ")";

                using (System.Data.SqlClient.SqlConnection sqlconn = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    sqlconn.Open();
                    System.Data.SqlClient.SqlCommand command = sqlconn.CreateCommand();
                    command.CommandText = strSql;
                    command.ExecuteNonQuery();
                    sqlconn.Close();
                }
                //如果excel读出来的dataset有空行，舍掉
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    int flag = 0;//记录每一行空值的个数，等于列数量时，舍掉该列
                    for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                    {
                        string val=ds.Tables[0].Rows[i][j].ToString();
                        if (ds.Tables[0].Rows[i][j].ToString()==""||val=="NULL")
                        {
                            flag++;
                        }
                    }
                    if (flag==ds.Tables[0].Columns.Count)
                    {
                        ds.Tables[0].Rows.Remove(ds.Tables[0].Rows[i]);
                    }
                }
                //用bcp导入数据
                using (System.Data.SqlClient.SqlBulkCopy bcp = new System.Data.SqlClient.SqlBulkCopy(connectionString))
                {
                    bcp.SqlRowsCopied += new System.Data.SqlClient.SqlRowsCopiedEventHandler(bcp_SqlRowsCopied);
                    bcp.BatchSize = 500;//每次传输的行数
                    bcp.NotifyAfter = 100;//进度提示的行数
                    bcp.DestinationTableName = sheetName;//目标表
                    bcp.WriteToServer(ds.Tables[0]);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        //进度显示
        void bcp_SqlRowsCopied(object sender, System.Data.SqlClient.SqlRowsCopiedEventArgs e)
        {
            this.Text = e.RowsCopied.ToString();
            this.Update();
        }
        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnScan_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "EXCEL文件(*.xls)|*.xls|所有文件（*.*）|*.*";
            if (DialogResult.OK==fd.ShowDialog())
            {
                txtFileName.Text = fd.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        /// <summary>
        /// 开始导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            if (txtFileName.Text.Trim()!=""&&txtConnStr.Text.Trim()!=""&&txtSheetName.Text!="")
            {
                TransferData(txtFileName.Text.Trim(), txtSheetName.Text.Trim(), txtConnStr.Text.Trim());
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmUpdate f = new FrmUpdate();
            f.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FrmHtmlSubString f = new FrmHtmlSubString();
            f.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FrmHtmlRegex f = new FrmHtmlRegex();
            f.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FrmHtmlAgilityPack f = new FrmHtmlAgilityPack();
            f.ShowDialog();
        }


    }
} 
    

