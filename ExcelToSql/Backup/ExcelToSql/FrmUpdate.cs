using System;
using System.Configuration;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ExcelToSql
{
    public partial class FrmUpdate : Form
    {
        public FrmUpdate()
        {
            InitializeComponent();
        }

        private void btnScan_Click(object sender, EventArgs e)
        {

            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "EXCEL文件(*.xls)|*.xls|所有文件（*.*）|*.*";
            if (DialogResult.OK == fd.ShowDialog())
            {
                txtFileName.Text = fd.FileName;
            }
        }
        //导入EXCEL
        public DataSet ImportExcel(string fileName)
        {
            //string fileName = "d:\\123.xls";
            string excelStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1'";//execl 2003
            //string excelStr = "Provider= Microsoft.Ace.OleDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";//execl 2007以上（需要装个AccessDatabaseEngine引擎，网上找找）
            DataSet ds = new DataSet();
            using (System.Data.OleDb.OleDbConnection cn = new OleDbConnection(excelStr))
            {
                using (OleDbDataAdapter dr = new OleDbDataAdapter("SELECT * FROM ["+txtSheetName.Text+"$]", excelStr))
                {
                    dr.Fill(ds);
                    return ds;
                }
            }
            //插入到数据库
        }
        /// <summary>
        /// 逐行导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked==true)
            {
                 Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);    
            //获取appSettings节点   
            AppSettingsSection appSettings = (AppSettingsSection)config.GetSection("appSettings");
                //删除name，然后添加新值    
                appSettings.Settings.Remove("connStr");
                appSettings.Settings.Add("connStr", txtConnStr.Text.Trim());
            //保存配置文件   
            config.Save();
            // 强制重新载入配置文件的ConnectionStrings配置节 
            ConfigurationManager.RefreshSection("appSettings");
            }
            if (txtFileName.Text.Trim() == "" || txtConnStr.Text.Trim() == "" || txtSheetName.Text.Trim() == "")
            {
                return;
            }
            DataTable dt= ImportExcel(txtFileName.Text.Trim()).Tables[0];
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                label4.Text = "完成：" + i.ToString() + "条 / 共" + (dt.Rows.Count-1).ToString() + "条 ";
                string sql_col = "";
                //确定sql列数
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j < dt.Columns.Count - 1)
                    {
                        sql_col += "'" + dt.Rows[i][j].ToString() + "',";
                    }
                    else
                    {
                        sql_col += "'" + dt.Rows[i][j].ToString() + "'";
                    }
                }
                if (sql_col == "")
                {
                    return;
                }
                string sql = "insert into " + txtSheetName.Text + " values(" + sql_col + ")";
                SqlConnection myConnection = new SqlConnection(txtConnStr.Text);
                string cmdText = sql;
                SqlCommand myCommand = new SqlCommand(cmdText, myConnection);
                try
                {
                    myConnection.Close();
                    myConnection.Open();
                    myCommand.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    throw new Exception(ex.Message, ex);
                }
                finally
                {
                    myConnection.Close();
                }
            }
            
        }

        private void FrmUpdate_Load(object sender, EventArgs e)
        {
            txtConnStr.Text = ConfigurationSettings.AppSettings["connStr"].ToString();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtSheetName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtConnStr_TextChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void txtFileName_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
