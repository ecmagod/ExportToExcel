using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;

namespace ExcelToSql
{
    public class DBClass
    {
        SqlConnection myConnection = new SqlConnection(System.Configuration.ConfigurationSettings.AppSettings["connStr"]);
        //public int nPublic;
        //SQL reader
        public SqlDataReader SqlGetinfo(string sCmd)
        {
            string cmdText = sCmd;
            SqlDataReader dr = null;
            SqlCommand myCommand = new SqlCommand(cmdText, myConnection);
            try
            {
                myConnection.Open();
                dr = myCommand.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (SqlException ex)
            {
                throw new Exception(ex.Message, ex);
            }
            return dr;
        }
        //DataSet 
        public DataSet DsGetinfo(string sCmd)
        {
            string cmdText = sCmd;
            SqlDataAdapter da = new SqlDataAdapter(cmdText, myConnection);
            DataSet ds = new DataSet();
            try
            {
                myConnection.Open();
                da.Fill(ds, "aa");
            }
            catch (SqlException ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                myConnection.Close();
            }
            return ds;
        }
        //Execute
        public int Executeinfo(string sCmd)
        {
            int nResult = -1;
            string cmdText = sCmd;
            SqlCommand myCommand = new SqlCommand(cmdText, myConnection);
            try
            {
                myConnection.Close();
                myConnection.Open();
                nResult = myCommand.ExecuteNonQuery();
            }
            catch (SqlException ex)
            {
                throw new Exception(ex.Message, ex);
            }
            finally
            {
                myConnection.Close();
            }
            return nResult;
        }
        //加密
        public string base64(string s, bool c)
        {
            if (c)
            {
                return System.Convert.ToBase64String(System.Text.Encoding.Default.GetBytes(s));
            }
            else
            {
                return System.Text.Encoding.Default.GetString(System.Convert.FromBase64String(s));
            }
        }


        /// <summary>
        /// datagridview导出excel
        /// </summary>
        /// <param name="m_DataView"></param>
        /// <param name="heji"></param>
        public void DataToExcel(DataGridView m_DataView, string heji)
        {
            try
            {
                SaveFileDialog kk = new SaveFileDialog();
                kk.AddExtension = true;
                kk.Title = "保存EXECL文件";
                kk.Filter = "EXECL文件(*.xls) |*.xls |所有文件(*.*) |*.*";
                kk.FilterIndex = 1;
                if (kk.ShowDialog() == DialogResult.OK)
                {
                    string FileName = kk.FileName;
                    //if (File.Exists(FileName))//默认覆盖重名文件
                    //    File.Delete(FileName);
                    FileStream objFileStream;
                    StreamWriter objStreamWriter;
                    string strLine = "";
                    objFileStream = new FileStream(FileName, FileMode.OpenOrCreate, FileAccess.Write);
                    objStreamWriter = new StreamWriter(objFileStream, System.Text.Encoding.Unicode);
                    for (int i = 0; i < m_DataView.Columns.Count; i++)
                    {
                        if (m_DataView.Columns[i].Visible == true && m_DataView.Columns[i] is DataGridViewTextBoxColumn) //判断该列是否为打印列
                        {
                            strLine = strLine + m_DataView.Columns[i].HeaderText.ToString() + Convert.ToChar(9);
                        }
                    }
                    objStreamWriter.WriteLine(strLine);
                    strLine = "";
                    for (int i = 0; i < m_DataView.Rows.Count; i++)
                    {
                        if (m_DataView.Columns[0].Visible == true)
                        {
                            if (m_DataView.Rows[i].Cells[0].Value == null)
                                strLine = strLine + " " + Convert.ToChar(9);
                            else
                                strLine = strLine + m_DataView.Rows[i].Cells[0].Value.ToString() + Convert.ToChar(9);
                        }
                        for (int j = 1; j < m_DataView.Columns.Count; j++)
                        {
                            if (m_DataView.Columns[j].Visible == true && m_DataView.Columns[j] is DataGridViewTextBoxColumn)
                            {
                                if (m_DataView.Rows[i].Cells[j].Value == null)
                                    strLine = strLine + " " + Convert.ToChar(9);
                                else
                                {
                                    string rowstr = "";
                                    rowstr = m_DataView.Rows[i].Cells[j].Value.ToString();
                                    if (rowstr.IndexOf("\r\n") > 0)
                                        rowstr = rowstr.Replace("\r\n", " ");
                                    if (rowstr.IndexOf("\t") > 0)
                                        rowstr = rowstr.Replace("\t", " ");
                                    strLine = strLine + rowstr + Convert.ToChar(9);
                                }
                            }
                        }
                        objStreamWriter.WriteLine(strLine);
                        strLine = "";
                    }
                    strLine = heji;
                    objStreamWriter.WriteLine(strLine);
                    strLine = "";
                    objStreamWriter.Close();
                    objFileStream.Close();
                    MessageBox.Show("保存EXCEL成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("文件可能正在打开，请关闭后导出！" + ex.Message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        /// <summary>
        /// 文本框只能输入数字和‘.’
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void txtzhongxiang_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(Char.IsNumber(e.KeyChar)) && e.KeyChar != (char)13 && e.KeyChar != (char)8)//只能输入数字
            {
                int i = ((TextBox)sender).Text.IndexOf('.');
                if (e.KeyChar == '.' && ((TextBox)sender).Text.IndexOf('.') == -1)
                {

                    e.Handled = false;
                }
                else
                {
                    e.Handled = true;
                }
            }
            else//判断回车
            {
                if (e.KeyChar == (char)13)
                {
                    // button3_Click(sender, e);
                }
                else
                    e.Handled = false;
            }
        }

    }
}
