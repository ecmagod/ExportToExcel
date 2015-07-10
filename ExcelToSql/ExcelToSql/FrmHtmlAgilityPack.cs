using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using HtmlAgilityPack;
using System.IO;
using System.Data.SqlClient;

namespace ExcelToSql
{
    public partial class FrmHtmlAgilityPack : Form
    {
        public FrmHtmlAgilityPack()
        {
            InitializeComponent();
            this.skinEngine1.SkinFile = "DeepCyan.ssk";
        }
        //进度条最大值，文件中的车数
        private int CountTemp = 10000;
        //导入前前数据库中的最大ID
        private string idTemp;
        //插入数据库、更改控件进程
        private Thread t;
        /// <summary>
        /// 选择要导入的文件，开始导入子线程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            //提示确认导入该文件
            DialogResult Dr = MessageBox.Show("确定将 " + cboFileName.Text.Trim() + " 导入到系统中？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            //按钮变化
            SetControl(btnImport, false);
            SetControl(btnExit, false);
            //SetControl(btnStop, true);
            if (Dr == DialogResult.No)
            {
                SetControl(btnImport, true);
                SetControl(btnExit, true);
                SetControl(btnStop, false);
                return;
            }
            t = new Thread(new ThreadStart(importToSql));
            t.Start();
        }

        /// <summary>
        /// 导入主方法
        /// </summary>
        private void importToSql()
        {
             GetTextValue(cboFileName);
             string FilePath = controlTextValue;
            if (File.Exists(FilePath))//存在此文件
            {
                SetSumLabelValue("正在确认导入的车辆总数···");
                //HtmlAgilityPack自带加载html为htmlDocument
                HtmlAgilityPack.HtmlDocument hd = new HtmlAgilityPack.HtmlDocument();
                hd.Load(FilePath, UTF8Encoding.UTF8);
                HtmlNode rootNode = hd.DocumentNode;
                HtmlNodeCollection categoryNodeList = rootNode.SelectNodes("//font[@*]");//根据xpath获取节点树
                if (categoryNodeList == null)
                {
                    return;
                }

                //foreach读取效率比for、while高
                //for (int i = categoryNodeList.Count-1; i >=0; i--)
                //{
                //    if (categoryNodeList[i].InnerText.Contains("车数"))
                //    {
                //         CountTemp = Int32.Parse(categoryNodeList[i + 1].InnerText.Trim());
                //         break;
                //    }
                //}
                foreach (HtmlNode item in categoryNodeList)
                {
                    if (item.InnerText.Contains("车数"))
                    {
                        CountTemp = Int32.Parse(categoryNodeList[categoryNodeList.IndexOf(item) + 1].InnerText.Trim());
                        break;
                    }
                }
                SetSumLabelValue("需要导入的总车数为： "+CountTemp);
                SetControl(btnStop, true);
                //从数据库中查询最大id作为统计标记
                string sqlId = "select max(id) as 'id' from lz_data";
                DBClass db = new DBClass();
                DataTable dt = db.DsGetinfo(sqlId).Tables[0];
                idTemp = dt.Rows[0]["id"].ToString();
                //从数据库中获取需要查对的料口数据
                List<string> liaokouData = getSqlData("select liaokou from lz_liaokou", "liaokou");
                //条数标记
                int flag = 0;

                //遍历HtmlNodeCollection中的每一个标签内的内容，获取需要插入的数据
                foreach (HtmlNode item in categoryNodeList)
                {
                    if (item.InnerText.Trim() != "")
                    {
                        //判断车号：三位整数
                        string strTemp = item.InnerText.Trim();
                        int intTemp = 0;
                        if (strTemp.Length == 3 && Int32.TryParse(strTemp, out intTemp))
                        {
                            //判断料口：去前两位与数据库中比较，如果是三堆，则在车号前加‘*’
                            if (categoryNodeList.IndexOf(item) + 2>=categoryNodeList.Count)
                            {
                                break;
                            }
                            string liaokouTemp1 = categoryNodeList[categoryNodeList.IndexOf(item) + 2].InnerText.Trim();
                            int pageDown=categoryNodeList.IndexOf(item) - 17;
                            if (pageDown>=categoryNodeList.Count)
                            {
                                break;
                            }
                            if (liaokouTemp1.Length == 4)//infoCheck(liaokouData,liaokouTemp1))//如果存在，则是料口
                            {
                                //如果料口和车号的索引固定（未翻页），可确定其他数据的索引
                                liaokouTemp1 = liaokouTemp1.Substring(0, 2);
                                if (liaokouTemp1 == "三堆")
                                {
                                    strTemp = "*" + strTemp;
                                }
                                //判断称重:长度为6位的小数
                                string chengzhongTemp = categoryNodeList[categoryNodeList.IndexOf(item) + 5].InnerText.Trim();
                                decimal czTemp = 0;
                                if (chengzhongTemp.Length == 6 && decimal.TryParse(chengzhongTemp, out czTemp))
                                {
                                    //判断日期：长度13位或者14位
                                    string riqiTemp = categoryNodeList[categoryNodeList.IndexOf(item) + 10].InnerText.Trim();
                                    string riqiIns = getDate(riqiTemp);
                                    bool b = InsertToSql(strTemp, liaokouTemp1, chengzhongTemp, riqiIns);
                                    if (b == true)
                                    {
                                        flag++;
                                        //在richtextbox显示导入的数据
                                        SetRichTextBoxValue(flag + " - " + strTemp + " - " + liaokouTemp1 + " - " + chengzhongTemp + " - " + riqiIns + "\n");
                                        SetLabelValue(flag);
                                        SetProcessBarValue(flag,CountTemp);
                                    }
                                    else
                                    {
                                        this.Text = "存在未导入数据，请审核！";
                                    }
                                }
                            }
                            else if (categoryNodeList[categoryNodeList.IndexOf(item) - 17].InnerText.Trim().Length == 4)//翻页时
                            {
                                liaokouTemp1 = categoryNodeList[categoryNodeList.IndexOf(item) - 17].InnerText.Trim();
                                //如果料口和车号的索引固定（未翻页），可确定其他数据的索引
                                liaokouTemp1 = liaokouTemp1.Substring(0, 2);
                                if (liaokouTemp1 == "三堆")
                                {
                                    strTemp = "*" + strTemp;
                                }
                                //判断称重:长度为6位的小数
                                string chengzhongTemp = categoryNodeList[categoryNodeList.IndexOf(item) - 14].InnerText.Trim();
                                decimal czTemp = 0;
                                if (chengzhongTemp.Length == 6 && decimal.TryParse(chengzhongTemp, out czTemp))
                                {
                                    //判断日期：长度13位或者14位
                                    string riqiTemp = categoryNodeList[categoryNodeList.IndexOf(item) + 4].InnerText.Trim();
                                    string riqiIns = getDate(riqiTemp);
                                    bool b = InsertToSql(strTemp, liaokouTemp1, chengzhongTemp, riqiIns);
                                    if (b == true)
                                    {
                                        flag++;
                                        //int valueTemp = flag / CountTemp * 100;
                                        //在richtextbox显示导入的数据
                                        SetRichTextBoxValue(flag + " - " + strTemp + " - " + liaokouTemp1 + " - " + chengzhongTemp + " - " + riqiIns + "\n");
                                        SetLabelValue(flag);
                                        SetProcessBarValue(flag,CountTemp);
                                    }
                                    else
                                    {
                                        this.Text = "存在未导入数据，请审核后导入！";
                                    }
                                }
                            }
                        }
                    }
                }

                //遍历完毕，统计数据库中本次录入的车数和吨数（根据ID）
                string sqlCount = "select count(id) as 'cheshu',sum(dunshu) as 'dunshu' from lz_data where id>'" + idTemp + "'";
                SqlDataReader dr = db.SqlGetinfo(sqlCount);
                if (dr.Read())
                {
                    SetSumLabelValue("导入合计：" + dr["cheshu"].ToString() + " 车  " + dr["dunshu"].ToString() + " 吨");
                }
                //如果操作有误，可放弃保存本次导入的数据
                DialogResult dr2 = MessageBox.Show("导入完成 "+dr["cheshu"].ToString()+" 车  "+dr["dunshu"].ToString()+"吨，\n请确定后保存？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr2 == DialogResult.No)
                {
                    string sqlDel = "delete from lz_data where id>'" + idTemp + "'";
                    int ires = db.Executeinfo(sqlDel);
                    if (ires > 0)
                    {
                        SetSumLabelValue( "已删除" + ires.ToString() + "条，未做任何更改");
                        SetProcessBarValue(0, CountTemp);
                        SetLabelValue(0);
                        SetRichTextBoxValue("已删除" + ires.ToString() + "条，未做任何更改");
                    }
                }
                SetControl(btnImport, true);
                SetControl(btnExit, true);
                SetControl(btnStop, false);
            }
        }
        /// <summary>
        /// 根据sql语句获取数据库中数据，并返回泛型数据数组
        /// </summary>
        /// <param name="cmd">查询语句</param>
        /// <returns>泛型数组</returns>
        private List<string> getSqlData(string cmd,string column)
        {
            List<string> sqlData = new List<string>();
            DBClass db = new DBClass();
            SqlDataReader sDR = db.SqlGetinfo(cmd);
            while (sDR.Read())
            {
                sqlData.Add(sDR[column].ToString());
            }
            sDR.Close();
            return sqlData;
        }

        /// <summary>
        /// 根据sql查询出来的数据sqldata，判断该数据str是否存在
        /// </summary>
        /// <param name="SqlData">查询出来的元数据sqldata</param>
        /// <param name="str">要判断的数据</param>
        /// <returns></returns>
        private bool infoCheck(List<string> SqlData, string str)
        {
            for (int i = 0; i < SqlData.Count; i++)
            {
                if (str==SqlData[i])
                {
                    return true;
                }
                else
                {
                    continue;
                }
            }
            return false;
        }
        #region 数据插入数据库表
        /// <summary>
        /// 每条数据插入到数据库中
        /// </summary>
        /// <param name="carNum">车号</param>
        /// <param name="liaoKou">料口</param>
        /// <param name="jingDun">净吨</param>
        /// <param name="riQi">日期</param>
        private bool InsertToSql(string carNum, string liaoKou, string dunshu, string riQi)
        {
            //根据车号查询车主姓名、超吨
            string sql_sel = "select chezhu,beizhu from lz_chezhu where chehao='" + carNum + "'";
            DBClass db = new DBClass();
            DataTable dt_chezhu = db.DsGetinfo(sql_sel).Tables[0];
            string chezhu = "新车主";
            string xiandun = "999";
            if (dt_chezhu.Rows.Count > 0)
            {
                chezhu = dt_chezhu.Rows[0]["chezhu"].ToString();
                xiandun = dt_chezhu.Rows[0]["beizhu"].ToString();
            }

            decimal xiandun_inp = 0;
            decimal dunshu_inp = 0;
            decimal chaodun = 0;
            //净吨和超吨必须是小数
            if (!decimal.TryParse(dunshu, out dunshu_inp) || !decimal.TryParse(xiandun, out xiandun_inp))
            {
                toolStripStatusLabel1.Text = "数据类型有误！"+DateTime.Now.ToShortTimeString();
                return false;
            }
            decimal jingdun = dunshu_inp;
            if (dunshu_inp > xiandun_inp)
            {
                jingdun = xiandun_inp;
                chaodun = dunshu_inp - xiandun_inp;
            }
            //插入数据库
            string sql_ins = "insert into lz_data values('" + chezhu + "','" + carNum + "','" + riQi + "','" + liaoKou + "','" + dunshu_inp.ToString() + "','" + chaodun + "','" + jingdun + "','0')";
            int i = db.Executeinfo(sql_ins);
            if (i > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion
        /// <summary>
        /// 选择文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnScan_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "htm 文件(*.htm)|*.htm|所有文件（*.*）|*.*";
            if (DialogResult.OK == fd.ShowDialog())
            {
                cboFileName.Text = fd.FileName;
            }
        }
        /// <summary>
        /// 根据htmlnode的innertext获取string类型日期
        /// </summary>
        /// <param name="item">传入的项目文本</param>
        /// <returns></returns>
        private string getDate(string item)
        {
            string[] riqis = item.Split('-');
            string riqiGet="1900/01/01";
            DateTime dtTemp;
            //默认为20xx年开始，年份在数组中的默认index为2，年份由20和文件中的两位年份组成
            riqiGet = "20" + riqis[2].Substring(0, 2).Trim() + "-" + riqis[0].Trim() + "-" + riqis[1].Trim();
             if (!DateTime.TryParse(riqiGet,out dtTemp))
            {
                riqiGet = "20" + riqis[0].Trim() + "-" + riqis[1].Trim() + "-" + riqis[2].Substring(0, 2).Trim();
            }
            if(!DateTime.TryParse(riqiGet,out dtTemp))
            {
                MessageBox.Show("日期类型有误！");
                t.Interrupt();
                Application.Exit();
                return "";
            }
            return riqiGet;
        }
        /// <summary>
        /// 子线程委托更改控件label richtextbox和processbar
        /// </summary>
        /// <param name="value"></param>
        #region   委托更改控件的值
        delegate void SetEnable(Control ctl,bool b);
        private void SetControl(Control ctl,bool b)
        {
            if (ctl.InvokeRequired)
            {
                SetEnable c = new SetEnable(SetControl);
                this.Invoke(c, new object[] { ctl, b });

            }
            else
            {
                ctl.Enabled = b;
            }
        }
        //委托获取控件的text值
        private string controlTextValue="";
        delegate void  GetTextValueCallback(Control c);
        private void  GetTextValue(Control c)
        {
            if (c.InvokeRequired)
            {
                GetTextValueCallback g = new GetTextValueCallback(GetTextValue);
                this.Invoke(g, new object[] { c });
            }
            else
            {
                controlTextValue = c.Text.Trim();
            }
        }
        delegate void SetValueCallback(int value);
        private void SetLabelValue(int value)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.statusStrip1.InvokeRequired)
            {
                SetValueCallback d = new SetValueCallback(SetLabelValue);
                this.Invoke(d, new object[] { value });
            }
            else
            {
                this.toolStripStatusLabel2.Text = "已完成：" + value.ToString() + " / "+CountTemp.ToString()+"  ";
            }
        }
        delegate void SetProcessBarValueCallback(int value, int sum);
        private void SetProcessBarValue(int value,int sum)
        {
            if (this.progressBar1.InvokeRequired)
            {
                SetProcessBarValueCallback d = new SetProcessBarValueCallback(SetProcessBarValue);
                this.Invoke(d, new object[] { value,sum });
            }
            else
            {
                this.progressBar1.Value = value;
                this.progressBar1.Maximum = sum;
            }

        }
        delegate void SetSumValueCallback(string value);
        private void SetSumLabelValue(string value)
        {
            if (statusStrip1.InvokeRequired)
            {
                SetSumValueCallback d = new SetSumValueCallback(SetSumLabelValue);
                this.Invoke(d, new object[] { value });

            }
            else
            {
                this.toolStripStatusLabel1.Text =value;
            }
        }
        delegate void SetStringValueCallback(string value);
        private void SetRichTextBoxValue(string value)
        {
            if (this.richTextBox1.InvokeRequired)
            {
                SetStringValueCallback d = new SetStringValueCallback(SetRichTextBoxValue);
                this.Invoke(d, new object[] { value });                
            }
            else
            {
                this.richTextBox1.Text = value + richTextBox1.Text; ;
            }
        }
#endregion
        /// <summary>
        /// 退出时终止未结束的线程
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExit_Click(object sender, EventArgs e)
        {
            try
            {
                Environment.Exit(0);
            }
            catch(Exception ex)
            {
                MessageBox.Show("程序已退出"+ex.Message);
            }
        }

        /// <summary>
        /// 终止导入，并删除已经导入的数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnStop_Click(object sender, EventArgs e)
        {
            //挂起导入线程
            t.Suspend();
            DialogResult drStop = MessageBox.Show("确定终止导入吗", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (drStop==DialogResult.Yes)
            {
                DBClass db = new DBClass();
                string sqlDel = "delete from lz_data where id>'" + idTemp + "'";
                int ires = db.Executeinfo(sqlDel);
                if (ires >= 0)
                {
                    SetSumLabelValue("已删除" + ires.ToString() + "条，未做任何更改");
                    SetRichTextBoxValue("已删除" + ires.ToString() + "条，未做任何更改\r");
                    SetControl(btnStop, false);
                    SetControl(btnImport, true);
                    SetControl(btnExit, true);
                    SetProcessBarValue(0, CountTemp);
                    SetLabelValue(0);
                }
            }
            else
            {
                t.Resume();
            }
        }

        private void FrmHtmlAgilityPack_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("程序已退出" + ex.Message);
            }
        }

    }
}
