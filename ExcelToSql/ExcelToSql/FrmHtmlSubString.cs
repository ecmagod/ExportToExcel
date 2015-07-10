using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Net;
using System.Xml.XPath;
using Sgml;
using System.Configuration;

namespace ExcelToSql
{
    public partial class FrmHtmlSubString : Form
    {
        public FrmHtmlSubString()
        {
            InitializeComponent();
        }
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
                txtFileName.Text = fd.FileName;
            }
        }

         #region 利用开源的sgmlreader读取html页面内容
        /// <summary>
        /// 读取html页面内容
        /// </summary>
        /// <param name="uri">网址</param>
        /// <param name="xpath">xpath标签</param>
        /// <returns></returns>
        private string GetWellFormedHTML(string uri, string xpath)
        {
            StreamReader sReader = null;//读取字节流
            StringWriter sw = null;//写入字符串
            SgmlReader reader = null;//sgml读取方法
            XmlTextWriter writer = null;//生成xml数据流
            try
            {
                if (uri == String.Empty) 
                    uri = "http://www.ypshop.net/list--91-940-940--search-1.html"; 
                WebClient webclient = new WebClient();
                webclient.Encoding = Encoding.UTF8;
                //页面内容
                string strWebContent = webclient.DownloadString(uri);


                reader = new SgmlReader();
                reader.DocType = "HTML";
                reader.InputStream = new StringReader(strWebContent);


                sw = new StringWriter();
                writer = new XmlTextWriter(sw);
                writer.Formatting = Formatting.Indented;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Whitespace)
                    {
                        writer.WriteNode(reader, true);
                    }
                }
                //return sw.ToString();
                if (xpath == null)
                {
                    return sw.ToString();
                }
                else
                { //Filter out nodes from HTML
                    StringBuilder sb = new StringBuilder();
                    XPathDocument doc = new XPathDocument(new StringReader(sw.ToString()));
                    XPathNavigator nav = doc.CreateNavigator();
                    XPathNodeIterator nodes = nav.Select(xpath);
                    while (nodes.MoveNext())
                    {
                        sb.Append(nodes.Current.Value + " ");
                    }
                    return sb.ToString();
                }
            }
            catch (Exception exp)
            {
                writer.Close();
                reader.Close();
                sw.Close();
                sReader.Close();
                return exp.Message;
            }
        }
        #endregion


                  #region             下面这个是参考的

/// <summary>
   
        private string GetWellFormedHTML_Handle(string uri)
        {
            StreamReader sReader = null;
            StringWriter sw = null;
            SgmlReader reader = null;
            XmlTextWriter writer = null;
            try
            {
                if (uri == String.Empty) uri = "http://www.ypshop.net/list--91-940-940--search-1.html";
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(uri);
                HttpWebResponse res = (HttpWebResponse)req.GetResponse();
                sReader = new StreamReader(res.GetResponseStream());


                reader = new SgmlReader();
                reader.DocType = "HTML";
                reader.InputStream = new StringReader(sReader.ReadToEnd());


                sw = new StringWriter();
                writer = new XmlTextWriter(sw);
                writer.Formatting = Formatting.Indented;
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Whitespace)
                    {
                        writer.WriteNode(reader, true);
                    }
                }


                StringBuilder sb = new StringBuilder();
                XPathDocument doc = new XPathDocument(new StringReader(sw.ToString()));
                XPathNavigator nav = doc.CreateNavigator();
                //XPathNodeIterator nodes = nav.Select(xpath);
                //while (nodes.MoveNext())
                //{
                //    sb.Append(nodes.Current.Value + " ");
                //}
                return sb.ToString();


            }
            catch (Exception exp)
            {
                writer.Close();
                reader.Close();
                sw.Close();
                sReader.Close();
                return exp.Message;
            }
        }
#endregion

                  #region   html字符串分析导入
        /// <summary>
        /// 执行导入
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImport_Click(object sender, EventArgs e)
        {
            //参数完整性验证
            if (txtFileName.Text.Trim() == "")
            {
                return;
            }


            //提示确认导入该文件
            DialogResult dr=MessageBox.Show("确定将 "+txtFileName.Text.Trim()+" 导入到系统中？","提示",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
            if (dr == DialogResult.No)
                return;

            /*
             * SUBSTRING识别通用的中间数据，识别字符串“<TR><TD HEIGHT=""19""></TD><TD></TD><TD></TD><TD COLSPAN=""3"" NOWRAP="""" VALIGN=""TOP""><FONT FACE=""宋体""”
             */

            //利用sgmlreader将htm文件读取到一个字符串中，注意：string长度受限于内存大小
            string html = GetWellFormedHTML(txtFileName.Text.Trim(),null);

            //每一个html中标识一行有效行的字符串，与xpath同义 注意：双引号转义
            string sub = @"<TR><TD HEIGHT=""19""></TD><TD></TD><TD></TD><TD COLSPAN=""3"" NOWRAP="""" VALIGN=""TOP""><FONT FACE=""宋体""";
            //获取截取字符串的起止范围
            int carNum_FirstIndex = html.IndexOf(sub);
            int carNum_LastIndex = html.LastIndexOf(sub);

            //定义每次插入的信息，执行第一次索引查询、插入
            string carNum = "";
            string liaokou = "";
            string jingdun = "";
            string riqi = "";

            //进度标记
            int flag = 1;
            //起始索引小于最终索引时，获取该范围内的车号、净吨、日期、料口，插入数据库，然后重新把下一个sub索引赋值给carNum_FirstIndex
            while (carNum_FirstIndex<carNum_LastIndex)
            {
                if (carNum_FirstIndex==carNum_LastIndex)//起止索引相等时为最后一条，插入完break
                {
                    //将该有效行内的所需信息导入sql
                    carNum = html.Substring(carNum_FirstIndex + sub.Length + 18, 3);
                    liaokou = html.Substring(carNum_FirstIndex + sub.Length + 223, 2);
                    jingdun = html.Substring(carNum_FirstIndex + sub.Length + 501, 6);
                    //日期时间有时候是个位数，有两种情况
                    riqi = DateTime.Now.Year.ToString() + "-" + html.Substring(carNum_FirstIndex + sub.Length + 1149, 5);
                    DateTime riqi_test1;
                    if (!DateTime.TryParse(riqi, out riqi_test1))
                    {
                        riqi = DateTime.Now.Year.ToString() + "-" + html.Substring(carNum_FirstIndex + sub.Length + 1148, 5);
                    }
                    //三堆料口车号前面加星号
                    if (liaokou == "三堆")
                    {
                        carNum = "*" + carNum;
                    }
                    //InsertToSql(carNum, liaokou, jingdun, riqi);
                    //进度提示，进行到多少条
                    label4.Text = "已经导入 " + flag.ToString() + " 条";
                    carNum_FirstIndex = html.IndexOf(sub, carNum_FirstIndex + 1);
                    flag++;
                    break;
                }
                //将该有效行内的所需信息导入sql
                carNum = html.Substring(carNum_FirstIndex + sub.Length + 18, 3);
                liaokou = html.Substring(carNum_FirstIndex + sub.Length + 223, 2);
                jingdun = html.Substring(carNum_FirstIndex + sub.Length + 501,6);
                //日期时间有时候是个位数，有两种情况
                riqi = DateTime.Now.Year.ToString()+"-"+html.Substring(carNum_FirstIndex + sub.Length + 1149, 5);
                DateTime riqi_test;
                if (!DateTime.TryParse(riqi,out riqi_test))
                {
                    riqi = DateTime.Now.Year.ToString() + "-" + html.Substring(carNum_FirstIndex + sub.Length + 1148, 5);
                }
                //三堆料口车号前面加星号
                if (liaokou=="三堆")
                {
                    carNum = "*" + carNum;
                }
                //InsertToSql(carNum, liaokou, jingdun, riqi);
                //进度提示，进行到多少条
                label4.Text = "已经导入 "+flag.ToString()+" 条";
                carNum_FirstIndex = html.IndexOf(sub, carNum_FirstIndex + 1);
                flag++;
            }

            /*
             *  识别翻页时特殊的数据格式，根据title“中铝股份山东分公司汽车衡物资检斤明细”从第二个到最后一个,每个title识别一条
             * */

            //获取title起止索引
            int title_FirstIndex = html.IndexOf("中铝股份山东分公司汽车衡物资检斤明细");
            int title_LastIndex = html.LastIndexOf("中铝股份山东分公司汽车衡物资检斤明细");
            //翻页标记
            int s_flag = 0;
            //逐条添加到数据库
            while (title_FirstIndex<=title_LastIndex)
            {
                //跳过第一次
                if (s_flag==0)
                {
                    title_FirstIndex = html.IndexOf("中铝股份山东分公司汽车衡物资检斤明细", title_FirstIndex + 50);
                    if (title_FirstIndex == -1)//最后一条break
                    {
                        break;
                    }
                    //标识递增
                    s_flag++;
                    continue;
                }
                carNum = html.Substring(title_FirstIndex + 1349, 3);
                liaokou = html.Substring(title_FirstIndex - 2198, 2);
                jingdun = html.Substring(title_FirstIndex - 2000, 500);
                //日期时间有时候是个位数，有两种情况
                riqi = DateTime.Now.Year.ToString() + "-" + html.Substring(carNum_FirstIndex + sub.Length + 1149, 5);
                DateTime riqi_test;
                if (!DateTime.TryParse(riqi, out riqi_test))
                {
                    riqi = DateTime.Now.Year.ToString() + "-" + html.Substring(carNum_FirstIndex + sub.Length + 1148, 5);
                }
                //三堆料口车号前面加星号
                if (liaokou == "三堆")
                {
                    carNum = "*" + carNum;
                }
                //InsertToSql(carNum, liaokou, jingdun, riqi);
                //进度提示，进行到多少条
                label4.Text = "已经导入 " + flag.ToString() + " 条";
                title_FirstIndex = html.IndexOf("中铝股份山东分公司汽车衡物资检斤明细", title_FirstIndex + 50);
                if (title_FirstIndex==-1)//最后一条break
                {
                    break;
                }
                //两个标识递增
                flag++;
                s_flag++;
            }

            /* 
             * 第一种识别标记字符串无法识别部分，用第二种识别字符串“<TR><TD HEIGHT="19"></TD><TD></TD><TD></TD><TD></TD><TD COLSPAN="5" NOWRAP="" VALIGN="TOP"><FONT FACE="宋体" ”   
             */


        }
        #endregion

                  #region 数据插入数据库表
        /// <summary>
        /// 每条数据插入到数据库中
        /// </summary>
        /// <param name="carNum">车号</param>
        /// <param name="liaoKou">料口</param>
        /// <param name="jingDun">净吨</param>
        /// <param name="riQi">手动输入的日期</param>
        private bool InsertToSql(string carNum,string liaoKou,string dunshu,string riQi)
        {
            //根据车号查询车主姓名、超吨
            string sql_sel = "select chezhu,beizhu from lz_chezhu where chehao='"+carNum+"'";
            DBClass db = new DBClass();
            DataTable dt_chezhu = db.DsGetinfo(sql_sel).Tables[0];
            string chezhu = "新车主";
            string xiandun = "999";
            if (dt_chezhu.Rows.Count>0)
            {
                chezhu = dt_chezhu.Rows[0]["chezhu"].ToString();
                xiandun = dt_chezhu.Rows[0]["beizhu"].ToString();
            }

            decimal xiandun_inp = 0;
            decimal dunshu_inp = 0;
            
            decimal chaodun=0;
            //净吨和超吨必须是小数
            if (!decimal.TryParse(dunshu,out dunshu_inp)||!decimal.TryParse(xiandun,out xiandun_inp))
            {
                MessageBox.Show("所读取的数据类型有误！","错误");
                return false;
            }
            decimal jingdun = dunshu_inp;
            if (dunshu_inp > xiandun_inp)
            {
                jingdun = xiandun_inp;
                chaodun = dunshu_inp - xiandun_inp;
            }
            //插入数据库
            string sql_ins = "insert into lz_data values('"+chezhu+"','"+carNum+"','"+riQi+"','"+liaoKou+"','"+dunshu_inp.ToString()+"','"+chaodun+"','"+jingdun+"','0')";
            int i = db.Executeinfo(sql_ins);
            if (i>0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        /// <summary>
        /// 日期
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmHtmlRead_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Now;
        }
    }
}
