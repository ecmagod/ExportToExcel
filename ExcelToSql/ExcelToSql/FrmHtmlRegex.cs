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
using System.Configuration;
using System.Text.RegularExpressions;
//引入sgmlreader的命名空间
using Sgml;
//引入HtmlParser的命名空间
using Winista.Text.HtmlParser;
using Winista.Text.HtmlParser.Lex;
using Winista.Text.HtmlParser.Util;
using Winista.Text.HtmlParser.Tags;
using Winista.Text.HtmlParser.Filters;

namespace ExcelToSql
{
    public partial class FrmHtmlRegex : Form
    {
        public FrmHtmlRegex()
        {
            InitializeComponent();
        }
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
            DialogResult dr = MessageBox.Show("确定将 " + txtFileName.Text.Trim() + " 导入到系统中？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dr == DialogResult.No)
                return;

            //利用sgmlreader将htm文件读取到一个字符串中，注意：string长度受限于内存大小
            string html = GetWellFormedHTML(txtFileName.Text.Trim(), null);

            //正则表达式验证标签<TR>***</TR>和<TD>***</TD>
            //也不太会用，有时候类似<TR WIDTH="">无法识别，正则表达式不固定，关键也不太会,对于复杂不标准的html效率低下
            Regex regTR = new Regex(@"(?is)<tr[^>]*>(?:(?!</tr>).)*</tr>");
            Regex regTD = new Regex(@"(?is)<t[dh][^>]*>((?:(?!</td>).)*)</t[dh]>");
            MatchCollection mcTR = regTR.Matches(html);
            foreach (Match mTR in mcTR)
            {
                if (mTR.ToString().Trim()!="")
                {
                    MatchCollection mcTD = regTD.Matches(mTR.Value);

                    foreach (Match mTD in mcTD)
                    {
                        if (mTD.Groups[1].Value.Trim() != "")
                        {
                            richTextBox1.Text += mTD.Groups[1].Value + "\n";
                        }
                    }
                }
            }

            
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


    }
}
