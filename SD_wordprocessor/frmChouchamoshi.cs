using clsBuiness;
using DCTS.CustomComponents;
using NPOI.OpenXmlFormats.Dml;
using NPOI.OpenXmlFormats.Dml.WordProcessing;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.XWPF.UserModel;
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using WeifenLuo.WinFormsUI.Docking;

namespace SD_wordprocessor
{
    public partial class frmChouchamoshi : DockContent
    {

        string txfind;
        List<clsword_info> Orderinfolist_Server;
        private PictureBox pictureBox1;
        private SortableBindingList<clsKeyWord_web_info> sortableLogList;

        int Fwidth;
        int Fheight;
        public string FPath;
        FontStyle Fstyle = FontStyle.Regular;
        float Fsize = 16;
        Color Fcolor = System.Drawing.Color.Black;
        FontFamily a = FontFamily.GenericMonospace;
        List<clsKeyWord_web_info> Word_web_Server;
        DataTable qtyTable;
        public ReportForm reportForm;
        bool is_AdminIS;
        string lastinput;
        public frmChouchamoshi(bool is_AdminIS1)
        {
            InitializeComponent();
            Word_web_Server = new List<clsKeyWord_web_info>();
            is_AdminIS = is_AdminIS1;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            // MessageBox.Show("没找到,请重新确认");



            Word_web_Server = new List<clsKeyWord_web_info>();

            clsAllnew BusinessHelp = new clsAllnew();


            // addtxt();

            string line = "ADDR=1234;NAME=ZHANG;PHONE=6789";

            Regex reg = new Regex("NAME=(.+);");

            string modified = reg.Replace(line, "NAME=WANG;");


            //this.textBox2.Text = "";
            txfind = this.textBox1.Text;
            if (lastinput != null && lastinput.Replace(" ", "") == txfind.Replace(" ", ""))
                return;

            lastinput = txfind;

            string alltxfind = this.textBox1.Text;
            #region 判断单独4 个代码是不是一个 字的功能
            int stepmethod = 0;
            if (stepmethod == 0 && alltxfind != "")
            {
                string[] fileTextall = System.Text.RegularExpressions.Regex.Split(alltxfind, " ");
                for (int iqqq = 0; iqqq < fileTextall.Length; iqqq++)
                {
                    string strSelect1 = "select * from Word_ku where zhengtidaima ='" + fileTextall[iqqq].ToString() + "'";
                    Orderinfolist_Server = new List<clsword_info>();
                    Orderinfolist_Server = BusinessHelp.findWord(strSelect1);
                    clsKeyWord_web_info item = new clsKeyWord_web_info();
                    if (Orderinfolist_Server.Count > 0)
                    {
                        item.word = Orderinfolist_Server[0].zi;
                        List<clsKeyWord_web_info> bew = new List<clsKeyWord_web_info>();
                        bew.Add(item);

                        Word_web_Server = Word_web_Server.Concat(bew).ToList();
                    }

                    //hanzi cha

                    strSelect1 = "select * from Word_ku where zi ='" + fileTextall[iqqq].ToString() + "'";
                    Orderinfolist_Server = new List<clsword_info>();
                    Orderinfolist_Server = BusinessHelp.findWord(strSelect1);
                    item = new clsKeyWord_web_info();
                    if (Orderinfolist_Server.Count > 0)
                    {
                        item.word = Orderinfolist_Server[0].zhengtidaima;
                        List<clsKeyWord_web_info> bew = new List<clsKeyWord_web_info>();
                        bew.Add(item);

                        Word_web_Server = Word_web_Server.Concat(bew).ToList();
                    }
                    else
                    {
                        if (txfind.Length == 4)
                        {
                            MessageBox.Show("未查到本条信息,或录入信息错误！" + fileTextall[iqqq]);

                            List<clsKeyWord_web_info> bew = new List<clsKeyWord_web_info>();

                            clsKeyWord_web_info addempty = new clsKeyWord_web_info();
                            addempty.word = "未找到";

                            bew.Add(addempty);
                            Word_web_Server = Word_web_Server.Concat(bew).ToList();
                        }

                    }
                }
            }



            #endregion
            int ddd = txfind.Replace(" ", "").Length / 4;
            int index = ddd / 5;
            string s = "7521 9991 8531 9991 7452 ";
            int ds = s.Length;
            int indexall = 0;
            for (int iqqq = 0; iqqq < index; iqqq++)
            {
                string[] fileTextall = System.Text.RegularExpressions.Regex.Split(alltxfind, " ");
                int weiindex = (iqqq + 1) * 5;
                string aldd = "";
                int startindex = indexall;
                for (int inli = startindex; inli < weiindex; inli++)
                {
                    aldd += fileTextall[inli] + " ";
                    indexall++;
                }
                txfind = aldd.Trim();


                string[] fileText = System.Text.RegularExpressions.Regex.Split(txfind, " ");

                string jiegou = "";
                string zucheng_jiegou = "";

                string jiegou2 = "";
                string zi1 = "";
                string zi2 = "";
                for (int i = 0; i < fileText.Length; i++)
                {
                    txfind = fileText[i];
                    if (i == 0)
                    {
                        if (txfind == "7521")
                            zucheng_jiegou = "左右";
                        if (txfind == "9991")
                            zucheng_jiegou = "左右";
                        if (txfind == "9992")
                            zucheng_jiegou = "上下";
                        if (txfind == "9993")
                            zucheng_jiegou = "内外";
                        if (txfind == "9994")
                            zucheng_jiegou = "左中右";
                        if (txfind == "9995")
                            zucheng_jiegou = "上中下";
                        if (txfind == "9996")
                            zucheng_jiegou = "复杂";
                    }
                    else if (i == 1)
                    {
                        jiegou = jiegouCheck(jiegou);

                    }

                    else if (i == 2)
                    {
                        string strSelect1 = "select * from Word_ku where zhengtidaima like'%" + txfind.ToString() + "%'";
                        Orderinfolist_Server = new List<clsword_info>();
                        Orderinfolist_Server = BusinessHelp.findWord(strSelect1);
                        zi1 = Orderinfolist_Server[0].zi;
                        if (jiegou == "上")
                        {
                            zi1 = Orderinfolist_Server[0].bushou1;
                        }
                        else if (jiegou == "下")
                        {
                            zi1 = Orderinfolist_Server[0].bushou4;

                        }
                    }
                    else if (i == 3)
                    {
                        if (txfind == "9991" || txfind == "0001")
                            jiegou2 = "上";
                        jiegou2 = jiegouCheck(jiegou2);
                    }
                    else if (i == 4)
                    {
                        string strSelect1 = "select * from Word_ku where zhengtidaima like'%" + txfind.ToString() + "%'";
                        Orderinfolist_Server = new List<clsword_info>();
                        Orderinfolist_Server = BusinessHelp.findWord(strSelect1);
                        zi2 = Orderinfolist_Server[0].zi;
                        if (jiegou2 == "上")
                        {
                            zi2 = Orderinfolist_Server[0].bushou1;
                        }
                        else if (jiegou2 == "下")
                        {
                            zi2 = Orderinfolist_Server[0].bushou4;

                        }

                    }
                    //7521 9991 8531 9991 7452
                }
                string WEBstrSelect = "select * from Word_web where pianpang like'%" + zi1 + zi2 + "%'";

                List<clsKeyWord_web_info> one_Word_web_Server = BusinessHelp.findWord_web(WEBstrSelect);
                List<clsKeyWord_web_info> Aging_CaseListResult = one_Word_web_Server.FindAll(so => so.mark2 != null && so.mark2 == zucheng_jiegou);
                if (Aging_CaseListResult.Count == 0)
                {
                    MessageBox.Show("未查到本条信息,或录入信息错误！" + txfind);
                    clsKeyWord_web_info addempty = new clsKeyWord_web_info();
                    addempty.word = "未找到";

                    Aging_CaseListResult.Add(addempty);

                }
                Word_web_Server = Word_web_Server.Concat(Aging_CaseListResult).ToList();

            }
            BindDataGridView();

        }

        private string jiegouCheck(string jiegou)
        {
            if (txfind == "9991" || txfind == "0001")
                jiegou = "上";
            if (txfind == "0001")
                jiegou = "上";
            if (txfind == "0002")
                jiegou = "下";
            if (txfind == "0003")
                jiegou = "左";
            if (txfind == "0004")
                jiegou = "右";
            if (txfind == "0005")
                jiegou = "内";
            if (txfind == "0006")
                jiegou = "外";
            return jiegou;
        }
        private void BindDataGridView()
        {

            if (Word_web_Server != null)
            {
                sortableLogList = new SortableBindingList<clsKeyWord_web_info>(Word_web_Server);
                bindingSource1.DataSource = new SortableBindingList<clsKeyWord_web_info>(Word_web_Server);
                dataGridView1.AutoGenerateColumns = true;
                qtyTable = new DataTable();
                for (int i = 0; i < 10; i++)
                    qtyTable.Columns.Add("name" + (i + 1).ToString(), System.Type.GetType("System.String"));
                qtyTable.Rows.Add(qtyTable.NewRow());
                for (int j = 0; j < 20; j++)
                {
                    qtyTable.Rows.Add(qtyTable.NewRow());
                }
                int jk = 0;
                {
                    for (int j = 0; j < 100; j++)
                    {
                        bool isjioushu = IsOdd(j);
                        if (isjioushu == false || j == 0)
                        {

                            //在库1也就是基础库里随机生成100组（50组文字（含标点符号和字母）转代码、50组代码转文字）代码以及文字供参与者填写给出代码的另一种代表格式
                            for (int i = 0; i < 10; i++)
                            {
                                if (Word_web_Server.Count > jk)
                                {
                                    isjioushu = IsOdd(i);

                                    if (isjioushu == false || i == 0)
                                    {
                                        qtyTable.Rows[j][i] = Word_web_Server[jk].word;
                                        jk++;
                                    }
                                    else
                                    {
                                        qtyTable.Rows[j][i] = Word_web_Server[jk].word;
                                        jk++;
                                    }
                                }
                            }


                        }
                    }
                }
                this.bindingSource1.DataSource = qtyTable;
                dataGridView1.DataSource = bindingSource1;
                this.toolStripLabel1.Text = "条数：" + Word_web_Server.Count.ToString();
            }
            else
            {
                dataGridView1.DataSource = null;


            }

        }
        public static bool IsOdd(int n)
        {
            while (true)
            {
                switch (n)
                {
                    case 1: return true;
                    case 0: return false;
                }
                n -= 2;
            }
        }
        private void 发信模板ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //if (textBox2.Text == "")
            //    return;

            clsAllnew BusinessHelp = new clsAllnew();
            //   BusinessHelp.downcsv(dataGridView1);
            string strFileName = "  信息" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + strFileName + ".docx";






            createWord(dir, "textBox2.Text");

            //   Button1_Click(this, EventArgs.Empty);

        }

        #region npoi
        public void createWord(string strFileName, string body)
        {

            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();
            XWPFParagraph p1 = doc.CreateParagraph();
            XWPFRun rUserHead = p1.CreateRun();
            rUserHead.SetText(body);
            rUserHead.SetColor("4F6B72");

            rUserHead.FontSize = 15;
            rUserHead.IsBold = true;
            rUserHead.FontFamily = "宋体";
            //r1.SetUnderline(UnderlinePatterns.DotDotDash);
            //位置
            rUserHead.SetTextPosition(20);

            //增加换行

            rUserHead.AddCarriageReturn();

            FileStream sw = File.OpenWrite(strFileName);
            doc.Write(sw);
            sw.Close();
            MessageBox.Show("下载完成！" + strFileName, "下载word");

        }
        protected void Button1_Click(object sender, EventArgs e)
        {
            //图片位置
            String m_PicPath = @"C:\ttt.jpeg";
            FileStream gfs = null;
            MemoryStream ms = new MemoryStream();
            XWPFDocument m_Docx = new XWPFDocument();
            //页面设置
            //A4:W=11906,h=16838
            //CT_SectPr m_SectPr = m_Docx.Document.body.AddNewSectPr();
            m_Docx.Document.body.sectPr = new CT_SectPr();
            CT_SectPr m_SectPr = m_Docx.Document.body.sectPr;
            //页面设置A4纵向
            m_SectPr.pgSz.h = (ulong)16838;
            m_SectPr.pgSz.w = (ulong)11906;
            XWPFParagraph gp = m_Docx.CreateParagraph();
            // gp.GetCTPPr().AddNewJc().val = ST_Jc.center; //水平居中
            XWPFRun gr = gp.CreateRun();
            gr.GetCTR().AddNewRPr().AddNewRFonts().ascii = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().eastAsia = "黑体";
            gr.GetCTR().AddNewRPr().AddNewRFonts().hint = ST_Hint.eastAsia;
            gr.GetCTR().AddNewRPr().AddNewSz().val = (ulong)44;//2号字体
            gr.GetCTR().AddNewRPr().AddNewSzCs().val = (ulong)44;
            gr.GetCTR().AddNewRPr().AddNewB().val = true; //加粗
            gr.GetCTR().AddNewRPr().AddNewColor().val = "red";//字体颜色
            gr.SetText("NPOI创建Word2007Docx");
            gp = m_Docx.CreateParagraph();
            //gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            // gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gp.IndentationFirstLine = 15;
            gr = gp.CreateRun();
            CT_RPr rpr = gr.GetCTR().AddNewRPr();
            CT_Fonts rfonts = rpr.AddNewRFonts();
            rfonts.ascii = "宋体";
            rfonts.eastAsia = "宋体";
            rpr.AddNewSz().val = (ulong)21;//5号字体
            rpr.AddNewSzCs().val = (ulong)21;
            gr.SetText("NPOI，顾名思义，就是POI的.NET版本。那POI又是什么呢？POI是一套用Java写成的库，能够帮助开 发者在没有安装微软Office的情况下读写Office 97-2003的文件，支持的文件格式包括xls, doc, ppt等 。目前POI的稳定版本中支持Excel文件格式(xls和xlsx)，其他的都属于不稳定版本（放在poi的scrachpad目录 中）。");
            //创建表
            XWPFTable table = m_Docx.CreateTable(1, 4);//创建一行4列表
            CT_Tbl m_CTTbl = m_Docx.Document.body.GetTblArray()[0];//获得文档第一张表
            CT_TblPr m_CTTblPr = m_CTTbl.AddNewTblPr();
            m_CTTblPr.AddNewTblW().w = "2000"; //表宽
            m_CTTblPr.AddNewTblW().type = ST_TblWidth.dxa;
            m_CTTblPr.tblpPr = new CT_TblPPr();//表定位
            m_CTTblPr.tblpPr.tblpX = "4003";//表左上角坐标
            m_CTTblPr.tblpPr.tblpY = "365";
            m_CTTblPr.tblpPr.tblpXSpec = ST_XAlign.center;//若不为“Null”，则优先tblpX，即表由tblpXSpec定位
            m_CTTblPr.tblpPr.tblpYSpec = ST_YAlign.center;//若不为“Null”，则优先tblpY，即表由tblpYSpec定位  
            m_CTTblPr.tblpPr.leftFromText = (ulong)180;
            m_CTTblPr.tblpPr.rightFromText = (ulong)180;
            m_CTTblPr.tblpPr.vertAnchor = ST_VAnchor.text;
            m_CTTblPr.tblpPr.horzAnchor = ST_HAnchor.page;


            //表1行4列充值：a,b,c,d
            table.GetRow(0).GetCell(0).SetText("a");
            table.GetRow(0).GetCell(1).SetText("b");
            table.GetRow(0).GetCell(2).SetText("c");
            table.GetRow(0).GetCell(3).SetText("d");
            CT_Row m_NewRow = new CT_Row();//创建1行
            XWPFTableRow m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row); //必须要！！！
            XWPFTableCell cell = m_Row.CreateCell();//创建单元格，也创建了一个CT_P
            CT_Tc cttc = cell.GetCTTc();
            CT_TcPr ctPr = cttc.AddNewTcPr();
            //ctPr.gridSpan.val = "3";//合并3列
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "666";
            cell = m_Row.CreateCell();//创建单元格，也创建了一个CT_P
            cell.GetCTTc().GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cell.GetCTTc().GetPList()[0].AddNewR().AddNewT().Value = "e";
            //合并3列，合并2行
            //1行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            cell.SetText("f");
            cell = m_Row.CreateCell();//从第2单元格开始合并
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            // ctPr.gridSpan.val = "3";//合并3列        
            ctPr.AddNewVMerge().val = ST_Merge.restart;//开始合并行
            ctPr.AddNewVAlign().val = ST_VerticalJc.center;//垂直居中
            cttc.GetPList()[0].AddNewPPr().AddNewJc().val = ST_Jc.center;
            cttc.GetPList()[0].AddNewR().AddNewT().Value = "777";
            //2行
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            cell.SetText("g");
            cell = m_Row.CreateCell();//第2单元格
            cttc = cell.GetCTTc();
            ctPr = cttc.AddNewTcPr();
            // ctPr.gridSpan.val = "3";//合并3列
            ctPr.AddNewVMerge().val = ST_Merge.@continue;//继续合并行


            //表插入图片
            m_NewRow = new CT_Row();
            m_Row = new XWPFTableRow(m_NewRow, table);
            table.AddRow(m_Row);
            cell = m_Row.CreateCell();//第1单元格
            //inline方式插入图片
            //gp = table.GetRow(table.Rows.Count - 1).GetCell(0).GetParagraph(table.GetRow(table.Rows.Count - 1).GetCell(0).GetCTTc().GetPList()[0]);//获得指定表单元格的段
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            gfs = new FileStream(m_PicPath, FileMode.Open, FileAccess.Read);//读取图片文件
            gr.AddPicture(gfs, (int)PictureType.PNG, "1.jpg", 500000, 500000);//插入图片
            gfs.Close();
            //Anchor方式插入图片
            CT_Anchor an = new CT_Anchor();
            an.distB = (uint)(0);
            an.distL = 114300u;
            an.distR = 114300U;
            an.distT = 0U;
            an.relativeHeight = 251658240u;
            an.behindDoc = false; //"0"
            an.locked = false;  //"0"
            an.layoutInCell = true;  //"1"
            an.allowOverlap = true;  //"1" 


            NPOI.OpenXmlFormats.Dml.CT_Point2D simplePos = new NPOI.OpenXmlFormats.Dml.CT_Point2D();
            simplePos.x = (long)0;
            simplePos.y = (long)0;
            CT_EffectExtent effectExtent = new CT_EffectExtent();
            effectExtent.b = 0L;
            effectExtent.l = 0L;
            effectExtent.r = 0L;
            effectExtent.t = 0L;
            //wrapSquare(四周)
            cell = m_Row.CreateCell();//第2单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            CT_WrapSquare wrapSquare = new CT_WrapSquare();
            wrapSquare.wrapText = ST_WrapText.bothSides;
            gfs = new FileStream(m_PicPath, FileMode.Open, FileAccess.Read);//读取图片文件
            // gr.AddPicture(gfs, (int)PictureType.PNG, "1.png", 500000, 500000, 0, 0, wrapSquare, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gr.AddPicture(gfs, (int)PictureType.PNG, "虚拟机", 500000, 500000);
            gfs.Close();
            //wrapTight（紧密）
            cell = m_Row.CreateCell();//第3单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            CT_WrapTight wrapTight = new CT_WrapTight();
            wrapTight.wrapText = ST_WrapText.bothSides;
            wrapTight.wrapPolygon = new CT_WrapPath();
            wrapTight.wrapPolygon.edited = false;
            wrapTight.wrapPolygon.start = new CT_Point2D();
            wrapTight.wrapPolygon.start.x = 0;
            wrapTight.wrapPolygon.start.y = 0;
            CT_Point2D lineTo = new CT_Point2D();
            wrapTight.wrapPolygon.lineTo = new List<CT_Point2D>();
            lineTo = new CT_Point2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapTight.wrapPolygon.lineTo.Add(lineTo);
            gfs = new FileStream(m_PicPath, FileMode.Open, FileAccess.Read);//读取图片文件
            //gr.AddPicture(gfs, (int)PictureType.PNG, "1.png", 500000, 500000, 0, 0, wrapTight, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gr.AddPicture(gfs, (int)PictureType.PNG, "虚拟机", 500000, 500000);
            gfs.Close();
            //wrapThrough(穿越)
            cell = m_Row.CreateCell();//第4单元格
            gp = cell.GetParagraph(cell.GetCTTc().GetPList()[0]);
            gr = gp.CreateRun();//创建run
            gfs = new FileStream(m_PicPath, FileMode.Open, FileAccess.Read);//读取图片文件
            CT_WrapThrough wrapThrough = new CT_WrapThrough();
            wrapThrough.wrapText = ST_WrapText.bothSides;
            wrapThrough.wrapPolygon = new CT_WrapPath();
            wrapThrough.wrapPolygon.edited = false;
            wrapThrough.wrapPolygon.start = new CT_Point2D();
            wrapThrough.wrapPolygon.start.x = 0;
            wrapThrough.wrapPolygon.start.y = 0;
            lineTo = new CT_Point2D();
            wrapThrough.wrapPolygon.lineTo = new List<CT_Point2D>();
            lineTo = new CT_Point2D();
            lineTo.x = 0;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 1343;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo = new CT_Point2D();
            lineTo.x = 21405;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            lineTo.x = 0;
            lineTo.y = 0;
            wrapThrough.wrapPolygon.lineTo.Add(lineTo);
            // gr.AddPicture(gfs, (int)PictureType.PNG, "15.png", 500000, 500000, 0, 0, wrapThrough, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gr.AddPicture(gfs, (int)PictureType.PNG, "虚拟机", 500000, 500000);
            gfs.Close();


            gp = m_Docx.CreateParagraph();
            //gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            // gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gp.IndentationFirstLine = 15;
            gr = gp.CreateRun();
            gr.SetText("NPOI是POI项目的.NET版本。POI是一个开源的Java读写Excel、WORD等微软OLE2组件文档的项目。使用NPOI你就可以在没有安装Office或者相应环境的机器上对WORD/EXCEL文档进行读写。NPOI是构建在POI3.x版本之上的，它可以在没有安装Office的情况下对Word/Excel文档进行读写操作。");
            gp = m_Docx.CreateParagraph();
            // gp.GetCTPPr().AddNewJc().val = ST_Jc.both;
            //gp.IndentationFirstLine = Indentation("宋体", 21, 2, FontStyle.Regular);//段首行缩进2字符
            gp.IndentationFirstLine = 15;
            gr = gp.CreateRun();
            gr.SetText("NPOI之所以强大，并不是因为它支持导出Excel，而是因为它支持导入Excel，并能“理解”OLE2文档结构，这也是其他一些Excel读写库比较弱的方面。通常，读入并理解结构远比导出来得复杂，因为导入你必须假设一切情况都是可能的，而生成你只要保证满足你自己需求就可以了，如果把导入需求和生成需求比做两个集合，那么生成需求通常都是导入需求的子集。");
            //在本段中插图-wrapSquare
            //gr = gp.CreateRun();//创建run
            wrapSquare = new CT_WrapSquare();
            wrapSquare.wrapText = ST_WrapText.bothSides;
            gfs = new FileStream(m_PicPath, FileMode.Open, FileAccess.Read);//读取图片文件
            // gr.AddPicture(gfs, (int)PictureType.PNG, "15.png", 500000, 500000, 900000, 200000, wrapSquare, an, simplePos, ST_RelFromH.column, ST_RelFromV.paragraph, effectExtent);
            gr.AddPicture(gfs, (int)PictureType.PNG, "虚拟机", 500000, 500000);
            gfs.Close();


            m_Docx.Write(ms);
            ms.Flush();
            SaveToFile(ms, Path.GetPathRoot(Directory.GetCurrentDirectory()) + "\\NPOI.docx");


        }
        static void SaveToFile(MemoryStream ms, string fileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                byte[] data = ms.ToArray();


                fs.Write(data, 0, data.Length);
                fs.Flush();
                data = null;
            }
        }
        //protected int Indentation(String fontname, int fontsize, int Indentationfonts, FontStyle fs)
        //{
        //    //字显示宽度，用于段首行缩进
        //    /*字号与fontsize关系
        //     * 初号（0号）=84，小初=72，1号=52，2号=44，小2=36，3号=32，小3=30，4号=28，小4=24，5号=21，小5=18，6号=15，小6=13，7号=11，8号=10
        //     */
        //    Graphics m_tmpGr = this.CreateGraphics();
        //    m_tmpGr.PageUnit = GraphicsUnit.Point;
        //    SizeF size = m_tmpGr.MeasureString("好", new Font(fontname, fontsize * 0.75F, fs));
        //    return (int)size.Width * Indentationfonts * 10;
        //}

        #endregion

        private void addtxt()
        {

            #region 转移汉字
            string strr = "女又";
            byte[] buffer1 = Encoding.Default.GetBytes(strr);
            string strr1 = Encoding.GetEncoding("GBK").GetString(buffer1);
            byte[] buffer11 = Encoding.GetEncoding("GBK").GetBytes(strr);

            string gb = Encoding.GetEncoding("GB2312").GetString(buffer1);
            //如
            string a = "如";
            string b = "女";
            string c = "籹";


            buffer1 = Encoding.GetEncoding("GBK").GetBytes(a); //200 231
            byte[] buffer2 = Encoding.GetEncoding("GBK").GetBytes(b); //197 174

            buffer1 = Encoding.Unicode.GetBytes(a);
            buffer2 = Encoding.Unicode.GetBytes(c);


            if (a.Contains("女"))
            {

            }
            #endregion
            string fullname = AppDomain.CurrentDomain.BaseDirectory + "System\\16.png";
            pictureBox1 = new PictureBox();

            pictureBox1.Image = Image.FromFile(fullname);

            string FPath = fullname;
            Image ig = pictureBox1.Image;


            clsAllnew BusinessHelp = new clsAllnew();
            addtxt(pictureBox1.Image, "文口", fullname);

            //  pictureBox1.Image =  pictureBox1;


            saveFileDialog1.Filter = "BMP|*.bmp|JPEG|*.jpeg|GIF|*.gif|PNG|*.png";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string picPath = saveFileDialog1.FileName;
                string picType = picPath.Substring(picPath.LastIndexOf(".") + 1, (picPath.Length - picPath.LastIndexOf(".") - 1));
                switch (picType)
                {
                    case "bmp":
                        Bitmap bt = new Bitmap(pictureBox1.Image);
                        Bitmap mybmp = new Bitmap(bt, ig.Width, ig.Height);
                        mybmp.Save(picPath, ImageFormat.Bmp); break;
                    case "jpeg":
                        Bitmap bt1 = new Bitmap(pictureBox1.Image);
                        Bitmap mybmp1 = new Bitmap(bt1, ig.Width, ig.Height);
                        mybmp1.Save(picPath, ImageFormat.Jpeg); break;
                    case "gif":
                        Bitmap bt2 = new Bitmap(pictureBox1.Image);
                        Bitmap mybmp2 = new Bitmap(bt2, ig.Width, ig.Height);
                        mybmp2.Save(picPath, ImageFormat.Gif); break;
                    case "png":
                        Bitmap bt3 = new Bitmap(pictureBox1.Image);
                        Bitmap mybmp3 = new Bitmap(bt3, ig.Width, ig.Height);
                        mybmp3.Save(picPath, ImageFormat.Png); break;
                }
            }

        }
        #region 水印
        public void addtxt(Image ig, string txtChar, string FPath1)
        {
            {
                FPath = FPath1;

                pictureBox1 = new PictureBox();

                pictureBox1.Image = ig;
                if (txtChar.Trim() != "")
                {

                    int x2 = (int)(ig.Width - Fwidth) / 2;
                    int y2 = (int)(ig.Height - Fheight) / 2;
                    makeWatermark(x2, y2, txtChar.Trim());

                }
            }

        }

        public void makeWatermark(int x, int y, string txt)
        {
            System.Drawing.Image image = Image.FromFile(FPath);
            System.Drawing.Graphics e = System.Drawing.Graphics.FromImage(image);
            System.Drawing.Font f = new System.Drawing.Font(a, Fsize, Fstyle);
            System.Drawing.Brush b = new System.Drawing.SolidBrush(Fcolor);
            e.DrawString(txt, f, b, x, y);
            SizeF XMaxSize = e.MeasureString(txt, f);

            Fwidth = (int)XMaxSize.Width;
            Fheight = (int)XMaxSize.Height;

            e.Dispose();
            pictureBox1.Image = image;
        }
        #endregion

        private void tableLayoutPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void 打印ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bool ishow = true;
            reportForm = new ReportForm();
            reportForm.InitializeDataSource(qtyTable, null);
            if (ishow == true)
                reportForm.ShowDialog();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Word_web_Server = new List<clsKeyWord_web_info>();
            this.textBox1.Text = "";

            BindDataGridView();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (is_AdminIS == false)
            {
                MessageBox.Show("权限不够");

            }
            var form = new frmInputCenter();

            if (form.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button1_Click(null, EventArgs.Empty);

        }


    }
}
