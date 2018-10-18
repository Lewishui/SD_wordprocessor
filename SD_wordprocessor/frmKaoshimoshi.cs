using clsBuiness;
using DCTS.CustomComponents;
using NPOI.XWPF.UserModel;
using SDdb;
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
using WeifenLuo.WinFormsUI.Docking;

namespace SD_wordprocessor
{
    public partial class frmKaoshimoshi : DockContent
    {
        List<clsword_info> Loglist_Server;
        private SortableBindingList<clsword_info> sortableLogList;
        string checkname;
        public ReportForm reportForm;
        DataTable qtyTable;

        public frmKaoshimoshi()
        {
            InitializeComponent();
            string sss = checkku();
            reportForm = new ReportForm();
        }
        
        public string checkku()
        {
            checkname = "";
            var form = new frmCheckKu();
            if (form.ShowDialog() == DialogResult.OK)
            {
                checkname = form.checkname;
            }
            if (checkname == "")
            {
                MessageBox.Show("请重新选择库数据！");
                return "";

            }
            return checkname;
        }

        private void frmKaoshimoshi_Load(object sender, EventArgs e)
        {


            randomM();


        }

        private void randomM()
        {
            clsAllnew BusinessHelp = new clsAllnew();
            if (checkname == "")
            {
                MessageBox.Show("请重新选择库数据！");
                return;

            }

            Random random = new Random(System.Guid.NewGuid().GetHashCode());
            int r = random.Next();


            string strSelect = "select * from Word_ku where ku_id like'%" + checkname + "%'" + "order by RANDOM() limit 100";
            if (checkname == "12")
            {
                List<clsword_info> allLoglist_Server = new List<clsword_info>();
                Loglist_Server = new List<clsword_info>();
                strSelect = "select * from Word_ku where ku_id like'%" + "1" + "%'" + "order by RANDOM() limit 50";

                Loglist_Server = BusinessHelp.findWord(strSelect);

                strSelect = "select * from Word_ku where ku_id like'%" + "2" + "%'" + "order by RANDOM() limit 50";

                allLoglist_Server = BusinessHelp.findWord(strSelect);
                allLoglist_Server = allLoglist_Server.Concat(Loglist_Server).ToList();

                Loglist_Server = new List<clsword_info>();
                Loglist_Server = allLoglist_Server;
            }

            else
            {
                //string strSelect = "select * from Word_ku where name='" + findtext + "'";
                Loglist_Server = new List<clsword_info>();
                Loglist_Server = BusinessHelp.findWord(strSelect);
            }


            BindDataGridView();
        }
        private void BindDataGridView()
        {

            if (Loglist_Server != null)
            {
                sortableLogList = new SortableBindingList<clsword_info>(Loglist_Server);
                bindingSource1.DataSource = new SortableBindingList<clsword_info>(Loglist_Server);
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
                                if (Loglist_Server.Count > jk)
                                {
                                    isjioushu = IsOdd(i);

                                    if (isjioushu == false || i == 0)
                                    {
                                        qtyTable.Rows[j][i] = Loglist_Server[jk].zi;
                                        jk++;
                                    }
                                    else
                                    {
                                        qtyTable.Rows[j][i] = Loglist_Server[jk].zhengtidaima;
                                        jk++;
                                    }
                                }
                            }


                        }
                    }
                }
                this.bindingSource1.DataSource = qtyTable;
                dataGridView1.DataSource = bindingSource1;
                this.toolStripLabel1.Text = "条数：" + Loglist_Server.Count.ToString();
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

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            checkku();

            randomM();


        }

        private void 发信模板ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string delimiter = "\t";
            string body = "";

            for (int j = 0; j < dataGridView1.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView1.Columns.Count; k++)
                {
                    if (dataGridView1.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += dataGridView1.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;

                    }
                    else
                    {
                        strRowValue += dataGridView1.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                body += delimiter + strRowValue;

                //sw.WriteLine(strRowValue);
            }


            if (body == "")
                return;

            clsAllnew BusinessHelp = new clsAllnew();
            //   BusinessHelp.downcsv(dataGridView1);
            string strFileName = "  信息" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + "\\" + strFileName + ".docx";

            createWord(dir, body);

            //   Button1_Click(this, EventArgs.Empty);

        }
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

        private void 保存打印ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DataTable qtyTable = PDFcheckdaadetail_Method();
            bool ishow = true;


            reportForm = new ReportForm();
            reportForm.InitializeDataSource(qtyTable, Loglist_Server);
            if (ishow == true)
                reportForm.ShowDialog();
        }
        private DataTable PDFcheckdaadetail_Method()
        {
            DataTable qtyTable = new DataTable();
            foreach (clsword_info item in Loglist_Server)
            {
                qtyTable = new DataTable();
                qtyTable.Columns.Add("name1", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name2", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name3", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name4", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name5", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name6", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name7", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name8", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name9", System.Type.GetType("System.String"));
                qtyTable.Columns.Add("name10", System.Type.GetType("System.String"));
                foreach (clsword_info k in Loglist_Server)
                {
                    qtyTable.Rows.Add(qtyTable.NewRow());
                }





            }
            return qtyTable;


        }
    }
}
