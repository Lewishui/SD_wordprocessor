using clsBuiness;
using GZ_DQCleaningCompany;
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SD_wordprocessor
{

   
    public partial class frmInputCenter : Form
    {
        [DllImport("wininet.dll")]
        List<clsword_info> userlist_Server;
        List<clswordpart_info> wordpart_Server;
        List<clsbushoudaima_info> bushoudaima_Server;
        List<clsKeyWord_web_info> Word_webResult;
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        int findtype = 0;
        bool isnet = false;
        private string txname;
    
        public frmInputCenter()
        {
            InitializeComponent();
            wordpart_Server = new List<clswordpart_info>();
            this.comboBox2.SelectedIndex = 0;

            TabPage tp = tabControl1.TabPages[2];//在这里先保存，以便以后还要显示

            tabControl1.TabPages.Remove(tp);//隐藏（删除）


            //  tabControl1.TabPages.Insert(0, tp);//显示（插入）
            
            isnet = IsConnectionInternet();
        }

        private extern static bool InternetGetConnectedState(int Description, int ReservedValue);
        private bool IsConnectionInternet()
        {
            isnet = false;
            if (InternetGetConnectedState(0, 0) == true)
            {
                MessageBox.Show("系统已连上网络");
                isnet = true;

            }
            else
            {
                MessageBox.Show("系统未连接网络");
                isnet = false;

            }
            return isnet;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            userlist_Server = new List<clsword_info>();
            clsword_info item = new clsword_info();

            if (textBox10.Text == "" || textBox10.Text == "")
            {
                MessageBox.Show("请填写完整信息然后创建！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (textBox11.Text.Trim() != textBox11.Text.Trim())
            {
                MessageBox.Show("请填写完整信息然后创建，请重新输入！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            item.zi = textBox10.Text.Trim();
            item.zhengtidaima = textBox11.Text.Trim();
            item.bushou1 = this.textBox1.Text.Trim();
            item.bushou2 = this.textBox6.Text.Trim();
            item.bushou3 = this.textBox9.Text.Trim();
            item.bushou4 = this.textBox2.Text.Trim();
            item.bushou5 = this.textBox5.Text.Trim();
            item.bushou6 = this.textBox8.Text.Trim();
            item.bushou7 = this.textBox3.Text.Trim();
            item.bushou8 = this.textBox4.Text.Trim();
            item.bushou9 = this.textBox7.Text.Trim();

            clswordpart_info maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox1");
            if (maichangmingchenglist != null)
                item.bushou1daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox2");
            if (maichangmingchenglist != null)
                item.bushou2daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox3");
            if (maichangmingchenglist != null)
                item.bushou3daima = maichangmingchenglist.bushou_daima.Trim();


            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox4");
            if (maichangmingchenglist != null)
                item.bushou4daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox5");
            if (maichangmingchenglist != null)
                item.bushou5daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox6");
            if (maichangmingchenglist != null)
                item.bushou6daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox7");
            if (maichangmingchenglist != null)
                item.bushou7daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox8");
            if (maichangmingchenglist != null)
                item.bushou8daima = maichangmingchenglist.bushou_daima.Trim();

            maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == "textBox9");
            if (maichangmingchenglist != null)
                item.bushou9daima = maichangmingchenglist.bushou_daima.Trim();

            item.jiegou = this.comboBox2.Text.Trim();
            item.jiegoudaima = this.textBox37.Text.Trim();

            item.ku_id = this.comboBox1.Text.Trim();


            item.Input_Date = DateTime.Now.ToString("yyyy/MM/dd/HH");
            //item.ku_id = "1";

            userlist_Server.Add(item);
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.createWordKU_Server(userlist_Server);
            if (isnet == true)
            {

            }
            MessageBox.Show("创建成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // this.Close();

        }

        private void textBox1_Click(object sender, EventArgs e)
        {

            txname = "textBox1";

            tx12();

        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            clswordpart_info maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == txname);
            if (maichangmingchenglist == null && txname != null)
            {
                clswordpart_info item = new clswordpart_info();

                item.bushou_name = txname;
                item.bushou_daima = textBox12.Text;


                wordpart_Server.Add(item);
            }
            else if (maichangmingchenglist != null && txname != null)
            {
                if (txname != null)
                    maichangmingchenglist.bushou_name = txname;
                maichangmingchenglist.bushou_daima = textBox12.Text;


            }
        }

        private void textBox6_Click(object sender, EventArgs e)
        {

            txname = "textBox2";
            tx12();

        }

        private void textBox9_Click(object sender, EventArgs e)
        {
            txname = "textBox3";

            tx12();
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            txname = "textBox4";

            tx12();
        }

        private void tx12()
        {
            clswordpart_info maichangmingchenglist = wordpart_Server.Find(o => o.bushou_name != null && o.bushou_name == txname);
            if (maichangmingchenglist != null)
                textBox12.Text = maichangmingchenglist.bushou_daima;
            else
                textBox12.Text = "";
        }

        private void textBox5_Click(object sender, EventArgs e)
        {
            txname = "textBox5";
            tx12();
        }

        private void textBox8_Click(object sender, EventArgs e)
        {
            txname = "textBox6";
            tx12();
        }

        private void textBox3_Click(object sender, EventArgs e)
        {
            txname = "textBox7";
            tx12();
        }

        private void textBox4_Click(object sender, EventArgs e)
        {
            txname = "textBox8";
            tx12();
        }

        private void textBox7_ClientSizeChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_Click(object sender, EventArgs e)
        {
            txname = "textBox9";
            tx12();
        }
        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                findtype = 1;
            if (checkBox2.Checked == true)
                findtype = 2;
            if (checkBox2.Checked == true && checkBox1.Checked == true)
                findtype = 3;

            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(KEYFile);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.AutoGenerateColumns = true;
                    this.dataGridView1.DataSource = userlist_Server;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void KEYFile(object sender, DoWorkEventArgs e)
        {
            userlist_Server = new List<clsword_info>();

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            string strSelect = "select * from Word_ku where ku_id like'%" + findtype.ToString() + "%'";
            if (findtype == 3)
                strSelect = "select * from Word_ku";

            DateTime oldDate = DateTime.Now;

            userlist_Server = BusinessHelp.findWord(strSelect);
            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            bushoudaima_Server = new List<clsbushoudaima_info>();

            clsbushoudaima_info item = new clsbushoudaima_info();
            item.shang = textBox25.Text.Trim();
            item.xia = this.textBox24.Text.Trim();
            item.zuo = this.textBox30.Text.Trim();
            item.you = textBox27.Text.Trim();
            item.zhong = this.textBox23.Text.Trim();
            item.shangzhongxia = this.textBox31.Text.Trim();
            item.zuozhongyou = this.textBox28.Text.Trim();
            item.nei = this.textBox29.Text.Trim();
            item.wai = this.textBox26.Text.Trim();
            item.zuoyou_jiegou = this.textBox32.Text.Trim();
            item.zuoyou_jiegoudaima = this.textBox33.Text.Trim();

            item.shangxia_jiegou = this.textBox35.Text.Trim();
            item.shangxia_jiegoudaima = this.textBox34.Text.Trim();



            bushoudaima_Server.Add(item);

            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.createbushoudaimaServer(bushoudaima_Server);
            MessageBox.Show("创建成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);



        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedIndex == 2)
            {
                clsAllnew BusinessHelp = new clsAllnew();
                string strSelect = "select * from BuShouDaima";

                List<clsbushoudaima_info> ClaimReport_Server = BusinessHelp.find_bushoudaima(strSelect);

                foreach (clsbushoudaima_info item in ClaimReport_Server)
                {
                    textBox25.Text = item.shang;
                    this.textBox24.Text = item.xia;
                    this.textBox30.Text = item.zuo;
                    textBox27.Text = item.you;
                    textBox23.Text = item.zhong;
                    textBox31.Text = item.shangzhongxia;
                    this.textBox28.Text = item.zuozhongyou;
                    this.textBox29.Text = item.nei;
                    this.textBox26.Text = item.wai;
                    this.textBox32.Text = item.zuoyou_jiegou;
                    this.textBox33.Text = item.zuoyou_jiegoudaima;

                    this.textBox35.Text = item.shangxia_jiegou;
                    this.textBox34.Text = item.shangxia_jiegoudaima;
                }
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            showOrHide(false);

            if (comboBox2.Text == "左右")
            {
                label1.Visible = true;
                label6.Visible = true;
                textBox1.Visible = true;
                textBox6.Visible = true;


                jiegoudaima("9991");

            }
            else if (comboBox2.Text == "上下")
            {
                label1.Visible = true;
                label2.Visible = true;
                textBox1.Visible = true;
                textBox2.Visible = true;
                jiegoudaima("9992");

            }
            else if (comboBox2.Text == "内外")
            {
                label1.Visible = true;
                label2.Visible = true;
                textBox1.Visible = true;
                textBox2.Visible = true;
                jiegoudaima("9993");

            }
            else if (comboBox2.Text == "左中右")
            {
                label1.Visible = true;
                label6.Visible = true;
                textBox1.Visible = true;
                textBox6.Visible = true;
                label9.Visible = true;
                textBox9.Visible = true;
                jiegoudaima("9994");

            }
            else if (comboBox2.Text == "上中下")
            {


                label1.Visible = true;
                label2.Visible = true;
                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                label3.Visible = true;
                jiegoudaima("9995");


            }
            else if (comboBox2.Text == "复杂")
            {
                showOrHide(true);
                jiegoudaima("9996");
            }


        }

        private void jiegoudaima(string inpu)
        {
            textBox37.Text = inpu;
        }


        private void showOrHide(bool show)
        {
            label1.Visible = show;
            label2.Visible = show;
            textBox1.Visible = show;
            textBox2.Visible = show;
            textBox3.Visible = show;
            label3.Visible = show;

            label6.Visible = show;
            textBox6.Visible = show;

            label5.Visible = show;
            textBox5.Visible = show;


            label4.Visible = show;
            textBox4.Visible = show;


            label9.Visible = show;
            textBox9.Visible = show;

            label8.Visible = show;
            textBox8.Visible = show;

            label7.Visible = show;
            textBox7.Visible = show;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void button7_Click(object sender, EventArgs e)
        {
            clsAllnew BusinessHelp = new clsAllnew();
            Word_webResult = new List<clsKeyWord_web_info>();
            BusinessHelp.tsStatusLabel2 = toolStripLabel1;
            Word_webResult = BusinessHelp.ReadWeb_Report111();
            int di = 0;
            // foreach (clsKeyWord_web_info item in Word_webResult)
            {
                BusinessHelp.createWord_web_Server(Word_webResult);
                di++;

            }
            MessageBox.Show("ok" + Word_webResult.Count());


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void frmInputCenter_Load(object sender, EventArgs e)
        {

        }



    }



}
