using clsBuiness;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SD_wordprocessor
{
    public partial class frmInputCenter : Form
    {
        List<clsword_info> userlist_Server;
        List<clswordpart_info> wordpart_Server;
        List<clsbushoudaima_info> bushoudaima_Server;

        private string txname;

        public frmInputCenter()
        {
            InitializeComponent();
            wordpart_Server = new List<clswordpart_info>();

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



            item.Input_Date = DateTime.Now.ToString("yyyy/MM/dd/HH");
            item.ku_id = "1";

            userlist_Server.Add(item);
            clsAllnew BusinessHelp = new clsAllnew();

            BusinessHelp.createWordKU_Server(userlist_Server);

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
            else
            {
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

        private void button4_Click(object sender, EventArgs e)
        {

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



    }



}
