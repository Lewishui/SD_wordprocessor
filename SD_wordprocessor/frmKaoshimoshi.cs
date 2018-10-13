using clsBuiness;
using DCTS.CustomComponents;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        public frmKaoshimoshi()
        {
            InitializeComponent();
            string sss = checkku();

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
            if (checkname == "")
            {
                MessageBox.Show("请重新选择库数据！");
                return;

            }

            Random random = new Random(System.Guid.NewGuid().GetHashCode());
            int r = random.Next();



            string strSelect = "select * from Word_ku where ku_id like'%" + checkname + "%'" + "order by RANDOM() limit 100";

            //string strSelect = "select * from Word_ku where name='" + findtext + "'";

            clsAllnew BusinessHelp = new clsAllnew();
            Loglist_Server = new List<clsword_info>();
            Loglist_Server = BusinessHelp.findWord(strSelect);
            BindDataGridView();
        }
        private void BindDataGridView()
        {

            if (Loglist_Server != null)
            {
                sortableLogList = new SortableBindingList<clsword_info>(Loglist_Server);
                bindingSource1.DataSource = new SortableBindingList<clsword_info>(Loglist_Server);
                dataGridView1.AutoGenerateColumns = true;
                var qtyTable = new DataTable();
                for (int i = 0; i < 10; i++)
                    qtyTable.Columns.Add((i + 1).ToString(), System.Type.GetType("System.String"));
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
                            for (int i = 0; i < 10; i++)
                            {
                                if (Loglist_Server.Count > jk)
                                {
                                    qtyTable.Rows[j][i] = Loglist_Server[jk].zi;
                                    jk++;
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

    }
}
