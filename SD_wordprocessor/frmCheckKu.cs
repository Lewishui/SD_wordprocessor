using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SD_wordprocessor
{
    public partial class frmCheckKu : Form
    {
        public string checkname;



        public frmCheckKu()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
                checkname = "1";
            if (checkBox2.Checked == true)
                checkname += "2";
            this.Close();

        }
    }
}
