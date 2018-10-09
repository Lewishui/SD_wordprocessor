using clsBuiness;
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.Linq;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace SD_wordprocessor
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);


            #region Noway
            bool success = NewMySqlHelper.DbConnectable();

            if (success == false)
            {
                MessageBox.Show("系统网络异常,请保持网络畅通或联系开发人员 !");
                return;
            }

            string strSelect = "select * from control_soft_time where name='" + "SD_wordprocessor" + "'";


            clsAllnew BusinessHelp = new clsAllnew();
            List<softTime_info> list_Server = new List<softTime_info>();
            list_Server = BusinessHelp.findsoftTime(strSelect);



            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            if (list_Server.Count == 0 && list_Server[0].endtime == null || list_Server[0].endtime == "")
            {
                MessageBox.Show("系统网络异常,请保持网络畅通或联系开发人员 !");
                return;
            }
            else
                dt2 = Convert.ToDateTime(list_Server[0].endtime);

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;

            if (timeTotal > 0 && timeTotal < 10)
            {
                MessageBox.Show("本系统【HTmail】服务即将到期,请及时续费以免影响使用 !\r\n\r\n温馨提示：联系方式网址：www.yhocn.com\r\nQQ：512250428\r\n微信：bqwl07910", "服务到期", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            if (timeTotal < 0)
            {
                MessageBox.Show("本系统【HTmail】服务到期,请及时续费 !\r\n\r\n温馨提示：联系方式网址：www.yhocn.com\r\nQQ：512250428\r\n微信：bqwl07910", "服务到期", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();

                return;
            }
            #endregion



            Application.Run(new frmlogin());
        }
    }
}
