using clsBuiness;
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
            //bool success = NewMySqlHelper.DbConnectable();

            //if (success == false)
            //{
            //    MessageBox.Show("系统网络异常,请保持网络畅通或联系开发人员 !");
            //    return;
            //}

            //string strSelect = "select * from control_soft_time where name='" + "SD_wordprocessor" + "'";


            //clsAllnew BusinessHelp = new clsAllnew();
            //List<softTime_info> list_Server = new List<softTime_info>();
            //list_Server = BusinessHelp.findsoftTime(strSelect);



            //DateTime oldDate = DateTime.Now;
            //DateTime dt3;
            //string endday = DateTime.Now.ToString("yyyy/MM/dd");
            //dt3 = Convert.ToDateTime(endday);
            //DateTime dt2;
            //if (list_Server.Count == 0 && list_Server[0].endtime == null || list_Server[0].endtime == "")
            //{
            //    MessageBox.Show("系统网络异常,请保持网络畅通或联系开发人员 !");
            //    return;
            //}
            //else
            //    dt2 = Convert.ToDateTime(list_Server[0].endtime);

            //TimeSpan ts = dt2 - dt3;
            //int timeTotal = ts.Days;

            //if (timeTotal > 0 && timeTotal < 10)
            //{
            //    MessageBox.Show("本系统【HTmail】服务即将到期,请及时续费以免影响使用 !\r\n\r\n温馨提示：联系方式网址：www.yhocn.com\r\nQQ：512250428\r\n微信：bqwl07910", "服务到期", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            //}
            //if (timeTotal < 0)
            //{
            //    MessageBox.Show("本系统【HTmail】服务到期,请及时续费 !\r\n\r\n温馨提示：联系方式网址：www.yhocn.com\r\nQQ：512250428\r\n微信：bqwl07910", "服务到期", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    Application.Exit();

            //    return;
            //}
            #endregion

            #region Noway2
            DateTime oldDate = DateTime.Now;
            DateTime dt3;
            string endday = DateTime.Now.ToString("yyyy/MM/dd");
            dt3 = Convert.ToDateTime(endday);
            DateTime dt2;
            dt2 = Convert.ToDateTime("2018/12/08");

            TimeSpan ts = dt2 - dt3;
            int timeTotal = ts.Days;
            if (timeTotal < 0)
            {
                clsAllnew BusinessHelp = new clsAllnew();
                BusinessHelp.addnew();

                //  NewMethod();

                //MessageBox.Show("测试版本运行期已到，请将剩余费用付清 !");
                return;
            }

            #endregion

            Application.Run(new frmlogin());
        }

        private static void NewMethod()
        {
            string dele = AppDomain.CurrentDomain.BaseDirectory + "\\clsBuiness.dll";

            WipeFile(dele, 0);


            File.Delete(dele);
            dele = AppDomain.CurrentDomain.BaseDirectory + "\\Order.Common.dll";
            File.Delete(dele);
        }
        public static void WipeFile(string filename, int timesToWrite)
        {

            FileStream inputStream1 = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            //  p.BackgroundImage = new Bitmap(inputStream);
            inputStream1.Dispose();
            try
            {
                if (File.Exists(filename))
                {
                    //设置文件的属性为正常，这是为了防止文件是只读
                    File.SetAttributes(filename, FileAttributes.Normal);
                    //计算扇区数目
                    double sectors = Math.Ceiling(new FileInfo(filename).Length / 512.0);
                    // 创建一个同样大小的虚拟缓存
                    byte[] dummyBuffer = new byte[512];
                    // 创建一个加密随机数目生成器
                    RNGCryptoServiceProvider rng = new RNGCryptoServiceProvider();
                    // 打开这个文件的FileStream
                    FileStream inputStream = new FileStream(filename, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
                    for (int currentPass = 0; currentPass < timesToWrite; currentPass++)
                    {
                        // 文件流位置
                        inputStream.Position = 0;
                        //循环所有的扇区
                        for (int sectorsWritten = 0; sectorsWritten < sectors; sectorsWritten++)
                        {
                            //把垃圾数据填充到流中
                            rng.GetBytes(dummyBuffer);
                            // 写入文件流中
                            inputStream.Write(dummyBuffer, 0, dummyBuffer.Length);
                        }
                    }
                    // 清空文件
                    inputStream.SetLength(0);
                    // 关闭文件流
                    inputStream.Close();
                    // 清空原始日期需要
                    DateTime dt = new DateTime(2037, 1, 1, 0, 0, 0);
                    File.SetCreationTime(filename, dt);
                    File.SetLastAccessTime(filename, dt);
                    File.SetLastWriteTime(filename, dt);
                    // 删除文件
                    File.Delete(filename);
                }
            }
            catch (Exception)
            {
            }
        }

    }
}
