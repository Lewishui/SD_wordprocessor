
using ISR_System;
using mshtml;
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace clsBuiness
{
    public enum ProcessStatus
    {
        初始化,
        登录界面,
        确认YES,
        第一页面,
        第二页面,
        Filter下拉,
        关闭页面,
        结束页面

    }
    public class clsAllnew
    {
        private string dataSource = "Data.sqlite";
        string newsth;
        public BackgroundWorker bgWorker1;
        private ProcessStatus isrun = ProcessStatus.初始化;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public PictureBox pictureBox1;
        int Fwidth;
        int Fheight;
        public string FPath;
        FontStyle Fstyle = FontStyle.Regular;
        float Fsize = 18;
        Color Fcolor = System.Drawing.Color.Yellow;
        FontFamily a = FontFamily.GenericMonospace;
        private WbBlockNewUrl MyWebBrower;
        private Form viewForm;
        WbBlockNewUrl myDoc = null;
        private int login;
        private bool isOneFinished = false;
        bool loading;
        bool loading_bushoujiansuo;
        private DateTime StopTime;
        List<clsKeyWord_web_info> ALLWord_webResult;

        clsKeyWord_web_info puitem;
        int ongoingIndex = 0;
        int jiegoudaima = 0;



        public clsAllnew()
        {
            newsth = AppDomain.CurrentDomain.BaseDirectory + "" + dataSource;

        }
        public List<softTime_info> findsoftTime(string findtext)
        {
            //    findtext = sqlAddPCID(findtext);
            MySql.Data.MySqlClient.MySqlDataReader reader = NewMySqlHelper.ExecuteReader(findtext);
            List<softTime_info> ClaimReport_Server = new List<softTime_info>();

            while (reader.Read())
            {
                softTime_info item = new softTime_info();
                if (reader.GetValue(0) != null && Convert.ToString(reader.GetValue(0)) != "")
                    item._id = Convert.ToString(reader.GetValue(0));

                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.name = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.starttime = reader.GetString(2);
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.endtime = reader.GetString(3);

                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;
        }

        public List<clsuserinfo> findUser(string findtext)
        {
            string strSelect = "select * from _User where name='" + findtext + "'";

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + dataSource);

            dbConn.Open();
            SQLiteCommand dbCmd = dbConn.CreateCommand();
            //dbCmd.CommandText = "SELECT * FROM tuijanhaoma";
            dbCmd.CommandText = strSelect;

            DbDataReader reader = SQLiteHelper.ExecuteReader("Data Source=" + newsth, dbCmd);
            List<clsuserinfo> ClaimReport_Server = new List<clsuserinfo>();

            while (reader.Read())
            {
                clsuserinfo item = new clsuserinfo();

                if (reader.GetValue(0) != null && Convert.ToString(reader.GetValue(0)) != "")
                    item.Order_id = reader.GetString(0);
                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.name = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.password = reader.GetString(2);
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.Createdate = reader.GetString(3);
                if (reader.GetValue(4) != null && Convert.ToString(reader.GetValue(4)) != "")
                    item.Btype = reader.GetString(4);
                if (reader.GetValue(5) != null && Convert.ToString(reader.GetValue(5)) != "")
                    item.denglushijian = reader.GetString(5);
                if (reader.GetValue(6) != null && Convert.ToString(reader.GetValue(6)) != "")
                    item.jigoudaima = reader.GetString(6);
                if (reader.GetValue(7) != null && Convert.ToString(reader.GetValue(7)) != "")
                    item.userTime = reader.GetString(7);
                if (reader.GetValue(8) != null && Convert.ToString(reader.GetValue(8)) != "")
                    item.AdminIS = reader.GetString(8);


                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }
        public void updateLoginTime_Server(List<clsuserinfo> AddMAPResult)
        {
            string sql = "update _User set denglushijian ='" + AddMAPResult[0].denglushijian.Trim() + "' where name ='" + AddMAPResult[0].name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;

        }
        public void update_userTime_Server(List<clsuserinfo> AddMAPResult)
        {
            string sql = "update _User set userTime ='" + AddMAPResult[0].userTime.Trim() + "' where name ='" + AddMAPResult[0].name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;

        }

        public void createUser_Server(List<clsuserinfo> AddMAPResult)
        {
            //string sql = "insert into _User(name,password,Createdate,Btype,denglushijian,jigoudaima,AdminIS) values ('" + AddMAPResult[0].name + "','" + AddMAPResult[0].password + "','" + AddMAPResult[0].Createdate + "','" + AddMAPResult[0].Btype + "','" + AddMAPResult[0].denglushijian + "','" + AddMAPResult[0].jigoudaima + "','" + AddMAPResult[0].AdminIS + "')";
            string sql = "INSERT INTO _User ( name, password,Createdate,Btype,denglushijian,jigoudaima,userTime,AdminIS ) " +

                      "VALUES (\"" + AddMAPResult[0].name + "\"" +

                             ",\"" + AddMAPResult[0].password + "\"" +
                                  ",\"" + AddMAPResult[0].Createdate + "\"" +
                                       ",\"" + AddMAPResult[0].Btype + "\"" +
                                            ",\"" + AddMAPResult[0].denglushijian + "\"" +
                                                 ",\"" + AddMAPResult[0].jigoudaima + "\"" +
                                                     ",\"" + AddMAPResult[0].userTime + "\"" +
                             ",\"" + AddMAPResult[0].AdminIS + "\")";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;



        }
        public void lock_Userpassword_Server(List<clsuserinfo> AddMAPResult)
        {
            string sql = "update _User set Btype ='" + AddMAPResult[0].Btype.Trim() + "' where name ='" + AddMAPResult[0].name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;

        }
        public List<clsuserinfo> ReadUserlistfromServer()
        {
            string conditions = "select * from _User";//成功

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + dataSource);

            dbConn.Open();
            SQLiteCommand dbCmd = dbConn.CreateCommand();
            dbCmd.CommandText = conditions;
            DbDataReader reader = SQLiteHelper.ExecuteReader("Data Source=" + newsth, dbCmd);
            List<clsuserinfo> ClaimReport_Server = new List<clsuserinfo>();
            while (reader.Read())
            {
                clsuserinfo item = new clsuserinfo();
                if (reader.GetValue(0) != null && Convert.ToString(reader.GetValue(0)) != "")
                    item.Order_id = reader.GetString(0);
                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.name = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.password = reader.GetString(2);
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.Createdate = reader.GetString(3);
                if (reader.GetValue(4) != null && Convert.ToString(reader.GetValue(4)) != "")
                    item.Btype = reader.GetString(4);
                if (reader.GetValue(5) != null && Convert.ToString(reader.GetValue(5)) != "")
                    item.denglushijian = reader.GetString(5);
                if (reader.GetValue(6) != null && Convert.ToString(reader.GetValue(6)) != "")
                    item.jigoudaima = reader.GetString(6);

                if (reader.GetValue(7) != null && Convert.ToString(reader.GetValue(7)) != "")
                    item.userTime = reader.GetString(7);
                if (reader.GetValue(8) != null && Convert.ToString(reader.GetValue(8)) != "")
                    item.AdminIS = reader.GetString(8);

                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }

        public void deleteUSER(string name)
        {
            string sql2 = "delete from _User where  name='" + name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql2, CommandType.Text, null);

            return;

        }
        public void changeUserpassword_Server(List<clsuserinfo> AddMAPResult)
        {
            string sql = "update _User set password ='" + AddMAPResult[0].password.Trim() + "' where name ='" + AddMAPResult[0].name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);
            return;

        }


        public void createWordKU_Server(List<clsword_info> AddMAPResult)
        {
            string sql = "INSERT INTO Word_ku ( zi, zhengtidaima,bushou1,bushou2,bushou3,bushou4,bushou5,bushou6,bushou7,bushou8,bushou9,Input_Date,ku_id,Message,bushou1daima,bushou2daima,bushou3daima,bushou4daima,bushou5daima,bushou6daima,bushou7daima,bushou8daima,bushou9daima,jiegou,jiegoudaima  ) " +

                      "VALUES (\"" + AddMAPResult[0].zi + "\"" +

                             ",\"" + AddMAPResult[0].zhengtidaima + "\"" +
                                  ",\"" + AddMAPResult[0].bushou1 + "\"" +
                                       ",\"" + AddMAPResult[0].bushou2 + "\"" +
                                            ",\"" + AddMAPResult[0].bushou3 + "\"" +
                                                 ",\"" + AddMAPResult[0].bushou4 + "\"" +
                                                     ",\"" + AddMAPResult[0].bushou5 + "\"" +
                                                            ",\"" + AddMAPResult[0].bushou6 + "\"" +
                                                                   ",\"" + AddMAPResult[0].bushou7 + "\"" +
                                                                          ",\"" + AddMAPResult[0].bushou8 + "\"" +
                            ",\"" + AddMAPResult[0].bushou9 + "\"" +
                         ",\"" + AddMAPResult[0].Input_Date + "\"" +
                            ",\"" + AddMAPResult[0].ku_id + "\"" +
                          ",\"" + AddMAPResult[0].Message + "\"" +
                           ",\"" + AddMAPResult[0].bushou1daima + "\"" +
                            ",\"" + AddMAPResult[0].bushou2daima + "\"" +
                         ",\"" + AddMAPResult[0].bushou3daima + "\"" +
                        ",\"" + AddMAPResult[0].bushou4daima + "\"" +
                      ",\"" + AddMAPResult[0].bushou5daima + "\"" +
                         ",\"" + AddMAPResult[0].bushou6daima + "\"" +
                          ",\"" + AddMAPResult[0].bushou7daima + "\"" +
                     ",\"" + AddMAPResult[0].bushou8daima + "\"" +
                             ",\"" + AddMAPResult[0].bushou9daima + "\"" +
                     ",\"" + AddMAPResult[0].jiegou + "\"" +

                             ",\"" + AddMAPResult[0].jiegoudaima + "\")";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;
        }

        public List<clsword_info> findWord(string findtext)
        {

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + dataSource);

            dbConn.Open();
            SQLiteCommand dbCmd = dbConn.CreateCommand();
            //dbCmd.CommandText = "SELECT * FROM tuijanhaoma";
            dbCmd.CommandText = findtext;

            DbDataReader reader = SQLiteHelper.ExecuteReader("Data Source=" + newsth, dbCmd);
            List<clsword_info> ClaimReport_Server = new List<clsword_info>();

            while (reader.Read())
            {
                clsword_info item = new clsword_info();

                item.Order_id = reader.GetInt32(0);
                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.zi = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.zhengtidaima = Convert.ToString(reader.GetString(2));
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.bushou1 = reader.GetString(3);
                if (reader.GetValue(4) != null && Convert.ToString(reader.GetValue(4)) != "")

                    item.bushou2 = reader.GetString(4);
                if (reader.GetValue(5) != null && Convert.ToString(reader.GetValue(5)) != "")

                    item.bushou3 = reader.GetString(5);
                if (reader.GetValue(6) != null && Convert.ToString(reader.GetValue(6)) != "")

                    item.bushou4 = reader.GetString(6);
                if (reader.GetValue(7) != null && Convert.ToString(reader.GetValue(7)) != "")

                    item.bushou5 = reader.GetString(7);
                if (reader.GetValue(8) != null && Convert.ToString(reader.GetValue(8)) != "")

                    item.bushou6 = reader.GetString(8);
                if (reader.GetValue(9) != null && Convert.ToString(reader.GetValue(9)) != "")
                    item.bushou7 = Convert.ToString(reader.GetString(9));

                if (reader.GetValue(10) != null && Convert.ToString(reader.GetValue(10)) != "")

                    item.bushou8 = Convert.ToString(reader.GetString(10));
                if (reader.GetValue(11) != null && Convert.ToString(reader.GetValue(11)) != "")

                    item.bushou9 = reader.GetString(11);
                if (reader.GetValue(12) != null && Convert.ToString(reader.GetValue(12)) != "")

                    item.Input_Date = reader.GetString(12);
                if (reader.GetValue(13) != null && Convert.ToString(reader.GetValue(13)) != "")

                    item.ku_id = reader.GetString(13);


                if (reader.GetString(14) != null && reader.GetString(14) != "")
                    item.Message = Convert.ToString(reader.GetString(14));
                if (reader.GetValue(15) != null && Convert.ToString(reader.GetValue(15)) != "")

                    item.bushou1daima = reader.GetString(15);
                if (reader.GetValue(16) != null && Convert.ToString(reader.GetValue(16)) != "")
                    item.bushou2daima = Convert.ToString(reader.GetString(16));

                if (reader.GetValue(17) != null && Convert.ToString(reader.GetValue(17)) != "")
                    item.bushou3daima = Convert.ToString(reader.GetString(17));


                if (reader.GetValue(18) != null && Convert.ToString(reader.GetValue(18)) != "")
                    item.bushou4daima = Convert.ToString(reader.GetString(18));

                if (reader.GetValue(19) != null && Convert.ToString(reader.GetValue(19)) != "")
                    item.bushou5daima = Convert.ToString(reader.GetString(19));

                if (reader.GetValue(20) != null && Convert.ToString(reader.GetValue(20)) != "")
                    item.bushou6daima = Convert.ToString(reader.GetString(20));
                if (reader.GetValue(21) != null && Convert.ToString(reader.GetValue(21)) != "")
                    item.bushou7daima = Convert.ToString(reader.GetString(21));
                if (reader.GetValue(22) != null && Convert.ToString(reader.GetValue(22)) != "")
                    item.bushou8daima = Convert.ToString(reader.GetString(22));
                if (reader.GetValue(23) != null && Convert.ToString(reader.GetValue(23)) != "")
                    item.bushou9daima = Convert.ToString(reader.GetString(23));

                if (reader.GetValue(24) != null && Convert.ToString(reader.GetValue(24)) != "")
                    item.jiegou = Convert.ToString(reader.GetString(24));


                if (reader.GetValue(25) != null && Convert.ToString(reader.GetValue(25)) != "")
                    item.jiegoudaima = Convert.ToString(reader.GetString(25));


                //if (reader.GetValue(26) != null && Convert.ToString(reader.GetValue(26)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(26));

                //if (reader.GetValue(27) != null && Convert.ToString(reader.GetValue(27)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(27));

                //if (reader.GetValue(28) != null && Convert.ToString(reader.GetValue(28)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(28));

                //if (reader.GetValue(29) != null && Convert.ToString(reader.GetValue(29)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(29));



                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }
        public void downcsv(DataGridView dataGridView)
        {
            if (dataGridView.Rows.Count == 0)
            {
                MessageBox.Show("Sorry , No Data Output !", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.Filter = "csv|*.csv";
            string strFileName = "  信息" + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            saveFileDialog.FileName = strFileName;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                strFileName = saveFileDialog.FileName.ToString();
            }
            else
            {
                return;
            }
            FileStream fa = new FileStream(strFileName, FileMode.Create);
            StreamWriter sw = new StreamWriter(fa, Encoding.Unicode);
            string delimiter = "\t";
            string strHeader = "";
            for (int i = 0; i < dataGridView.Columns.Count; i++)
            {
                strHeader += dataGridView.Columns[i].HeaderText + delimiter;
            }
            sw.WriteLine(strHeader);

            //output rows data
            for (int j = 0; j < dataGridView.Rows.Count; j++)
            {
                string strRowValue = "";

                for (int k = 0; k < dataGridView.Columns.Count; k++)
                {
                    if (dataGridView.Rows[j].Cells[k].Value != null)
                    {
                        strRowValue += dataGridView.Rows[j].Cells[k].Value.ToString().Replace("\r\n", " ").Replace("\n", "") + delimiter;


                    }
                    else
                    {
                        strRowValue += dataGridView.Rows[j].Cells[k].Value + delimiter;
                    }
                }
                sw.WriteLine(strRowValue);
            }
            sw.Close();
            fa.Close();
            MessageBox.Show("下载完成 ！", "System", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void createbushoudaimaServer(List<clsbushoudaima_info> AddMAPResult)
        {
            string sql = "INSERT INTO BuShouDaima ( shang, xia,zuo,you,zhong,shangzhongxia,zuozhongyou,nei,wai,jiegou,jiegoudaiam) " +

                      "VALUES (\"" + AddMAPResult[0].shang + "\"" +

                             ",\"" + AddMAPResult[0].xia + "\"" +
                                  ",\"" + AddMAPResult[0].zuo + "\"" +
                                       ",\"" + AddMAPResult[0].you + "\"" +
                                            ",\"" + AddMAPResult[0].zhong + "\"" +
                                                 ",\"" + AddMAPResult[0].shangzhongxia + "\"" +
                                                     ",\"" + AddMAPResult[0].zuozhongyou + "\"" +
                                                            ",\"" + AddMAPResult[0].nei + "\"" +
       ",\"" + AddMAPResult[0].wai + "\"" +
                                                            ",\"" + AddMAPResult[0].zuoyou_jiegou + "\"" +
                                                               ",\"" + AddMAPResult[0].zuoyou_jiegoudaima + "\"" +
                                                            ",\"" + AddMAPResult[0].shangxia_jiegou + "\"" +


                             ",\"" + AddMAPResult[0].shangxia_jiegoudaima + "\")";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);

            return;
        }

        public List<clsbushoudaima_info> find_bushoudaima(string findtext)
        {

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + dataSource);

            dbConn.Open();
            SQLiteCommand dbCmd = dbConn.CreateCommand();
            //dbCmd.CommandText = "SELECT * FROM tuijanhaoma";
            dbCmd.CommandText = findtext;

            DbDataReader reader = SQLiteHelper.ExecuteReader("Data Source=" + newsth, dbCmd);
            List<clsbushoudaima_info> ClaimReport_Server = new List<clsbushoudaima_info>();

            while (reader.Read())
            {
                clsbushoudaima_info item = new clsbushoudaima_info();

                item.status_id = reader.GetInt32(0);
                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.shang = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.xia = Convert.ToString(reader.GetString(2));
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.zuo = reader.GetString(3);
                if (reader.GetValue(4) != null && Convert.ToString(reader.GetValue(4)) != "")

                    item.you = reader.GetString(4);
                if (reader.GetValue(5) != null && Convert.ToString(reader.GetValue(5)) != "")

                    item.zhong = reader.GetString(5);
                if (reader.GetValue(6) != null && Convert.ToString(reader.GetValue(6)) != "")

                    item.shangzhongxia = reader.GetString(6);
                if (reader.GetValue(7) != null && Convert.ToString(reader.GetValue(7)) != "")

                    item.zuozhongyou = reader.GetString(7);
                if (reader.GetValue(8) != null && Convert.ToString(reader.GetValue(8)) != "")

                    item.nei = reader.GetString(8);
                if (reader.GetValue(9) != null && Convert.ToString(reader.GetValue(9)) != "")
                    item.wai = Convert.ToString(reader.GetString(9));

                if (reader.GetValue(10) != null && Convert.ToString(reader.GetValue(10)) != "")

                    item.zuoyou_jiegou = reader.GetString(10);
                if (reader.GetValue(11) != null && Convert.ToString(reader.GetValue(11)) != "")
                    item.zuoyou_jiegoudaima = Convert.ToString(reader.GetString(11));

                if (reader.GetValue(12) != null && Convert.ToString(reader.GetValue(12)) != "")

                    item.shangxia_jiegou = reader.GetString(12);
                if (reader.GetValue(13) != null && Convert.ToString(reader.GetValue(13)) != "")
                    item.shangxia_jiegoudaima = Convert.ToString(reader.GetString(13));



                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }
        public static bool HasChineseTest(string text)
        {
            //string text = "是不是汉字，ABC,keleyi.com";
            char[] c = text.ToCharArray();
            bool ischina = false;

            for (int i = 0; i < c.Length; i++)
            {
                if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                {
                    ischina = true;

                }
                else
                {
                    //  ischina = false;
                }
            }
            return ischina;

        }

        public List<clsKeyWord_web_info> ReadWeb_Report111()
        {

            string path1 = AppDomain.CurrentDomain.BaseDirectory + "System\\all word .txt";

            string[] fileText1 = File.ReadAllLines(path1);
            ALLWord_webResult = new List<clsKeyWord_web_info>();
            for (int i = 2; i < fileText1.Length; i++)
            {
                string[] tatile1 = System.Text.RegularExpressions.Regex.Split(fileText1[i], "\t");
                for (int iq = 1; iq < tatile1.Length; iq++)
                {
                    clsKeyWord_web_info item = new clsKeyWord_web_info();
                    string name = tatile1[iq];
                    //   name = "我";
                    bool ishanzi = HasChineseTest(name);
                    if (tatile1[iq] != null && tatile1[iq] != "" && tatile1[iq] != "鼻")
                    {
                        item.word = tatile1[iq];
                        ALLWord_webResult.Add(item);
                    }
                }
            }
            // string ssd = ALLWord_webResult[71111].word;
            ongoingIndex = 0;
            //㻩 䯊 图 膔 鄠 落 励 𠄝
            List<clsKeyWord_web_info> Reaad_ALLWord_webResult = new List<clsKeyWord_web_info>();

            foreach (clsKeyWord_web_info item in ALLWord_webResult)
            {

               // Thread.Sleep(2000); 
                #region 繁体转简体
                if (string.IsNullOrEmpty(ALLWord_webResult[ongoingIndex].word))
                {
                    continue;
                }
                else
                {
                    string value = ALLWord_webResult[ongoingIndex].word.Trim();
                    string newValue = StringConvert(value, "2");
                    if (!string.IsNullOrEmpty(newValue)&&!newValue.Contains("?"))
                    {
                        ALLWord_webResult[ongoingIndex].word = newValue;

                    }
                }
                #endregion
                //   ALLWord_webResult[0].word = "机";
                isOneFinished = false;
             //  jiegoudaima = 0;//被屏蔽的
                jiegoudaima = 2;//查字网

                login = 0;
                puitem = new clsKeyWord_web_info();
                puitem = item;
                InitialWebbroswerIE2();
                StopTime = DateTime.Now;
                int time = 0;

                while (!isOneFinished)
                {
                    time++;
                    if (time > 200000)
                        time = 0;

                    System.Windows.Forms.Application.DoEvents();
                    DateTime rq2 = DateTime.Now;  //结束时间
                    int a = rq2.Second - StopTime.Second;
                    TimeSpan ts = rq2 - StopTime;
                    int timeTotal = ts.Seconds;

                    if (timeTotal >= 20)
                    {
                        //tsStatusLabel1.Text = "超出时间 正在退出....";
                        isOneFinished = true;
                        //StopTime = DateTime.Now;
                    }
                    if (jiegoudaima == 2 && isrun == ProcessStatus.第二页面)
                    {
                        getPIANpang2();
                    }
                }
                if (viewForm != null)
                {
                    MyWebBrower = null;
                    viewForm.Close();

                }
                Thread.Sleep(1000);
                jiegoudaima = 1;
                isOneFinished = false;
                login = 0;
                //获取结构
                puitem = new clsKeyWord_web_info();
                puitem = item;
                InitialWebbroswerIE2();
                StopTime = DateTime.Now;
                time = 0;

                while (!isOneFinished)
                {
                    time++;
                    if (time > 200000)
                        time = 0;

                    System.Windows.Forms.Application.DoEvents();
                    DateTime rq2 = DateTime.Now;  //结束时间
                    int a = rq2.Second - StopTime.Second;
                    TimeSpan ts = rq2 - StopTime;
                    int timeTotal = ts.Seconds;

                    if (timeTotal >= 20)
                    {
                        // tsStatusLabel1.Text = "超出时间 正在退出....";
                        isOneFinished = true;
                        //StopTime = DateTime.Now;
                    }
                    if (jiegoudaima == 1 && isrun == ProcessStatus.第二页面)
                    {
                        getjiegou();
                    }
                }
                if (viewForm != null)
                {
                    MyWebBrower = null;
                    viewForm.Close();

                }
                Reaad_ALLWord_webResult.Add(ALLWord_webResult[ongoingIndex]);
                login = 0;
                ongoingIndex++;


            }
            //2793 3306 2804 10916 3984
            return Reaad_ALLWord_webResult;




        }
        public string StringConvert(string x, string type)
        {
            String value = String.Empty;
            switch (type)
            {
                case "1"://转繁体
                    value = Microsoft.VisualBasic.Strings.StrConv(x, Microsoft.VisualBasic.VbStrConv.TraditionalChinese, 0);
                    break;
                case "2":
                    value = Microsoft.VisualBasic.Strings.StrConv(x, Microsoft.VisualBasic.VbStrConv.SimplifiedChinese, 0);
                    break;
                default:
                    break;
            }
            return value;
        }

        public void InitialWebbroswerIE2()
        {
            try
            {

                MyWebBrower = new WbBlockNewUrl();
                //不显示弹出错误继续运行框（HP方可）
                MyWebBrower.ScriptErrorsSuppressed = true;
                MyWebBrower.BeforeNewWindow += new EventHandler<WebBrowserExtendedNavigatingEventArgs>(MyWebBrower_BeforeNewWindow2);
                MyWebBrower.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(AnalysisWebInfo2);
                MyWebBrower.Dock = DockStyle.Fill;
                MyWebBrower.IsWebBrowserContextMenuEnabled = true;
                //显示用的窗体
                viewForm = new Form();
                //viewForm.Icon=
                viewForm.ClientSize = new System.Drawing.Size(550, 600);
                viewForm.StartPosition = FormStartPosition.CenterScreen;
                viewForm.Controls.Clear();
                viewForm.Controls.Add(MyWebBrower);
                viewForm.FormClosing += new FormClosingEventHandler(viewForm_FormClosing);
                viewForm.Show();
                if (jiegoudaima == 0)
                {
                    MyWebBrower.Url = new Uri("http://tool.httpcn.com/Zi/BuShou.html");
                    //MyWebBrower.Url = new Uri("http://www.haomeili.net/HanZi/SuoYouHanZi");
                }
                else if (jiegoudaima == 1)
                    MyWebBrower.Url = new Uri("https://zidian.911cha.com/");
                else if (jiegoudaima == 2)
                    MyWebBrower.Url = new Uri("http://www.chaziwang.com/bushou.html");


                tsStatusLabel1.Text = "接入 ...." + MyWebBrower.Url;

            }
            catch (Exception ex)
            {

                return;
                throw ex;
            }

        }
        void MyWebBrower_BeforeNewWindow2(object sender, WebBrowserExtendedNavigatingEventArgs e)
        {
            #region 在原有窗口导航出新页
            //e.Cancel = true;//http://pro.wwpack-crest.hp.com/wwpak.online/regResults.aspx
            //MyWebBrower.Navigate(e.Url);
            #endregion
        }
        private void viewForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (isrun != ProcessStatus.关闭页面)
            {
                //if (MessageBox.Show("正在进行，是否中止?", "关闭", MessageBoxButtons.OKCancel) == DialogResult.OK)
                //{
                //    if (MyWebBrower != null)
                //    {
                //        if (MyWebBrower.IsBusy)
                //        {
                //            MyWebBrower.Stop();
                //        }
                //        MyWebBrower.Dispose();
                //        MyWebBrower = null;
                //    }
                //}
                //else
                //{
                //    e.Cancel = true;
                //}
            }
        }
        protected void AnalysisWebInfo2(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            myDoc = sender as WbBlockNewUrl;
            if (jiegoudaima == 0)
            {
                #region 读取字形结构
                if (myDoc != null && myDoc.Url.ToString().IndexOf("http://tool.httpcn.com/Zi/BuShou.html") >= 0 && login == 0)
                {
                    // MyWebBrower.Navigate("http://tool.httpcn.com/Html/zi/BuShou/1_1.html");

                    HtmlElement userName = null;
                    HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("input");
                    HtmlElement submit = null;


                    foreach (HtmlElement item in atab)
                    {
                        if (item.GetAttribute("name") == "wd")
                            userName = item;
                        if (item.GetAttribute("Id") == "zisubmit")
                            submit = item;
                    }
                    if (userName != null)
                    {
                        userName.SetAttribute("Value", ALLWord_webResult[ongoingIndex].word);

                        submit.InvokeMember("Click");
                        isrun = ProcessStatus.第二页面;
                    }
                    //login++;
                    //loading = true;
                    //while (loading == true)
                    //{

                    //    Application.DoEvents();
                    //    Get_Kaijiang();

                    //}
                    login++;


                }
                if (myDoc.Url.ToString().IndexOf("http://tool.httpcn.com/Html/Zi/") >= 0 && isrun == ProcessStatus.第二页面)
                {

                    //loading = true;
                    //while (loading == true)
                    //{

                    //    Application.DoEvents();
                    getPIANpang();

                    //}
                    //isrun = ProcessStatus.关闭页面;
                    //isOneFinished = true;
                    //login = 0;


                }

                #endregion
            }
            if (jiegoudaima == 2)//之前网站屏蔽了
            {
                #region 读取字形结构
                if (myDoc != null && myDoc.Url.ToString().IndexOf("http://www.chaziwang.com/bushou.html") >= 0 && login == 0)
                {

                    HtmlElement userName = null;
                    HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("input");
                    HtmlElement submit = null;


                    foreach (HtmlElement item in atab)
                    {
                        if (item.GetAttribute("name") == "q")
                            userName = item;
                        if (item.GetAttribute("Id") == "su")
                            submit = item;
                    }
                    HtmlElementCollection atab1 = myDoc.Document.GetElementsByTagName("button");
                    foreach (HtmlElement item in atab1)
                    {

                        if (item.GetAttribute("id") == "su")
                            submit = item;
                    }

                    if (userName != null && submit != null)
                    {
                        userName.SetAttribute("Value", ALLWord_webResult[ongoingIndex].word);

                        submit.InvokeMember("Click");
                        isrun = ProcessStatus.第二页面;
                    }
                    //login++;
                    //loading = true;
                    //while (loading == true)
                    //{

                    //    Application.DoEvents();
                    //    Get_Kaijiang();

                    //}
                    login++;


                }
                if (myDoc.Url.ToString().IndexOf("http://www.chaziwang.com/bushou.html") >= 0 && isrun == ProcessStatus.第二页面)
                {

                    // loading = true;
                    //  while (loading == true)
                    {

                        //   Application.DoEvents();
                        // getPIANpang2();

                    }
                    //isrun = ProcessStatus.关闭页面;
                    //isOneFinished = true;
                    //login = 0;


                }

                #endregion
            }
            else if (jiegoudaima == 1)
            {
                if (myDoc != null && myDoc.Url.ToString().IndexOf("https://zidian.911cha.com") >= 0 && login == 0)
                {
                    HtmlElement userName = null;
                    HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("input");
                    HtmlElement submit = null;

                    foreach (HtmlElement item in atab)
                    {
                        if (item.GetAttribute("name") == "q")
                            userName = item;
                        if (item.GetAttribute("value") == "查询")
                            submit = item;
                    }
                    if (userName != null)
                    {
                        userName.SetAttribute("Value", ALLWord_webResult[ongoingIndex].word);

                        submit.InvokeMember("Click");
                        isrun = ProcessStatus.第二页面;
                        login++;
                        return;

                    }
                }
                else if (myDoc.Url.ToString().IndexOf("https://zidian.911cha.com") >= 0 && isrun == ProcessStatus.第二页面)
                {
                    //loading = true;
                    //while (loading == true)
                    //{

                    //    Application.DoEvents();
                    getjiegou();

                    //}
                    //isrun = ProcessStatus.关闭页面;
                    //isOneFinished = true;
                    //login = 0;


                }

            }


        }

        private void getPIANpang()
        {
            HtmlElement userNames = myDoc.Document.GetElementById("div_a1");
            HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("div");
            if (userNames != null)
            {

                string dd = userNames.InnerText;


                HtmlElementCollection atab11 = userNames.Document.GetElementsByTagName("div");

                foreach (HtmlElement item in atab)
                {
                    if (item.OuterText != null && item.OuterText.Contains("字形结构"))
                    {
                        int st = item.OuterText.IndexOf("[ 首尾分解查字 ]：");
                        int end = item.OuterText.IndexOf("[ 笔顺编号 ]：");
                        if (st > 0)
                        {
                            string body = item.OuterText.Substring(st, end - st);
                            body = body.Replace("[ 首尾分解查字 ]：", "").Trim();
                            ALLWord_webResult[ongoingIndex].pianpang = body;
                            loading = false;

                            isrun = ProcessStatus.关闭页面;
                            isOneFinished = true;
                            login = 0;
                            break;
                        }

                    }

                }
            }
        }
        private void getPIANpang2()
        {
            HtmlElementCollection atab;
            //   HtmlElement userNames = myDoc.Document.GetElementById("wrap");
            //try
            //{
            if (myDoc.IsDisposed == false)
            {

                atab = myDoc.Document.GetElementsByTagName("p");

                //}
                //catch (Exception ex)
                //{
                //    return;
                //    throw;
                //}
                // if (userNames != null)
                {

                    //    string dd = userNames.InnerText;


                    // HtmlElementCollection atab11 = userNames.Document.GetElementsByTagName("div");

                    foreach (HtmlElement item in atab)
                    {
                        if (item.OuterText != null && item.OuterText.Contains("部件分解"))
                        {

                            string body = item.OuterText;
                            body = body.Replace("部件分解：", "").Trim();
                            ALLWord_webResult[ongoingIndex].pianpang = body;
                            loading = false;

                            isrun = ProcessStatus.关闭页面;
                            isOneFinished = true;
                            login = 0;
                            break;


                        }

                    }
                }
            }
        }

        private void getjiegou()
        {
            //HtmlElement userNames = myDoc.Document.GetElementById("div_a1");

            //HtmlElementCollection atab11 = userNames.Document.GetElementsByTagName("div");

            //if (userNames != null)
            //{

            //    string dd = userNames.InnerText;

            //}
            if (myDoc.IsDisposed == false)
            {
                HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("p");

                foreach (HtmlElement item in atab)
                {
                    if (item.OuterText != null && item.OuterText.Contains("结构"))
                    {
                        HtmlElementCollection atab11 = item.Document.GetElementsByTagName("span");
                        bool isjiegou = false;
                        foreach (HtmlElement item1 in atab11)
                        {
                            if (item1.OuterText != null && item1.OuterText.Contains("结构"))
                            {
                                isjiegou = true;
                                continue;
                            }
                            if (isjiegou == true)
                            {
                                ALLWord_webResult[ongoingIndex].mark2 = item1.OuterText;
                                loading = false;
                                isrun = ProcessStatus.关闭页面;
                                isOneFinished = true;
                                login = 0;
                                break;

                            }

                        }
                        if (isOneFinished == true)
                            break;

                    }
                }
            }
        }
        private void Get_Kaijiang()
        {
            if (MyWebBrower == null)
            {
                loading = false;
                return;

            }
            try
            {
                List<clsKeyWord_web_info> Word_webResult = new List<clsKeyWord_web_info>();

                IHTMLDocument2 doc = (IHTMLDocument2)MyWebBrower.Document.DomDocument;
                HTMLDocument myDoc1 = doc as HTMLDocument;
                if (myDoc1 != null)
                {
                    IHTMLElementCollection Tablelin = myDoc1.getElementsByTagName("table");
                    if (Tablelin != null)
                        foreach (IHTMLElement items in Tablelin)
                        {

                            if (items != null && items.outerText != null && items.outerText.Contains("笔画"))
                            {

                                HTMLTable materialTable = items as HTMLTable;
                                IHTMLElementCollection rows = materialTable.rows as IHTMLElementCollection;
                                int KeyInfoRowIndex = 2;
                                int KeyInfoCellCIDIndex = 1, KeyInfoCellPN = 2, KeyInfoCellLocation = 3, KeyInfoCellDataSource = 4, KeyInfoCellOrder = 5, KeyInfoCell_haomaileixing = 0, KeyInfoCell_disanyilou = 6, KeyInfoCell_dieryilou = 7, KeyInfoCell_shangciyilou = 8, KeyInfoCell_dangqianyilou = 9, KeyInfoCell_yuchujilv = 10;
                                int KeyInfoCellCIDIndex1 = 0;
                                for (int i = 0; i < rows.length - 1; i++)
                                {
                                    clsKeyWord_web_info item = new clsKeyWord_web_info();
                                    #region MyRegion

                                    HTMLTableRowClass KeyRow = rows.item(KeyInfoCellCIDIndex1, null) as HTMLTableRowClass;
                                    HTMLTableCell shiming = KeyRow.cells.item(KeyInfoCell_haomaileixing, null) as HTMLTableCell;
                                    item.mark1 = shiming.innerText;

                                    HTMLTableCell shimingLocation = KeyRow.cells.item(KeyInfoCellCIDIndex, null) as HTMLTableCell;
                                    item.pianpang = shimingLocation.innerText;


                                    #endregion
                                    string[] tatile = System.Text.RegularExpressions.Regex.Split(item.pianpang, " ");
                                    char[] cc = item.pianpang.ToCharArray();
                                    for (int i11 = 0; i11 < cc.Length; i11++)
                                    {
                                        HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("a");
                                        string tx = cc[i11].ToString();
                                        ///Html/zi/BuShou/1_1.html
                                        ///http://tool.httpcn.com/Zi/BuShou.html
                                        //http://tool.httpcn.com/Html/zi/BuShou/1_1.html


                                        string filter_qishu = (KeyInfoCellCIDIndex1 + 1).ToString() + "_" + (i11 + 1).ToString();


                                        MyWebBrower.Navigate("http://tool.httpcn.com/Html/zi/BuShou/" + filter_qishu.ToString() + ".html");
                                        //  return;

                                        loading_bushoujiansuo = true;
                                        while (loading_bushoujiansuo == true)
                                        {
                                            System.Windows.Forms.Application.DoEvents();
                                            //  Application.DoEvents();
                                            Get_bushoujiansuo();

                                        }
                                    }

                                    loading = false;

                                    Word_webResult.Add(item);
                                    KeyInfoCellCIDIndex1++;
                                }

                            }
                        }
                    isrun = ProcessStatus.关闭页面;
                    isOneFinished = true;
                }
            }
            catch (Exception ex)
            {
                return;
                throw;
            }
        }
        private void Get_Kaijiangbefore()
        {
            if (MyWebBrower == null)
            {
                loading = false;
                return;

            }
            try
            {
                List<clsKeyWord_web_info> Word_webResult = new List<clsKeyWord_web_info>();

                IHTMLDocument2 doc = (IHTMLDocument2)MyWebBrower.Document.DomDocument;
                HTMLDocument myDoc1 = doc as HTMLDocument;
                if (myDoc1 != null)
                {
                    IHTMLElementCollection Tablelin = myDoc1.getElementsByTagName("table");
                    if (Tablelin != null)
                        foreach (IHTMLElement items in Tablelin)
                        {

                            if (items != null && items.outerText != null && items.outerText.Contains("笔画"))
                            {

                                HTMLTable materialTable = items as HTMLTable;
                                IHTMLElementCollection rows = materialTable.rows as IHTMLElementCollection;
                                int KeyInfoRowIndex = 2;
                                int KeyInfoCellCIDIndex = 1, KeyInfoCellPN = 2, KeyInfoCellLocation = 3, KeyInfoCellDataSource = 4, KeyInfoCellOrder = 5, KeyInfoCell_haomaileixing = 0, KeyInfoCell_disanyilou = 6, KeyInfoCell_dieryilou = 7, KeyInfoCell_shangciyilou = 8, KeyInfoCell_dangqianyilou = 9, KeyInfoCell_yuchujilv = 10;
                                int KeyInfoCellCIDIndex1 = 0;
                                for (int i = 0; i < rows.length - 1; i++)
                                {
                                    clsKeyWord_web_info item = new clsKeyWord_web_info();
                                    #region MyRegion

                                    HTMLTableRowClass KeyRow = rows.item(KeyInfoCellCIDIndex1, null) as HTMLTableRowClass;
                                    HTMLTableCell shiming = KeyRow.cells.item(KeyInfoCell_haomaileixing, null) as HTMLTableCell;
                                    item.mark1 = shiming.innerText;

                                    HTMLTableCell shimingLocation = KeyRow.cells.item(KeyInfoCellCIDIndex, null) as HTMLTableCell;
                                    item.pianpang = shimingLocation.innerText;


                                    #endregion
                                    string[] tatile = System.Text.RegularExpressions.Regex.Split(item.pianpang, " ");
                                    char[] cc = item.pianpang.ToCharArray();
                                    for (int i11 = 0; i11 < cc.Length; i11++)
                                    {
                                        HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("a");
                                        string tx = cc[i11].ToString();
                                        ///Html/zi/BuShou/1_1.html
                                        ///http://tool.httpcn.com/Zi/BuShou.html
                                        //http://tool.httpcn.com/Html/zi/BuShou/1_1.html


                                        string filter_qishu = (KeyInfoCellCIDIndex1 + 1).ToString() + "_" + (i11 + 1).ToString();


                                        MyWebBrower.Navigate("http://tool.httpcn.com/Html/zi/BuShou/" + filter_qishu.ToString() + ".html");
                                        //  return;

                                        loading_bushoujiansuo = true;
                                        while (loading_bushoujiansuo == true)
                                        {
                                            System.Windows.Forms.Application.DoEvents();
                                            //  Application.DoEvents();
                                            Get_bushoujiansuo();

                                        }
                                    }

                                    loading = false;

                                    Word_webResult.Add(item);
                                    KeyInfoCellCIDIndex1++;
                                }

                            }
                        }
                    isrun = ProcessStatus.关闭页面;
                    isOneFinished = true;
                }
            }
            catch (Exception ex)
            {
                return;
                throw;
            }
        }

        #region 首页 > 新华字典 > 部首检索 >





        #endregion

        private void Get_bushoujiansuo()
        {
            if (MyWebBrower == null)
            {
                loading_bushoujiansuo = false;
                return;

            }
            try
            {
                List<clsKeyWord_web_info> Word_webResult = new List<clsKeyWord_web_info>();

                IHTMLDocument2 doc = (IHTMLDocument2)MyWebBrower.Document.DomDocument;
                HTMLDocument myDoc1 = doc as HTMLDocument;
                if (myDoc1 != null)
                {
                    IHTMLElementCollection Tablelin = myDoc1.getElementsByTagName("table");
                    if (Tablelin != null)
                        foreach (IHTMLElement items in Tablelin)
                        {

                            if (items != null && items.outerText != null && items.outerText.Contains("拼音"))
                            {

                                HTMLTable materialTable = items as HTMLTable;
                                IHTMLElementCollection rows = materialTable.rows as IHTMLElementCollection;
                                int KeyInfoRowIndex = 2;
                                int KeyInfoCellCIDIndex = 1, KeyInfoCellPN = 2, KeyInfoCellLocation = 3, KeyInfoCellDataSource = 4, KeyInfoCellOrder = 5, KeyInfoCell_haomaileixing = 0, KeyInfoCell_disanyilou = 6, KeyInfoCell_dieryilou = 7, KeyInfoCell_shangciyilou = 8, KeyInfoCell_dangqianyilou = 9, KeyInfoCell_yuchujilv = 10;
                                int KeyInfoCellCIDIndex1 = 0;
                                for (int i = 0; i < rows.length - 1; i++)
                                {
                                    clsKeyWord_web_info item = new clsKeyWord_web_info();
                                    #region MyRegion

                                    HTMLTableRowClass KeyRow = rows.item(KeyInfoCellCIDIndex1, null) as HTMLTableRowClass;
                                    HTMLTableCell shiming = KeyRow.cells.item(KeyInfoCell_haomaileixing, null) as HTMLTableCell;
                                    item.mark1 = shiming.innerText;

                                    HTMLTableCell shimingLocation = KeyRow.cells.item(KeyInfoCellCIDIndex, null) as HTMLTableCell;
                                    item.pianpang = shimingLocation.innerText;


                                    #endregion
                                    string[] tatile = System.Text.RegularExpressions.Regex.Split(item.pianpang, " ");
                                    char[] cc = item.pianpang.ToCharArray();
                                    for (int i11 = 0; i11 < cc.Length; i11++)
                                    {
                                        HtmlElementCollection atab = myDoc.Document.GetElementsByTagName("a");
                                        string tx = cc[i11].ToString();
                                        ///Html/zi/BuShou/1_1.html
                                        ///http://tool.httpcn.com/Zi/BuShou.html
                                        //http://tool.httpcn.com/Html/zi/BuShou/1_1.html

                                        string filter_qishu = (KeyInfoCellCIDIndex1 + 1).ToString() + "_" + (i11 + 1).ToString();

                                        MyWebBrower.Navigate("http://tool.httpcn.com/Html/zi/BuShou/" + filter_qishu.ToString() + ".html");

                                    }

                                    loading_bushoujiansuo = false;

                                    Word_webResult.Add(item);
                                    KeyInfoCellCIDIndex1++;
                                }

                            }
                        }
                    //isrun = ProcessStatus.关闭页面;
                    //isOneFinished = true;
                }
            }
            catch (Exception ex)
            {
                return;
                throw;
            }
        }


        public void createWord_web_Server(List<clsKeyWord_web_info> AddMAPResult)
        {
            foreach (clsKeyWord_web_info item in AddMAPResult)
            {
                if (item != null && item.word != null && item.word != "")
                {
                    if (item.word == "膔")
                    {
                        MessageBox.Show("到了", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                    deleteWord_web(item.word);

                    string sql = "INSERT INTO Word_web ( word, pianpang,jiegou,mark1,mark2,mark3,mark4,mark5) " +

                              "VALUES (\"" + item.word + "\"" +
                                     ",\"" + item.pianpang + "\"" +
                                          ",\"" + item.jiegou + "\"" +
                                               ",\"" + item.mark1 + "\"" +
                                                    ",\"" + item.mark2 + "\"" +
                                                         ",\"" + item.mark3 + "\"" +
                                                             ",\"" + item.mark4 + "\"" +
                                     ",\"" + item.mark5 + "\")";
                    //宻 亻
                    int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql, CommandType.Text, null);
                }
            }
            return;
        }
        public void deleteWord_web(string name)
        {
            string sql2 = "delete from Word_web where  word='" + name + "'";

            int result = SQLiteHelper.ExecuteNonQuery(SQLiteHelper.CONNECTION_STRING_BASE, sql2, CommandType.Text, null);

            return;

        }
        public List<clsKeyWord_web_info> findWord_web(string findtext)
        {

            SQLiteConnection dbConn = new SQLiteConnection("Data Source=" + dataSource);

            dbConn.Open();
            SQLiteCommand dbCmd = dbConn.CreateCommand();
            //dbCmd.CommandText = "SELECT * FROM tuijanhaoma";
            dbCmd.CommandText = findtext;

            DbDataReader reader = SQLiteHelper.ExecuteReader("Data Source=" + newsth, dbCmd);
            List<clsKeyWord_web_info> ClaimReport_Server = new List<clsKeyWord_web_info>();

            while (reader.Read())
            {
                clsKeyWord_web_info item = new clsKeyWord_web_info();

                item.Order_id = reader.GetInt32(0);
                if (reader.GetValue(1) != null && Convert.ToString(reader.GetValue(1)) != "")
                    item.word = reader.GetString(1);
                if (reader.GetValue(2) != null && Convert.ToString(reader.GetValue(2)) != "")
                    item.pianpang = Convert.ToString(reader.GetString(2));
                if (reader.GetValue(3) != null && Convert.ToString(reader.GetValue(3)) != "")
                    item.jiegou = reader.GetString(3);
                if (reader.GetValue(4) != null && Convert.ToString(reader.GetValue(4)) != "")

                    item.mark1 = reader.GetString(4);
                if (reader.GetValue(5) != null && Convert.ToString(reader.GetValue(5)) != "")

                    item.mark2 = reader.GetString(5);
                if (reader.GetValue(6) != null && Convert.ToString(reader.GetValue(6)) != "")

                    item.mark3 = reader.GetString(6);
                if (reader.GetValue(7) != null && Convert.ToString(reader.GetValue(7)) != "")

                    item.mark4 = reader.GetString(7);
                if (reader.GetValue(8) != null && Convert.ToString(reader.GetValue(8)) != "")

                    item.mark5 = reader.GetString(8);

                ClaimReport_Server.Add(item);

                //这里做数据处理....
            }
            return ClaimReport_Server;

        }


    }
}
