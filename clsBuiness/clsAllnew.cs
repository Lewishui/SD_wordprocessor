
using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace clsBuiness
{
    public class clsAllnew
    {
        private string dataSource = "Data.sqlite";
        string newsth;
        public BackgroundWorker bgWorker1;

        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }

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
            string sql = "INSERT INTO Word_ku ( zi, zhengtidaima,bushou1,bushou2,bushou3,bushou4,bushou5,bushou6,bushou7,bushou8,bushou9,Input_Date,ku_id,Message,bushou1daima,bushou2daima,bushou3daima,bushou4daima,bushou5daima,bushou6daima,bushou7daima,bushou8daima,bushou9daima ) " +

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


                             ",\"" + AddMAPResult[0].bushou9daima + "\")";

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

                //if (reader.GetValue(24) != null && Convert.ToString(reader.GetValue(24)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(24));


                //if (reader.GetValue(25) != null && Convert.ToString(reader.GetValue(25)) != "")
                //    item.bushou5daima = Convert.ToString(reader.GetString(25));


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



    }
}
