using Order.Common;
using SDdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SQLite;
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
    
  
        
    }
}
