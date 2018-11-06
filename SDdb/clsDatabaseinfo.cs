using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;

namespace SDdb
{
    public class softTime_info
    {
        public string _id { get; set; }//玩法种类

        public string starttime { get; set; }//玩法种类
        public string name { get; set; }//玩法种类
        public string endtime { get; set; }//玩法种类
    }
    public class clsuserinfo
    {
        public string Order_id { get; set; }
        public string name { get; set; }
        public string password { get; set; }
        public string Btype { get; set; }
        public string denglushijian { get; set; }
        public string Createdate { get; set; }
        public string AdminIS { get; set; }
        public string jigoudaima { get; set; }
        public string userTime { get; set; }
    }
    public class clsword_info
    {
        public int Order_id { get; set; }
        public string zi { get; set; }
        public string zhengtidaima { get; set; }
        public string bushou1 { get; set; }
        public string bushou2 { get; set; }
        public string bushou3 { get; set; }
        public string bushou4 { get; set; }
        public string bushou5 { get; set; }
        public string bushou6 { get; set; }
        public string bushou7 { get; set; }
        public string bushou8 { get; set; }
        public string bushou9 { get; set; }
        public string Input_Date { get; set; }
        public string ku_id { get; set; }
        public string Message { get; set; }
        public string bushou1daima { get; set; }
        public string bushou2daima { get; set; }
        public string bushou3daima { get; set; }
        public string bushou4daima { get; set; }
        public string bushou5daima { get; set; }
        public string bushou6daima { get; set; }
        public string bushou7daima { get; set; }
        public string bushou8daima { get; set; }
        public string bushou9daima { get; set; }
        //
       
        public string jiegou { get; set; }
        public string jiegoudaima { get; set; }
      


    }
    public class clswordpart_info
    {
        public string bushou_name { get; set; }
        public string bushou_daima { get; set; }
      
    
    }
    public class clsbushoudaima_info
    {
        public int status_id { get; set; }      
        public string shang { get; set; }
        public string xia { get; set; }
        public string zuo { get; set; }
        public string you { get; set; }
        public string zhong { get; set; }
        public string shangzhongxia { get; set; }
        public string zuozhongyou { get; set; }
        public string nei { get; set; }
        public string wai { get; set; }
        public string zuoyou_jiegou { get; set; }
        public string zuoyou_jiegoudaima { get; set; }

        public string shangxia_jiegou { get; set; }
        public string shangxia_jiegoudaima { get; set; }

    }

    public class clsKeyWord_web_info
    {
        public int Order_id { get; set; }
        public string word { get; set; }
        public string pianpang { get; set; }
        public string jiegou { get; set; }
        public string mark1 { get; set; }
        public string mark2{ get; set; }
        public string mark3 { get; set; }
        public string mark4{ get; set; }
        public string mark5 { get; set; }
    }
}
