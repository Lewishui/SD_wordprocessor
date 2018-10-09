using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;

namespace Order.Common
{
   public class clsCommHelp
    { 
        #region NullToString
        public static string NullToString(object obj)
        {
            string strResult = "";
            if (obj != null)
            {
                strResult = obj.ToString().Trim();
            }
            return strResult;
        }
        #endregion

        #region StringToDecimal
        /// <summary>
        /// 转换字符串，将字符串转换成数字，并且将空字符串转换成0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static decimal StringToDecimal(string s)
        {
            decimal result = 0;

            if (s != null && s != "")
            {
                result = Decimal.Parse(s);
            }
            return result;
        }
        #endregion

        #region StringToInt
        /// <summary>
        /// 转换字符串，将字符串转换成数字，并且将空字符串转换成0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int StringToInt(string s)
        {
            int result = 0;

            if (s != null && s != "")
            {
                result = Convert.ToInt32(s.Trim());
            }
            return result;
        }
        #endregion

        #region 日期转换(objToDateTime)
        /// <summary>
        /// 将excel里取得的日期转化成String数据
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string objToDateTime<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("MM/dd/yyyy");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("MM/dd/yyyy");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(0, 4) + "-" +
                                                       obj.ToString().Substring(4, 2) + "-" +
                                                       obj.ToString().Substring(6, 2)).ToString("MM/dd/yyyy");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            return strResult;
        }

        public static string objToDateTime1<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("yyyy/MM/dd");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("yyyy/MM/dd");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(4, 4) + "-" +
                                                       obj.ToString().Substring(0, 2) + "-" +
                                                       obj.ToString().Substring(2, 2)).ToString("yyyy/MM/dd");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            return strResult;
        }
        public static string objToDateTime2<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("yyyy/MM/dd/HH/mm");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("yyyy/MM/dd/HH/mm");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(4, 4) + "-" +
                                                       obj.ToString().Substring(0, 2) + "-" +
                                                       obj.ToString().Substring(2, 2)).ToString("yyyy/MM/dd/HH/mm");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            return strResult;
        }
        #endregion

        #region 字符串简单加密解密

        /// <summary>
        /// 简单加密解密

        /// </summary>
        /// <param name="str">需要加密、解密的字符串</param>
        /// <returns>加密、解密后的字符串</returns>
        public static string encryptString(string str)
        {
            string strResult = "";
            char[] charMessage = str.ToCharArray();
            foreach (char c in charMessage)
            {
                char newChar = changerChar(c);
                strResult += newChar.ToString();
            }
            return strResult;
        }

        private static char changerChar(char c)
        {
            char resutlt;
            int intStrLength = 0;
            string twoString = Convert.ToString(c, 2).PadLeft(8, '0');
            if (twoString.Length > 8)
            {
                twoString = Convert.ToString(c, 2).PadLeft(16, '0');
            }
            intStrLength = twoString.Length;
            string newTwoString = twoString.Substring(intStrLength / 2) + twoString.Substring(0, intStrLength / 2);
            resutlt = Convert.ToChar(Convert.ToInt32(newTwoString, 2));
            return resutlt;
        }
        #endregion

        #region 将字符串日期转换为时间类型

        public static DateTime GetDateByString(string dateString)
        {
            return DateTime.Parse(dateString.Substring(0, 4) + "-" + dateString.Substring(4, 2) + "-" + dateString.Substring(6, 2));
        }
        #endregion

        #region 关闭打开的Excel
        public static void CloseExcel(Microsoft.Office.Interop.Excel.Application ExcelApplication, Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook)
        {
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
            ExcelWorkbook = null;
            ExcelApplication.Quit();
            GC.Collect();
            clsKeyMyExcelProcess.Kill(ExcelApplication);
        }
        #endregion

        #region 得到Sap连接字符串

        #endregion

        #region 判断2个日期是否是整年

        public static bool CheckThroughoutTheYear(string data1, string date2)
        {
            bool blnResult = false;
            string dtStart = "";
            string dtEnd = "";
            if (Convert.ToDateTime(date2).CompareTo(Convert.ToDateTime(data1)) > 0)
            {
                dtStart = data1;
                dtEnd = date2;
            }
            else
            {
                dtStart = date2;
                dtEnd = data1;
            }
            string strTheoryDate = Convert.ToDateTime(dtEnd).ToString("yyyy")
                                 + Convert.ToDateTime(dtStart).ToString("MMdd");
            strTheoryDate = Convert.ToDateTime(objToDateTime<string>(strTheoryDate)).AddDays(-1).ToString("MM/dd/yyyy");
            if (objToDateTime<string>(strTheoryDate) == objToDateTime<string>(dtEnd))
            {
                blnResult = true;
            }
            return blnResult;
        }

        #endregion

        #region 判断汉字
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
        #endregion
        public static string GetMacAddress()
        {
            try
            {
                string strMac = string.Empty;
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    if ((bool)mo["IPEnabled"] == true)
                    {
                        strMac = mo["MacAddress"].ToString();
                    }
                }
                moc = null;
                mc = null;
                return strMac;
            }
            catch
            {
                return "unknown";
            }
        }
        public static string GetDiskID()
        {
            try
            {
                string strDiskID = string.Empty;
                ManagementClass mc = new ManagementClass("Win32_DiskDrive");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    strDiskID = mo.Properties["Model"].Value.ToString();
                }
                moc = null;
                mc = null;
                return strDiskID;
            }
            catch
            {
                return "unknown";
            }
        }
        public static string GetCpuID()
        {
            try
            {
                string strCpuID = string.Empty;
                ManagementClass mc = new ManagementClass("Win32_Processor");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (ManagementObject mo in moc)
                {
                    strCpuID = mo.Properties["ProcessorId"].Value.ToString();
                }
                moc = null;
                mc = null;
                return strCpuID;
            }
            catch
            {
                return "unknown";
            }
        }  

        public static string getMacAddr_Local()
        {
            string madAddr = null;
            try
            {
                ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
                ManagementObjectCollection moc2 = mc.GetInstances();
                foreach (ManagementObject mo in moc2)
                {
                    if (Convert.ToBoolean(mo["IPEnabled"]) == true)
                    {
                        madAddr = mo["MacAddress"].ToString();
                        madAddr = madAddr.Replace(':', '-');
                    }
                    mo.Dispose();
                }
                if (madAddr == null)
                {
                    return "unknown";
                }
                else
                {
                    return madAddr;
                }
            }
            catch (Exception)
            {
                return "unknown";
            }
        }
        /// <summary>
        /// 取网卡信息
        /// </summary>
        /// <returns></returns>
        public static string GetNetAdapterInfo()
        {
            StringBuilder sb = new StringBuilder();
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();
            if (adapters != null)
            {
                foreach (NetworkInterface ni in adapters)
                {
                    string fCardType = "未知网卡";
                    IPInterfaceProperties ips = ni.GetIPProperties();

                    PhysicalAddress pa = ni.GetPhysicalAddress();
                    if (pa == null) continue;
                    string pastr = pa.ToString();
                    if (pastr.Length < 7) continue;
                    if (pastr.Substring(0, 6) == "000000") continue;
                    if (ni.Name.ToLower().IndexOf("vmware") > -1) continue;
                    string fRegistryKey = "SYSTEM\\CurrentControlSet\\Control\\Network\\{4D36E972-E325-11CE-BFC1-08002BE10318}\\" + ni.Id + "\\Connection";
                    RegistryKey rk = Registry.LocalMachine.OpenSubKey(fRegistryKey, false);
                    if (rk != null)
                    {
                        // 区分 PnpInstanceID     
                        // 如果前面有 PCI 就是本机的真实网卡    
                        // MediaSubType 为 01 则是常见网卡，02为无线网卡。    
                        string fPnpInstanceID = rk.GetValue("PnpInstanceID", "").ToString();
                        int fMediaSubType = Convert.ToInt32(rk.GetValue("MediaSubType", 0));
                        if (fPnpInstanceID.Length > 3 && fPnpInstanceID.Substring(0, 3) == "PCI")
                            if (ni.NetworkInterfaceType.ToString().ToLower().IndexOf("wireless") == -1)
                                fCardType = "物理网卡";
                            else
                                fCardType = "无线网卡";
                        else if (fMediaSubType == 1 || fMediaSubType == 0)
                            fCardType = "虚拟网卡";
                        else if (fMediaSubType == 2 || ni.NetworkInterfaceType.ToString().ToLower().IndexOf("wireless") > -1)
                            fCardType = "无线网卡";
                        else if (fMediaSubType == 7)
                            fCardType = "蓝牙";
                    }
                    StringBuilder isb = new StringBuilder();
                    UnicastIPAddressInformationCollection UnicastIPAddressInformationCollection = ips.UnicastAddresses;
                    foreach (UnicastIPAddressInformation UnicastIPAddressInformation in UnicastIPAddressInformationCollection)
                    {
                        if (UnicastIPAddressInformation.Address.AddressFamily == AddressFamily.InterNetwork)
                            isb.Append(string.Format("Ip Address: {0}", UnicastIPAddressInformation.Address) + "\r\n"); // Ip 地址    
                    }

                    // IPAddressCollection ipc = ips.DnsAddresses;
                    //if (ipc.Count > 0)
                    //{
                    //    foreach (IPAddress ip in ipc)
                    //    {
                    //        isb.Append(string.Format("DNS服务器地址：{0}\r\n",ip));
                    //    }
                    //}
                    string s = string.Format("{0}\r\n描述信息：{1}\r\n类型：{2}\r\n速度：{3} MB\r\nMac地址：{4}\r\n{5}", fCardType, ni.Name, ni.NetworkInterfaceType, ni.Speed / 1024 / 1024, ni.GetPhysicalAddress(), isb.ToString());
                    sb.Append(s + "\r\n");
                }
            }
            return sb.ToString();
        }

    }
}
