using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Data.OracleClient;
using System.Xml;
using System.Drawing;

namespace SajetClass
{
    class SajetCommon
    {
        public static string g_sFileVersion = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();  //郎セ 
        public static string g_sFileName = Path.GetFileName(FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileName); //郎嘿          
        public static string g_sExeName = ClientUtils.fCurrentProject;

        public static DialogResult Show_Message(string sKeyMsg, int iType)
        {
            string sXMLFile = Path.GetFileNameWithoutExtension(g_sFileName);
            return ClientUtils.ShowMessage(SetLanguage(sKeyMsg), iType, g_sExeName, sXMLFile);
        }

        public static string SetLanguage(string sSearchText, string sDefaultTxt, int iTransType)
        {
            string sText = SetLanguage(sSearchText, iTransType);
            if (sText != sSearchText)
                return sText;
            else
                return sDefaultTxt;
        }
        public static string SetLanguage(string sSearchText)
        {
            string sXMLFile = "";
            string sText = "";

            sXMLFile = Path.GetFileNameWithoutExtension(g_sFileName); //Dll.xml
            sText = ClientUtils.SetLanguage(sSearchText, g_sExeName, sXMLFile);
            if (sSearchText == sText)
            {
                sXMLFile = g_sExeName; //Program.xml
                sText = ClientUtils.SetLanguage(sSearchText, g_sExeName, sXMLFile);
            }
            return sText;
        }
        public static string SetLanguage(string sSearchText, int iTransType)
        {
            string sXMLFile = "";
            switch (iTransType)
            {
                case 1:  //Dll.xml
                    sXMLFile = Path.GetFileNameWithoutExtension(g_sFileName);
                    break;
                case 2:  //Program.xml
                    sXMLFile = g_sExeName;
                    break;
            }
            string sText = ClientUtils.SetLanguage(sSearchText, g_sExeName, sXMLFile);
            return sText;
        }
        public static void SetLanguageControl(Control c)
        {
            //锣传じンTxt瓣粂ē
            string sXMLFile = Path.GetFileNameWithoutExtension(g_sFileName);
            ClientUtils.SetLanguage(c, g_sExeName, sXMLFile);
            ClientUtils.SetLanguage(c, g_sExeName, g_sExeName);
        }

        public static string GetSysBaseData(string sProgram, string sParamName, ref string sErrorMsg)
        {
            //弄SYS_BASE砞﹚
            string sSQL = "";
            sSQL = " SELECT PARAM_VALUE "
                 + "   FROM SAJET.SYS_BASE "
                 + "  WHERE Upper(PROGRAM) = '" + sProgram.ToUpper() + "' "
                 + "    and Upper(PARAM_NAME) = '" + sParamName.ToUpper() + "' ";
            DataSet dsTemp = ClientUtils.ExecuteSQL(sSQL);
            if (dsTemp.Tables[0].Rows.Count > 0)
                return dsTemp.Tables[0].Rows[0]["PARAM_VALUE"].ToString();
            else
            {
                sErrorMsg = sErrorMsg + sParamName + Environment.NewLine;
                return "";
            }
        }

        public static string GetMaxID(string sTable, string sField, int iIDLength)
        {
            string sMaxID = "0";
            try
            {
                object[][] Params = new object[5][];
                Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TFIELD", sField };
                Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TTABLE", sTable };
                Params[2] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TNUM", iIDLength.ToString() };
                Params[3] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
                Params[4] = new object[] { ParameterDirection.Output, OracleType.VarChar, "T_MAXID", "" };
                DataSet dsTemp = ClientUtils.ExecuteProc("SAJET.SJ_GET_MAXID", Params);

                string sRes = dsTemp.Tables[0].Rows[0]["TRES"].ToString();
                sMaxID = dsTemp.Tables[0].Rows[0]["T_MAXID"].ToString();

                if (sRes != "OK")
                {
                    SajetCommon.Show_Message(sRes, 0);
                    return "0";
                }
            }
            catch (Exception ex)
            {
                SajetCommon.Show_Message("SAJET.SJ_GET_MAXID" + Environment.NewLine + ex.Message, 0);
                return "0";
            }
            return sMaxID;
        }

        public static string GetID(string sTable, string sFieldID, string sFieldName, string sValue)
        {
            return GetID(sTable, sFieldID, sFieldName, sValue, "");
        }
        public static string GetID(string sTable, string sFieldID, string sFieldName, string sValue, string sEnabled)
        {
            //т逆ID
            if (string.IsNullOrEmpty(sValue))
                return "0";
            string sSQL = "select " + sFieldID + " from " + sTable + " "
                        + "where " + sFieldName + " = '" + sValue + "' ";
            if (!string.IsNullOrEmpty(sEnabled))
                sSQL = sSQL + " and ENABLED = '" + sEnabled + "' ";
            sSQL = sSQL + " and Rownum = 1 ";

            DataSet dsTemp = ClientUtils.ExecuteSQL(sSQL);
            if (dsTemp.Tables[0].Rows.Count > 0)
                return dsTemp.Tables[0].Rows[0][sFieldID].ToString();
            else
                return "0";
        }

        public static Image LoadImage(string sFileName)
        {
            string sPath = Application.StartupPath + "\\";
            if (File.Exists(sPath + sFileName))
                return Image.FromFile(sPath + sFileName);
            else
                return null;
        }

        /// <summary>
        /// 将CSV文件的数据读取到DataTable中
        /// </summary>
        /// <param name="fileName">CSV文件路径</param>
        /// <returns>返回读取了CSV数据的DataTable</returns>
        public static DataTable OpenCSV(string filePath)
        {
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);
                if (strLine.Substring(0, 1) == ";")
                {
                    continue;
                }
                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
                    for (int i = 0; i < columnCount; i++)
                    {
                        DataColumn dc = new DataColumn(tableHead[i]);
                        dt.Columns.Add(dc);
                    }
                }
                else
                {
                    aryLine = strLine.Split(',');
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < columnCount; j++)
                    {
                        dr[j] = aryLine[j];
                    }
                    dt.Rows.Add(dr);
                }
            }
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            fs.Close();
            return dt;
        }

        //存Sajet_Log
        public static void SaveLog(string sStaus, string sMessage)
        {
            //抓取月份
            string sMonth = "";
            if (DateTime.Now.Month.ToString().Length > 1)
            {
                sMonth = DateTime.Now.Month.ToString();
            }
            else
            {
                sMonth = "0" + DateTime.Now.Month.ToString();
            }
            //抓取日期
            string sDay = "";
            if (DateTime.Now.Month.ToString().Length > 1)
            {
                sDay = DateTime.Now.Day.ToString();
            }
            else
            {
                sDay = "0" + DateTime.Now.Day.ToString();
            }

            //排檔名
            string sLogName = DateTime.Today.Year + "_" + sMonth + "_" + sDay + ".log";

            //判斷Sajet_Log是否存在
            if (!File.Exists(Directory.GetCurrentDirectory() + @"\Sajet_Log"))
            {
                Directory.CreateDirectory(Directory.GetCurrentDirectory() + @"\Sajet_Log");
            }

            //判斷sLogName是否存在
            if (!File.Exists(Directory.GetCurrentDirectory() + @"\Sajet_Log\" + sLogName))
            {
                //File.Create(Directory.GetCurrentDirectory() + @"\Sajet_Log\" + sLogName);


                FileStream fs = new FileStream(Directory.GetCurrentDirectory() + @"\Sajet_Log\" + sLogName, FileMode.Create, FileAccess.Write);//创建写入文件

                //存訊息
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(DateTime.Now.ToString() + sStaus + sMessage + "\r\n");

                //sw.AutoFlush = true;
                sw.Flush();
                sw.Close();
                sw.Dispose();
            }
            else
            {

                FileStream fs = new FileStream(Directory.GetCurrentDirectory() + @"\Sajet_Log\" + sLogName, FileMode.Append, FileAccess.Write);//创建写入文件
                //存訊息
                StreamWriter sw = new StreamWriter(fs);
                sw.Write(DateTime.Now.ToString() + sStaus + sMessage + "\r\n");

                //sw.AutoFlush = true;
                sw.Flush();
                sw.Close();
                sw.Dispose();
            }
        }
        /// <summary>
        /// 对输入治具管控工单跨线别操作
        /// </summary>
        /// <param name="sn"></param>
        /// <param name="userId"></param>
        /// <param name="terminalId"></param>
        /// <returns></returns>
        public static bool CheckWorkLineByTooling(string toolingSN, string userId, string terminalId)
        {
            string str = " SELECT PDLINE_ID FROM SAJET.SYS_TERMINAL  WHERE TERMINAL_ID = '" + terminalId + "'";
            DataSet dsTemp = ClientUtils.ExecuteSQL(str);
            string pdline = dsTemp.Tables[0].Rows[0][0].ToString();
            str = "SELECT DEFAULT_PDLINE_ID  FROM SAJET.G_WO_BASE B,SAJET.G_SN_TOOLING_AC33B S INNER JOIN SAJET.SYS_TOOLING_SN T ON S.TOOLING_SN_ID=T.TOOLING_SN_ID WHERE B.WORK_ORDER=S.WORK_ORDER AND T.TOOLING_SN='" + toolingSN + "' ";
            dsTemp = ClientUtils.ExecuteSQL(str);
            //当前操作的sn不在工单所属的线别，并且没有跨线别操作的权限，则管控不允许操作
            if (dsTemp.Tables[0] != null && dsTemp.Tables[0].Rows.Count > 0 && dsTemp.Tables[0].Rows[0][0].ToString() != pdline)
            {
                int g_iPrivilege = ClientUtils.GetPrivilege(userId, "Work Order By Line", "Data Center");
                if (g_iPrivilege <= 0) //只有只读的操作权限
                {
                    Show_Message("没有跨线别操作的权限！", 0);
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// 对输入大板号管控工单跨线别操作
        /// </summary>
        /// <param name="sn"></param>
        /// <param name="userId"></param>
        /// <param name="terminalId"></param>
        /// <returns></returns>
        public static bool CheckWorkLineByPanel(string panelNo, string userId, string terminalId)
        {
            string str = " SELECT PDLINE_ID FROM SAJET.SYS_TERMINAL  WHERE TERMINAL_ID = '" + terminalId + "'";
            DataSet dsTemp = ClientUtils.ExecuteSQL(str);
            string pdline = dsTemp.Tables[0].Rows[0][0].ToString();
            str = "SELECT DEFAULT_PDLINE_ID  FROM SAJET.G_WO_BASE B,SAJET.G_SN_STATUS S WHERE B.WORK_ORDER=S.WORK_ORDER AND S.PANEL_NO='" + panelNo + "' ";
            dsTemp = ClientUtils.ExecuteSQL(str);
            //当前操作的sn不在工单所属的线别，并且没有跨线别操作的权限，则管控不允许操作
            if (dsTemp.Tables[0] != null && dsTemp.Tables[0].Rows.Count > 0 && dsTemp.Tables[0].Rows[0][0].ToString() != pdline)
            {
                int g_iPrivilege = ClientUtils.GetPrivilege(userId, "Work Order By Line", "Data Center");
                if (g_iPrivilege <= 0) //只有只读的操作权限
                {
                    Show_Message("没有跨线别操作的权限！", 0);
                    return false;
                }
            }
            return true;
        }
    }
}
