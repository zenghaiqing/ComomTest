using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SajetClass;
using System.IO;
using System.Data.OracleClient;
using System.Reflection;
using System.Xml;
using System.Data.OleDb;
using System.Runtime.InteropServices;

namespace ComomTest
{
    public partial class fMain : Form
    {
        public static string g_sExeName;
        String g_sFunctionType;
        public static string g_sUserID, g_sUserNo;
        String g_sIniFile = Application.StartupPath + "\\sajet.ini";
        String g_sIniFactoryID;
        String G_sTerminalID, g_Processid, g_sStageID, g_sLineID;
        int g_iPrivilege = 0;
        string Wo;
        bool bLotQtyOK = true;
        bool g_bConfiguration;
        public static string g_sProgram, g_sFunction;
        string TCOMMAND = "";
        string Part_no, SFC_QRCODE, SFC_HC, SFC_IMEI, RetMsg;
        int Qty;
        bool wo_finish;
        DataSet dsTemp;
        public static string IsSwitch;//是否启用员工权限管理
        SajetInifile sajetInifile1;
        string mainPartNo = "";
        string mainPartId = "";
        string routeId = "";
        string mainPartVersion = "";
        string mainWoType = "";
        //接受不良代码
        string errorCode = "N/A"; 
        //根据kpsn查g_sn_status表的工单
        string kpsnWorkOrder = string.Empty; 
        //工单
        string workOrder = string.Empty;
        //记录主SN
        string mainSn = string.Empty;
        private Color[] LogMsgTypeColor = { Color.Black, Color.Green, Color.Orange, Color.Red };

        public  const string OK = "OK";
        public  string DEFECT_CODE = "BURNER";
        const string BACKUP_ERROR = "Backup_Error";
        const string BACKUP = "Backup";

        string qfilename;
        string qsuffix;
        string qtestPath;
        string qresultok;
        string qresultng;
        string qfunmethod;
        string qdefectcode;


        //自定义定时器
        private System.Timers.Timer timer = new System.Timers.Timer();

        /// <summary>
        /// 初始化timer设置
        /// </summary>
        private void InitializeTimer()
        {
            //设置timer可用
            timer.Enabled = false;

            //设置timer
            timer.Interval = 2000; //ms

            //设置是否重复计时，如果该属性设为False,则只执行timer_Elapsed方法一次。
            timer.AutoReset = true;

            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);

            //防止GC回收内存
            GC.KeepAlive(timer);
        }

        private void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {

            timer.Stop();
            try
            {
                ///TODO: 执行内容
                Type thisType = GetType();
                MethodInfo theMethod = thisType.GetMethod(qfunmethod);
                theMethod.Invoke(this, null);
                //Log(LogType.Debug, "测试1");
                //Thread.Sleep(6000);
            }
            catch (Exception ex)
            {
                string msg = ex.ToString();
                //logs.logError(msg);
                //ShowLog(msg);
            }
            if ("停止".Equals(tsbStart.Text))
            {
                timer.Start();
            }
        }

        public fMain()
        {
            InitializeComponent();
            InitializeTimer();
        }

        public bool GetTerminalID()
        {
            sajetInifile1 = new SajetInifile();
            G_sTerminalID = sajetInifile1.ReadIniFile(g_sIniFile, g_sFunction, "Terminal", "");

            if (string.IsNullOrEmpty(G_sTerminalID))
            {
                ShowMsg("Terminal not be assign", 0);
                return false;
            }

            string sSQL = "SELECT A.TERMINAL_NAME,B.PROCESS_NAME,C.PDLINE_NAME "
                         + "      ,A.PDLINE_ID,A.STAGE_ID,A.PROCESS_ID,D.STAGE_NAME "
                         + " From SAJET.SYS_TERMINAL A "
                         + "     ,SAJET.SYS_PROCESS B "
                         + "     ,SAJET.SYS_PDLINE C "
                         + "     ,SAJET.SYS_STAGE D "
                         + "Where A.TERMINAL_ID = '" + G_sTerminalID + "' "
                         + "AND A.PROCESS_ID = B.PROCESS_ID "
                         + "AND A.PDLINE_ID = C.PDLINE_ID "
                         + " AND A.STAGE_ID = D.STAGE_ID ";
            DataSet dsTemp = ClientUtils.ExecuteSQL(sSQL);
            if (dsTemp.Tables[0].Rows.Count == 0)
            {
                ShowMsg("Terminal data Error", 0);
                return false;
            }
            g_Processid = dsTemp.Tables[0].Rows[0]["PROCESS_ID"].ToString();
            this.Text = this.Text + " ("
                      + dsTemp.Tables[0].Rows[0]["PROCESS_NAME"].ToString() + " / "
                      + dsTemp.Tables[0].Rows[0]["TERMINAL_NAME"].ToString() + ")";

            // g_sProcessID = dsTemp.Tables[0].Rows[0]["PROCESS_ID"].ToString();
            g_sStageID = dsTemp.Tables[0].Rows[0]["STAGE_ID"].ToString();
            g_sLineID = dsTemp.Tables[0].Rows[0]["PDLINE_ID"].ToString();

            lablPDLine.Text = dsTemp.Tables[0].Rows[0]["PDLINE_NAME"].ToString();
            lablProcess.Text = dsTemp.Tables[0].Rows[0]["PROCESS_NAME"].ToString();
            lablTerminal.Text = dsTemp.Tables[0].Rows[0]["TERMINAL_NAME"].ToString();
            lablStage.Text = dsTemp.Tables[0].Rows[0]["STAGE_NAME"].ToString();
            return true;
        }

        public void Check_Privilege()
        {
            btnConfig.Enabled = false;
            this.p_Input.Enabled = false;
            g_iPrivilege = ClientUtils.GetPrivilege(g_sUserID, g_sFunction, g_sProgram);
            this.p_Input.Enabled = (g_iPrivilege >= 1);
            g_bConfiguration = (g_iPrivilege >= 2);
            btnConfig.Enabled = g_bConfiguration;
        }

        /// <summary>
        /// iType(0:Error;1:Warning;default:Normal)
        /// </summary>
        /// <param name="sText"></param>
        /// <param name="iType"></param>
        private void Show_Message(string sText, int iType)
        {

            switch (iType)
            {
                case 0: //Error
                    Log(LogType.Error, sText);
                    break;
                case 1: //Warning
                    Log(LogType.Warning, sText);
                    break;
                default:
                    Log(LogType.Normal, sText);
                    break;
            }
        }

        /// <summary>
        /// iType(0:Error;1:Warning;default:Normal)
        /// </summary>
        /// <param name="sText"></param>
        /// <param name="iType"></param>
        private void ShowMsg(string sText, int iType)
        {
            this.BeginInvoke(new Action(() =>
            {
                sText = SajetCommon.SetLanguage(sText);
                switch (iType)
                {
                    case 0: //Error
                        this.txtMsg.Text = SajetCommon.SetLanguage(sText);
                        this.txtMsg.BackColor = Color.Yellow;
                        this.txtMsg.ForeColor = Color.Red;
                        break;
                    case 1: //Warning
                        this.txtMsg.Text = SajetCommon.SetLanguage(sText);
                        break;
                    default:
                        this.txtMsg.Text = SajetCommon.SetLanguage(sText);
                        this.txtMsg.BackColor = Color.White;
                        this.txtMsg.ForeColor = Color.Green;
                        break;
                }
            }));
            
        }

        private void btnConfig_Click_1(object sender, EventArgs e)
        {
            fTerminal f = new fTerminal(g_bConfiguration);
            try
            {
                if (f.ShowDialog() == DialogResult.OK)
                {
                    GetTerminalID();
                    Check_Privilege();
                    this.Close();
                    this.Dispose();

                }
            }
            finally
            {
                f.Dispose();
            }
        }

        private void fMain_Load(object sender, EventArgs e)
        {
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            g_sExeName = ClientUtils.fCurrentProject;
            g_sFunction = ClientUtils.fFunctionName;
            g_sProgram = ClientUtils.fProgramName;
            ClientUtils.SetLanguage(this, g_sExeName);
            this.Text = this.Text + " (" + SajetCommon.g_sFileVersion + ")";

            SajetCommon.SetLanguageControl(this);
            this.BackgroundImage = ClientUtils.LoadImage("ImgMain.jpg");
            this.BackgroundImageLayout = ImageLayout.Stretch;

            g_sUserID = ClientUtils.UserPara1;
            g_sUserNo = ClientUtils.fLoginUser;

            /**
             * 权限按钮
             * zenghaiqing
             * 2020年12月22日17:02:26
             */
        



            string sSQL = " select value  from sajet.g_setup_switch WHERE ID='1000000001'";
            DataSet dst = ClientUtils.ExecuteSQL(sSQL);
            if (dst.Tables[0].Rows.Count > 0)
            {
                IsSwitch = dst.Tables[0].Rows[0]["value"].ToString();
            }
            Wo = "";
            Part_no = "";
            Qty = 0;
            Check_Privilege();
            DataSet ds = ClientUtils.GetFunction(g_sProgram, "N");
            foreach (DataRow dr in ds.Tables[0].Rows)
            {
                if (g_sFunction == dr["FUNCTION"].ToString())
                {
                    g_sFunctionType = dr["FUN_TYPE"].ToString();
                    break;
                }
            }
            SajetInifile sajetInifile1 = new SajetInifile();
            g_sIniFactoryID = sajetInifile1.ReadIniFile(g_sIniFile, "System", "Factory", "0");
            G_sTerminalID = sajetInifile1.ReadIniFile(g_sIniFile, g_sFunction, "Terminal", "0");

            //弄セTerminal
            if (!GetTerminalID())
            {
                //this.p_Input.Enabled = false;
                //this.groupBox_Details.Enabled = false;
                return;
            }

            //初始化读取配置
            GetTerminalConfigInfo(g_Processid);

            ClearData();
            this.txt_wo.Focus();
            this.txt_wo.SelectAll();
            wo_finish = false;
        }


        private void GetTerminalConfigInfo(string process_id)
        {
            string sql = $"select PROCESS_ID,FILE_NAME,SUFFIX,TEST_PATH ,RESULT_OK ,RESULT_NG,FUN_METHOD,DEFECT_CODE from sajet.G_TEST_TERMINAL_CONFIG where process_id = {process_id}";
            DataTable dataTable = ClientUtils.ExecuteSQL(sql).Tables[0];

            if(dataTable.Rows.Count == 0)
            {
                ShowMsg("站点测试文件未配置，请联系管理员。", 0);
            }

            qfilename = dataTable.Rows[0]["FILE_NAME"].ToString();
            qsuffix = dataTable.Rows[0]["SUFFIX"].ToString();
            qtestPath = dataTable.Rows[0]["TEST_PATH"].ToString();
            qresultok = dataTable.Rows[0]["RESULT_OK"].ToString().ToUpper();
            qresultng = dataTable.Rows[0]["RESULT_NG"].ToString().ToUpper();
            qdefectcode = dataTable.Rows[0]["DEFECT_CODE"].ToString();
            qfunmethod = dataTable.Rows[0]["FUN_METHOD"].ToString();

            tslPath.Text = qtestPath;
            labelfilename.Text = qfilename + qsuffix;
        }

        private void ClearData()
        {
            this.txt_wo.Text = "";
            this.txt_wo.SelectAll();
        }

        private void tsbSetup_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                tslPath.Text = folderBrowserDialog1.SelectedPath;
                sajetInifile1.WriteIniFile(g_sIniFile, g_sFunctionType, "PATH", tslPath.Text);
            }
        }

        private void tsbStart_Click(object sender, EventArgs e)
        {
            if ("启动".Equals(tsbStart.Text))
            {
                timer.Stop();
                timer.Start();
                tsbStart.Text = "停止";
                ShowMsg("开始抓取文件...", 3);
                //pictureBox1.BackgroundImage = ClientUtils.LoadImage("green.png");
            }
            else
            {
                timer.Stop();
                tsbStart.Text = "启动";
                ShowMsg("已停止抓取文件", 2);
                //pictureBox1.BackgroundImage = ClientUtils.LoadImage("red.png");
            }


            //if (tsbStart.Text == "启动")
            //{
            //    if (txt_wo.Text.Trim() == "")
            //    {
            //        ShowMsg("请输入工单！",0);
            //        return;
            //    }
            //    if (labelpartno.Text.Trim() == "N/A")
            //    {
            //        ShowMsg("请输入工单,并按回车键！", 0);
            //        return;
            //    }

            //    tsbStart.Text = "停止";
            //    timer1.Start();
            //    ShowMsg("开始抓取文件...", 3);
            //}
            //else
            //{
            //    tsbStart.Text = "启动";
            //    timer1.Stop();
            //    ShowMsg("已停止抓取文件", 2);
            //}
        }


        /// <summary>
        /// 文件是否被占用
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <returns>True:被使用;Flase:未被使用</returns>
        public static bool IsFileInUse(string fileName)
        {
            bool inUse = true;

            if (File.Exists(fileName))
            {
                inUse = false;
            }
            else
            {
                FileStream fs = null;
                try
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    inUse = false;
                }
                catch
                {
                    inUse = false;
                }
                finally
                {
                    if (fs != null)

                        fs.Close();
                }
            }

            return inUse;//true表示正在使用,false没有使用

        }

        /// <summary>
        /// 将指定的文件复制到指定的文件路径中
        /// </summary>
        /// <param name="localfilepath">要复制文件路径</param>
        /// <param name="targetfilepath">要目标文件复制到指定的路径</param>
        public void CopyFile(string localfilepath, string copyfilepath)
        {
            List<string[]> lscopy = new List<string[]>();
            try
            {
                if (string.IsNullOrEmpty(localfilepath))
                {
                    return;
                }

                if (string.IsNullOrEmpty(copyfilepath))
                {
                    return;
                }

                ///源文件存在
                if (File.Exists(localfilepath))
                {
                    if (!File.Exists(copyfilepath))
                    {
                        try
                        {
                            if (IsFileInUse(localfilepath))
                            {
                                SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + localfilepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                                return;
                            }
                            File.Copy(localfilepath, copyfilepath);
                            File.Delete(localfilepath);
                        }
                        catch
                        {
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + localfilepath + " " + SajetCommon.SetLanguage("copy or delete wrong", 1));
                        }
                    }
                    else
                    {
                        try
                        {
                            if (IsFileInUse(localfilepath))
                            {
                                SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + localfilepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                                return;
                            }
                            File.Delete(localfilepath);
                        }
                        catch
                        {
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + localfilepath + " " + SajetCommon.SetLanguage("delete wrong", 1));
                        }
                    }
                }
                else
                {
                    SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + localfilepath + " " + SajetCommon.SetLanguage("not exist", 1));
                }
            }
            catch (Exception ex)
            {
                SajetCommon.SaveLog("[Error]", ex.Message);
            }
        }

       
        private void InitDirectory(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //timer1.Stop();

            //Type thisType = GetType();
            //MethodInfo theMethod = thisType.GetMethod(qfunmethod);
            //theMethod.Invoke(this, null);
            
            //if (tsbStart.Text.Equals("停止"))
            //{
            //    timer1.Start();
            //}
        }

         /// <summary>
         /// 
         /// 烧录 昆山供应商
         /// 
         /// </summary>
        public void BURNER()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            DirectoryInfo folder = new DirectoryInfo(sFolderPath);

            DateTime nowdate = DateTime.Now.AddSeconds(-1);

            //var fileInfo = Directory.GetFiles(sFolderPath).Where(p => (File.GetCreationTime(p) <= nowdate || File.GetLastWriteTime(p) <= nowdate && p.ToUpper().EndsWith($"*{qsuffix}")));

            //var files = folder.GetFiles().Where(p => (File.GetCreationTime(p.FullName) <= nowdate || File.GetLastWriteTime(p.FullName) <= nowdate) && p.FullName.ToUpper().EndsWith($"*{qsuffix.ToUpper()}"));

            var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));

            try
            {
                if (files.Count() > 0)
                {
                    Log(LogType.Debug, "==============================================================");
                    foreach (FileInfo file in files)  //遍历文件
                    {
                        string filepath = file.FullName;

                        //检查文件是否被使用
                        if (IsFileInUse(filepath))
                        {
                            Log(LogType.Error, filepath + " in use!");
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                            continue;
                        }
                        DataTable dt = new DataTable();
                        dt = SajetCommon.OpenCSV(filepath);
                        foreach (DataRow row in dt.Rows)
                        {
                            //产品SN	芯片SN	测试站台	测试人员	测试时间	文件信息	CRC	测试结果	备注
                            string sn = row[0].ToString().Replace("\t","");
                            string row2 = row[1].ToString();
                            string row3 = row[2].ToString();
                            string row4 = row[3].ToString();
                            string row5 = row[4].ToString();
                            string row6 = row[5].ToString();
                            string row7 = row[6].ToString();
                            string result = row[7].ToString().ToUpper();
                            string row9 = row[8].ToString();

                            //保存测试记录
                            InsertTestInfo(sn, result, row4, row3, row2, row5, row6, row7, row9);
                            if (!result.Equals(qresultok))
                            {
                                //errorCode = dt.Rows[0][6].ToString();
                                //if (string.IsNullOrEmpty(errorCode))
                                //{
                                errorCode = qdefectcode;
                                //}

                                if (!checkDefectCode(errorCode))
                                {
                                    CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                    continue;
                                }
                            }
                            else
                            {
                                errorCode = "N/A";
                            }
                            string msg = SJ_CHK_Sync_SMT_SN(txt_wo.Text, sn, errorCode);
                            if (!msg.StartsWith(OK))
                            {
                                Log(LogType.Error, sn + "," + msg);
                                CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                if ("N/A".Equals(errorCode))
                                {
                                    Log(LogType.Normal, sn + "," + msg);
                                    CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                }
                                else
                                {
                                    Log(LogType.Error, sn + "," + msg);
                                    CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                }
                            }
                        }
                    }
                }
                else
                {
                    ShowMsg("正在抓取文件...", 3);
                }
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);
            }
        }

        /// <summary>
        /// 
        /// 半成品测试
        /// 
        /// </summary>
        public void MANUAL()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            DirectoryInfo folder = new DirectoryInfo(sFolderPath);
            DateTime nowdate = DateTime.Now.AddSeconds(-1);
            var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));
            try
            {
                if (files.Count() > 0)
                {
                    Log(LogType.Debug, "==============================================================");
                    foreach (FileInfo file in files)  //遍历文件
                    {
                        string filepath = file.FullName;

                        //检查文件是否被使用
                        if (IsFileInUse(filepath))
                        {
                            Log(LogType.Error, filepath + " in use!");
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                            continue;
                        }

                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.Load(filepath);

                        XmlNodeList topM = xmldoc.SelectNodes("//test");
                        //XmlNodeList topM = xmldoc.DocumentElement.ChildNodes;

                        

                        string sn = "";
                        string stime = "";
                        string result = "";
                        //string result = file.Name.ToUpper();
                        //string resultok = result.Substring(result.Length - qsuffix.Length - qresultok.Length, qresultok.Length);
                        //string resultng = result.Substring(result.Length - qsuffix.Length - qresultng.Length, qresultng.Length);

                        foreach (XmlElement element in topM)
                        {
                            sn = element.GetElementsByTagName("runningCard")[0].InnerText;
                            stime = element.GetElementsByTagName("testingDate")[0].InnerText;
                            result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                        }

                        //查询子项是否有ng项目
                        if (qresultok.Equals(result))
                        {
                            XmlNodeList teststepxmls = xmldoc.SelectNodes("//testStep");
                            foreach(XmlElement element in teststepxmls)
                            {
                                result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                                if (qresultng.Equals(result))
                                {
                                    break;
                                }
                            }
                        }

                        //查询子项是否有ng项目
                        if (qresultok.Equals(result))
                        {
                            XmlNodeList teststepxmls = xmldoc.SelectNodes("//testItem");
                            foreach (XmlElement element in teststepxmls)
                            {
                                result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                                if (qresultng.Equals(result))
                                {
                                    break;
                                }
                            }
                        }


                        //判空
                        if (string.IsNullOrEmpty(result))
                        {
                            Log(LogType.Error, sn + ",测试结果未识别或文件格式错误！请查看测试文件");
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}{qsuffix}");
                            continue;
                        }

                        //保存测试记录
                        InsertTestInfo(sn, result, "", "", stime, "", "", "", "");
                        if (!result.Equals(qresultok))
                        {
                            //errorCode = dt.Rows[0][6].ToString();
                            //if (string.IsNullOrEmpty(errorCode))
                            //{
                            errorCode = qdefectcode;
                            //}

                            if (!checkDefectCode(errorCode))
                            {
                                CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff")  +$"_{sn}_{result}{qsuffix}");
                                continue;
                            }
                        }
                        else
                        {
                            errorCode = "N/A";
                        }
                        string msg = SJ_CHK_SN_GO(txt_wo.Text, sn, errorCode);
                        if (!msg.StartsWith(OK))
                        {
                            Log(LogType.Error, sn + "," + msg);
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                        }
                        else
                        {
                            if ("N/A".Equals(errorCode))
                            {
                                Log(LogType.Normal, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                Log(LogType.Error, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                        }
                    }
                }
                else
                {
                    ShowMsg("正在抓取文件...", 3);
                }
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);
            }
        }

        private  IEnumerable<string> GetfileList(string path, string suffix)
        {
            IEnumerable<string> fileInfo = null;

            fileInfo = Directory.GetFiles(path).Where(p => File.GetCreationTime(p) < DateTime.Now.AddSeconds(-1) && p.EndsWith(suffix));


            return fileInfo;

        }

        private  IEnumerable<string> GetfileList(string path, string suffix, int time)
        {
            IEnumerable<string> fileInfo = null;

            fileInfo = Directory.GetFiles(path).Where(p => File.GetCreationTime(p) < DateTime.Now.AddSeconds(-1) && File.GetCreationTime(p) >= DateTime.Now.AddSeconds(time) && p.EndsWith(suffix));

            return fileInfo;

        }

        private static IEnumerable<string> GetDirList(string path)
        {
            IEnumerable<string> fileInfo = null;

            fileInfo = Directory.GetDirectories(path).Where(p => File.GetCreationTime(p) < DateTime.Now.AddSeconds(-1));

            return fileInfo;
        }

        private static IEnumerable<string> GetDirList(string path, int day)
        {
            IEnumerable<string> fileInfo = null;

            fileInfo = Directory.GetDirectories(path).Where(p => File.GetCreationTime(p) < DateTime.Now.AddSeconds(-1) && File.GetCreationTime(p) >= DateTime.Now.AddDays(day));

            return fileInfo;
        }

        private  void AccessConn(string pathJob, string suffix, string intervalDate, string intervalTime)
        {
            IEnumerable<string> dirList0;
            if (!string.IsNullOrEmpty(intervalDate))
            {
                dirList0 = GetDirList(pathJob, 0 - Convert.ToInt32(intervalDate));
            }
            else
            {
                dirList0 = GetDirList(pathJob);
            }
            if (dirList0 == null) return;
            foreach (string dir0 in dirList0)
            {
                IEnumerable<string> dirList = null;
                if (!string.IsNullOrEmpty(intervalDate))
                {
                    dirList = GetDirList(dir0, 0 - Convert.ToInt32(intervalDate));
                }
                else
                {
                    dirList = GetDirList(dir0);
                }
                if (dirList == null) continue;
                foreach (string dir in dirList)
                {
                    IEnumerable<string> fileList = null;
                    if (!string.IsNullOrEmpty(intervalTime))
                    {
                        fileList = GetfileList(dir, suffix, 0 - Convert.ToInt32(intervalTime));
                    }
                    else
                    {
                        fileList = GetfileList(dir, suffix);
                    }
                    if (fileList == null) continue;
                    foreach (string file in fileList)
                    {
                        string dirName1 = dir.Substring(dir.Length - 23, 23);

                        //DataTable data = ClientUtils.ExecuteSQL(string.Format("SELECT * FROM SAJET.G_SN_DEFECT_TEST_ITEM WHERE DIR_NAME = '{0}'", dirName1)).Tables[0];
                        //if (data.Rows.Count > 0) continue;

                        //FT_RESULT1   FT_RESULT2   FT_RESULT3
                        string Con = @"Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + file;//第二个参数为文件的路径 
                        OleDbConnection dbconn = new OleDbConnection(Con);
                        try
                        {
                            dbconn.Open();//建立连接
                            string sql = @"SELECT
		                            BarCode,IsOK
	                            FROM
		                            Result1 ";
                            //
                            OleDbDataAdapter inst = new OleDbDataAdapter(sql, dbconn);//选择全部内容
                            DataSet ds = new DataSet();//临时存储
                            inst.Fill(ds);//用inst填充ds
                            if (ds == null) continue;
                            DataTable dataTable = ds.Tables[0];//展示ds第一张表到dataGridView1控件
                            for (int i = 0; i < dataTable.Rows.Count; i++)
                            {
                                //string dirName = dir.Substring(dir.Length - 23, 23);
                                string item_id = dataTable.Rows[i]["BarCode"].ToString();
                                string ateName = dataTable.Rows[i]["IsOK"].ToString();
                                //string serialNumber = dataTable.Rows[i]["SERIAL_NUMBER"].ToString();
                                //string workOrder = dataTable.Rows[i]["WORK_ORDER"].ToString();
                                //string startTime = dataTable.Rows[i]["START_TIME"].ToString();
                                //string stopTime = dataTable.Rows[i]["STOP_TIME"].ToString();
                                //string result = dataTable.Rows[i]["RESULT"].ToString();
                                //string productLine = dataTable.Rows[i]["PRODUCT_LINE"].ToString();
                                //string product = dataTable.Rows[i]["PRODUCT"].ToString();
                                //string tpsName = dataTable.Rows[i]["TPS_NAME"].ToString();
                                //string tpsFlag = dataTable.Rows[i]["TPS_FLAG"].ToString();
                                //string tpsDesc = dataTable.Rows[i]["TPS_DESC"].ToString();
                                //string uutName = dataTable.Rows[i]["UUT_NAME"].ToString();
                                //string itemName = dataTable.Rows[i]["ITEM_NAME"].ToString();
                                //string testItemName = dataTable.Rows[i]["TEST_ITEM_NAME"].ToString();
                                //string testSerialNumber = dataTable.Rows[i]["TEST_SERIAL_NUMBER"].ToString();
                                //string resultDesc = dataTable.Rows[0]["RESULT_DESC"].ToString();
                                ////增加操作：利用insert方法，在dbconn.Open(); 后添加以下代码，然后将所有代码复制到对应按钮的click事件下
                                ////保存不良数据
                                //InsertTestData(dirName, item_id, ateName, serialNumber, workOrder, startTime, stopTime,
                                // result, productLine, product, tpsName, tpsFlag, tpsDesc, uutName, itemName,
                                // testItemName, testSerialNumber, resultDesc);
                            }

                            //Log($"收集测试数据条数：{dataTable.Rows.Count}");
                        }

                        finally
                        {
                            dbconn.Close();//关闭连接
                        }

                    }
                }
            }


        }


        private  void AccessConnALL()
        {

            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            string sqldate = $"select TO_CHAR(nvl(max(EXTEND_D1),sysdate - 1),'YYYY/MM/DD HH24:MI:SS')  MAXTIME from sajet.G_SN_TEST_2 where terminal_id = {G_sTerminalID}";
            DataTable data3 = ClientUtils.ExecuteSQL(sqldate).Tables[0];
            //STOP_TIME   2021 / 1 / 14 0:44:50
            string readTime = DateTime.Now.ToString(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            string result1idtest = "0";
            if(data3.Rows.Count > 0)
            {
                readTime = data3.Rows[0]["MAXTIME"].ToString();

                string sqldate2 = $"select nvl(MAX(EXTEND_N1),0) RESULT1ID from sajet.G_SN_TEST_2 where terminal_id = {G_sTerminalID} AND EXTEND_D1 = TO_DATE('{readTime}', 'YYYY/MM/DD HH24:MI:SS')";
                DataTable data4 = ClientUtils.ExecuteSQL(sqldate2).Tables[0];
                result1idtest = data4.Rows[0]["RESULT1ID"].ToString();
            }

            //EXTEND_N1


            IEnumerable<string> fileList = GetfileList(sFolderPath, qsuffix);
            if (fileList == null) return;
            
            foreach (string file in fileList)
            {
                
                string filepath = file;

                //检查文件是否被使用
                if (IsFileInUse(filepath))
                {
                    Log(LogType.Error, filepath + " in use!");
                    SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                    continue;
                }
                //string dirName1 = dir.Substring(dir.Length - 23, 23);

                //DataTable data = ClientUtils.ExecuteSQL(string.Format("SELECT * FROM SAJET.G_SN_DEFECT_TEST_ITEM WHERE DIR_NAME = '{0}'", dirName1)).Tables[0];
                //if (data.Rows.Count > 0) continue;

                //FT_RESULT1   FT_RESULT2   FT_RESULT3
                string Con = @"Provider=Microsoft.Jet.Oledb.4.0;Data Source=" + file;//第二个参数为文件的路径 
                OleDbConnection dbconn = new OleDbConnection(Con);
                try
                {
                    dbconn.Open();//建立连接
                    string sql = string.Format(@"
                    SELECT BarCode,IsOK,ToTime,Result1ID FROM (SELECT
		            BarCode,IsOK,ToTime,Result1ID
	            FROM
		            Result1 WHERE ToTime >= #{0}# and  Result1ID > {1}
                    union
                    SELECT
		            BarCode,IsOK,ToTime,Result1ID
	            FROM
		            Result1 WHERE ToTime > #{0}# and Result1ID > 0)  ORDER BY ToTime ASC", readTime,result1idtest);
                    //
                    OleDbDataAdapter inst = new OleDbDataAdapter(sql, dbconn);//选择全部内容
                    DataSet ds = new DataSet();//临时存储
                    inst.Fill(ds);//用inst填充ds
                    if (ds == null) continue;
                    DataTable dataTable = ds.Tables[0];//展示ds第一张表到dataGridView1控件

                    if (dataTable.Rows.Count == 0) continue;
                    Log(LogType.Debug, "==============================================================");
                    for (int i = 0; i < dataTable.Rows.Count; i++)
                    {
                        //string dirName = dir.Substring(dir.Length - 23, 23);
                        string sn = dataTable.Rows[i]["BarCode"].ToString();
                        string result = dataTable.Rows[i]["IsOK"].ToString();
                        string stime = dataTable.Rows[i]["ToTime"].ToString();
                        string result1id = dataTable.Rows[i]["Result1ID"].ToString();
                        DateTime datetest;
                        if (!DateTime.TryParse(stime, out datetest)) {
                            datetest = DateTime.Now.AddMinutes(-10);
                        };
                        //DateTime datetest = DateTime.ParseExact(stime, "yyyy/M/d H:mm:ss", System.Globalization.CultureInfo.CurrentCulture);

                        //判空
                        if (string.IsNullOrEmpty(result))
                        {
                            Log(LogType.Error, sn + ",测试结果未识别或文件格式错误！请查看测试文件");
                            //CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}{qsuffix}");
                            continue;
                        }


                        //保存测试记录
                        InsertTestInfo(sn, result, "", "", stime, "", "", "", "", datetest, result1id);
                        if (!result.Equals(qresultok))
                        {
                            //errorCode = dt.Rows[0][6].ToString();
                            //if (string.IsNullOrEmpty(errorCode))
                            //{
                            errorCode = qdefectcode;
                            //}

                            if (!checkDefectCode(errorCode))
                            {
                                //CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                continue;
                            }
                        }
                        else
                        {
                            errorCode = "N/A";
                        }
                        string msg = SJ_CHK_SN_GO(txt_wo.Text, sn, errorCode);

                        msg += $";{i + 1}";
                        if (!msg.StartsWith(OK))
                        {
                            Log(LogType.Error, sn + "," + msg);
                            //CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                        }
                        else
                        {
                            if ("N/A".Equals(errorCode))
                            {
                                Log(LogType.Normal, sn + "," + msg);
                                //CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                Log(LogType.Error, sn + "," + msg);
                                //CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                        }
                    }
                        

                }
                finally
                {
                    dbconn.Close();//关闭连接
                    //CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + qsuffix);
                }

            }
        }


        private  void InsertTestData(string dirName, string item_id, string ateName, string serialNumber, string workOrder, string startTime, string stopTime,
            string result, string productLine, string product, string tpsName, string tpsFlag, string tpsDesc, string uutName, string itemName,
            string testItemName, string testSerialNumber, string resultDesc)
        {
            string recID = "0";
            string sql = string.Format(@"INSERT INTO SAJET.G_SN_DEFECT_TEST_ITEM (
	                        DIR_NAME,
	                        RECID,
	                        ITEM_ID,
	                        ATE_NAME,
	                        SERIAL_NUMBER,
	                        WORK_ORDER,
	                        START_TIME,
	                        STOP_TIME,
	                        RESULT,
	                        PRODUCT_LINE,
	                        PRODUCT,
	                        TPS_NAME,
	                        TPS_FLAG,
	                        TPS_DESC,
	                        UUT_NAME,
	                        ITEM_NAME,
	                        TEST_ITEM_NAME,
	                        TEST_SERIAL_NUMBER,
	                        RESULT_DESC 
                        )
                        VALUES
	                        (
		                        '{0}',
		                        '{1}',
		                        '{2}',
		                        '{3}',
		                        '{4}',
		                        '{5}',
		                        TO_DATE( '{6}', 'YYYY/MM/DD HH24:MI:SS' ),
		                        TO_DATE( '{7}', 'YYYY/MM/DD HH24:MI:SS' ),
		                        '{8}',
		                        '{9}',
		                        '{10}',
		                        '{11}',
		                        '{12}',
		                        '{13}',
		                        '{14}',
		                        '{15}',
		                        '{16}',
		                        '{17}',
		                        '{18}' 
	                        )", dirName, recID, item_id, ateName, serialNumber, workOrder, startTime, stopTime,
             result, productLine, product, tpsName, tpsFlag, tpsDesc, uutName, itemName,
             testItemName, testSerialNumber, resultDesc);

            ClientUtils.ExecuteSQL(sql);

        }

        /// <summary>
        /// 
        /// 荣耀半成品测试
        /// 
        /// </summary>
        public void RYMANUAL()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            //DirectoryInfo folder = new DirectoryInfo(sFolderPath);
            //DateTime nowdate = DateTime.Now.AddSeconds(-1);
            //var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));
            try
            {
                AccessConnALL();
                   
                ShowMsg("正在抓取文件...", 3);
 
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);

            }
        }


        /// <summary>
        /// 
        /// ATE测试
        /// 
        /// </summary>
        public void ATSTEST()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            DirectoryInfo folder = new DirectoryInfo(sFolderPath);
            DateTime nowdate = DateTime.Now.AddSeconds(-1);
            var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));
            try
            {
                if (files.Count() > 0)
                {
                    Log(LogType.Debug, "==============================================================");
                    foreach (FileInfo file in files)  //遍历文件
                    {
                        string filepath = file.FullName;

                        //检查文件是否被使用
                        if (IsFileInUse(filepath))
                        {
                            Log(LogType.Error, filepath + " in use!");
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                            continue;
                        }

                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.Load(filepath);

                        XmlNodeList topM = xmldoc.SelectNodes("//test");
                        //XmlNodeList topM = xmldoc.DocumentElement.ChildNodes;



                        string sn = "";
                        string stime = "";
                        string result = "";
                        //string result = file.Name.ToUpper();
                        //string resultok = result.Substring(result.Length - qsuffix.Length - qresultok.Length, qresultok.Length);
                        //string resultng = result.Substring(result.Length - qsuffix.Length - qresultng.Length, qresultng.Length);

                        foreach (XmlElement element in topM)
                        {
                            sn = element.GetElementsByTagName("runningCard")[0].InnerText;
                            stime = element.GetElementsByTagName("testingDate")[0].InnerText;
                            result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                        }
                        //查询子项是否有ng项目
                        if (qresultok.Equals(result))
                        {
                            XmlNodeList teststepxmls = xmldoc.SelectNodes("//testStep");
                            foreach (XmlElement element in teststepxmls)
                            {
                                result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                                if (qresultng.Equals(result))
                                {
                                    break;
                                }
                            }
                        }
                        //查询子项是否有ng项目
                        if (qresultok.Equals(result))
                        {
                            XmlNodeList teststepxmls = xmldoc.SelectNodes("//testItem");
                            foreach (XmlElement element in teststepxmls)
                            {
                                result = element.GetElementsByTagName("result")[0].InnerText.ToUpper();
                                if (qresultng.Equals(result))
                                {
                                    break;
                                }
                            }
                        }


                        //判空
                        if (string.IsNullOrEmpty(result))
                        {
                            Log(LogType.Error, sn + ",测试结果未识别或文件格式错误！请查看测试文件");
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}{qsuffix}");
                            continue;
                        }


                        //保存测试记录
                        InsertTestInfo(sn, result, "", "", stime, "", "", "", "");
                        if (!result.Equals(qresultok))
                        {
                            //errorCode = dt.Rows[0][6].ToString();
                            //if (string.IsNullOrEmpty(errorCode))
                            //{
                            errorCode = qdefectcode;
                            //}

                            if (!checkDefectCode(errorCode))
                            {
                                CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff")  +$"_{sn}_{result}{qsuffix}");
                                continue;
                            }
                        }
                        else
                        {
                            errorCode = "N/A";
                        }
                        string msg = SJ_CHK_SN_GO(txt_wo.Text, sn, errorCode);
                        if (!msg.StartsWith(OK))
                        {
                            Log(LogType.Error, sn + "," + msg);
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                        }
                        else
                        {
                            if ("N/A".Equals(errorCode))
                            {
                                Log(LogType.Normal, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                Log(LogType.Error, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                        }
                    }
                }
                else
                {
                    ShowMsg("正在抓取文件...", 3);
                }
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);
            }
        }

        bool RunStatus;
        private void tb_Tolling_M_SN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Return)
                return;
            RunStatus = true;
            if (G_sTerminalID == "101012")//等于单板测试
            {

                txtPanelNO.SelectAll();
                txtPanelNO.Focus();
                return;
            }
            tb_TollingSN.SelectAll();
            tb_TollingSN.Focus();
        }

        string tollingSN;

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
                    SajetCommon.Show_Message("没有跨线别操作的权限！", 0);
                    return false;
                }
            }
            return true;
        }
        private bool checkTooling(string ToolingSN)
        {
            try
            {
                object[][] Params = new object[4][];
                Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TTERMINALID", G_sTerminalID };
                Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TREV", ToolingSN };
                Params[2] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TEMP", g_sUserNo };
                Params[3] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
                DataSet ds = ClientUtils.ExecuteProc("SAJET.SJ_ckrt_TOOLINGR", Params);
                string sTres = ds.Tables[0].Rows[0]["TRES"].ToString().Trim();
                if (sTres.Substring(0, 2) == "OK")
                {
                    return true;
                }
                else
                {
                    Log(LogType.Error,  ToolingSN + " : " + sTres);
                    SajetCommon.SaveLog("[Error]",  ToolingSN + " : " + sTres);
                    return false;
                }

            }
            catch (System.Exception ex)
            {
                Log(LogType.Error, ex.Message);
                SajetCommon.SaveLog("[Error]", ex.Message);
                return false;
            }
        }
     

        


        public const int KEYEVENTF_KEYUP = 2;
        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);
        [System.Runtime.InteropServices.DllImport("user32")]
        private static extern int mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        [DllImport("user32.dll", EntryPoint = "keybd_event", SetLastError = true)]
        public static extern void keybd_event(Keys bVk, byte bScan, uint dwFlags, uint dwExtraInfo);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool SwitchToThisWindow(IntPtr hWnd, bool fAltTab);
        [DllImport("user32.dll", EntryPoint = "FindWindow")]
        private extern static IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll ")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private void tb_TollingSN_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Return)
                return;
            tollingSN = tb_TollingSN.Text.ToString().Trim();
            if (!SajetCommon.CheckWorkLineByTooling(tollingSN, g_sUserID, G_sTerminalID))
            {
                return;
            }
            if (!checkTooling(tollingSN))
            {
                //modify by hua 使全选 2015/04/17
                this.tb_TollingSN.Focus();
                this.tb_TollingSN.Text = "";

                //end hua 
                return;
            }
            //int y = 219;
            //if (checkBox1.Checked)
            //{
            //    y = 239;
            //}

            //if (checkBox2.Checked)
            //{
            //    y = 200;
            //}

            #region 验证治具是否锁定
            string sSQL = "SELECT TOOLING_STATUS FROM sajet.SYS_TOOLING_SN WHERE TOOLING_SN='" + tollingSN + "'";
            DataTable dtTollingSN = ClientUtils.ExecuteSQL(sSQL).Tables[0];
            if (dtTollingSN.Rows.Count == 0)
            {
                this.tb_TollingSN.Focus();
                this.tb_TollingSN.SelectAll();
                Log(LogType.Error, tb_TollingSN.Text.Trim() + "治具不存在！");
                return;
            }
            if ("H".Equals(dtTollingSN.Rows[0]["TOOLING_STATUS"]))
            {
                this.tb_TollingSN.Focus();
                this.tb_TollingSN.SelectAll();
                Log(LogType.Error, tb_TollingSN.Text.Trim() + "治具已锁定！");
                return;
            }
            #endregion




            #region 把成品号输入到华为测试系统界面
            string sql = "SELECT TS.SERIAL_NUMBER,TS.TOOLING_SEQ FROM SAJET.G_SN_TOOLING_AC33B TS INNER JOIN SAJET.SYS_TOOLING_SN T ON TS.TOOLING_SN_ID=T.TOOLING_SN_ID WHERE T.TOOLING_SN='" + tollingSN + "' and TS.process_id = 101055  ORDER BY TS.TOOLING_SEQ ";
            DataTable dt = ClientUtils.ExecuteSQL(sql).Tables[0];
            string sn = "";
            if (dt.Rows.Count > 0)
            {
                //提前验证是否存在不良记录
                bool isBool = false;
                int j = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    j++;
                    string strSN = dr["SERIAL_NUMBER"].ToString().ToUpper();
                    string strRes = "";
                    string status = CheckRoute(G_sTerminalID, strSN, ref strRes);
                    if (status == "NG")
                    {
                        isBool = true;
                        Log(LogType.Error, j + " " + strSN + " : " + SajetCommon.SetLanguage(strRes.Replace("Next:", "下一站：").Replace("/Current:", "/当前站：")));
                    }
                }
                if (isBool)
                {
                    tb_TollingSN.SelectAll();
                    tb_TollingSN.Focus();
                    return;
                }


                SetCursorPos(838, 219);
                //mouse_event(dwFlags.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                //mouse_event(dwFlags.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                mouse_event(2, 0, 0, 0, 0);
                mouse_event(4, 0, 0, 0, 0);

                DataTable dt8 = dt.Clone();

                for (int i = 0; i < 8; i++)
                {
                    DataRow newRow = dt8.NewRow();
                    newRow["SERIAL_NUMBER"] = "";
                    newRow["TOOLING_SEQ"] = i + 1;
                    dt8.Rows.Add(newRow);
                }
                dt8.AcceptChanges();

                for (int i = 0; i < dt8.Rows.Count; i++)
                {
                    DataRow[] drArray = dt.Select("TOOLING_SEQ = " + (i + 1));
                    if (drArray.Length > 0)
                    {
                        DataRow dr = drArray[0];
                        dt8.Rows[i]["SERIAL_NUMBER"] = dr["SERIAL_NUMBER"];
                    }
                }
                dt8.AcceptChanges();

                //一个治具号绑定8个小板号
                for (int i = 0; i < dt8.Rows.Count; i++)
                {
                    sn = dt8.Rows[i]["SERIAL_NUMBER"].ToString().ToUpper();
                    if (sn != "" && sn.StartsWith("[)>"))
                    {
                        //SendKeys.Send("{CAPSLOCK}");
                        //SendKeys.Send("{[}");
                        //SendKeys.Send("{)}");
                        //SendKeys.Send("{>}");
                        //String[] strs = sn.Replace("[)>", "").Split('-');
                        ////SendKeys.Send(strs[0]);
                        //foreach (char ch in strs[0])
                        //{
                        //    keybd_event(Keys.CapsLock, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                        //    keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        //}
                        //SendKeys.Send("{-}");
                        //foreach (char ch in strs[1])
                        //{
                        //    keybd_event(Keys.CapsLock, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                        //    keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        //}

                        SendKeys.Send(sn.Replace("[", "{[}").Replace(")", "{)}").Replace(">", "{>}").Replace("-", "{-}"));
                        SendKeys.Send("{Enter}");

                        //兼容45码
                        //if (sn != "" && sn.StartsWith("[)>"))
                        //{
                        //    //SendKeys.Send("{CAPSLOCK}");
                        //    SendKeys.Send("{[}");
                        //    SendKeys.Send("{)}");
                        //    SendKeys.Send("{>}");
                        //    String[] strs = sn.Replace("[)>", "").Split('-');
                        //    SendKeys.Send(strs[0]);
                        //    SendKeys.Send("{-}");
                        //    SendKeys.Send(strs[1]);
                        //    SendKeys.Send("{Enter}");
                        //    //keybd_event(Keys.CapsLock, 0, 0, 0);
                        //    //keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);

                        //}


                        //兼容45码另外一种写法
                        //keybd_event((Keys)(219), 0, 0, 0);
                        //keybd_event((Keys)(219), 0, KEYEVENTF_KEYUP, 0);

                        //keybd_event((Keys)(16), 0, 0, 0);
                        //keybd_event((Keys)(48), 0, 0, 0);
                        //keybd_event((Keys)(48), 0, KEYEVENTF_KEYUP, 0);
                        //keybd_event((Keys)(16), 0, KEYEVENTF_KEYUP, 0);

                        //keybd_event((Keys)(16), 0, 0, 0);
                        //keybd_event((Keys)(190), 0, 0, 0);
                        //keybd_event((Keys)(190), 0, KEYEVENTF_KEYUP, 0);
                        //keybd_event((Keys)(16), 0, KEYEVENTF_KEYUP, 0);

                        //String[] strs = sn.Replace("[)>", "").Split('-');
                        //foreach (char ch in strs[0])
                        //{
                        //    keybd_event(Keys.CapsLock, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                        //    keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        //}

                        //keybd_event((Keys)(189), 0, 0, 0);
                        //keybd_event((Keys)(189), 0, KEYEVENTF_KEYUP, 0);

                        //foreach (char ch in strs[1])
                        //{
                        //    keybd_event(Keys.CapsLock, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, 0, 0);
                        //    keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                        //    keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        //}

                        //keybd_event(Keys.Enter, 0, 0, 0);

                    }
                    else
                    {
                        //SendKeys.Send(sn);
                        foreach (char ch in sn)
                        {
                            keybd_event(Keys.CapsLock, 0, 0, 0);
                            keybd_event((Keys)ch, 0, 0, 0);
                            keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                            keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        }
                    }
                    keybd_event(Keys.Enter, 0, 0, 0);

                }
            }
            else
            {
                Log(LogType.Error, "治具号：" + tb_TollingSN + "没有查询到绑定的小板号，请确认！");
            }
            #endregion

            tb_TollingSN.Focus();
            tb_TollingSN.Clear();
            RunStatus = true;
        }

        private string CheckRoute(string tterminalid, string tsn, ref string tres)
        {
            object[][] Params = new object[3][];
            Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TERMINALID", tterminalid };
            Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TSN", tsn };
            Params[2] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
            dsTemp = ClientUtils.ExecuteProc("SAJET.SJ_CKRT_ROUTE_ZHANG", Params);
            string sRes = dsTemp.Tables[0].Rows[0]["TRES"].ToString();
            if (sRes.Substring(0, 2) == "OK")
            {
                tres = sRes;
                return "OK";
            }
            else
            {
                tres = sRes;
                return "NG";
            }

        }

        private string CheckRoute(string tterminalid, string tsn, ref string tmsg, ref string tres)
        {
            //string tres = null;
            object[][] Params = new object[4][];
            Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TERMINALID", tterminalid };
            Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TSN", tsn };
            Params[2] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TMSG", "" };
            Params[3] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
            dsTemp = ClientUtils.ExecuteProc("SAJET.SJ_CKRT_ROUTE_TEST2", Params);
            //SJ_CKRT_ROUTE(TERMINALID IN VARCHAR2, TSN IN VARCHAR2 ,TRES OUT VARCHAR2)
            string sRes = dsTemp.Tables[0].Rows[0]["TRES"].ToString();
            string sMsg = dsTemp.Tables[0].Rows[0]["TMSG"].ToString();
            if (sRes.Substring(0, 2) == "OK")
            {
                tres = sRes;
                return "OK";
            }
            else
            {
                tmsg = sMsg;
                tres = sRes;
                return "NG";
            }

        }

        string g_sPanelno;
        private void txtPanelNO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                //  检查母治具是否锁定
                var estooling = tb_Tolling_M_SN.Text.Trim();
                var csql = "   SELECT *  FROM SAJET.SYS_TOOLING_SN_TEMP  WHERE TOOLING_SN=:TOOLING_SN  AND PROCESS_ID=:PROCESS_ID  AND DEL_FLAG=0";
                var cparms = new object[2][];
                cparms[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TOOLING_SN", estooling };
                cparms[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "PROCESS_ID", g_Processid };
                var cdata = ClientUtils.ExecuteSQL(csql, cparms).Tables[0];
                if (cdata.Rows.Count > 0)
                {
                    Log(LogType.Error, estooling + "治具已锁定！");
                    return;
                }


                //这里就应该拼接Log信息   收到请求：带出小板号 时间 
                #region 把小板号输入到华为测试系统界面
                string sn = "";
                string sql = null;
                if (!SajetCommon.CheckWorkLineByPanel(txtPanelNO.Text.Trim(), g_sUserID, G_sTerminalID))
                {
                    return;
                }

            
                sql = "SELECT S.SERIAL_NUMBER FROM SAJET.G_SN_PRINT_LABEL@SMTDBLINK S WHERE S.RC_NO ='" + txtPanelNO.Text.Trim() + "'  ";
                var orderbystr = "ORDER BY ROWID ASC"; //镭雕默认正序
                    
                sql += orderbystr;
               

                //Log(LogType.Error, fpar);
                //Log(LogType.Error, sql);

                //if (b)
                //    sql = "SELECT S.SERIAL_NUMBER FROM SAJET.G_SN_STATUS S WHERE S.PANEL_NO ='" + txtPanelNO.Text.Trim() + "' ORDER BY ROWID DESC";
                //else
                //    sql = "SELECT S.SERIAL_NUMBER FROM SAJET.G_SN_STATUS S WHERE S.PANEL_NO ='" + txtPanelNO.Text.Trim() + "'";
                //继续拼接log信息   完成请求  时间     响应时间：  

                //获取大板数量
                g_sPanelno = txtPanelNO.Text.Trim();
                txtPanelNO.Focus();
                txtPanelNO.SelectAll();

                DataTable dt = ClientUtils.ExecuteSQL(sql).Tables[0];
                if (dt.Rows.Count > 0)
                {
                    //提前严重整盘大板是否存在不良记录
                    bool isBool = false;
                    int i = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        i++;
                        string strSN = dr["SERIAL_NUMBER"].ToString().ToUpper();
                        string strRes = "";
                        string status = CheckRoute(G_sTerminalID, strSN, ref strRes);
                        if (status == "NG")
                        {
                            isBool = true;
                            Log(LogType.Error, i + " " + strSN + " : " + SajetCommon.SetLanguage(strRes.Replace("Next:", "下一站：").Replace("/Current:", "/当前站：")));
                        }
                    }
                    if (isBool)
                    {
                        tb_TollingSN.SelectAll();
                        tb_TollingSN.Focus();
                        return;
                    }


                    SetCursorPos(838, 219);
                    mouse_event(2, 0, 0, 0, 0);
                    mouse_event(4, 0, 0, 0, 0);

                    foreach (DataRow dr in dt.Rows)
                    {
                        sn = dr["SERIAL_NUMBER"].ToString().ToUpper();
                        foreach (char ch in sn)
                        {
                            keybd_event(Keys.CapsLock, 0, 0, 0);
                            keybd_event((Keys)ch, 0, 0, 0);
                            keybd_event((Keys)ch, 0, KEYEVENTF_KEYUP, 0);
                            keybd_event(Keys.CapsLock, 0, KEYEVENTF_KEYUP, 0);
                        }
                        keybd_event(Keys.Enter, 0, 0, 0);
                    }
                    snCount = dt.Rows.Count;
                    fullSnCount = dt.Rows.Count;
                    failSnCount = 0;
                    exSnCount = 0;

                    txtPanelNO.Focus();
                    txtPanelNO.SelectAll();
                    Log(LogType.Normal, "====大板号" + txtPanelNO.Text.Trim() + ": " + +snCount + "pcs小板过站中...====");
                }
                else
                {
                    Log(LogType.Error, "大板号：" + txtPanelNO.Text.Trim() + "没有查询到小板号，请确认！");
                }
                #endregion

            }
        }

        int snCount, fullSnCount, failSnCount, exSnCount;

        private void fMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            timer.Stop();
        }

        /// <summary>
        /// 
        /// Q值校准
        /// 
        /// </summary>
        public void QVALUETEST()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            DirectoryInfo folder = new DirectoryInfo(sFolderPath);
            DateTime nowdate = DateTime.Now.AddSeconds(-1);
            var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));
            try
            {
                if (files.Count() > 0)
                {
                    Log(LogType.Debug, "==============================================================");
                    foreach (FileInfo file in files)  //遍历文件
                    {
                        string filepath = file.FullName;

                        //检查文件是否被使用
                        if (IsFileInUse(filepath))
                        {
                            Log(LogType.Error, filepath + " in use!");
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                            continue;
                        }

                        XmlDocument xmldoc = new XmlDocument();
                        xmldoc.Load(filepath);

                        XmlNode oNode = xmldoc.DocumentElement;
                        string xmlns = oNode.Attributes["xmlns"].Value;

                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmldoc.NameTable);
                        nsmgr.AddNamespace("ts", xmlns);

                        XmlNodeList topM = xmldoc.SelectNodes("descendant::ts:ResultSet", nsmgr);

                        string sn = "";
                        string sstarttime = "";
                        string sendtime = "";
                        string result = "";
              
                        //测试时间
                        foreach (XmlElement element in topM)
                        {
                            if (element.Name.Equals("ResultSet"))
                            {
                                sstarttime = element.Attributes["startDateTime"].Value;
                                sendtime = element.Attributes["endDateTime"].Value;
                                break;
                            }
                        }
                        //测试sn
                        XmlNodeList snxmls = xmldoc.SelectNodes("descendant::ts:TestResults/ts:UUT/ts:SerialNumber", nsmgr);
                        foreach (XmlElement element in snxmls)
                        {
                            if (element.Name.Equals("SerialNumber"))
                            {
                                sn = element.InnerText;
                                break;
                            }
                        }

                       
                        XmlNodeList testresultxmls = xmldoc.SelectNodes("descendant::ts:Outcome", nsmgr);
                        foreach (XmlElement element in testresultxmls)
                        {
                            if (element.Name.Equals("Outcome"))
                            {
                                result = element.Attributes["value"].Value.ToUpper();
                                if (qresultng.Equals(result))
                                {
                                    break;
                                }
                            }
                        }

                        //判空
                        if (string.IsNullOrEmpty(result))
                        {
                            Log(LogType.Error, sn + ",测试结果未识别或文件格式错误！请查看测试文件");
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}{qsuffix}");
                            continue;
                        }


                        //保存测试记录
                        InsertTestInfo(sn, result, "", "", sstarttime, sendtime, "", "", "");
                        if (!result.Equals(qresultok))
                        {
                            //errorCode = dt.Rows[0][6].ToString();
                            //if (string.IsNullOrEmpty(errorCode))
                            //{
                            errorCode = qdefectcode;
                            //}

                            if (!checkDefectCode(errorCode))
                            {
                                CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                continue;
                            }
                        }
                        else
                        {
                            errorCode = "N/A";
                        }
                        string msg = SJ_CHK_SN_GO(txt_wo.Text, sn, errorCode);
                        if (!msg.StartsWith(OK))
                        {
                            Log(LogType.Error, sn + "," + msg);
                            CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                        }
                        else
                        {
                            if ("N/A".Equals(errorCode))
                            {
                                Log(LogType.Normal, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                Log(LogType.Error, sn + "," + msg);
                                CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                            }
                        }
                    }
                }
                else
                {
                    ShowMsg("正在抓取文件...", 3);
                }
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);
            }
        }

        /// <summary>
        /// 
        /// 老化测试
        /// 
        /// </summary>
        public void BURNTEST()
        {
            string sFolderPath = tslPath.Text;
            string sdata = DateTime.Now.ToString("yyyyMMddHH");
            string sBackupPath = sFolderPath + "\\" + BACKUP + "\\" + sdata;
            string sBackupErrorPath = sFolderPath + "\\" + BACKUP_ERROR + "\\" + sdata;


            InitDirectory(sFolderPath);
            InitDirectory(sBackupPath);
            InitDirectory(sBackupErrorPath);

            DirectoryInfo folder = new DirectoryInfo(sFolderPath);
            DateTime nowdate = DateTime.Now.AddSeconds(-1);
            var files = folder.GetFiles().Where(p => File.GetCreationTime(p.FullName) <= nowdate && p.FullName.ToUpper().EndsWith($"{qsuffix.ToUpper()}"));
            try
            {
                if (files.Count() > 0)
                {
                    Log(LogType.Debug, "==============================================================");
                    foreach (FileInfo file in files)  //遍历文件
                    {
                        string filepath = file.FullName;

                        //检查文件是否被使用
                        if (IsFileInUse(filepath))
                        {
                            Log(LogType.Error, filepath + " in use!");
                            SajetCommon.SaveLog("[Error]", SajetCommon.SetLanguage("The original file path", 1) + " " + filepath + " " + SajetCommon.SetLanguage("be occupied", 1));
                            continue;
                        }
                        DataTable dt = new DataTable();
                        dt = SajetCommon.OpenCSV(filepath);
                        foreach (DataRow row in dt.Rows)
                        {
                            //bar code	result	start time	end time
                            string sn = row[0].ToString().Replace("\t", "");
                            string result = row[1].ToString().ToUpper();
                            string starttime = row[2].ToString();
                            string endtime = row[3].ToString();

                            //保存测试记录
                            InsertTestInfo(sn, result, "", "", starttime, endtime, "", "", "");
                            if (!result.Equals(qresultok))
                            {
                                //errorCode = dt.Rows[0][6].ToString();
                                //if (string.IsNullOrEmpty(errorCode))
                                //{
                                errorCode = qdefectcode;
                                //}

                                if (!checkDefectCode(errorCode))
                                {
                                    CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff")  +$"_{sn}_{result}{qsuffix}");
                                    continue;
                                }
                            }
                            else
                            {
                                errorCode = "N/A";
                            }
                            string msg = SJ_CHK_SN_GO(txt_wo.Text, sn, errorCode);
                            if (!msg.StartsWith(OK))
                            {
                                Log(LogType.Error, sn + "," + msg);
                                CopyFile(filepath, sBackupErrorPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff")  +$"_{sn}_{result}{qsuffix}");
                            }
                            else
                            {
                                if ("N/A".Equals(errorCode))
                                {
                                    Log(LogType.Normal, sn + "," + msg);
                                    CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                }
                                else
                                {
                                    Log(LogType.Error, sn + "," + msg);
                                    CopyFile(filepath, sBackupPath + "\\" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + $"_{sn}_{result}{qsuffix}");
                                }
                            }
                        }
                    }
                }
                else
                {
                    ShowMsg("正在抓取文件...", 3);
                }
            }
            catch (Exception ex)
            {
                Show_Message(ex.Message, 0);
            }
        }

        private void InsertTestInfo(string sn,string result,string testuser,string testmac,string s1,string s2,string s3,string s4,string s5,DateTime d1,string result1id)
        {
            string sql = "INSERT INTO sajet.G_SN_TEST_2(SN, TERMINAL_ID, RESULT, TEST_USER, TEST_MAC,EXTEND_S1,EXTEND_S2,EXTEND_S3,EXTEND_S4,EXTEND_S5,EXTEND_D1,EXTEND_N1) " +
                $"VALUES ('{sn}',{G_sTerminalID},'{result}','{testuser}','{testmac}','{s1}','{s2}','{s3}','{s4}','{s5}',TO_DATE('{d1.ToString("yyyyMMddHHmmss")}','YYYYMMDDHH24MISS'),{result1id} )";
            ClientUtils.ExecuteSQL(sql);
        }

        private void InsertTestInfo(string sn, string result, string testuser, string testmac, string s1, string s2, string s3, string s4, string s5)
        {
            string sql = "INSERT INTO sajet.G_SN_TEST_2(SN, TERMINAL_ID, RESULT, TEST_USER, TEST_MAC,EXTEND_S1,EXTEND_S2,EXTEND_S3,EXTEND_S4,EXTEND_S5) " +
                $"VALUES ('{sn}',{G_sTerminalID},'{result}','{testuser}','{testmac}','{s1}','{s2}','{s3}','{s4}','{s5}' )";
            ClientUtils.ExecuteSQL(sql);
        }

        /// <summary>
        /// 检查是否有此ErrorCode
        /// </summary>
        /// <param name="defectCode"></param>
        /// <returns></returns>
        private bool checkDefectCode(string defectCode)
        {
            try
            {
                string sSQL = " select * from sajet.sys_defect t where t.defect_code = '" + defectCode + "' and t.enabled = 'Y' ";
                dsTemp = ClientUtils.ExecuteSQL(sSQL);

                if (dsTemp.Tables[0].Rows.Count > 0)
                {
                    return true;
                }
                else
                {
                    Log(LogType.Error, defectCode + " : " + SajetCommon.SetLanguage("请维护不良现象代码"));
                    SajetCommon.SaveLog("[Error]", defectCode + " : " + SajetCommon.SetLanguage("ErrorCode ERROR"));
                    return false;
                }
            }
            catch (System.Exception ex)
            {
                Log(LogType.Error, ex.Message);
                SajetCommon.SaveLog("[Error]", ex.Message);
                return false;
            }
        }

        private string SJ_CHK_Sync_SMT_SN(string wo, string sn,string defect)
        {
            object[][] Params = new object[8][];
            Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TTERMINALID", G_sTerminalID };
            Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TWO", wo };
            Params[2] = new object[] { ParameterDirection.Input, OracleType.DateTime, "TNOW", DateTime.Now };
            Params[3] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TEMP", g_sUserNo };
            Params[4] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TDEFECT", defect };
            Params[5] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TREV", sn };
            Params[6] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
            Params[7] = new object[] { ParameterDirection.Output, OracleType.VarChar, "tnextproc", "" };
            DataSet ds = ClientUtils.ExecuteProc("SAJET.SJ_CHK_SN_WO_GO", Params);
            return ds.Tables[0].Rows[0]["TRES"].ToString();
        }

        private string SJ_CHK_SN_GO(string wo, string sn, string defect)
        {
            object[][] Params = new object[8][];
            Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TTERMINALID", G_sTerminalID };
            Params[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TWO", wo };
            Params[2] = new object[] { ParameterDirection.Input, OracleType.DateTime, "TNOW", DateTime.Now };
            Params[3] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TEMP", g_sUserNo };
            Params[4] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TDEFECT", defect };
            Params[5] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TREV", sn };
            Params[6] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
            Params[7] = new object[] { ParameterDirection.Output, OracleType.VarChar, "tnextproc", "" };
            DataSet ds = ClientUtils.ExecuteProc("SAJET.SJ_CHK_SN_GO", Params);
            return ds.Tables[0].Rows[0]["TRES"].ToString();
        }


        private void txt_wo_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Enter)
            {
                return;
            }
            workOrder = txt_wo.Text.Trim();
            if (string.IsNullOrEmpty(workOrder)) return;

            if (!CheckAndGetWOInfo(workOrder))
            {
                txt_wo.SelectAll();
                txt_wo.Focus();
                return;
            }
            txt_wo.Enabled = false;
            ShowMsg("Work Order OK", 3);
        }

        /// <summary>
        /// 根据SN或工单检查工单状态
        /// </summary>
        /// <param name="trev">工单</param>
        /// <param name="msg">消息</param>
        /// <returns>工单OK 返回true</returns>
        private bool CheckAndGetWOInfo(string workorder)
        {
            try
            {
                object[][] Params = new object[6][];
                Params[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TREV", workorder };
                Params[1] = new object[] { ParameterDirection.Output, OracleType.Number, "TTARGET_QTY", 0 };
                Params[2] = new object[] { ParameterDirection.Output, OracleType.Number, "TINPUT_QTY", 0 };
                Params[3] = new object[] { ParameterDirection.Output, OracleType.Number, "TOUTPUT_QTY",0 };
                Params[4] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TPART_NO", "" };
                Params[5] = new object[] { ParameterDirection.Output, OracleType.VarChar, "TRES", "" };
                DataSet ds = ClientUtils.ExecuteProc("SAJET.Sj_Chk_Wo", Params);

                string msg = ds.Tables[0].Rows[0]["TRES"].ToString();
                mainPartNo = ds.Tables[0].Rows[0]["TPART_NO"].ToString();
                this.labeltargetqty.Text = ds.Tables[0].Rows[0]["TTARGET_QTY"].ToString();
                labelpartno.Text = mainPartNo;
                if (!msg.StartsWith(OK))
                {
                    ShowMsg(workorder + " : "+ msg, 0);
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Show_Message(workorder + " : " + ex.Message, 0);
                return false;
            }
        }

        private void Log(LogType msgtype, string msg)
        {
            string sType = "Warning";
            if (msgtype.ToString() == "Normal")
            {
                sType = "PASS";
            }
            else if (msgtype.ToString() == "Error")
            {
                sType = "FAIL";
            }else if(msgtype == LogType.Debug)
            {
                sType = "START TEST";
            }
            this.log1.Invoke(new EventHandler(delegate
            {
                string msgtemp = DateTime.Now.ToLongTimeString() + "【" + sType.ToString() + "】  " + msg;
                //log记录方式倒序排列
                log1.Items.Insert(0, msgtemp);
                log1.Items[0].Font = new System.Drawing.Font(log1.Items[0].Font, FontStyle.Bold);
                int ss = (int)msgtype;
                log1.Items[0].ForeColor = LogMsgTypeColor[(int)msgtype];


            }));
            this.log1.Invoke(new EventHandler(delegate
            {
                if (log1.Items.Count > 500)
                {
                    log1.Items.Remove(log1.Items[log1.Items.Count - 1]);
                }
            }));
            ClientUtils.addConnLog(msgtype, g_sFunction, msg);
        }

    }
}
